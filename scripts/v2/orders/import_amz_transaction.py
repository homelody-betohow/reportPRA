from __future__ import annotations

"""
Amazon 交易明细 Excel -> 表 amz_transaction。

取代 v1「B4_1_合并_transaction交易明细.py」中与本脚本重叠的部分：
  - 读取「已发放订单」「已推迟订单」两份表（本脚本为目录下分别导入，用 source_kind 区分）；
  - 全表字符串 trim；
  - seller sku 清洗规则与 B4_1 的 extract_values 一致（便于与订单统计 SKU / 订单号识别码 对齐）。

与 B4_1 的差异：B4_1 另生成「仅 order id 非空、fba≠0、按 order-id识别码 汇总」的 Excel 供 B4_2 映射；
  本脚本将「全量明细行」写入 MySQL；汇总映射请改为从 amz_transaction 查询或后续步骤实现。

读取目录：python/excel/daily/amazon/ 下 transaction交易明细-*.xlsx
（已发放订单 / 已推迟订单 表头一致，source_kind 由文件名区分：released / deferred）。

line_hash：对 LINE_HASH_KEYS 子集做 stable_line_hash（SHA-256），与 uk_amz_txn_line_hash 一致；
重复导入 UPSERT。键选取说明见脚本内 LINE_HASH_KEYS 注释。

用法（在 python/v2 目录下，或将 orders 与 warehouse-rent 加入 PYTHONPATH）：
  python orders/import_amz_transaction.py
  python orders/import_amz_transaction.py --file ../excel/daily/amazon/transaction交易明细-已发放订单.xlsx
  python orders/import_amz_transaction.py --dir ../excel/daily/amazon
"""

import argparse
import os
import sys
from pathlib import Path
from typing import Any

_ORDERS_DIR = Path(__file__).resolve().parent
_V2_DIR = _ORDERS_DIR.parent
_WR_DIR = _V2_DIR / "warehouse-rent"
for _p in (_ORDERS_DIR, _V2_DIR, _WR_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

import pandas as pd
from logger import get_logger, setup_stdout_utf8

from db import connect, load_db_config
from excel_common import (
    cell_decimal,
    cell_dt,
    cell_str,
    row_subset_for_line_hash,
    stable_line_hash,
    upsert_rows,
)

TABLE = "amz_transaction"
_LOG = get_logger("AMZ-TXN")

# Excel 列名（与 python/excel/daily/amazon 下样例一致）-> 库字段 -> 类型
_AMZ_TXN_MAP: list[tuple[str, str, str]] = [
    ("期间", "period_label", "s16"),
    ("报表原日期", "report_row_at", "dt"),
    ("店铺名称", "shop_name", "s128"),
    ("站点", "site_code", "s16"),
    ("币种", "currency", "s16"),
    ("划款账单对账状态", "payout_reconcile_status", "s64"),
    ("划款时间", "payout_at", "dt"),
    ("结算时间-开始", "settlement_start_at", "dt"),
    ("结算时间-结束", "settlement_end_at", "dt"),
    ("结算时间", "settlement_at", "dt"),
    ("发货时间", "shipped_at", "dt"),
    ("发货仓库", "ship_warehouse", "s255"),
    ("group id", "group_id", "s128"),
    ("type", "transaction_type", "s64"),
    ("order id", "amazon_order_id", "s64"),
    ("原销售订单号", "original_sales_order_no", "s64"),
    ("merchantOrderId", "merchant_order_id", "s64"),
    ("配送方式", "fulfillment_channel", "s32"),
    ("seller sku", "seller_sku", "s255"),
    ("子ASIN", "child_asin", "s32"),
    ("父ASIN", "parent_asin", "s32"),
    ("warehouse sku", "warehouse_sku", "s255"),
    ("description", "line_description", "s512"),
    ("quantity", "quantity", "dec"),
    ("marketplace", "marketplace", "s64"),
    ("product sales", "product_sales", "dec"),
    ("product sales tax", "product_sales_tax", "dec"),
    ("shipping credits", "shipping_credits", "dec"),
    ("shipping credits tax", "shipping_credits_tax", "dec"),
    ("gift wrap credits", "gift_wrap_credits", "dec"),
    ("gift wrap credits tax", "gift_wrap_credits_tax", "dec"),
    ("regulatory fee", "regulatory_fee", "dec"),
    ("promotional rebates", "promotional_rebates", "dec"),
    ("promotional rebates tax", "promotional_rebates_tax", "dec"),
    ("marketplace withheld tax", "marketplace_withheld_tax", "dec"),
    ("sales tax collected", "sales_tax_collected", "dec"),
    ("low value goods", "low_value_goods", "dec"),
    ("amazon point costs", "amazon_point_costs", "dec"),
    ("selling fees", "selling_fees", "dec"),
    ("fba fees", "fba_fees", "dec"),
    ("other transaction fees", "other_transaction_fees", "dec"),
    ("other", "other_amount", "dec"),
    ("total", "total_amount", "dec"),
    ("采购成本", "purchase_cost", "dec"),
    ("采购运费", "purchase_shipping", "dec"),
    ("采购税费", "purchase_tax", "dec"),
    ("头程运费", "first_leg_shipping", "dec"),
    ("头程税费", "first_leg_tax", "dec"),
    ("转人民币汇率", "fx_rate_cny", "dec"),
]

LINE_HASH_SOURCE_FIELDS: tuple[str, ...] = ("source_kind",) + tuple(c for _, c, _ in _AMZ_TXN_MAP)

# line_hash 参与键（改后需重导或接受新旧 hash 并存）：
#   纳入：来源、店铺/期间、结算分组与时间定位、订单与 SKU 维度、description、数量、
#         Amazon 侧金额列（至 total_amount）——保证「不同明细行」不易撞 hash。
#   排除：采购/头程/汇率（常为后补或内部表，不应因补数产生新行）；
#         划款对账状态、划款时间（易随对账变化，同一明细应 UPSERT 更新而非新行）。
_LINE_HASH_EXCLUDE: frozenset[str] = frozenset(
    {
        "purchase_cost",
        "purchase_shipping",
        "purchase_tax",
        "first_leg_shipping",
        "first_leg_tax",
        "fx_rate_cny",
        "payout_reconcile_status",
        "payout_at",
    }
)
LINE_HASH_KEYS: tuple[str, ...] = tuple(
    k for k in LINE_HASH_SOURCE_FIELDS if k not in _LINE_HASH_EXCLUDE
)

_src_set = frozenset(LINE_HASH_SOURCE_FIELDS)
_bad_keys = tuple(k for k in LINE_HASH_KEYS if k not in _src_set)
if _bad_keys:
    raise ValueError(f"LINE_HASH_KEYS 含未知列: {_bad_keys}")


def default_amz_transaction_excel_dir() -> Path:
    """python/excel/daily/amazon"""
    return _ORDERS_DIR.parents[1] / "excel" / "daily" / "amazon"


def _source_kind_from_filename(path: Path) -> str:
    name = path.name
    if "已推迟订单" in name:
        return "deferred"
    if "已发放订单" in name:
        return "released"
    _LOG.warn(f"文件名未含「已发放订单/已推迟订单」，source_kind=unknown: {name}")
    return "unknown"


def _convert(v: Any, kind: str) -> Any:
    if kind == "dt":
        return cell_dt(v)
    if kind == "dec":
        return cell_decimal(v)
    if kind.startswith("s"):
        n = int(kind[1:])
        return cell_str(v, n)
    return cell_str(v)


def _strip_frame_strings(df: pd.DataFrame) -> None:
    """与 v1 交易明细脚本一致：去掉单元格字符串首尾空格。"""
    for col in df.columns:
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)


def _normalize_seller_sku_like_b4_1(s: Any) -> Any:
    """
    与 v1 B4_1 extract_values 一致：清洗 seller sku 后再落库 seller_sku，
    便于与订单统计里「订单号 + SKU」拼接键对齐。
    """
    if pd.isna(s):
        return None
    if isinstance(s, str):
        t = s.strip()
    else:
        t = str(s).strip()
    if not t:
        return None
    if "amzn.gr." in t:
        return t.split("amzn.gr.")[-1].split("-")[0].split("_")[0]
    return t.split("#")[0].split("BCFBAFL")[0]


def _read_transaction_frame(xlsx: Path) -> pd.DataFrame:
    _LOG.warn(f"读取 Excel：{xlsx} sheet=0 header=第1行")
    df = pd.read_excel(xlsx, sheet_name=0, header=0, engine="openpyxl", dtype=object)
    df.columns = [("" if c is None else str(c)).replace("\n", " ").strip() for c in df.columns]
    df = df.dropna(how="all")
    _strip_frame_strings(df)
    if "seller sku" in df.columns:
        df["seller sku"] = df["seller sku"].apply(_normalize_seller_sku_like_b4_1)
    _LOG.info(f"读取完成：行数={len(df)} 列数={len(df.columns)}")
    return df


def _row_dict(series: pd.Series, source_kind: str) -> dict[str, Any]:
    out: dict[str, Any] = {"source_kind": source_kind}
    for excel_col, db_col, kind in _AMZ_TXN_MAP:
        if excel_col not in series.index:
            out[db_col] = None
            continue
        out[db_col] = _convert(series[excel_col], kind)
    return out


def _insert_columns() -> list[str]:
    return ["line_hash", "source_kind"] + [c for _, c, _ in _AMZ_TXN_MAP]


def _missing_excel_columns(df: pd.DataFrame) -> list[str]:
    required = [zh for zh, _, _ in _AMZ_TXN_MAP]
    return [c for c in required if c not in df.columns]


def import_file(conn, xlsx: Path) -> tuple[int, int, int]:
    """
    Returns:
        (executemany 累计行数, 跳过行数, Excel 有效行数)
    """
    source_kind = _source_kind_from_filename(xlsx)
    df = _read_transaction_frame(xlsx)
    missing = _missing_excel_columns(df)
    if missing:
        _LOG.error(f"Excel 缺少列（共 {len(missing)}）：{missing[:12]}{'...' if len(missing) > 12 else ''}")
        return 0, 0, len(df)

    insert_cols = _insert_columns()
    rows: list[tuple[Any, ...]] = []
    skipped = 0
    hash_body_keys = tuple(k for k in LINE_HASH_KEYS if k != "source_kind")
    for _, series in df.iterrows():
        d = _row_dict(series, source_kind)
        if all(
            d.get(k) is None or (isinstance(d.get(k), str) and not str(d.get(k)).strip())
            for k in hash_body_keys
        ):
            skipped += 1
            continue
        h_in = row_subset_for_line_hash(d, LINE_HASH_KEYS)
        d["line_hash"] = stable_line_hash(h_in)
        rows.append(tuple(d[c] for c in insert_cols))

    n_excel = len(df)
    if not rows:
        _LOG.warn(f"无有效行可写：Excel 行数={n_excel} 跳过={skipped}")
        return 0, skipped, n_excel

    _LOG.info(
        f"准备写入 MySQL：表={TABLE} 行数={len(rows)}（跳过空行={skipped}，"
        f"line_hash 键数={len(LINE_HASH_KEYS)}）"
    )
    n = upsert_rows(conn, table=TABLE, columns=insert_cols, rows=rows)
    _LOG.info(f"MySQL executemany 完成：批次累计行数={n}")
    return n, skipped, n_excel


def main() -> int:
    setup_stdout_utf8()
    ap = argparse.ArgumentParser(description="导入 transaction交易明细-*.xlsx -> amz_transaction")
    ap.add_argument("--file", type=Path, default=None, help="指定单个 xlsx")
    ap.add_argument(
        "--dir",
        type=Path,
        default=None,
        help=f"Excel 目录，默认 {default_amz_transaction_excel_dir()}",
    )
    args = ap.parse_args()
    base = args.dir or default_amz_transaction_excel_dir()
    _LOG.info(f"任务：Amazon 交易明细导入 -> {TABLE}")
    if os.environ.get("ORDER_IMPORT_VERBOSE") == "1":
        _LOG.info("line_hash 键列表 LINE_HASH_KEYS: " + ", ".join(LINE_HASH_KEYS))
    if not base.is_dir():
        _LOG.error(f"目录不存在: {base}")
        return 2

    if args.file:
        files = [args.file]
    else:
        files = sorted(
            p
            for p in base.glob("transaction交易明细*.xlsx")
            if p.is_file() and not p.name.startswith("~$")
        )
    if not files:
        _LOG.error(f"未找到 xlsx: {base}（匹配 transaction交易明细*.xlsx，排除 ~$）")
        return 1

    _LOG.info(f"待处理文件数={len(files)} 目录={base}")

    cfg = load_db_config()
    _LOG.info(f"连接数据库：host={cfg.host} port={cfg.port} database={cfg.database} user={cfg.user}")
    conn = connect(cfg)
    total_upsert = 0
    total_skip = 0
    try:
        for p in files:
            if not p.is_file():
                _LOG.warn(f"跳过（非文件）: {p}")
                continue
            n, skipped, n_excel = import_file(conn, p)
            conn.commit()
            _LOG.info(f"已提交：{p.name} Excel行={n_excel} 写入UPSERT累计={n} 跳过={skipped}")
            total_upsert += n
            total_skip += skipped
        _LOG.info(f"全部完成：写入UPSERT累计行数={total_upsert} 总跳过={total_skip} 文件数={len(files)}")
        return 0
    except Exception:
        conn.rollback()
        _LOG.error("已回滚事务")
        raise
    finally:
        conn.close()
        _LOG.info("数据库连接已关闭")


if __name__ == "__main__":
    raise SystemExit(main())
