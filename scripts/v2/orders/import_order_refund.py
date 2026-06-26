from __future__ import annotations

"""
RMA 退款 Excel -> 表 sales_order_refund。

落库列见 _REFUND_MAP；line_hash 仅使用下方 LINE_HASH_KEYS（须为 LINE_HASH_SOURCE_FIELDS 子集）。
  改规则后需按业务重导或清表；字段越少越容易不同行撞同一 hash。
  算法：excel_common.row_subset_for_line_hash + stable_line_hash。

日志风格对齐 warehouse-rent/import_provider_4px_detail.py：读取、行数、写入、提交。
可选：环境变量 ORDER_IMPORT_VERBOSE=1 启动时额外打印全部 line_hash 字段名。
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
    cell_str_or_empty,
    default_order_excel_dir,
    row_subset_for_line_hash,
    stable_line_hash,
    upsert_rows,
)

TABLE = "sales_order_refund"
SOURCE_TYPE = "Excel"
_LOG = get_logger("ORDER-REFUND")

_REFUND_MAP: list[tuple[str, str, str]] = [
    ("平台", "platform", "s64"),
    ("店铺别名", "shop_alias", "s128"),
    ("店铺英文名", "shop_name_en", "s128"),
    ("站点", "platform_site", "s64"),
    ("仓库名称", "warehouse_name", "s255"),
    ("订单目的国家", "order_dest_country", "s64"),
    ("RMA创建时间", "rma_created_at", "dt"),
    ("RMA审核时间", "rma_audit_at", "dt"),
    ("RMA退款时间", "rma_refund_at", "dt"),
    ("原订单付款时间", "orig_order_paid_at", "dt"),
    ("退款原订单号", "refund_orig_order_no", "s128"),
    ("退款原订单参考号", "refund_orig_ref_no", "s128e"),
    ("退款原订单跟踪号", "refund_orig_track_no", "s255"),
    ("PayPal退款交易号", "paypal_refund_txn_no", "s255"),
    ("退款类型", "refund_type", "s64"),
    ("运输方式", "shipping_method", "s128"),
    ("运输方式名称", "shipping_method_name", "s255"),
    ("退款状态", "refund_status", "s64"),
    ("退款方式", "refund_method", "s64"),
    ("RMA产品", "rma_product_sku", "s128"),
    ("RMA产品数量", "rma_product_qty", "dec"),
    ("币种", "currency_code", "s16"),
    ("产品名称", "product_name", "s512"),
    ("一级品类", "category_lv1", "s128"),
    ("二级品类", "category_lv2", "s128"),
    ("三级品类", "category_lv3", "s128"),
    ("产品款式", "product_style", "s512"),
    ("退款金额", "refund_amount", "dec"),
    ("退款原因", "refund_reason", "s512"),
    ("平台退款原因", "platform_refund_reason", "s512"),
    ("创建人", "created_by", "s128"),
    ("退款备注", "refund_remark", "s2048"),
    ("财务备注", "finance_remark", "s2048"),
    ("产品默认采购员账号", "default_buyer_acct", "s128"),
    ("产品默认采购员", "default_buyer_name", "s128"),
    ("销售负责人账号", "sales_owner_acct", "s128"),
    ("销售负责人", "sales_owner", "s128"),
    ("开发负责人账号", "dev_owner_acct", "s128"),
    ("开发负责人", "dev_owner", "s128"),
    ("运营负责人", "ops_owner", "s128"),
    ("产品问题类型", "product_issue_type", "s128"),
    ("问题分类", "issue_category", "s128"),
    ("产品问题", "product_issue", "s2048"),
]

# 落库行里会出现的全部业务键（不含 line_hash）；用于校验 LINE_HASH_KEYS
LINE_HASH_SOURCE_FIELDS: tuple[str, ...] = tuple(c for _, c, _ in _REFUND_MAP) + ("source_type",)
# 参与 line_hash 的键, [请谨慎修改，否则历史行的 line_hash 与库内不一致，需按业务重导或清表]
# ("refund_orig_order_no", "refund_orig_ref_no", "rma_product_sku", "refund_amount")
LINE_HASH_KEYS: tuple[str, ...] = (
    "platform",
    "platform_site",
    "shop_name_en",
    "warehouse_name",
    "refund_orig_order_no",
    "refund_orig_ref_no",
    "rma_product_sku",
    "refund_amount",
    "rma_created_at"
)

_src_set = frozenset(LINE_HASH_SOURCE_FIELDS)
_bad_keys = tuple(k for k in LINE_HASH_KEYS if k not in _src_set)
if _bad_keys:
    raise ValueError(f"LINE_HASH_KEYS 含未知列（非 LINE_HASH_SOURCE_FIELDS 子集）: {_bad_keys}")


def _convert(v: Any, kind: str) -> Any:
    if kind == "dt":
        return cell_dt(v)
    if kind == "dec":
        return cell_decimal(v)
    if kind.startswith("s") and kind.endswith("e"):
        n = int(kind[1:-1])
        return cell_str_or_empty(v, n)
    if kind.startswith("s"):
        n = int(kind[1:])
        return cell_str(v, n)
    return cell_str(v)


def _row_dict(series: pd.Series) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for zh, col, kind in _REFUND_MAP:
        if zh not in series.index:
            continue
        out[col] = _convert(series[zh], kind)
    out["source_type"] = SOURCE_TYPE
    return out


def _read_refund_frame(xlsx: Path) -> pd.DataFrame:
    last: Exception | None = None
    sheet_used: str | int | None = None
    df: pd.DataFrame | None = None
    for sn in ("RMA退款", 0):
        try:
            df = pd.read_excel(xlsx, sheet_name=sn, header=2, engine="openpyxl", dtype=object)
            sheet_used = sn
            break
        except Exception as e:
            last = e
            continue
    else:
        raise RuntimeError(f"无法读取工作表（需「RMA退款」或第一个 sheet）: {xlsx}") from last
    _LOG.warn(f"读取 Excel：{xlsx} sheet={sheet_used!r} header=第3行")
    df.columns = [("" if c is None else str(c)).replace("\n", " ").strip() for c in df.columns]
    df = df.dropna(how="all")
    _LOG.info(f"读取完成：行数={len(df)} 列数={len(df.columns)}")
    return df


def _insert_columns() -> list[str]:
    cols = ["line_hash"]
    for _, c, _ in _REFUND_MAP:
        cols.append(c)
    cols.append("source_type")
    return cols


def import_file(conn, xlsx: Path) -> tuple[int, int, int]:
    df = _read_refund_frame(xlsx)
    insert_cols = _insert_columns()
    rows: list[tuple[Any, ...]] = []
    skipped = 0
    for _, series in df.iterrows():
        d = _row_dict(series)
        order_no = d.get("refund_orig_order_no")
        sku = d.get("rma_product_sku")
        if not order_no or not str(order_no).strip() or not sku or not str(sku).strip():
            skipped += 1
            continue
        if d.get("refund_orig_ref_no") is None:
            d["refund_orig_ref_no"] = ""
        h_in = row_subset_for_line_hash(d, LINE_HASH_KEYS)
        d["line_hash"] = stable_line_hash(h_in)
        rows.append(tuple(d[c] for c in insert_cols))
    n_excel = len(df)
    if not rows:
        _LOG.warn(f"无有效行可写：Excel 行数={n_excel} 跳过={skipped}（缺退款原订单号或 RMA产品）")
        return 0, skipped, n_excel
    _LOG.info(
        f"准备写入 MySQL：表={TABLE} 行数={len(rows)}（跳过={skipped}，line_hash 键数={len(LINE_HASH_KEYS)}）"
    )
    n = upsert_rows(conn, table=TABLE, columns=insert_cols, rows=rows)
    _LOG.info(f"MySQL executemany 完成：批次累计行数={n}")
    return n, skipped, n_excel


def main() -> int:
    setup_stdout_utf8()
    ap = argparse.ArgumentParser(description="导入 RMA-*.xlsx -> sales_order_refund")
    ap.add_argument("--file", type=Path, default=None, help="指定单个 xlsx；默认导入目录下全部 RMA-*.xlsx")
    ap.add_argument("--dir", type=Path, default=None, help=f"Excel 目录，默认 {default_order_excel_dir()}")
    args = ap.parse_args()
    base = args.dir or default_order_excel_dir()
    _LOG.info(f"任务：RMA 退款明细导入 -> {TABLE}")
    _LOG.info(
        f"line_hash 参与键共 {len(LINE_HASH_KEYS)} 个（LINE_HASH_KEYS）；"
        "说明见模块文档字符串；算法 row_subset_for_line_hash + stable_line_hash"
    )
    if os.environ.get("ORDER_IMPORT_VERBOSE") == "1":
        _LOG.info("line_hash 键列表 LINE_HASH_KEYS: " + ", ".join(LINE_HASH_KEYS))
    if not base.is_dir():
        _LOG.error(f"目录不存在: {base}")
        return 2

    if args.file:
        files = [args.file]
    else:
        files = sorted(base.glob("RMA*.xlsx"))
    if not files:
        _LOG.error(f"未找到 RMA*.xlsx: {base}")
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
