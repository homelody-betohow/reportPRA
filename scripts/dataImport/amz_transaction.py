from __future__ import annotations

"""
从 ERP 共享目录读取「transaction交易明细-*.xlsx」，写入 MySQL 表 amz_transaction。

路径：config.path_config.TRANSACTION_PATH + MODE_PATTERN + 日期子目录
  每天 -> .../每天/transaction交易明细/2026-06-09/transaction交易明细-*.xlsx
  每月 -> .../每月/transaction交易明细/2026-06/transaction交易明细-*.xlsx

line_hash：对 LINE_HASH_KEYS 子集做 stable_line_hash（键排序 JSON + SHA-256），
与线上去重键 uk_amz_txn_line_hash 一致。

source_kind：由文件名区分，不从 Excel 读取
  transaction交易明细-已发放订单* -> released
  transaction交易明细-已推迟订单* -> deferred

用法：
  cd d:\\py-project\\report
  python scripts\\dataImport\\amz_transaction.py
  python scripts\\dataImport\\amz_transaction.py --date 2026-06-09
  python scripts\\dataImport\\amz_transaction.py --file "\\\\Betohow\\...\\transaction交易明细-*.xlsx"

未指定 --date / --dir / --file 时，日期子目录默认为当天
（日报=今天 YYYY-MM-DD，月报=当月 YYYY-MM；与其它导入脚本的 DATE_PATH 不同）。
"""

import argparse
import os
import sys
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

import pandas as pd

_REPORT_ROOT = Path(__file__).resolve().parents[2]
_DATA_IMPORT_DIR = Path(__file__).resolve().parent
for _p in (_REPORT_ROOT, _DATA_IMPORT_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from config.path_config import MODE_PATTERN, TRANSACTION_PATH  # noqa: E402
from database.db_connection import DatabaseConfig, get_db_manager  # noqa: E402
from import_common import (  # noqa: E402
    cell_decimal,
    cell_dt,
    cell_int,
    cell_str,
    cell_str_or_empty,
    row_subset_for_line_hash,
    stable_line_hash,
    upsert_rows,
)

TABLE = "amz_transaction"
SOURCE_KIND_RELEASED = "released"
SOURCE_KIND_DEFERRED = "deferred"
_RELEASED_NAME_MARK = "transaction交易明细-已发放订单"
_DEFERRED_NAME_MARK = "transaction交易明细-已推迟订单"

# seller_sku/platform_sku 清洗等级：
# - strict：严格（默认），与订单统计对齐口径一致（兼容 # / BCFBAFL / amzn.gr.）
# - loose ：宽松，仅清洗 amzn.gr*（其余原样保留）
SKU_CLEAN_LEVEL_STRICT = "strict"
SKU_CLEAN_LEVEL_LOOSE = "loose"
SKU_CLEAN_LEVEL_DEFAULT = SKU_CLEAN_LEVEL_LOOSE
_SKU_CLEAN_LEVEL = SKU_CLEAN_LEVEL_DEFAULT

# ANSI 终端颜色（与 order_shipped.py、run_batch.py 风格一致）
_RESET = "\033[0m"
_ANSI = {
    "RED": "\033[91m",
    "GREEN": "\033[92m",
    "YELLOW": "\033[93m",
    "CYAN": "\033[96m",
    "BOLD": "\033[1m",
}


def _enable_windows_ansi() -> None:
    if os.name != "nt":
        return
    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32
        for handle_id in (-11, -12):
            handle = kernel32.GetStdHandle(handle_id)
            mode = ctypes.c_uint32()
            if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
                kernel32.SetConsoleMode(handle, mode.value | 0x0004)
    except Exception:
        pass


def _use_color() -> bool:
    if os.getenv("NO_COLOR"):
        return False
    if not hasattr(sys.stdout, "isatty") or not sys.stdout.isatty():
        return False
    return True


def _c(text: str, *styles: str) -> str:
    if not _use_color():
        return text
    codes = "".join(_ANSI.get(s, s) for s in styles)
    return f"{codes}{text}{_RESET}"


def _cell_display(cell: Any) -> str:
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return ""
    return str(cell).strip()


def _strip_frame_strings(df: pd.DataFrame) -> None:
    """去掉单元格字符串首尾空格。"""
    for col in df.columns:
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)


def _normalize_platform_sku(s: Any) -> str | None:
    """
    标准化 seller sku，用于后续与订单统计表对齐。
    注意：不影响 line_hash（line_hash 仍使用原 seller_sku 字段）。
    """
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return None
    t = s.strip() if isinstance(s, str) else str(s).strip()
    if not t:
        return None
    # 宽松：仅处理 amzn.gr*（其余不动）
    level = (_SKU_CLEAN_LEVEL or SKU_CLEAN_LEVEL_DEFAULT).lower()
    t_lower = t.lower()
    is_amzn_gr = t_lower.startswith("amzn.gr") or ("amzn.gr." in t_lower)
    if level == SKU_CLEAN_LEVEL_LOOSE:
        if not is_amzn_gr:
            return t
        # amzn.gr. 规则（大小写无关）
        # 以 "amzn.gr." 为切分点（若不存在但以 amzn.gr 开头，则用原串继续截断）
        tail = t
        if "amzn.gr." in t_lower:
            # 用原字符串按不区分大小写的位置切分
            # 简化：按 lower 的切分长度在原串截取
            idx = t_lower.rfind("amzn.gr.")
            tail = t[idx + len("amzn.gr.") :]
        return tail.split("-")[0].split("_")[0]

    # 严格模式
    if "amzn.gr." in t_lower:
        idx = t_lower.rfind("amzn.gr.")
        tail = t[idx + len("amzn.gr.") :]
        return tail.split("-")[0].split("_")[0]
    return t.split("#")[0].split("BCFBAFL")[0]


# Excel 表头映射（表头第 1 行 header=0；列名以 ERP 导出为准）
_TRANSACTION_MAP: list[tuple[str, str, str]] = [
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

# ======================== transaction交易明细 唯一性逻辑键 =============================================== 
# 注意：这里不可随意修改，否则历史行的 line_hash 与库内不一致，需按业务重导或清表
# 注意：这里不可随意修改，否则历史行的 line_hash 与库内不一致，需按业务重导或清表
# 注意：这里不可随意修改，否则历史行的 line_hash 与库内不一致，需按业务重导或清表
LINE_HASH_KEYS: tuple[str, ...] = (
    "shop_name",
    "marketplace",
    "transaction_type",
    "amazon_order_id",
    "group_id",
    "seller_sku",
    "child_asin"
)
# ================================================================================ 


def _log(level: str, msg: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] [{level}] {msg}", flush=True)


def _log_success(msg: str) -> None:
    """任务全部完成时的提示（绿色粗体）。"""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] [INFO] {msg}"
    print(_c(line, "BOLD", "GREEN"), flush=True)


def _excel_row_no(pandas_idx: int, header_row: int = 0) -> int:
    """计算 Excel 实际行号（表头在第 header_row+1 行）。"""
    return int(pandas_idx) + header_row + 2


def _log_duplicate_line_hashes(hash_groups: dict[str, list[dict[str, Any]]], header_row: int = 0) -> None:
    """输出本批导入中 line_hash 重复的分组明细。"""
    dups = {h: rows for h, rows in hash_groups.items() if len(rows) > 1}
    if not dups:
        _log("INFO", "line_hash 检查：本批无重复")
        return

    extra = sum(len(rows) - 1 for rows in dups.values())
    unique = len(hash_groups)
    total = unique + extra
    _log(
        "WARN",
        f"line_hash 重复：有效行={total}，唯一={unique}，重复组={len(dups)}，多出行数={extra}",
    )
    for i, (h, rows) in enumerate(
        sorted(dups.items(), key=lambda x: x[1][0].get("_excel_row", 0)),
        1,
    ):
        # _log("WARN", f"  重复组 {i}/{len(dups)} line_hash={h}")
        for d in rows:
            key_bits = ", ".join(f"{k}={d.get(k)!r}" for k in LINE_HASH_KEYS)
            excel_row = d.get("_excel_row", "?")
            # _log( "WARN",f"    Excel行{excel_row}: {key_bits}",)


def _convert(v: Any, kind: str) -> Any:
    if kind == "dt":
        return cell_dt(v)
    if kind == "int":
        return cell_int(v)
    if kind == "dec":
        return cell_decimal(v)
    if kind.startswith("s") and kind.endswith("e"):
        return cell_str_or_empty(v, int(kind[1:-1]))
    if kind.startswith("s"):
        return cell_str(v, int(kind[1:]))
    return cell_str(v)


def _row_dict(series: pd.Series) -> dict[str, Any]:
    out: dict[str, Any] = {col: None for _, col, _ in _TRANSACTION_MAP}
    for zh, col, kind in _TRANSACTION_MAP:
        if zh not in series.index:
            continue
        out[col] = _convert(series[zh], kind)
    out["platform_sku"] = _normalize_platform_sku(out.get("seller_sku"))
    return out


def _print_excel_info(xlsx: Path, *, source_kind: str | None) -> None:
    """输出 Excel 文件信息。"""
    sep = _c("=" * 60, "CYAN")
    print(sep, flush=True)
    print(
        f"{_c('[源文件]', 'BOLD', 'CYAN')} {_c(xlsx.name, 'BOLD')}",
        flush=True,
    )
    if source_kind:
        print(
            f"  {_c('source_kind', 'YELLOW')}：{_c(source_kind, 'BOLD', 'GREEN')}",
            flush=True,
        )
    else:
        print(
            f"  {_c('WARN', 'YELLOW')}："
            f"文件名未匹配「已发放订单」或「已推迟订单」，source_kind 将为空",
            flush=True,
        )
    print(sep, flush=True)


def _read_transaction_frame(xlsx: Path, header_row: int = 0) -> pd.DataFrame:
    """读取 transaction 交易明细 Excel。"""
    # _log("INFO", f"读取 Excel：{xlsx}（表头第 {header_row+1} 行）")
    df = pd.read_excel(xlsx, sheet_name=0, header=header_row, engine="openpyxl", dtype=object)
    df.columns = [("" if c is None else str(c)).replace("\n", " ").strip() for c in df.columns]
    df = df.dropna(how="all")
    _strip_frame_strings(df)
    _log("INFO", f"读取完成：行数={len(df)} 列数={len(df.columns)}")
    return df


def _insert_columns() -> list[str]:
    cols = ["line_hash", "source_kind"]
    cols.extend(col for _, col, _ in _TRANSACTION_MAP)
    # 新增列：标准化 SKU（不参与 line_hash；用于后续与订单统计对齐）
    if "seller_sku" in cols and "platform_sku" not in cols:
        cols.insert(cols.index("seller_sku") + 1, "platform_sku")
    return cols


def resolve_source_kind(xlsx: Path) -> str | None:
    """
    根据文件名识别 source_kind。
    transaction交易明细-已发放订单* -> released
    transaction交易明细-已推迟订单* -> deferred
    """
    name = xlsx.name
    if name.startswith(_RELEASED_NAME_MARK) or _RELEASED_NAME_MARK in name:
        return SOURCE_KIND_RELEASED
    if name.startswith(_DEFERRED_NAME_MARK) or _DEFERRED_NAME_MARK in name:
        return SOURCE_KIND_DEFERRED
    return None


def _is_blank_str(v: Any) -> bool:
    if v is None:
        return True
    return not str(v).strip()


def import_file(
    conn,
    xlsx: Path,
    *,
    header_row: int = 0,
) -> tuple[int, int, int]:
    """返回 (UPSERT 行数, 跳过行数, Excel 总行数)。"""
    source_kind = resolve_source_kind(xlsx)
    _print_excel_info(xlsx, source_kind=source_kind)
    if source_kind is None:
        _log(
            "WARN",
            f"无法从文件名识别 source_kind：{xlsx.name}，"
            f"期望前缀 {_RELEASED_NAME_MARK!r} 或 {_DEFERRED_NAME_MARK!r}",
        )
    else:
        _log("INFO", f"source_kind={source_kind}（来自文件名）")
    df = _read_transaction_frame(xlsx, header_row=header_row)
    insert_cols = _insert_columns()
    dicts: list[dict[str, Any]] = []
    skipped = 0

    for idx, series in df.iterrows():
        d = _row_dict(series)
        # 有效行判断：至少需要有 amazon_order_id 或 group_id
        amazon_order_id = d.get("amazon_order_id")
        group_id = d.get("group_id")
        if _is_blank_str(amazon_order_id) and _is_blank_str(group_id):
            skipped += 1
            continue
        d["source_kind"] = source_kind
        d["_excel_row"] = _excel_row_no(idx, header_row)
        dicts.append(d)

    if not dicts:
        _log("WARN", f"无有效行：Excel 行数={len(df)} 跳过={skipped}")
        return 0, skipped, len(df)

    # ================= 去重 / 汇总：同 line_hash 的数值字段做累计 =================
    # 业务确认：若 line_hash 相同，quantity 与金额类字段都需要累计相加；
    # 其余字段保留首次出现的那一行（便于追溯）。
    dec_cols = {col for _, col, kind in _TRANSACTION_MAP if kind == "dec"}
    # 保护：platform_sku 不是从 Excel 读取的，但也应随 seller_sku 的首次值保留
    dec_cols.discard("line_hash")
    dec_cols.discard("source_kind")

    aggregated: dict[str, dict[str, Any]] = {}
    hash_groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for d in dicts:
        h_in = row_subset_for_line_hash(d, LINE_HASH_KEYS)
        h = stable_line_hash(h_in)
        d["line_hash"] = h
        hash_groups[h].append(d)

        if h not in aggregated:
            aggregated[h] = d
            continue

        base = aggregated[h]
        # 数值字段累计（Decimal 或 None）
        for c in dec_cols:
            a = base.get(c)
            b = d.get(c)
            if a is None:
                base[c] = b
            elif b is None:
                continue
            else:
                base[c] = a + b

    _log_duplicate_line_hashes(hash_groups, header_row)
    if len(aggregated) != len(dicts):
        _log(
            "WARN",
            f"同 line_hash 行已累计合并：原有效行={len(dicts)} 合并后={len(aggregated)} "
            f"（减少 {len(dicts) - len(aggregated)} 行）",
        )

    rows: list[tuple[Any, ...]] = [tuple(d.get(c) for c in insert_cols) for d in aggregated.values()]

    _log(
        "INFO",
        f"准备写入 {TABLE}：写入行={len(rows)}（合并前有效行={len(dicts)}）跳过={skipped} "
        f"line_hash 键数={len(LINE_HASH_KEYS)}",
    )
    n = upsert_rows(conn, table=TABLE, columns=insert_cols, rows=rows)
    return n, skipped, len(df)


def transaction_base_dir(mode: str | None = None) -> Path:
    pattern = mode or MODE_PATTERN
    return Path(TRANSACTION_PATH.format(MODE_PATTERN=pattern))


def resolve_date_dir(base: Path, mode: str, on_date: date) -> Path:
    if mode == "每月":
        return base / f"{on_date.year:04d}-{on_date.month:02d}"
    return base / on_date.isoformat()


def default_date_dir(base: Path, mode: str) -> Path:
    """默认日期目录：当天（日报 YYYY-MM-DD，月报 YYYY-MM）。"""
    return resolve_date_dir(base, mode, date.today())


def resolve_work_dir(base: Path, mode: str, on_date: date | None) -> Path:
    if on_date is not None:
        return resolve_date_dir(base, mode, on_date)
    return default_date_dir(base, mode)


def discover_transaction_files(directory: Path) -> list[Path]:
    """查找 transaction交易明细-*.xlsx 文件。"""
    files = sorted(directory.glob("transaction交易明细*.xlsx"))
    if not files:
        files = sorted(directory.glob("*transaction交易明细*.xlsx"))
    return [p for p in files if p.is_file() and not p.name.startswith("~$")]


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="transaction交易明细 Excel -> amz_transaction")
    ap.add_argument(
        "--date",
        type=date.fromisoformat,
        default=None,
        metavar="YYYY-MM-DD",
        help="覆盖日期子目录（默认当天）",
    )
    ap.add_argument(
        "--mode",
        choices=("每天", "每月"),
        default=None,
        help=f"路径模式，默认 path_config.MODE_PATTERN（{MODE_PATTERN}）",
    )
    ap.add_argument("--dir", type=Path, default=None, help="直接指定 Excel 目录")
    ap.add_argument("--file", type=Path, default=None, help="指定单个 xlsx")
    ap.add_argument(
        "--header-row",
        type=int,
        default=0,
        metavar="N",
        help="Excel 表头所在行号（从 0 开始，默认 0 表示第 1 行）",
    )
    ap.add_argument(
        "--sku-clean",
        choices=(SKU_CLEAN_LEVEL_STRICT, SKU_CLEAN_LEVEL_LOOSE),
        default=SKU_CLEAN_LEVEL_DEFAULT,
        help=(
            "platform_sku 清洗等级："
            "strict=严格（默认，与订单统计对齐）；"
            "loose=宽松（仅清洗 amzn.gr*）"
        ),
    )
    return ap.parse_args()


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    _enable_windows_ansi()

    args = parse_args()
    mode = args.mode or MODE_PATTERN
    global _SKU_CLEAN_LEVEL
    _SKU_CLEAN_LEVEL = getattr(args, "sku_clean", SKU_CLEAN_LEVEL_DEFAULT) or SKU_CLEAN_LEVEL_DEFAULT
    _log("INFO", f"platform_sku 清洗等级：{_SKU_CLEAN_LEVEL}")

    if args.file:
        files = [args.file.resolve()]
        work_dir = args.file.parent
    elif args.dir:
        work_dir = args.dir.resolve()
        if not work_dir.is_dir():
            _log("ERROR", f"目录不存在：{work_dir}")
            return 2
        files = discover_transaction_files(work_dir)
    else:
        if args.date is None:
            _log("INFO", f"默认日期：当天 {date.today().isoformat()}")
        work_dir = resolve_work_dir(transaction_base_dir(mode), mode, args.date)
        if not work_dir.is_dir():
            _log("ERROR", f"日期目录不存在：{work_dir}")
            return 2
        files = discover_transaction_files(work_dir)

    if not files:
        _log("ERROR", f"未找到 transaction交易明细-*.xlsx：{work_dir}")
        return 1

    _log("INFO", f"任务：导入 -> {TABLE}")
    _log("INFO", f"模式={mode} 目录={work_dir} 文件数={len(files)}")

    db = get_db_manager(DatabaseConfig())
    conn = db.get_connection()
    total_upsert = 0
    total_skip = 0
    try:
        for fp in files:
            n, skipped, n_excel = import_file(
                conn,
                fp,
                header_row=args.header_row,
            )
            conn.commit()
            _log(
                "INFO",
                f"已提交：{fp.name} Excel行={n_excel} UPSERT={n} 跳过={skipped}",
            )
            total_upsert += n
            total_skip += skipped
        _log_success(
            f"全部完成：UPSERT累计={total_upsert} 总跳过={total_skip}",
        )
        sepLine = _c("=" * 80, "YELLOW")
        print(f"{sepLine}\n\n")
        return 0
    except Exception:
        conn.rollback()
        _log("ERROR", "导入失败，已回滚")
        raise
    finally:
        conn.close()


if __name__ == "__main__":
    raise SystemExit(main())
