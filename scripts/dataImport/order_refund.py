from __future__ import annotations

"""
退款信息录入   Excel -> sales_order_refund
从 ERP 共享目录读取「RMA-*.xlsx」，写入 MySQL 表 sales_order_refund。

路径与 order_shipped 相同：ERP_ORDER_STA_PATH + MODE_PATTERN + DATE_PATH
  每天 -> .../ERP订单、RMA下载/2026-06-09/RMA-*.xlsx

line_hash：LINE_HASH_KEYS 子集经 stable_line_hash（键排序 JSON + SHA-256）。

币种校验：导入前读取 Excel A1 单元格，解析币种代码须为 EUR，否则中止导入，
并以彩色输出表格元信息。

report_hash：未指定 --import-batch 时，默认取 run_batch.lock 的 import_batch。

用法：
  cd d:\\py-project\\report
  python scripts\\dataImport\\order_refund.py
  python scripts\\dataImport\\order_refund.py --date 2026-06-09
  python scripts\\dataImport\\order_refund.py --file "\\\\Betohow\\...\\RMA-6.1-6.9.xlsx"
"""

import argparse
import os
import re
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd

_REPORT_ROOT = Path(__file__).resolve().parents[2]
_DATA_IMPORT_DIR = Path(__file__).resolve().parent
for _p in (_REPORT_ROOT, _DATA_IMPORT_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from batch_lock import read_import_batch_from_lock  # noqa: E402
from config.path_config import DATE_PATH, ERP_ORDER_STA_PATH, MODE_PATTERN  # noqa: E402
from database.db_connection import DatabaseConfig, get_db_manager  # noqa: E402
from import_common import (  # noqa: E402
    cell_decimal,
    cell_dt,
    cell_str,
    cell_str_or_empty,
    row_subset_for_line_hash,
    stable_line_hash,
    upsert_rows,
)

TABLE = "sales_order_refund"
SOURCE_TYPE = "Excel"
EXPECTED_CURRENCY = "EUR"  # RMA Excel A1 单元格须为该币种
_REFUND_SHEET_NAMES: tuple[str | int, ...] = ("RMA退款", 0)

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

# Excel 表头在第 3 行 -> pandas header=2；工作表优先「RMA退款」
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

# ======================== RMA退款 唯一性逻辑键 =============================================== 
# 注意：这里不可随意修改，否则历史行的 line_hash 与库内不一致，需按业务重导或清表
# 注意：这里不可随意修改，否则历史行的 line_hash 与库内不一致，需按业务重导或清表
LINE_HASH_KEYS: tuple[str, ...] = (
    "platform",
    "platform_site",
    "shop_name_en",
    "warehouse_name",
    "refund_orig_order_no",
    "refund_orig_ref_no",
    "rma_product_sku",
    "refund_amount",
    "rma_created_at",
)
# ================================================================================ 


def _log(level: str, msg: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] [{level}] {msg}")


def _convert(v: Any, kind: str) -> Any:
    if kind == "dt":
        return cell_dt(v)
    if kind == "dec":
        return cell_decimal(v)
    if kind.startswith("s") and kind.endswith("e"):
        return cell_str_or_empty(v, int(kind[1:-1]))
    if kind.startswith("s"):
        return cell_str(v, int(kind[1:]))
    return cell_str(v)


def _row_dict(series: pd.Series) -> dict[str, Any]:
    out: dict[str, Any] = {col: None for _, col, _ in _REFUND_MAP}
    for zh, col, kind in _REFUND_MAP:
        if zh not in series.index:
            continue
        out[col] = _convert(series[zh], kind)
    out["source_type"] = SOURCE_TYPE
    return out


def _parse_currency_code(cell: Any) -> str:
    """
    从 A1/A3 类单元格解析 ISO 币种代码（大写）。
    支持：EUR、币种:EUR、Currency: eur、币种： EUR 等常见写法。
    """
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return ""
    s = str(cell).strip()
    if not s:
        return ""
    compact = re.sub(r"\s+", "", s.upper())
    if re.fullmatch(r"[A-Z]{3}", compact):
        return compact
    for sep in (":", "："):
        if sep not in s:
            continue
        tail = re.sub(r"\s+", "", s.split(sep)[-1].strip().upper())
        m = re.match(r"^([A-Z]{3})\b", tail)
        if m:
            return m.group(1)
    return ""


def _resolve_refund_sheet(xlsx: Path) -> str | int:
    """解析 RMA 工作表名（优先「RMA退款」，否则第一个 sheet）。"""
    last: Exception | None = None
    for sn in _REFUND_SHEET_NAMES:
        try:
            pd.read_excel(xlsx, sheet_name=sn, nrows=1, engine="openpyxl")
            return sn
        except Exception as e:
            last = e
    raise RuntimeError(f"无法打开工作表（需「RMA退款」或第一个 sheet）：{xlsx}") from last


def _read_excel_a1_cell(xlsx: Path, *, sheet: str | int) -> Any:
    """读取 RMA Excel A1 单元格原始值。"""
    header_df = pd.read_excel(
        xlsx,
        sheet_name=sheet,
        header=None,
        usecols=[0],
        nrows=1,
        engine="openpyxl",
    )
    return header_df.iloc[0, 0] if len(header_df) > 0 else None


def _print_excel_header_info(
    xlsx: Path,
    *,
    sheet: str | int,
    a1_raw: Any,
    currency_code: str,
) -> None:
    """币种校验通过后，彩色输出 A1 表格元信息。"""
    a1_text = _cell_display(a1_raw) or "(空)"
    sep = _c("=" * 60, "CYAN")
    print(sep, flush=True)
    print(
        f"{_c('[源文件]', 'BOLD', 'CYAN')} {_c(xlsx.name, 'BOLD')} "
        f"{_c(f'(sheet={sheet!r})', 'CYAN')}",
        flush=True,
    )
    print(
        f"  {_c('币种校验：', 'YELLOW')}"
        f"{_c(a1_text, 'GREEN')} "
        f"→ 解析={_c(currency_code, 'BOLD', 'GREEN')} "
        f"{_c('✓ 校验通过', 'BOLD', 'GREEN')}",
        flush=True,
    )
    print(sep, flush=True)


def _validate_excel_currency(xlsx: Path) -> str | int:
    """
    校验 Excel A1 单元格币种为 EXPECTED_CURRENCY（默认 EUR）。
    校验通过后彩色输出 A1 信息；不符则中止导入。
    返回实际使用的工作表名。
    """
    sheet = _resolve_refund_sheet(xlsx)
    a1_raw = _read_excel_a1_cell(xlsx, sheet=sheet)
    code = _parse_currency_code(a1_raw)
    if code != EXPECTED_CURRENCY:
        print(
            f"{_c('[币种校验失败]', 'BOLD', 'RED')} "
            f"文件={xlsx.name} "
            f"sheet={sheet!r} "
            f"A1={_cell_display(a1_raw)!r} "
            f"解析={code!r} "
            f"要求={EXPECTED_CURRENCY}",
            flush=True,
        )
        raise RuntimeError(
            f"币种非 {EXPECTED_CURRENCY}（文件={xlsx.name}，A1={a1_raw!r}，解析={code!r}），导入已中止"
        )
    _print_excel_header_info(
        xlsx,
        sheet=sheet,
        a1_raw=a1_raw,
        currency_code=code,
    )
    return sheet


def _read_refund_frame(xlsx: Path, *, sheet: str | int | None = None) -> pd.DataFrame:
    last: Exception | None = None
    sheet_used: str | int | None = sheet
    df: pd.DataFrame | None = None
    candidates = (sheet_used,) if sheet_used is not None else _REFUND_SHEET_NAMES
    for sn in candidates:
        if sn is None:
            continue
        try:
            df = pd.read_excel(xlsx, sheet_name=sn, header=2, engine="openpyxl", dtype=object)
            sheet_used = sn
            break
        except Exception as e:
            last = e
    if df is None:
        raise RuntimeError(f"无法读取工作表（需「RMA退款」或第一个 sheet）：{xlsx}") from last

    _log("INFO", f"读取 Excel：{xlsx} sheet={sheet_used!r}（表头第 3 行）")
    df.columns = [("" if c is None else str(c)).replace("\n", " ").strip() for c in df.columns]
    df = df.dropna(how="all")
    _log("INFO", f"读取完成：行数={len(df)} 列数={len(df.columns)}")
    return df


def _insert_columns(*, with_report_hash: bool = False) -> list[str]:
    cols = ["line_hash"]
    cols.extend(col for _, col, _ in _REFUND_MAP)
    if with_report_hash:
        cols.append("report_hash")
    cols.append("source_type")
    return cols


def import_file(conn, xlsx: Path, *, import_batch: str | None = None) -> tuple[int, int, int]:
    """返回 (UPSERT 行数, 跳过行数, Excel 总行数)。"""
    sheet = _validate_excel_currency(xlsx)
    df = _read_refund_frame(xlsx, sheet=sheet)
    insert_cols = _insert_columns(with_report_hash=import_batch is not None)
    dicts: list[dict[str, Any]] = []
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
        dicts.append(d)

    if not dicts:
        _log("WARN", f"无有效行：Excel 行数={len(df)} 跳过={skipped}")
        return 0, skipped, len(df)

    rows: list[tuple[Any, ...]] = []
    for d in dicts:
        h_in = row_subset_for_line_hash(d, LINE_HASH_KEYS)
        d["line_hash"] = stable_line_hash(h_in)
        if import_batch is not None:
            d["report_hash"] = import_batch
        rows.append(tuple(d[c] for c in insert_cols))

    _log(
        "INFO",
        f"准备写入 {TABLE}：有效行={len(rows)} 跳过={skipped} line_hash 键数={len(LINE_HASH_KEYS)}",
    )
    n = upsert_rows(conn, table=TABLE, columns=insert_cols, rows=rows)
    return n, skipped, len(df)


def erp_base_dir(mode: str | None = None) -> Path:
    return Path(ERP_ORDER_STA_PATH.format(MODE_PATTERN=mode or MODE_PATTERN))


def resolve_date_dir(base: Path, mode: str, on_date: date) -> Path:
    if mode == "每月":
        return base / f"{on_date.year:04d}-{on_date.month:02d}"
    return base / on_date.isoformat()


def default_date_dir(base: Path) -> Path:
    return base / DATE_PATH


def resolve_work_dir(base: Path, mode: str, on_date: date | None) -> Path:
    if on_date is not None:
        return resolve_date_dir(base, mode, on_date)
    return default_date_dir(base)


def discover_refund_files(directory: Path) -> list[Path]:
    return sorted(
        p for p in directory.glob("RMA*.xlsx") if p.is_file() and not p.name.startswith("~$")
    )


def _resolve_import_batch(cli_batch: str | None) -> str | None:
    """report_hash 优先取命令行 --import-batch，否则读 run_batch.lock 的 import_batch。"""
    if cli_batch and cli_batch.strip():
        return cli_batch.strip()
    batch = read_import_batch_from_lock()
    if batch:
        _log("INFO", f"从 run_batch.lock 读取 report_hash：{batch}")
    return batch


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="ERP RMA Excel -> sales_order_refund")
    ap.add_argument(
        "--date",
        type=date.fromisoformat,
        default=None,
        metavar="YYYY-MM-DD",
        help=f"覆盖日期子目录（默认 path_config.DATE_PATH={DATE_PATH}）",
    )
    ap.add_argument(
        "--mode",
        choices=("每天", "每月"),
        default=None,
        help=f"路径模式，默认 path_config.MODE_PATTERN（{MODE_PATTERN}）",
    )
    ap.add_argument("--dir", type=Path, default=None, help="直接指定 Excel 目录")
    ap.add_argument("--file", type=Path, default=None, help="指定单个 RMA xlsx")
    ap.add_argument(
        "--import-batch",
        "--batch",
        dest="import_batch",
        default=None,
        metavar="BATCH",
        help="导入批次号写入 report_hash（默认读 run_batch.lock 的 import_batch）",
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
    import_batch = _resolve_import_batch(args.import_batch)

    if args.file:
        files = [args.file.resolve()]
        work_dir = args.file.parent
    elif args.dir:
        work_dir = args.dir.resolve()
        if not work_dir.is_dir():
            _log("ERROR", f"目录不存在：{work_dir}")
            return 2
        files = discover_refund_files(work_dir)
    else:
        work_dir = resolve_work_dir(erp_base_dir(mode), mode, args.date)
        if not work_dir.is_dir():
            _log("ERROR", f"日期目录不存在：{work_dir}")
            return 2
        files = discover_refund_files(work_dir)

    if not files:
        _log("ERROR", f"未找到 RMA-*.xlsx：{work_dir}")
        return 1

    _log("INFO", f"任务：导入 -> {TABLE}")
    _log("INFO", f"模式={mode} 目录={work_dir} 文件数={len(files)}")

    db = get_db_manager(DatabaseConfig())
    conn = db.get_connection()
    total_upsert = 0
    total_skip = 0
    try:
        for fp in files:
            n, skipped, n_excel = import_file(conn, fp, import_batch=import_batch)
            conn.commit()
            _log("INFO", f"已提交：{fp.name} Excel行={n_excel} UPSERT={n} 跳过={skipped}")
            total_upsert += n
            total_skip += skipped
        _log("INFO", f"全部完成：UPSERT累计={total_upsert} 总跳过={total_skip}")
        return 0
    except Exception:
        conn.rollback()
        _log("ERROR", "导入失败，已回滚")
        raise
    finally:
        conn.close()


if __name__ == "__main__":
    raise SystemExit(main())
