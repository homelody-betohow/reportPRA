from __future__ import annotations

import hashlib
import json
import math
import warnings
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Any, Iterable, Sequence

import pandas as pd

# openpyxl 读部分 ERP 导出 xlsx 时会提示缺少默认样式，不影响数据读取
warnings.filterwarnings("ignore", message="Workbook contains no default style")


def _is_blank(v: Any) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return True
    if isinstance(v, str) and not v.strip():
        return True
    return False


def cell_str(v: Any, max_len: int | None = None) -> str | None:
    if _is_blank(v):
        return None
    s = str(v).strip()
    if not s:
        return None
    if max_len is not None and len(s) > max_len:
        return s[:max_len]
    return s


def cell_str_or_empty(v: Any, max_len: int | None = None) -> str:
    s = cell_str(v, max_len=max_len)
    return s if s is not None else ""


def cell_int(v: Any) -> int | None:
    if _is_blank(v):
        return None
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def cell_decimal(v: Any) -> Decimal | None:
    if _is_blank(v):
        return None
    if isinstance(v, Decimal):
        return v
    try:
        return Decimal(str(v).strip().replace(",", ""))
    except (InvalidOperation, ValueError):
        return None


def cell_dt(v: Any) -> datetime | None:
    if _is_blank(v):
        return None
    if isinstance(v, datetime):
        return v
    ts = pd.to_datetime(v, errors="coerce")
    if pd.isna(ts):
        return None
    return ts.to_pydatetime().replace(tzinfo=None)


def cell_margin_rate(v: Any) -> Decimal | None:
    if _is_blank(v):
        return None
    s = str(v).strip().replace(",", "")
    if s.endswith("%"):
        try:
            return Decimal(s[:-1].strip()) / Decimal(100)
        except InvalidOperation:
            return None
    return cell_decimal(s)


def norm_for_hash(v: Any) -> Any:
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.isoformat(sep=" ", timespec="seconds")
    if isinstance(v, Decimal):
        return format(v.normalize(), "f")
    if isinstance(v, bool):
        return v
    if isinstance(v, int):
        return v
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return None
        return format(Decimal(str(v)).normalize(), "f")
    if isinstance(v, str):
        t = v.strip()
        return t if t else None
    return str(v).strip() or None


def row_subset_for_line_hash(row: dict[str, Any], keys: Sequence[str]) -> dict[str, Any]:
    return {k: row.get(k) for k in keys}


def stable_line_hash(field_values: dict[str, Any]) -> str:
    payload = {k: norm_for_hash(field_values[k]) for k in sorted(field_values.keys())}
    raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":"), sort_keys=True)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def upsert_rows(
    conn,
    *,
    table: str,
    columns: Sequence[str],
    rows: Iterable[tuple[Any, ...]],
    chunk_size: int = 300,
) -> int:
    cols_sql = ", ".join(f"`{c}`" for c in columns)
    placeholders = ", ".join(["%s"] * len(columns))
    updates = ", ".join(f"`{c}`=VALUES(`{c}`)" for c in columns if c != "line_hash")
    sql = (
        f"INSERT INTO `{table}` ({cols_sql}) VALUES ({placeholders}) "
        f"ON DUPLICATE KEY UPDATE {updates}"
    )
    cur = conn.cursor()
    n = 0
    buf: list[tuple[Any, ...]] = []
    try:
        for row in rows:
            buf.append(row)
            if len(buf) >= chunk_size:
                cur.executemany(sql, buf)
                n += len(buf)
                buf.clear()
        if buf:
            cur.executemany(sql, buf)
            n += len(buf)
    finally:
        cur.close()
    return n


def insert_ignore_rows(
    conn,
    *,
    table: str,
    columns: Sequence[str],
    rows: Iterable[tuple[Any, ...]],
    chunk_size: int = 300,
) -> int:
    """INSERT IGNORE，遇唯一键冲突跳过；返回累计尝试插入行数。"""
    if not columns:
        return 0
    cols_sql = ", ".join(f"`{c}`" for c in columns)
    placeholders = ", ".join(["%s"] * len(columns))
    sql = f"INSERT IGNORE INTO `{table}` ({cols_sql}) VALUES ({placeholders})"
    cur = conn.cursor()
    n = 0
    buf: list[tuple[Any, ...]] = []
    try:
        for row in rows:
            buf.append(row)
            if len(buf) >= chunk_size:
                cur.executemany(sql, buf)
                n += cur.rowcount if cur.rowcount and cur.rowcount > 0 else 0
                buf.clear()
        if buf:
            cur.executemany(sql, buf)
            if cur.rowcount and cur.rowcount > 0:
                n += cur.rowcount
    finally:
        cur.close()
    return n
