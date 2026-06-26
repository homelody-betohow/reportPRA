from __future__ import annotations

import hashlib
import json
import math
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable, Sequence

import mysql.connector


def v2_root() -> Path:
    return Path(__file__).resolve().parents[1]


def default_order_excel_dir() -> Path:
    return v2_root().parent / "excel" / "daily" / "order"


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
        s = s[:max_len]
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
    try:
        import pandas as pd

        ts = pd.to_datetime(v, errors="coerce")
        if pd.isna(ts):
            return None
        return datetime.fromisoformat(ts.to_pydatetime().isoformat(sep=" "))
    except Exception:
        return None


def cell_margin_rate(v: Any) -> Decimal | None:
    """毛利率：支持小数 0.225 或百分数字符串 22.5%。"""
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
    """
    只取指定键构造 stable_line_hash 的入参；缺失键显式为 None，保证键集合固定。
    自定义 line_hash 时：在导入脚本里定义 LINE_HASH_KEYS 元组，再对本函数返回值调用 stable_line_hash。
    """
    return {k: row.get(k) for k in keys}


def stable_line_hash(field_values: dict[str, Any]) -> str:
    """
    与表 uk_*_line_hash 一致：对「即将落库的业务字段」生成稳定 SHA-256（64 位小写 hex）。

    步骤：
      1. 仅使用传入 dict 的键值（调用方应已排除 line_hash、id、created_at、updated_at、文件名等）。
      2. 每个值经 norm_for_hash：None/空串归一化为 JSON null；datetime → isoformat；
         Decimal/float → 规范化十进制字符串；str → strip 后空则 null。
      3. 按键名字母序排序后 json.dumps(..., ensure_ascii=False, separators=(',', ':'), sort_keys=True)。
      4. UTF-8 编码后 SHA-256 hexdigest。

    参与字段集合由各导入脚本定义：常用 row_subset_for_line_hash(row, LINE_HASH_KEYS)，
    LINE_HASH_KEYS 可为全量列或自定义子集；改子集即改变 hash 语义。
    """
    payload = {k: norm_for_hash(field_values[k]) for k in sorted(field_values.keys())}
    raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":"), sort_keys=True)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def upsert_rows(
    conn: mysql.connector.MySQLConnection,
    *,
    table: str,
    columns: Sequence[str],
    rows: Iterable[tuple[Any, ...]],
    chunk_size: int = 300,
) -> int:
    """按 line_hash 唯一键执行 INSERT ... ON DUPLICATE KEY UPDATE（MySQL 8 行别名 src）。"""
    cols_sql = ", ".join(f"`{c}`" for c in columns)
    placeholders = ", ".join(["%s"] * len(columns))
    updates = ", ".join(f"`{c}`=src.`{c}`" for c in columns if c != "line_hash")
    sql = (
        f"INSERT INTO `{table}` ({cols_sql}) VALUES ({placeholders}) AS src "
        f"ON DUPLICATE KEY UPDATE {updates}"
    )
    cur = conn.cursor()
    n = 0
    buf: list[tuple[Any, ...]] = []
    for row in rows:
        buf.append(row)
        if len(buf) >= chunk_size:
            cur.executemany(sql, buf)
            n += len(buf)
            buf.clear()
    if buf:
        cur.executemany(sql, buf)
        n += len(buf)
    cur.close()
    return n
