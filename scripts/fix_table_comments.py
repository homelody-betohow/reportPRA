"""
修复 MySQL 字段/表 COMMENT 乱码（建库时客户端为 latin1 导致 UTF-8 被写坏）。
在 report 目录执行：python scripts/fix_table_comments.py
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

_root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_root))

import pymysql
from database.db_connection import DatabaseConfig

SCHEMA = _root / "database" / "schema.sql"
TABLE_COMMENT_RE = re.compile(r"COMMENT='([^']*)'\s*;\s*$", re.M)
COL_COMMENT_RE = re.compile(
    r"^\s*`?(\w+)`?\s+(.+?)\s+COMMENT\s+'((?:[^'\\]|\\.)*)'\s*,?\s*$",
    re.M,
)


def _parse_schema(text: str) -> dict[str, dict[str, str]]:
    """返回 {表名: {列名: 备注, '__TABLE__': 表备注}}"""
    result: dict[str, dict[str, str]] = {}
    blocks = re.split(r"CREATE TABLE\s+(\w+)\s*\(", text, flags=re.I)
    for i in range(1, len(blocks), 2):
        table = blocks[i]
        body = blocks[i + 1]
        end = body.find(") ENGINE")
        if end < 0:
            continue
        section = body[:end]
        cols: dict[str, str] = {}
        for line in section.splitlines():
            line = line.strip()
            if not line or line.startswith(("PRIMARY", "UNIQUE", "INDEX", "KEY", "CONSTRAINT")):
                continue
            m = re.search(r"COMMENT\s+'((?:[^'\\]|\\.)*)'\s*,?\s*$", line)
            if not m:
                continue
            comment = m.group(1)
            col_part = line[: m.start()].strip().rstrip(",")
            col_name = col_part.split()[0].strip("`")
            cols[col_name] = comment
        tail = body[end:] if end >= 0 else body
        tm = re.search(r"COMMENT='([^']*)'", tail)
        if tm:
            cols["__TABLE__"] = tm.group(1)
        if cols:
            result[table] = cols
    return result


def _column_ddl(cur, schema: str, table: str, col: str) -> str:
    cur.execute(
        """
        SELECT COLUMN_TYPE, IS_NULLABLE, COLUMN_DEFAULT, EXTRA
        FROM information_schema.COLUMNS
        WHERE TABLE_SCHEMA=%s AND TABLE_NAME=%s AND COLUMN_NAME=%s
        """,
        (schema, table, col),
    )
    row = cur.fetchone()
    if not row:
        raise RuntimeError(f"列不存在: {table}.{col}")
    ctype, nullable, default, extra = row
    extra = (extra or "").replace("DEFAULT_GENERATED", "").strip()
    null_sql = "NOT NULL" if nullable == "NO" else "NULL"
    parts = [f"`{col}`", ctype, null_sql]
    if default is not None:
        d = str(default)
        if d.upper() in ("CURRENT_TIMESTAMP", "CURRENT_TIMESTAMP()"):
            parts.append("DEFAULT CURRENT_TIMESTAMP")
        elif ctype.split("(")[0].lower() in (
            "int", "tinyint", "smallint", "bigint", "decimal", "float", "double", "bit"
        ):
            parts.append(f"DEFAULT {d}")
        else:
            parts.append(f"DEFAULT '{d}'")
    elif nullable == "YES" and "auto_increment" not in extra.lower():
        parts.append("DEFAULT NULL")
    if extra:
        parts.append(extra)
    return " ".join(parts)


def main() -> int:
    cfg = DatabaseConfig()
    meta = _parse_schema(SCHEMA.read_text(encoding="utf-8"))
    conn = pymysql.connect(
        host=cfg.host,
        port=cfg.port,
        user=cfg.user,
        password=cfg.password,
        database=cfg.database,
        charset="utf8mb4",
        use_unicode=True,
    )
    try:
        with conn.cursor() as cur:
            cur.execute("SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci")
            fixed = 0
            for table, cols in meta.items():
                table_comment = cols.pop("__TABLE__", None)
                for col, comment in cols.items():
                    ddl = _column_ddl(cur, cfg.database, table, col)
                    sql = f"ALTER TABLE `{table}` MODIFY COLUMN {ddl} COMMENT %s"
                    cur.execute(sql, (comment,))
                    fixed += 1
                if table_comment:
                    cur.execute(
                        f"ALTER TABLE `{table}` COMMENT %s",
                        (table_comment,),
                    )
                    fixed += 1
            conn.commit()
        print(f"[OK] 已修复 {fixed} 条字段/表备注（库: {cfg.database}）")
        print("请在 Navicat 中断开重连后刷新表结构。")
        return 0
    finally:
        conn.close()


if __name__ == "__main__":
    raise SystemExit(main())
