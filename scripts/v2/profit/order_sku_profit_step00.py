from __future__ import annotations

"""
step00 — 回填 sales_order_sku_profit.product_sku（仓库 SKU → 产品 SKU）。

1. 若库中 product_sku_pricing 含 warehouse_sku 列：用定价表非空 warehouse_sku 映射；
   同一 TRIM(warehouse_sku) 多行时取 id 最大的一条（与最新导入一致）。
2. 若定价表已取消 warehouse_sku 列（与 import_product_pricing 一致）：改用 product_sku_mapping
   的 warehouse_sku → product_sku，且仅保留在 product_sku_pricing 中存在的 product_sku（与定价主数据对齐）。

与利润表一律按 TRIM(warehouse_sku) 关联。

默认仅填充「当前 product_sku 为空或仅空白」的行；加 --overwrite 则凡能命中映射即覆盖。

用法：
    python python/v2/profit/order_sku_profit_step00.py
    python python/v2/profit/order_sku_profit_step00.py --date-from 2026-05-01 --date-to 2026-05-08
    python python/v2/profit/order_sku_profit_step00.py -o M2605506023397
    python python/v2/profit/order_sku_profit_step00.py --overwrite
    python python/v2/profit/order_sku_profit_step00.py --dry-run
"""

import argparse
import sys
from datetime import date
from pathlib import Path
from typing import Any

_PROFIT_DIR = Path(__file__).resolve().parent
_V2_DIR = _PROFIT_DIR.parent
_WR_DIR = _V2_DIR / "warehouse-rent"
_ORDERS_DIR = _V2_DIR / "orders"
for _p in (_PROFIT_DIR, _V2_DIR, _WR_DIR, _ORDERS_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from db import connect, load_db_config
from logger import get_logger, setup_stdout_utf8  # type: ignore[import]
from order_sku_profit_constants import TABLE, build_date_filter

_LOG = get_logger("PROFIT-STEP00")

PRICING_TABLE = "product_sku_pricing"
MAPPING_TABLE = "product_sku_mapping"


def _order_nos_from_arg(order_arg: str | None) -> list[str]:
    if not order_arg:
        return []
    s = str(order_arg).strip()
    if not s:
        return []
    parts = [p.strip() for p in s.replace("，", ",").split(",")]
    return [p for p in parts if p]


def _profit_has_product_sku_column(conn) -> bool:
    cur = conn.cursor()
    try:
        cur.execute(
            """
            SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_SCHEMA = DATABASE()
              AND TABLE_NAME = %s
              AND COLUMN_NAME = 'product_sku'
            """,
            (TABLE,),
        )
        row = cur.fetchone()
        return int(row[0]) > 0 if row and row[0] is not None else False
    finally:
        cur.close()


def _psp_has_warehouse_sku_column(conn) -> bool:
    cur = conn.cursor()
    try:
        cur.execute(
            """
            SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_SCHEMA = DATABASE()
              AND TABLE_NAME = %s
              AND COLUMN_NAME = 'warehouse_sku'
            """,
            (PRICING_TABLE,),
        )
        row = cur.fetchone()
        return int(row[0]) > 0 if row and row[0] is not None else False
    finally:
        cur.close()


def _join_map_from_pricing_warehouse() -> str:
    """定价表含 warehouse_sku：ws_norm → psku。"""
    return f"""
            SELECT TRIM(psp.`warehouse_sku`) AS ws_norm, psp.`product_sku` AS psku
            FROM `{PRICING_TABLE}` psp
            INNER JOIN (
                SELECT TRIM(`warehouse_sku`) AS ws_norm, MAX(`id`) AS mid
                FROM `{PRICING_TABLE}`
                WHERE `warehouse_sku` IS NOT NULL AND TRIM(`warehouse_sku`) <> ''
                GROUP BY TRIM(`warehouse_sku`)
            ) pick ON TRIM(psp.`warehouse_sku`) = pick.ws_norm AND psp.`id` = pick.mid
    """


def _join_map_from_mapping_with_pricing() -> str:
    """定价表无 warehouse_sku：映射表 + 定价表存在性。"""
    return f"""
            SELECT TRIM(psm.`warehouse_sku`) AS ws_norm, psm.`product_sku` AS psku
            FROM `{MAPPING_TABLE}` psm
            INNER JOIN (
                SELECT TRIM(`warehouse_sku`) AS ws_norm, MAX(`id`) AS mid
                FROM `{MAPPING_TABLE}`
                WHERE TRIM(`warehouse_sku`) <> ''
                GROUP BY TRIM(`warehouse_sku`)
            ) pick ON TRIM(psm.`warehouse_sku`) = pick.ws_norm AND psm.`id` = pick.mid
            INNER JOIN `{PRICING_TABLE}` psp ON psm.`product_sku` = psp.`product_sku`
    """


def _scope_parts(
    date_from: date | None,
    date_to: date | None,
    order_no: str | None,
) -> tuple[str, str, list[Any]]:
    date_clause, params = build_date_filter("p", date_from, date_to)
    order_filter = ""
    nos = _order_nos_from_arg(order_no)
    if len(nos) == 1:
        order_filter = "AND p.`order_no` = %s"
        params = [*params, nos[0]]
    elif len(nos) > 1:
        ph = ", ".join(["%s"] * len(nos))
        order_filter = f"AND p.`order_no` IN ({ph})"
        params = [*params, *nos]
    return date_clause, order_filter, params


def _only_missing_sql(overwrite: bool) -> str:
    if overwrite:
        return ""
    return "AND (p.`product_sku` IS NULL OR TRIM(p.`product_sku`) = '')"


def count_candidates(
    conn,
    *,
    date_from: date | None,
    date_to: date | None,
    order_no: str | None,
    overwrite: bool,
    use_psp_wh: bool,
) -> int:
    date_clause, order_filter, params = _scope_parts(date_from, date_to, order_no)
    miss = _only_missing_sql(overwrite)
    inner = _join_map_from_pricing_warehouse() if use_psp_wh else _join_map_from_mapping_with_pricing()
    sql = f"""
        SELECT COUNT(*)
        FROM `{TABLE}` p
        INNER JOIN ({inner}) m ON TRIM(p.`warehouse_sku`) = m.ws_norm
        WHERE 1=1
          {miss}
          {date_clause}
          {order_filter}
    """
    cur = conn.cursor()
    try:
        cur.execute(sql, params)
        row = cur.fetchone()
        return int(row[0]) if row and row[0] is not None else 0
    finally:
        cur.close()


def run_step00(
    conn,
    *,
    date_from: date | None,
    date_to: date | None,
    order_no: str | None,
    overwrite: bool,
    dry_run: bool,
) -> dict[str, int]:
    if not _profit_has_product_sku_column(conn):
        raise RuntimeError(
            f"表 `{TABLE}` 缺少列 `product_sku`：请先执行迁移（见 docs/database/025_sales_order_sku_profit.sql 或 ALTER TABLE 增加该列）"
        )
    use_psp_wh = _psp_has_warehouse_sku_column(conn)
    mode = (
        f"`{PRICING_TABLE}.warehouse_sku`"
        if use_psp_wh
        else f"`{MAPPING_TABLE}.warehouse_sku`→product_sku 且存在于 `{PRICING_TABLE}`"
    )
    _LOG.info(f"step00：映射来源 {mode}")

    date_clause, order_filter, params = _scope_parts(date_from, date_to, order_no)
    miss = _only_missing_sql(overwrite)
    inner = _join_map_from_pricing_warehouse() if use_psp_wh else _join_map_from_mapping_with_pricing()
    n_before = count_candidates(
        conn,
        date_from=date_from,
        date_to=date_to,
        order_no=order_no,
        overwrite=overwrite,
        use_psp_wh=use_psp_wh,
    )
    _LOG.info(
        f"step00：→ `{TABLE}.product_sku`；候选行数={n_before}；overwrite={overwrite}"
    )
    if n_before == 0:
        _LOG.warn(
            "无候选行（日期/订单范围、映射为空，或默认模式下 product_sku 均已非空且未命中）"
        )
        return {"updated": 0, "candidates": 0}

    if dry_run:
        _LOG.warn("--dry-run：不写库")
        return {"updated": 0, "candidates": n_before}

    sql_upd = f"""
        UPDATE `{TABLE}` p
        INNER JOIN ({inner}) m ON TRIM(p.`warehouse_sku`) = m.ws_norm
        SET p.`product_sku` = m.psku
        WHERE 1=1
          {miss}
          {date_clause}
          {order_filter}
    """
    cur = conn.cursor()
    try:
        cur.execute(sql_upd, params)
        n = int(cur.rowcount or 0)
    finally:
        cur.close()
    conn.commit()
    _LOG.info(f"step00：已提交 UPDATE，rowcount={n}")
    return {"updated": n, "candidates": n_before}


def main() -> int:
    setup_stdout_utf8()
    ap = argparse.ArgumentParser(
        description=(
            f"step00：按仓库 SKU 映射回填 `{TABLE}.product_sku`（优先 `{PRICING_TABLE}.warehouse_sku`；"
            f"若无该列则用 `{MAPPING_TABLE}` + 定价表校验）"
        )
    )
    ap.add_argument("--date-from", type=date.fromisoformat, default=None, metavar="YYYY-MM-DD")
    ap.add_argument("--date-to", type=date.fromisoformat, default=None, metavar="YYYY-MM-DD")
    ap.add_argument(
        "-o",
        "--order-no",
        "--order-number",
        dest="order_no",
        default=None,
        metavar="ORDER",
        help="只处理指定订单号；多个单号用逗号分隔",
    )
    ap.add_argument(
        "--overwrite",
        action="store_true",
        help="凡 warehouse_sku 能命中定价映射即覆盖 product_sku（默认仅填空白）",
    )
    ap.add_argument("--dry-run", action="store_true", help="不写库，仅统计候选行数")
    args = ap.parse_args()

    _LOG.info("=" * 60)
    _LOG.info("step00：warehouse_sku → product_sku（见首条 INFO 映射来源）")

    cfg = load_db_config()
    _LOG.info(f"连接：{cfg.host}:{cfg.port}  db={cfg.database}")
    conn = connect(cfg)
    try:
        stats = run_step00(
            conn,
            date_from=args.date_from,
            date_to=args.date_to,
            order_no=args.order_no,
            overwrite=args.overwrite,
            dry_run=args.dry_run,
        )
        _LOG.info(f"完成：{stats}")
        return 0
    except RuntimeError as e:
        conn.rollback()
        _LOG.error(str(e))
        return 1
    except Exception:
        conn.rollback()
        _LOG.error("发生异常，已回滚事务")
        raise
    finally:
        conn.close()
        _LOG.info("数据库连接已关闭")


if __name__ == "__main__":
    raise SystemExit(main())
