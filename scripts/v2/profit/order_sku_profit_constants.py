from __future__ import annotations

"""
订单 SKU 利润流水线常量（表名、列清单、SQL 片段生成）。

可单独执行：
    # 仅自检 / SQL 预览（不连库，仅标准库）
    python python/v2/profit/order_sku_profit_constants.py
    python python/v2/profit/order_sku_profit_constants.py --print-select

    # 从 sales_order_shipped 写入 sales_order_sku_profit（calc_node=init），并把 shipped.profit_calc_node 标为 init
    # （需 mysql-connector、.env 数据库配置，与 run_order_sku_profit 相同依赖）
    python python/v2/profit/order_sku_profit_constants.py --init-sync
    python python/v2/profit/order_sku_profit_constants.py --init-sync --date-from 2026-05-01 --date-to 2026-05-08
    python python/v2/profit/order_sku_profit_constants.py --init-sync --dry-run
"""

# 单独执行 --init-sync 时回写 sales_order_shipped.profit_calc_node 的标记值
PROFIT_CALC_NODE_INIT = "init"

from datetime import date, datetime
from decimal import Decimal
from typing import Any

TABLE = "sales_order_sku_profit"
SHIPPED_TABLE = "sales_order_shipped"

D0 = Decimal("0")

# 与 sales_order_sku_profit.distribution_lev 一致：0=自营，1=分销（见表注释）
_DISTRIBUTION_WH_MARK = "分销"


def distribution_lev_from_warehouse_name(warehouse_name: Any) -> int:
    """warehouse_name 含「分销」时视为分销渠道，返回 1；否则 0。"""
    if warehouse_name is None:
        return 0
    name = str(warehouse_name).strip()
    if not name:
        return 0
    return 1 if _DISTRIBUTION_WH_MARK in name else 0


UK_COLS: frozenset[str] = frozenset({"line_hash"})

SUM_PAY_COLS: tuple[str, ...] = (
    "order_total_pay",
    "order_goods_pay",
    "platform_shipping_pay",
    "payment_fee_pay",
    "platform_fee_pay",
    "fba_fee_pay",
    "platform_subsidy_pay",
    "tax_pay",
    "other_fee_pay",
    "purchase_cost_pay",
    "purchase_shipping_pay",
    "purchase_tax_pay",
    "first_leg_shipping_pay",
    "first_leg_tax_pay",
    "packaging_fee_pay",
    "delivery_shipping_pay",
)

SUM_BASE_COLS: tuple[str, ...] = (
    "order_total_base",
    "order_goods_base",
    "platform_shipping_base",
    "payment_fee_base",
    "platform_fee_base",
    "fba_fee_base",
    "platform_subsidy_base",
    "tax_base",
    "other_fee_base",
    "purchase_cost_base",
    "purchase_shipping_base",
    "purchase_tax_base",
    "first_leg_shipping_base",
    "first_leg_tax_base",
    "packaging_fee_base",
    "delivery_shipping_base",
    "total_fee_base",
    "total_cost_base",
    "gross_profit_base",
)

INSERT_COLS: tuple[str, ...] = (
    "line_hash",
    "platform",
    "shop_name_en",
    "platform_site",
    "order_type",
    "ref_no",
    "order_no",
    "warehouse_sku",
    "platform_sku",
    "warehouse_name",
    "shipping_method",
    "pay_currency",
    "base_currency",
    "pay_time",
    "ship_time",
    "shipped_qty",
) + SUM_PAY_COLS + SUM_BASE_COLS + (
    "gross_margin_rate",
    "refund_qty",
    "refund_amount_base",
    "net_profit_base",
    "net_margin_rate",
    "distribution_lev",
    "calc_node",
    "source_note",
)


def build_date_filter(
    alias: str, date_from: date | None, date_to: date | None
) -> tuple[str, list[Any]]:
    col = f"`{alias}`.`pay_time`" if alias else "`pay_time`"
    parts: list[str] = []
    params: list[Any] = []
    if date_from:
        parts.append(f"{col} >= %s")
        params.append(datetime(date_from.year, date_from.month, date_from.day, 0, 0, 0))
    if date_to:
        parts.append(f"{col} <= %s")
        params.append(datetime(date_to.year, date_to.month, date_to.day, 23, 59, 59))
    clause = ("AND " + " AND ".join(parts)) if parts else ""
    return clause, params


def shipped_select_sql() -> str:
    pay_cols = ", ".join(f"`{c}`" for c in SUM_PAY_COLS)
    base_cols = ", ".join(f"`{c}`" for c in SUM_BASE_COLS)
    return (
        f"`line_hash`, `platform`, `shop_name_en`, `platform_site`, `order_type`, "
        f"COALESCE(`ref_no`,'') AS `ref_no`, `order_no`, `warehouse_sku`, "
        f"`platform_sku`, `warehouse_name`, `shipping_method`, "
        f"`pay_currency`, `base_currency`, `pay_time`, `ship_time`, "
        f"COALESCE(`warehouse_sku_qty`,0) AS `shipped_qty`, "
        f"`fx_rate_to_base`, {pay_cols}, {base_cols}, `gross_margin_rate`, "
        f"`profit_calc_node`"
    )


def _ensure_profit_runtime_path() -> None:
    """供 --init-sync 使用：将 v2/profit、v2、warehouse-rent、orders 加入 sys.path。"""
    import sys
    from pathlib import Path

    profit_dir = Path(__file__).resolve().parent
    v2_dir = profit_dir.parent
    for p in (profit_dir, v2_dir, v2_dir / "warehouse-rent", v2_dir / "orders"):
        s = str(p)
        if s not in sys.path:
            sys.path.insert(0, s)


def sync_shipped_to_profit_and_mark_init(
    *,
    date_from: date | None,
    date_to: date | None,
    order_no: str | None,
    only_unmarked: bool,
    dry_run: bool,
) -> int:
    """
    读 sales_order_shipped → 组装行 UPSERT 至 sales_order_sku_profit（calc_node 与将写入的 profit_calc_node 一致，均为 init）；
    非 dry_run 时按 line_hash 将 shipped.profit_calc_node 更新为 PROFIT_CALC_NODE_INIT。
    返回 UPSERT 行数；dry_run 时为 0。
    """
    _ensure_profit_runtime_path()
    from db import connect, load_db_config
    from logger import get_logger, setup_stdout_utf8  # type: ignore[import]
    from order_sku_profit_steps import (
        step_build_profit_rows,
        step_fetch_shop_name_en_skip_profit,
        step_fetch_shipped_lines,
        step_mark_shipped_profit_node,
        step_upsert_profit_rows,
    )

    setup_stdout_utf8()
    log = get_logger("ORDER-SKU-PROFIT-INIT")
    node = (PROFIT_CALC_NODE_INIT or "init")[:24]

    cfg = load_db_config()
    log.info(f"[init-sync] 连接 {cfg.host}:{cfg.port} db={cfg.database}")
    conn = None
    try:
        conn = connect(cfg)
        lines = step_fetch_shipped_lines(
            conn, date_from, date_to, order_no, only_unmarked=only_unmarked
        )
        if not lines:
            log.warn("[init-sync] sales_order_shipped 无有效行，结束")
            return 0

        skip_shops = step_fetch_shop_name_en_skip_profit(conn)
        profit_rows = step_build_profit_rows(
            lines, skip_shop_name_en_if_disabled=skip_shops
        )
        for r in profit_rows:
            r["calc_node"] = node
        log.info(f"[init-sync] 组装利润行 {len(profit_rows)} 条 → {TABLE}（calc_node={node!r}）")

        if dry_run:
            log.warn("[init-sync] --dry-run：不写库、不回写 profit_calc_node")
            for r in profit_rows[:3]:
                log.info(
                    f"  预览 line_hash={str(r.get('line_hash'))[:16]}… "
                    f"order={r.get('order_no')} net={r.get('net_profit_base')}"
                )
            return 0

        n = step_upsert_profit_rows(conn, profit_rows)
        hashes = [str(r["line_hash"]) for r in profit_rows if r.get("line_hash")]
        step_mark_shipped_profit_node(
            conn, hashes, node, date_from, date_to, order_no
        )
        conn.commit()
        log.info(f"[init-sync] 已提交：UPSERT {TABLE} {n} 行；profit_calc_node={node!r}")
        return n
    except Exception:
        if conn is not None:
            conn.rollback()
        log.error("[init-sync] 已回滚")
        raise
    finally:
        if conn is not None:
            conn.close()
            log.info("[init-sync] 数据库连接已关闭")


def validate_insert_cols() -> list[str]:
    """若 INSERT_COLS 存在重复列名，返回重复项列表；否则返回空列表。"""
    seen: set[str] = set()
    dup: set[str] = set()
    for c in INSERT_COLS:
        if c in seen:
            dup.add(c)
        seen.add(c)
    return sorted(dup)


def main() -> int:
    """CLI：自检列定义并打印摘要；--print-select 输出发货表 SELECT 列片段。"""
    import argparse
    import sys

    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    ap = argparse.ArgumentParser(
        description="order_sku_profit 常量：自检 / SQL 预览；--init-sync 从发货表写利润表并标记 profit_calc_node=init"
    )
    ap.add_argument(
        "--init-sync",
        action="store_true",
        help=f"从 {SHIPPED_TABLE} 写入 {TABLE}，并将 shipped 行的 profit_calc_node 设为 {PROFIT_CALC_NODE_INIT!r}",
    )
    ap.add_argument("--date-from", type=date.fromisoformat, default=None, metavar="YYYY-MM-DD")
    ap.add_argument("--date-to", type=date.fromisoformat, default=None, metavar="YYYY-MM-DD")
    ap.add_argument("--order-no", default=None, help="只处理指定订单号")
    ap.add_argument("--only-unmarked", action="store_true", help="仅 profit_calc_node 为空的发货明细")
    ap.add_argument("--dry-run", action="store_true", help="与 --init-sync 合用：不写库、不回写标记")
    ap.add_argument(
        "--print-select",
        action="store_true",
        help="仅打印 shipped_select_sql() 生成的 SELECT 列清单（单行）",
    )
    ap.add_argument(
        "--list-insert-cols",
        action="store_true",
        help="逐行打印 INSERT_COLS（与 UPSERT 列顺序一致）",
    )
    args = ap.parse_args()

    dupes = validate_insert_cols()
    if dupes:
        print(f"错误：INSERT_COLS 存在重复列：{dupes}", file=sys.stderr)
        return 1

    if args.init_sync:
        return sync_shipped_to_profit_and_mark_init(
            date_from=args.date_from,
            date_to=args.date_to,
            order_no=args.order_no,
            only_unmarked=args.only_unmarked,
            dry_run=args.dry_run,
        )

    if args.print_select:
        print(shipped_select_sql())
        return 0

    if args.list_insert_cols:
        for c in INSERT_COLS:
            print(c)
        return 0

    clause, params = build_date_filter("", date(2026, 1, 1), date(2026, 1, 31))
    print("order_sku_profit_constants（自检通过）")
    print(f"  TABLE={TABLE!r}  SHIPPED_TABLE={SHIPPED_TABLE!r}")
    print(f"  SUM_PAY_COLS={len(SUM_PAY_COLS)}  SUM_BASE_COLS={len(SUM_BASE_COLS)}  INSERT_COLS={len(INSERT_COLS)}")
    print(f"  UK_COLS={sorted(UK_COLS)!r}")
    print(f"  build_date_filter 示例（2026-01-01~2026-01-31）：{clause!r}  占位参数个数={len(params)}")
    print("  更多：--init-sync（写库+标记 init）| --print-select | --list-insert-cols | --help")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
