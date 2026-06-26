from __future__ import annotations

"""
按 sales_order_shipped.line_hash 一行一单写入 sales_order_sku_profit（与发货明细一一对应）。

订单退款（sales_order_refund）不参与毛利：refund_qty / refund_amount_base 恒为 0，
net_profit_base 与 gross_profit_base 一致（同发货明细行毛利率）。

实现拆分为：
  - order_sku_profit_constants.py — 表名、列清单、日期 SQL 片段
  - order_sku_profit_steps.py     — 各步骤：读发货、组装行、UPSERT、回写标记
  - run_order_sku_profit.py       — run_order_sku_profit() 总控编排 + CLI main()

用法：
  python run_order_sku_profit.py
  python run_order_sku_profit.py --date-from 2026-05-01 --date-to 2026-05-08
  python run_order_sku_profit.py --order-no M2605506023397
  python run_order_sku_profit.py --date-from 2026-05-01 --dry-run
  python run_order_sku_profit.py --mark-shipped --profit-node batch20260511
  python run_order_sku_profit.py --only-unmarked --mark-shipped

sales_order_shipped.profit_calc_node（可选）：
  成功写入利润表后，--mark-shipped 按 line_hash 精确更新对应发货行；同时 sales_order_sku_profit.calc_node 写入同一节点值。
  未 --mark-shipped 时 calc_node 取自发货行当前 profit_calc_node。
  --only-unmarked 仅处理 profit_calc_node 为空的明细行。
"""

import argparse
import sys
from datetime import date
from pathlib import Path

_PROFIT_DIR = Path(__file__).resolve().parent
_V2_DIR = _PROFIT_DIR.parent
_WR_DIR = _V2_DIR / "warehouse-rent"
_ORDERS_DIR = _V2_DIR / "orders"
for _p in (_PROFIT_DIR, _V2_DIR, _WR_DIR, _ORDERS_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from order_sku_profit_constants import TABLE
from order_sku_profit_steps import (
    default_profit_node,
    step_build_profit_rows,
    step_fetch_shop_name_en_skip_profit,
    step_fetch_shipped_lines,
    step_mark_shipped_profit_node,
    step_upsert_profit_rows,
)
from db import connect, load_db_config
from logger import get_logger, setup_stdout_utf8  # type: ignore[import]

_LOG = get_logger("ORDER-SKU-PROFIT")


def run_order_sku_profit(
    conn,
    date_from: date | None,
    date_to: date | None,
    order_no: str | None,
    dry_run: bool = False,
    *,
    only_unmarked: bool = False,
    mark_shipped: bool = False,
    profit_node: str | None = None,
) -> int:
    """
    总控：读发货 → 组装利润行 →（可选）UPSERT →（可选）回写 profit_calc_node。
    返回 UPSERT 行数；dry_run 时为 0 且不写库、不 commit。
    """
    if only_unmarked:
        _LOG.warn("已启用 --only-unmarked：仅处理 profit_calc_node 为空的发货明细行。")

    lines = step_fetch_shipped_lines(
        conn, date_from, date_to, order_no, only_unmarked=only_unmarked
    )
    if not lines:
        _LOG.warn("sales_order_shipped 无有效行，任务结束")
        return 0

    skip_shops = step_fetch_shop_name_en_skip_profit(conn)
    profit_rows = step_build_profit_rows(
        lines, skip_shop_name_en_if_disabled=skip_shops
    )
    if mark_shipped and not dry_run:
        node = default_profit_node(profit_node)
        for r in profit_rows:
            r["calc_node"] = node
    _LOG.info(f"合并摘要：发货利润行={len(profit_rows)}（不含 sales_order_refund）")

    if dry_run:
        _LOG.warn("--dry-run：不写库，预览前 3 个 line_hash")
        for r in profit_rows[:3]:
            _LOG.info(
                f"  line_hash={r['line_hash'][:16]}… order={r['order_no']} "
                f"gross={r['gross_profit_base']} refund={r['refund_amount_base']} net={r['net_profit_base']}"
            )
        return 0

    n = step_upsert_profit_rows(conn, profit_rows)
    if mark_shipped:
        hashes = [str(r["line_hash"]) for r in profit_rows if r.get("line_hash")]
        step_mark_shipped_profit_node(
            conn, hashes, default_profit_node(profit_node), date_from, date_to, order_no
        )
    conn.commit()
    _LOG.info(f"事务已提交：{TABLE} UPSERT {n} 行")
    return n


# 兼容旧名：外部若 import run 仍可用
run = run_order_sku_profit


def main() -> int:
    setup_stdout_utf8()
    ap = argparse.ArgumentParser(
        description="按 shipped.line_hash 写入 sales_order_sku_profit（订单退款不计入毛利）"
    )
    ap.add_argument("--date-from", type=date.fromisoformat, default=None, metavar="YYYY-MM-DD")
    ap.add_argument("--date-to", type=date.fromisoformat, default=None, metavar="YYYY-MM-DD")
    ap.add_argument("--order-no", default=None, help="只处理指定订单号")
    ap.add_argument("--dry-run", action="store_true", help="不写库")
    ap.add_argument("--only-unmarked", action="store_true", help="仅 profit_calc_node 为空的明细")
    ap.add_argument("--mark-shipped", action="store_true", help="写利润后按 line_hash 回写 profit_calc_node")
    ap.add_argument("--profit-node", default=None, metavar="STR", help="profit_calc_node 标记，最长 24")
    args = ap.parse_args()

    _LOG.info("=" * 60)
    _LOG.info(f"任务：{TABLE}（line_hash 行级）→ run_order_sku_profit 总控")
    cfg = load_db_config()
    _LOG.info(f"连接：{cfg.host}:{cfg.port} db={cfg.database}")
    conn = connect(cfg)
    try:
        run_order_sku_profit(
            conn,
            date_from=args.date_from,
            date_to=args.date_to,
            order_no=args.order_no,
            dry_run=args.dry_run,
            only_unmarked=args.only_unmarked,
            mark_shipped=args.mark_shipped and not args.dry_run,
            profit_node=args.profit_node,
        )
        return 0
    except Exception:
        conn.rollback()
        _LOG.error("已回滚")
        raise
    finally:
        conn.close()
        _LOG.info("数据库连接已关闭")


if __name__ == "__main__":
    raise SystemExit(main())
