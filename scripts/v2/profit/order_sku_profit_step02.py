from __future__ import annotations

"""
step02 — 仅计算 VAT 费用，写入 sales_order_sku_profit.vat_fee_pay / vat_fee_base。

税率来源：platform_shop_config.vat_rate（小数，如 0.19 表示 19%）。
匹配键：与 import_order_shipped 写入店铺配置一致 —— TRIM 后的 platform、platform_site、shop_name_en
三列相等即视为同一店铺；配置表多行同键时取 MAX(vat_rate)。

未命中配置的利润行：不更新其 vat_fee_*（保留库内已有值或默认 0）。
tax_pay / tax_base 仍表示发货明细中的「税费」等，不由本脚本改写。

公式：vat_fee_pay = order_total_pay × vat_rate，vat_fee_base = order_total_base × vat_rate（结果量化到 6 位小数）。
写库时同时将 `calc_node` 置为 `step02_vat`（须在表字段长度限制内，当前为 VARCHAR(24)）。

用法：
    python profit/order_sku_profit_step02.py
    python profit/order_sku_profit_step02.py --date-from 2026-05-01 --date-to 2026-05-08
    python profit/order_sku_profit_step02.py --dry-run
"""

import argparse
import sys
from datetime import date
from decimal import Decimal, InvalidOperation
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
from order_sku_profit_constants import TABLE
from order_sku_profit_step01 import (  # type: ignore[import]
    _executemany_chunked,
    _scope_parts,
    _to_decimal_unit,
)

_LOG = get_logger("PROFIT-STEP02")

_Q6 = Decimal("0.000001")
_D0 = Decimal("0").quantize(_Q6)
_PLATFORM_SHOP_TABLE = "platform_shop_config"
CALC_NODE_STEP02_VAT = "step02_vat"


def _dec(v: Any) -> Decimal:
    d = _to_decimal_unit(v)
    return d if d is not None else _D0


def _dec_rate(v: Any) -> Decimal | None:
    if v is None:
        return None
    try:
        d = v if isinstance(v, Decimal) else Decimal(str(v))
    except (InvalidOperation, ValueError, TypeError):
        return None
    return d.quantize(_Q6)


def _fetch_profit_vat_updates(
    conn,
    date_from: date | None,
    date_to: date | None,
    order_no: str | None,
) -> list[tuple[Decimal, Decimal, str, int]]:
    """
    返回 [(vat_fee_pay, vat_fee_base, calc_node, id), ...]：仅含能在 platform_shop_config 命中税率的利润行。
    """
    date_clause, order_filter, scope_params = _scope_parts(date_from, date_to, order_no)
    sql = f"""
        SELECT p.`id`,
               p.`order_total_pay`,
               p.`order_total_base`,
               psc.`vat_rate` AS cfg_vat_rate
        FROM `{TABLE}` p
        INNER JOIN (
            SELECT TRIM(`platform`) AS k_platform,
                   TRIM(`platform_site`) AS k_site,
                   TRIM(`shop_name_en`) AS k_shop,
                   MAX(`vat_rate`) AS vat_rate
            FROM `{_PLATFORM_SHOP_TABLE}`
            GROUP BY TRIM(`platform`), TRIM(`platform_site`), TRIM(`shop_name_en`)
        ) psc
          ON TRIM(COALESCE(p.`platform`, '')) = psc.k_platform
         AND TRIM(COALESCE(p.`platform_site`, '')) = psc.k_site
         AND TRIM(COALESCE(p.`shop_name_en`, '')) = psc.k_shop
        WHERE 1=1
          AND TRIM(COALESCE(p.`shop_name_en`, '')) <> ''
        {date_clause}
        {order_filter}
        ORDER BY p.`id`
    """
    cur = conn.cursor(dictionary=True)
    try:
        cur.execute(sql, scope_params)
        raw_rows = list(cur.fetchall())
    finally:
        cur.close()

    out: list[tuple[Decimal, Decimal, int]] = []
    for r in raw_rows:
        rate = _dec_rate(r.get("cfg_vat_rate"))
        if rate is None:
            continue
        pid = int(r["id"])
        otp = _dec(r.get("order_total_pay"))
        otb = _dec(r.get("order_total_base"))
        vat_p = (otp * rate).quantize(_Q6)
        vat_b = (otb * rate).quantize(_Q6)
        out.append((vat_p, vat_b, CALC_NODE_STEP02_VAT, pid))
    return out


def run_step02(
    conn,
    date_from: date | None,
    date_to: date | None,
    order_no: str | None,
    *,
    dry_run: bool = False,
) -> dict[str, int]:
    updates = _fetch_profit_vat_updates(conn, date_from, date_to, order_no)
    if not updates:
        _LOG.warn(
            "无待更新行：范围内无 sales_order_sku_profit，或与 platform_shop_config "
            "（platform + platform_site + shop_name_en）无交集"
        )
        return {"updated": 0, "candidates": 0}

    _LOG.info(
        f"step02（仅 VAT）：待更新行数={len(updates)}（vat_fee_* + calc_node={CALC_NODE_STEP02_VAT!r}，税率来自 {_PLATFORM_SHOP_TABLE}）"
    )
    if dry_run:
        _LOG.warn("--dry-run：不写库，预览前 3 条")
        for vat_p, vat_b, node, pid in updates[:3]:
            _LOG.info(
                f"  [dry-run] id={pid} calc_node={node!r} vat_fee_pay={vat_p} vat_fee_base={vat_b}"
            )
        return {"updated": 0, "candidates": len(updates)}

    sql_upd = (
        f"UPDATE `{TABLE}` SET `vat_fee_pay` = %s, `vat_fee_base` = %s, `calc_node` = %s "
        f"WHERE `id` = %s"
    )
    n = _executemany_chunked(conn, sql_upd, updates)
    conn.commit()
    _LOG.info(f"step02：已提交 UPDATE 批次数累计行={n}")
    return {"updated": n, "candidates": len(updates)}


def main() -> int:
    setup_stdout_utf8()
    ap = argparse.ArgumentParser(
        description="step02：回写 vat_fee_* 并将 calc_node 设为 step02_vat（税率来自 platform_shop_config）"
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
    ap.add_argument("--dry-run", action="store_true", help="不写库")
    args = ap.parse_args()

    _LOG.info("=" * 60)
    _LOG.info("step02：VAT 费用 + calc_node=step02_vat ← platform_shop_config.vat_rate")

    cfg = load_db_config()
    _LOG.info(f"连接：{cfg.host}:{cfg.port}  db={cfg.database}")
    conn = connect(cfg)
    try:
        stats = run_step02(
            conn,
            date_from=args.date_from,
            date_to=args.date_to,
            order_no=args.order_no,
            dry_run=args.dry_run,
        )
        _LOG.info(f"完成：{stats}")
        return 0
    except Exception:
        conn.rollback()
        _LOG.error("发生异常，已回滚事务")
        raise
    finally:
        conn.close()
        _LOG.info("数据库连接已关闭")


if __name__ == "__main__":
    raise SystemExit(main())
