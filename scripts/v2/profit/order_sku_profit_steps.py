from __future__ import annotations

"""
从 sales_order_shipped 读取、组装 sales_order_sku_profit 行并 UPSERT，可选回写 profit_calc_node。

platform_shop_config.shop_status=0 的店铺：若发货行 shop_name_en（TRIM）与该配置表中的 shop_name_en 一致，
则不写入 sales_order_sku_profit（也不进入 mark_shipped 的 line_hash 列表）。

供 run_order_sku_profit.py、order_sku_profit_constants.py --init-sync 使用。
依赖 sys.path 已包含 python/v2/orders（以便 from excel_common import upsert_rows）。
"""

from datetime import date, datetime
from decimal import Decimal
from typing import Any

import mysql.connector

from order_sku_profit_constants import (
    D0,
    INSERT_COLS,
    SHIPPED_TABLE,
    SUM_BASE_COLS,
    SUM_PAY_COLS,
    TABLE,
    build_date_filter,
    distribution_lev_from_warehouse_name,
    shipped_select_sql,
)
from excel_common import upsert_rows
from logger import get_logger

_LOG = get_logger("ORDER-SKU-PROFIT-STEPS")

_SOURCE_NOTE = "line_level:sales_order_shipped_no_refund"
_PLATFORM_SHOP_TABLE = "platform_shop_config"


def step_fetch_shop_name_en_skip_profit(conn: mysql.connector.MySQLConnection) -> frozenset[str]:
    """
    读取 platform_shop_config 中 shop_status=0 的店铺英文名（TRIM 后、非空去重）。
    发货/组装阶段若 shop_name_en 命中该集合，则不写入 sales_order_sku_profit。
    """
    sql = (
        f"SELECT DISTINCT TRIM(COALESCE(`shop_name_en`, '')) AS k "
        f"FROM `{_PLATFORM_SHOP_TABLE}` "
        "WHERE `shop_status` = 0 "
        "AND TRIM(COALESCE(`shop_name_en`, '')) <> ''"
    )
    cur = conn.cursor()
    try:
        cur.execute(sql)
        rows = cur.fetchall() or []
    finally:
        cur.close()
    out = frozenset(str(r[0]).strip() for r in rows if r and r[0] is not None and str(r[0]).strip())
    _LOG.info(
        f"读 `{_PLATFORM_SHOP_TABLE}`：shop_status=0 需跳过利润的 shop_name_en 共 {len(out)} 个"
    )
    return out


def default_profit_node(profit_node: str | None) -> str:
    """未指定时生成 batch 前缀时间戳；最长 24 字符（与表字段一致）。"""
    if profit_node and str(profit_node).strip():
        s = str(profit_node).strip()
    else:
        s = "batch" + datetime.now().strftime("%Y%m%d%H%M")
    return s[:24]


def _dec_or_zero(v: Any) -> Decimal:
    if v is None:
        return D0
    if isinstance(v, Decimal):
        return v
    try:
        return Decimal(str(v).strip().replace(",", ""))
    except Exception:
        return D0


def _maybe_decimal(v: Any) -> Decimal | None:
    if v is None:
        return None
    if isinstance(v, Decimal):
        return v
    try:
        return Decimal(str(v).strip().replace(",", ""))
    except Exception:
        return None


def step_fetch_shipped_lines(
    conn: mysql.connector.MySQLConnection,
    date_from: date | None,
    date_to: date | None,
    order_no: str | None,
    *,
    only_unmarked: bool = False,
) -> list[dict[str, Any]]:
    sel = shipped_select_sql()
    clause, params = build_date_filter("", date_from, date_to)
    sql = f"SELECT {sel} FROM `{SHIPPED_TABLE}` WHERE 1=1 {clause}"
    if only_unmarked:
        sql += " AND (`profit_calc_node` IS NULL OR TRIM(`profit_calc_node`) = '')"
    if order_no and str(order_no).strip():
        sql += " AND `order_no` = %s"
        params = list(params) + [str(order_no).strip()]
    cur = conn.cursor(dictionary=True)
    try:
        cur.execute(sql, params)
        rows = cur.fetchall() or []
    finally:
        cur.close()
    _LOG.info(f"读 {SHIPPED_TABLE}：{len(rows)} 行（only_unmarked={only_unmarked}）")
    return rows


def step_build_profit_rows(
    lines: list[dict[str, Any]],
    *,
    skip_shop_name_en_if_disabled: frozenset[str] | None = None,
) -> list[dict[str, Any]]:
    """
    skip_shop_name_en_if_disabled：与 platform_shop_config.shop_name_en（TRIM）一致且该配置 shop_status=0 时，
    不生成 sales_order_sku_profit 行（也不参与后续 mark_shipped 的 hashes）。
    """
    out: list[dict[str, Any]] = []
    skip_set = skip_shop_name_en_if_disabled or frozenset()
    skipped = 0
    for line in lines:
        if skip_set:
            sn = str(line.get("shop_name_en") or "").strip()
            if sn and sn in skip_set:
                skipped += 1
                continue
        wh_name = line.get("warehouse_name")
        dist = distribution_lev_from_warehouse_name(wh_name)
        gross_base = _dec_or_zero(line.get("gross_profit_base"))
        gm_rate = _maybe_decimal(line.get("gross_margin_rate"))
        pnode = line.get("profit_calc_node")
        if isinstance(pnode, str):
            pnode = pnode.strip() or None
        elif pnode is not None:
            pnode = str(pnode).strip() or None

        row: dict[str, Any] = {
            "line_hash": line.get("line_hash"),
            "platform": line.get("platform") or "",
            "shop_name_en": line.get("shop_name_en"),
            "platform_site": line.get("platform_site"),
            "order_type": line.get("order_type"),
            "ref_no": line.get("ref_no") or "",
            "order_no": line.get("order_no") or "",
            "warehouse_sku": line.get("warehouse_sku") or "",
            "platform_sku": line.get("platform_sku"),
            "warehouse_name": wh_name,
            "shipping_method": line.get("shipping_method"),
            "pay_currency": line.get("pay_currency"),
            "base_currency": line.get("base_currency"),
            "pay_time": line.get("pay_time"),
            "ship_time": line.get("ship_time"),
            "shipped_qty": int(line.get("shipped_qty") or 0),
            "gross_margin_rate": gm_rate,
            "refund_qty": 0,
            "refund_amount_base": D0,
            "net_profit_base": gross_base,
            "net_margin_rate": gm_rate,
            "distribution_lev": dist,
            "calc_node": pnode,
            "source_note": _SOURCE_NOTE,
        }
        for c in SUM_PAY_COLS:
            row[c] = _dec_or_zero(line.get(c))
        for c in SUM_BASE_COLS:
            row[c] = _dec_or_zero(line.get(c))

        lh = row.get("line_hash")
        if not lh:
            _LOG.warn("跳过无 line_hash 行")
            continue
        out.append(row)
    if skipped:
        _LOG.info(
            f"因 platform_shop_config.shop_status=0 且 shop_name_en 匹配，跳过 {skipped} 条发货行（不写 {TABLE}）"
        )
    return out


def step_upsert_profit_rows(
    conn: mysql.connector.MySQLConnection, profit_rows: list[dict[str, Any]]
) -> int:
    if not profit_rows:
        return 0
    tuples: list[tuple[Any, ...]] = []
    for r in profit_rows:
        tuples.append(tuple(r[c] for c in INSERT_COLS))
    n = upsert_rows(conn, table=TABLE, columns=INSERT_COLS, rows=tuples)
    _LOG.info(f"UPSERT `{TABLE}`：executemany 累计 {n} 行")
    return n


def step_mark_shipped_profit_node(
    conn: mysql.connector.MySQLConnection,
    hashes: list[str],
    node: str,
    date_from: date | None,
    date_to: date | None,
    order_no: str | None,
) -> None:
    """
    按 line_hash 回写 sales_order_shipped.profit_calc_node。
    date_from / date_to / order_no 保留与调用方一致，便于以后加范围校验；当前仅按 hashes 更新。
    """
    _ = (date_from, date_to, order_no)
    if not hashes:
        return
    node_s = (node or "")[:24]
    cur = conn.cursor()
    chunk = 400
    try:
        for i in range(0, len(hashes), chunk):
            part = hashes[i : i + chunk]
            placeholders = ", ".join(["%s"] * len(part))
            sql = (
                f"UPDATE `{SHIPPED_TABLE}` SET `profit_calc_node` = %s "
                f"WHERE `line_hash` IN ({placeholders})"
            )
            cur.execute(sql, [node_s, *part])
    finally:
        cur.close()
    _LOG.info(f"已回写 {SHIPPED_TABLE}.profit_calc_node={node_s!r}，行数={len(hashes)}")
