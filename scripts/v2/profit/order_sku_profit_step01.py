from __future__ import annotations

"""
step01 — 补填头程运费/税费：付款额 + 按汇率折算本位币（first_leg_shipping_* / first_leg_tax_*）
(仅计算 distribution_lev = 0 的自营数据，分销订单无需计算头程、关税)

1. 将 distribution_lev = 1 的分销行 `calc_node` 设为 "step01_done"（跳过后续步骤）。
    1.1 读取 distribution_lev = 0 且 first_leg_shipping_base 与 first_leg_tax_base 均已非 0 且非 NULL 的行，
   将 calc_node 设为 "step01_done"。
   1.2（仅 --calc-mode strict）读取 distribution_lev = 0 且 first_leg_shipping_pay 与 first_leg_tax_pay 均已非 0 且非 NULL 的行，
   使用 python/v2/config/fx_rates.json（rmb_per_eur、rates_to_eur）得到「1 单位 pay_currency = ? EUR」，
   更新 first_leg_shipping_base 与 first_leg_tax_base（base = pay × 该汇率），将 calc_node 设为 "step01_done"。
   --calc-mode loose 时跳过 1.2，不根据 pay 覆盖已有 base。

2. 读取 distribution_lev = 0 且 first_leg_shipping_base 为 0 或 NULL 的行，
   按产品 SKU（product_sku_mapping.product_sku，无映射时回退 warehouse_sku）+ platform_site 从 product_sku_pricing 取 first_leg_*_cny * shipped_qty（人民币），
   先用 fx_rates.json 将人民币合计换成本位 EUR，再按同一文件中的 pay_currency→EUR 汇率反算付款币 pay
   （满足 base = pay × 汇率），回写 first_leg_shipping_base，calc_node 设为 "step01_shipping"。

3. 读取 distribution_lev = 0 且 first_leg_tax_base 为 0 或 NULL 的行，
   按 duty_*_cny * shipped_qty 同上（与步骤 2 相同规则），
   回写 first_leg_tax_*，并将 calc_node 设为 "step01_tax"。
   若定价行已命中但对应站点的 duty_*_cny 为空或 NULL（无法解析为金额），
   则视为无关税：first_leg_tax_pay / first_leg_tax_base 置 0；
   若此时 first_leg_shipping_base 已大于 0 且非 NULL，则 calc_node 设为 "step01_done"，
   否则仍为 "step01_tax"（待步骤 2 补全头程 base 后，步骤 3 再次执行时可收口为 done）。

4. 读取 distribution_lev = 0 且 first_leg_shipping_base 与 first_leg_tax_base 均已非 0 且非 NULL 的行，
   将 calc_node 设为 "step01_done"。

platform_site → 定价列映射（头程 / 关税含税，数据来自 product_sku_pricing）：
  定价表仅按 product_sku 关联；利润行通过 line_hash 左连 product_sku_mapping 取 product_sku，
  无映射行时回退为 warehouse_sku（与历史订单「产品 SKU 默认等于仓库 SKU」一致）。
  含 'US'  → first_leg_us_cny     / duty_us_cny
  含 'UK'  → first_leg_uk_cny     / duty_uk_cny
  含 'CA'  → first_leg_ca_cny     / duty_ca_au_cny
  含 'JP'  → first_leg_jp_cny     / duty_jp_cny
  含 'AU'  → first_leg_eu_au_cny  / duty_ca_au_cny
  其他(EU) → first_leg_eu_au_cny  / duty_eu_cny

用法：
    python python/v2/profit/order_sku_profit_step01.py
    python python/v2/profit/order_sku_profit_step01.py --through-step 0
    python python/v2/profit/order_sku_profit_step01.py --through-step 2
    python python/v2/profit/order_sku_profit_step01.py --date-from 2026-05-01 --date-to 2026-05-08
    python python/v2/profit/order_sku_profit_step01.py --order-no M2605506023397
    python python/v2/profit/order_sku_profit_step01.py -o WEC0542605050093
    python python/v2/profit/order_sku_profit_step01.py --order-number A001,B002
    python python/v2/profit/order_sku_profit_step01.py --dry-run

--through-step：0（默认）或 4 = 执行步骤 1～4 全部；1 = 仅步骤 1；2 = 步骤 1～2；3 = 步骤 1～3。

--calc-mode：
    strict（默认）  执行步骤 1.2：双 pay 已齐时按 fx_rates.json 用 pay 回写 base；
    loose           跳过 1.2，保留已有 base 与 pay 的数值关系（仍执行 1.1 与步骤 2～4）。
"""

import argparse
import sys
from collections.abc import Callable
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

_PROFIT_DIR = Path(__file__).resolve().parent
_V2_DIR = _PROFIT_DIR.parent
_FX_JSON_PATH = _V2_DIR / "config" / "fx_rates.json"
_WR_DIR = _V2_DIR / "warehouse-rent"
_ORDERS_DIR = _V2_DIR / "orders"
for _p in (_PROFIT_DIR, _V2_DIR, _WR_DIR, _ORDERS_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from config.fx_rates import FxRates, load_rates  # type: ignore[import]
from db import connect, load_db_config  # type: ignore[import]
from logger import get_logger, setup_stdout_utf8  # type: ignore[import]
from order_sku_profit_constants import TABLE, build_date_filter

_LOG = get_logger("PROFIT-STEP01")

CALC_MODE_STRICT = "strict"
CALC_MODE_LOOSE = "loose"

CALC_NODE_STEP01_DONE = "step01_done"
CALC_NODE_STEP01_SHIPPING = "step01_shipping"
CALC_NODE_STEP01_TAX = "step01_tax"
PRICING_TABLE = "product_sku_pricing"
MAPPING_TABLE = "product_sku_mapping"
_Q6 = Decimal("0.000001")
_DEC_ZERO_Q6 = Decimal("0").quantize(_Q6)

_PRICING_FETCH_COLS: tuple[str, ...] = (
    "product_sku",
    "first_leg_eu_au_cny",
    "first_leg_us_cny",
    "first_leg_ca_cny",
    "first_leg_jp_cny",
    "first_leg_uk_cny",
    "duty_eu_cny",
    "duty_us_cny",
    "duty_ca_au_cny",
    "duty_jp_cny",
    "duty_uk_cny",
)


def _order_nos_from_arg(order_arg: str | None) -> list[str]:
    """从命令行解析订单号列表（支持英文/中文逗号分隔）。"""
    if not order_arg:
        return []
    s = str(order_arg).strip()
    if not s:
        return []
    parts = [p.strip() for p in s.replace("，", ",").split(",")]
    return [p for p in parts if p]


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


def _last_phase_index(through_step: int) -> int:
    if through_step in (0, 4):
        return 3
    if through_step < 0 or through_step > 4:
        raise ValueError("through_step 必须为 0～4")
    return through_step - 1


def _need_fx_rates_json(*, last_phase: int, calc_mode: str) -> bool:
    """步骤 2～3 必用汇率；步骤 1.2 仅在 strict 且执行到阶段 0 时需要。"""
    if last_phase >= 1:
        return True
    return last_phase >= 0 and calc_mode == CALC_MODE_STRICT


def _norm_site(platform_site: str | None) -> str:
    return (platform_site or "").strip().upper()


def _site_to_first_leg_col(platform_site: str | None) -> str:
    s = _norm_site(platform_site)
    if "US" in s:
        return "first_leg_us_cny"
    if "UK" in s:
        return "first_leg_uk_cny"
    if "CA" in s:
        return "first_leg_ca_cny"
    if "JP" in s:
        return "first_leg_jp_cny"
    if "AU" in s:
        return "first_leg_eu_au_cny"
    return "first_leg_eu_au_cny"


def _site_to_duty_col(platform_site: str | None) -> str:
    s = _norm_site(platform_site)
    if "US" in s:
        return "duty_us_cny"
    if "UK" in s:
        return "duty_uk_cny"
    if "CA" in s or "AU" in s:
        return "duty_ca_au_cny"
    if "JP" in s:
        return "duty_jp_cny"
    return "duty_eu_cny"


def _to_decimal_unit(raw: Any) -> Decimal | None:
    if raw is None:
        return None
    try:
        return Decimal(str(raw))
    except (InvalidOperation, ValueError, TypeError):
        return None


def _pay_ccy_iso(pay_currency: str | None) -> str:
    """归一化付款币种；空视为 CNY。"""
    c = (pay_currency or "").strip().upper()
    if c in ("", "CNY", "RMB", "CNH"):
        return "CNY"
    return c


def _fx_one_pay_unit_to_eur(ref: FxRates, pay_currency: str | None) -> Decimal | None:
    """
    来自 fx_rates.json：1 单位 pay_currency 折合多少 EUR（与 base = pay × 本系数一致）。
    CNY：1 / rmb_per_eur；EUR：1；其余查 rates_to_eur。
    """
    iso = _pay_ccy_iso(pay_currency)
    if iso == "CNY":
        if ref.rmb_per_eur <= 0:
            return None
        return (Decimal("1") / ref.rmb_per_eur).quantize(Decimal("0.00000001"))
    if iso == "EUR":
        return Decimal("1")
    r = ref.rates_to_eur.get(iso)
    if r is None:
        return None
    d = r if isinstance(r, Decimal) else Decimal(str(r))
    if d <= 0:
        return None
    return d


def _load_pricing_map(conn, skus: list[str]) -> dict[str, dict[str, Any]]:
    """批量读取 product_sku_pricing，按 product_sku 建索引（与表 uk_psp_product_sku 一致）。"""
    unique_skus = [s for s in dict.fromkeys(skus) if s]
    if not unique_skus:
        return {}

    cols_sql = ", ".join(f"`{c}`" for c in _PRICING_FETCH_COLS)
    result: dict[str, dict[str, Any]] = {}
    cur = conn.cursor()
    try:
        chunk = 500
        for i in range(0, len(unique_skus), chunk):
            batch = unique_skus[i : i + chunk]
            ph = ", ".join(["%s"] * len(batch))
            sql = f"SELECT {cols_sql} FROM `{PRICING_TABLE}` WHERE `product_sku` IN ({ph})"
            cur.execute(sql, batch)
            for row in cur.fetchall():
                row_dict = {_PRICING_FETCH_COLS[j]: row[j] for j in range(len(_PRICING_FETCH_COLS))}
                psku = str(row_dict.get("product_sku") or "").strip()
                if psku and psku not in result:
                    result[psku] = row_dict
    finally:
        cur.close()
    return result


def _executemany_chunked(
    conn, sql: str, rows: list[tuple[Any, ...]], *, chunk: int = 300
) -> int:
    if not rows:
        return 0
    cur = conn.cursor()
    total = 0
    try:
        for i in range(0, len(rows), chunk):
            batch = rows[i : i + chunk]
            cur.executemany(sql, batch)
            total += len(batch)
    finally:
        cur.close()
    return total


def _count_rows(conn, sql: str, params: list[Any]) -> int:
    cur = conn.cursor()
    try:
        cur.execute(sql, params)
        row = cur.fetchone()
        return int(row[0]) if row and row[0] is not None else 0
    finally:
        cur.close()


def step_mark_distribution_calc_done(
    conn,
    date_from: date | None,
    date_to: date | None,
    order_no: str | None,
    *,
    dry_run: bool = False,
) -> int:
    date_clause, order_filter, scope_params = _scope_parts(date_from, date_to, order_no)
    node = CALC_NODE_STEP01_DONE
    where_rest = f"""
        WHERE p.`distribution_lev` = 1
          AND (p.`calc_node` IS NULL OR p.`calc_node` <> %s)
          {date_clause}
          {order_filter}
    """
    count_sql = f"SELECT COUNT(*) FROM `{TABLE}` p {where_rest}"
    count_params = [node, *scope_params]
    n = _count_rows(conn, count_sql, count_params)
    _LOG.info(f"[step01-1-dist] 待标记 calc_node={node!r} 的分销行：{n}")
    if dry_run:
        _LOG.warn("[step01-1-dist] --dry-run：不写库")
        return 0
    if n == 0:
        return 0
    upd_sql = f"UPDATE `{TABLE}` p SET p.`calc_node` = %s {where_rest}"
    upd_params = [node, node, *scope_params]
    cur = conn.cursor()
    try:
        cur.execute(upd_sql, upd_params)
        return int(cur.rowcount or 0)
    finally:
        cur.close()


def step_sync_bases_from_pays_mark_done(
    conn,
    date_from: date | None,
    date_to: date | None,
    order_no: str | None,
    fx_rates: FxRates,
    *,
    dry_run: bool = False,
) -> int:
    """步骤 1.2：双 pay 已齐 → 用 fx_rates.json 的 pay→EUR 汇率回写双 base → step01_done。"""
    date_clause, order_filter, scope_params = _scope_parts(date_from, date_to, order_no)
    node = CALC_NODE_STEP01_DONE
    where_tail = f"""
        WHERE p.`distribution_lev` = 0
          AND p.`first_leg_shipping_pay` IS NOT NULL AND p.`first_leg_shipping_pay` <> 0
          AND p.`first_leg_tax_pay` IS NOT NULL AND p.`first_leg_tax_pay` <> 0
          AND (p.`calc_node` IS NULL OR p.`calc_node` <> %s)
          {date_clause}
          {order_filter}
    """
    sql_fetch = f"""
        SELECT p.`id`, p.`first_leg_shipping_pay`, p.`first_leg_tax_pay`, p.`pay_currency`
        FROM `{TABLE}` p
        {where_tail.strip()}
    """
    params = [node, *scope_params]
    cur = conn.cursor(dictionary=True)
    try:
        cur.execute(sql_fetch, params)
        rows = cur.fetchall()
    finally:
        cur.close()

    updates: list[tuple[Decimal, Decimal, str, int]] = []
    bad_rate: int = 0
    for row in rows:
        rate = _fx_one_pay_unit_to_eur(fx_rates, row.get("pay_currency"))
        if rate is None:
            bad_rate += 1
            continue
        sp = _to_decimal_unit(row.get("first_leg_shipping_pay"))
        tp = _to_decimal_unit(row.get("first_leg_tax_pay"))
        if sp is None or tp is None or sp <= 0 or tp <= 0:
            bad_rate += 1
            continue
        ship_b = (sp * rate).quantize(_Q6)
        tax_b = (tp * rate).quantize(_Q6)
        updates.append((ship_b, tax_b, node, int(row["id"])))

    n = len(updates)
    _LOG.info(f"[step01-1-2-sync-pay] 可同步 base 的自营行：{n}（原始候选 {len(rows)}；无汇率配置跳过 {bad_rate}）")
    if bad_rate:
        _LOG.warn(
            "[step01-1-2-sync-pay] pay_currency 在 fx_rates.json 中无对应 EUR 汇率时无法折算，"
            "请在 rates_to_eur 中补充该币种或核对 pay_currency 拼写"
        )
    if dry_run:
        _LOG.warn("[step01-1-2-sync-pay] --dry-run：不写库")
        return 0
    if not updates:
        return 0
    sql_upd = (
        f"UPDATE `{TABLE}` SET `first_leg_shipping_base` = %s, `first_leg_tax_base` = %s, "
        f"`calc_node` = %s WHERE `id` = %s"
    )
    return _executemany_chunked(conn, sql_upd, updates)


def step_mark_self_both_base_done(
    conn,
    date_from: date | None,
    date_to: date | None,
    order_no: str | None,
    *,
    dry_run: bool = False,
    log_tag: str = "step01-4-final",
) -> int:
    """双 base 已齐且尚未标记 done → 仅更新 calc_node（步骤 1.1 与步骤 4 共用逻辑）。"""
    date_clause, order_filter, scope_params = _scope_parts(date_from, date_to, order_no)
    node = CALC_NODE_STEP01_DONE
    where_rest = f"""
        WHERE p.`distribution_lev` = 0
          AND p.`first_leg_shipping_base` IS NOT NULL AND p.`first_leg_shipping_base` <> 0
          AND p.`first_leg_tax_base` IS NOT NULL AND p.`first_leg_tax_base` <> 0
          AND (p.`calc_node` IS NULL OR p.`calc_node` <> %s)
          {date_clause}
          {order_filter}
    """
    count_sql = f"SELECT COUNT(*) FROM `{TABLE}` p {where_rest}"
    count_params = [node, *scope_params]
    n = _count_rows(conn, count_sql, count_params)
    _LOG.info(f"[{log_tag}] 双 base 已齐、待标记 calc_node={node!r}：{n}")
    if dry_run:
        _LOG.warn(f"[{log_tag}] --dry-run：不写库")
        return 0
    if n == 0:
        return 0
    upd_sql = f"UPDATE `{TABLE}` p SET p.`calc_node` = %s {where_rest}"
    upd_params = [node, node, *scope_params]
    cur = conn.cursor()
    try:
        cur.execute(upd_sql, upd_params)
        return int(cur.rowcount or 0)
    finally:
        cur.close()


@dataclass(frozen=True)
class _LegFillSpec:
    pay_col: str
    base_col: str
    calc_node: str
    site_to_pricing_col: Callable[[Any], str]
    log_tag: str


def _fill_leg_pay_base_from_pricing(
    conn,
    date_from: date | None,
    date_to: date | None,
    order_no: str | None,
    spec: _LegFillSpec,
    fx_rates: FxRates,
    *,
    dry_run: bool = False,
) -> int:
    """
    步骤 2 / 3：待补行以 base 列为准（first_leg_shipping_base / first_leg_tax_base 为 0 或 NULL），
    与表头一致；利润行解析出的产品 SKU 与定价表 product_sku 关联取 *_cny。
    步骤 3 特例：关税列为空则税费写 0；头程 base 已齐时 calc_node 置 step01_done，否则 step01_tax（见模块说明）。
    使用 fx_rates.json：人民币合计 → EUR，再 pay_currency→EUR 得 pay = base / 汇率，
    回写 pay 与 base。
    """
    allowed = {
        ("first_leg_shipping_pay", "first_leg_shipping_base"),
        ("first_leg_tax_pay", "first_leg_tax_base"),
    }
    if (spec.pay_col, spec.base_col) not in allowed:
        raise ValueError(f"非法列组合: {spec.pay_col} / {spec.base_col}")
    date_clause, order_filter, params = _scope_parts(date_from, date_to, order_no)
    sql_fetch = f"""
        SELECT p.`id`, p.`warehouse_sku`, p.`platform_site`, p.`shipped_qty`, p.`pay_currency`,
               p.`first_leg_shipping_base`,
               COALESCE(m.`product_sku`, p.`warehouse_sku`) AS `product_sku_for_pricing`
        FROM `{TABLE}` p
        LEFT JOIN `{MAPPING_TABLE}` m ON m.`line_hash` = p.`line_hash`
        WHERE p.`distribution_lev` = 0
          AND (p.`{spec.base_col}` IS NULL OR p.`{spec.base_col}` = 0)
          {date_clause}
          {order_filter}
    """
    cur = conn.cursor(dictionary=True)
    try:
        cur.execute(sql_fetch, params)
        rows = cur.fetchall()
    finally:
        cur.close()

    _LOG.info(f"[{spec.log_tag}] {spec.base_col}=0/NULL 的自营行数：{len(rows)}")
    if not rows:
        return 0

    skus = [str(r.get("product_sku_for_pricing") or "").strip() for r in rows]
    pricing_map = _load_pricing_map(conn, skus)
    n_distinct = len(set(s for s in skus if s))
    n_rows_with_pricing = sum(
        1 for r in rows if pricing_map.get(str(r.get("product_sku_for_pricing") or "").strip())
    )
    _LOG.info(
        f"[{spec.log_tag}] product_sku_pricing：按 product_sku 命中 {len(pricing_map)} 条；"
        f"本批利润行有定价 {n_rows_with_pricing}/{len(rows)}（产品 SKU 去重 {n_distinct}）"
    )

    updates: list[tuple[Decimal, Decimal, str, int]] = []
    missing: list[str] = []
    bad_unit: int = 0
    tax_no_duty_done: int = 0
    tax_no_duty_pending_tax_node: int = 0
    bad_cny_to_eur: int = 0
    bad_pay_to_eur: int = 0

    for row in rows:
        psku = str(row.get("product_sku_for_pricing") or "").strip()
        pricing = pricing_map.get(psku)
        if not pricing:
            missing.append(psku)
            continue
        pcol = spec.site_to_pricing_col(row.get("platform_site"))
        unit = _to_decimal_unit(pricing.get(pcol))
        if unit is None:
            if spec.base_col == "first_leg_tax_base":
                ship_b = _to_decimal_unit(row.get("first_leg_shipping_base"))
                if ship_b is not None and ship_b > 0:
                    node = CALC_NODE_STEP01_DONE
                    tax_no_duty_done += 1
                else:
                    node = CALC_NODE_STEP01_TAX
                    tax_no_duty_pending_tax_node += 1
                updates.append((_DEC_ZERO_Q6, _DEC_ZERO_Q6, node, int(row["id"])))
                continue
            bad_unit += 1
            continue
        qty = int(row.get("shipped_qty") or 0)
        cny_total = (unit * Decimal(qty)).quantize(_Q6)
        base_eur_raw = fx_rates.rmb_to_eur(cny_total)
        if base_eur_raw is None:
            bad_cny_to_eur += 1
            continue
        base_eur = base_eur_raw.quantize(_Q6)
        rate = _fx_one_pay_unit_to_eur(fx_rates, row.get("pay_currency"))
        if rate is None:
            bad_pay_to_eur += 1
            continue
        pay_amt = (base_eur / rate).quantize(_Q6)
        updates.append((pay_amt, base_eur, spec.calc_node, int(row["id"])))

    if missing:
        u = sorted(set(missing))
        _LOG.warn(f"[{spec.log_tag}] {len(u)} 个利润行产品 SKU 在定价表无记录（前10）：{u[:10]}")
    if bad_unit or bad_cny_to_eur or bad_pay_to_eur:
        _LOG.warn(
            f"[{spec.log_tag}] 跳过：定价单价无效 {bad_unit} 行；"
            f"人民币→EUR(rmb_per_eur)失败 {bad_cny_to_eur} 行；"
            f"pay_currency 无 EUR 汇率(rates_to_eur) {bad_pay_to_eur} 行"
        )
    if tax_no_duty_done or tax_no_duty_pending_tax_node:
        _LOG.info(
            f"[{spec.log_tag}] duty_*_cny 为空/NULL：无关税写 0；"
            f"calc_node→{CALC_NODE_STEP01_DONE!r} {tax_no_duty_done} 行，"
            f"→{CALC_NODE_STEP01_TAX!r} {tax_no_duty_pending_tax_node} 行（待头程 base）"
        )

    _LOG.info(f"[{spec.log_tag}] 待更新 {spec.pay_col}/{spec.base_col}：{len(updates)} 行")

    if dry_run:
        _LOG.warn(f"[{spec.log_tag}] --dry-run：预览前 3 条")
        for pay_amt, base_eur, node, rid in updates[:3]:
            _LOG.info(
                f"  [dry-run] id={rid} {spec.pay_col}={pay_amt} {spec.base_col}={base_eur} "
                f"calc_node={node!r}"
            )
        return 0

    sql_upd = (
        f"UPDATE `{TABLE}` SET `{spec.pay_col}` = %s, `{spec.base_col}` = %s, "
        f"`calc_node` = %s WHERE `id` = %s"
    )
    return _executemany_chunked(conn, sql_upd, list(updates))


def step_fill_first_leg_shipping(
    conn,
    date_from: date | None,
    date_to: date | None,
    order_no: str | None,
    fx_rates: FxRates,
    *,
    dry_run: bool = False,
) -> int:
    return _fill_leg_pay_base_from_pricing(
        conn,
        date_from,
        date_to,
        order_no,
        _LegFillSpec(
            pay_col="first_leg_shipping_pay",
            base_col="first_leg_shipping_base",
            calc_node=CALC_NODE_STEP01_SHIPPING,
            site_to_pricing_col=_site_to_first_leg_col,
            log_tag="step01-2-shipping",
        ),
        fx_rates,
        dry_run=dry_run,
    )


def step_fill_first_leg_tax(
    conn,
    date_from: date | None,
    date_to: date | None,
    order_no: str | None,
    fx_rates: FxRates,
    *,
    dry_run: bool = False,
) -> int:
    return _fill_leg_pay_base_from_pricing(
        conn,
        date_from,
        date_to,
        order_no,
        _LegFillSpec(
            pay_col="first_leg_tax_pay",
            base_col="first_leg_tax_base",
            calc_node=CALC_NODE_STEP01_TAX,
            site_to_pricing_col=_site_to_duty_col,
            log_tag="step01-3-tax",
        ),
        fx_rates,
        dry_run=dry_run,
    )


def run_step01(
    conn,
    date_from: date | None,
    date_to: date | None,
    order_no: str | None,
    *,
    dry_run: bool = False,
    through_step: int = 0,
    calc_mode: str = CALC_MODE_STRICT,
) -> dict[str, int]:
    last = _last_phase_index(through_step)
    if calc_mode not in (CALC_MODE_STRICT, CALC_MODE_LOOSE):
        raise ValueError(f"calc_mode 须为 {CALC_MODE_STRICT!r} 或 {CALC_MODE_LOOSE!r}")

    stats: dict[str, int] = {
        "through_step": through_step,
        "distribution_calc_done": 0,
        "bases_ready_1_1_mark": 0,
        "sync_base_from_pay_1_2": 0,
        "first_leg_shipping": 0,
        "first_leg_tax": 0,
        "self_final_done": 0,
    }

    fx_rates: FxRates | None = None
    if _need_fx_rates_json(last_phase=last, calc_mode=calc_mode):
        fx_rates = load_rates(_FX_JSON_PATH)
        if fx_rates.issues:
            for msg in fx_rates.issues:
                _LOG.warn(f"[fx_rates] {msg}")
        _LOG.info(
            f"[fx_rates] 文件={_FX_JSON_PATH} source={fx_rates.source!r} "
            f"rmb_per_eur={fx_rates.rmb_per_eur} "
            f"rates_to_eur 币种数={len(fx_rates.rates_to_eur)}"
        )

    if last >= 0:
        stats["distribution_calc_done"] = step_mark_distribution_calc_done(
            conn, date_from, date_to, order_no, dry_run=dry_run
        )
        stats["bases_ready_1_1_mark"] = step_mark_self_both_base_done(
            conn,
            date_from,
            date_to,
            order_no,
            dry_run=dry_run,
            log_tag="step01-1-1-base-done",
        )
        if calc_mode == CALC_MODE_STRICT:
            if fx_rates is None:
                raise RuntimeError("strict 模式需要加载 fx_rates.json")
            stats["sync_base_from_pay_1_2"] = step_sync_bases_from_pays_mark_done(
                conn, date_from, date_to, order_no, fx_rates, dry_run=dry_run
            )
        else:
            _LOG.info("[step01] 宽松模式（loose）：跳过步骤 1.2（不按 pay 回写 base）")
            stats["sync_base_from_pay_1_2"] = 0

    if last >= 1 and fx_rates is not None:
        stats["first_leg_shipping"] = step_fill_first_leg_shipping(
            conn, date_from, date_to, order_no, fx_rates, dry_run=dry_run
        )
    if last >= 2 and fx_rates is not None:
        stats["first_leg_tax"] = step_fill_first_leg_tax(
            conn, date_from, date_to, order_no, fx_rates, dry_run=dry_run
        )
    if last >= 3:
        stats["self_final_done"] = step_mark_self_both_base_done(
            conn,
            date_from,
            date_to,
            order_no,
            dry_run=dry_run,
            log_tag="step01-4-final",
        )

    if not dry_run:
        conn.commit()
        _LOG.info(
            "[step01] 事务已提交："
            f"calc_mode={calc_mode}；分销 {stats['distribution_calc_done']}；"
            f"1.1 打标 {stats['bases_ready_1_1_mark']}；"
            f"1.2 按 pay 同步 base {stats['sync_base_from_pay_1_2']}；"
            f"头程运费 {stats['first_leg_shipping']}；"
            f"头程税费 {stats['first_leg_tax']}；"
            f"最终完成 {stats['self_final_done']}（through_step={through_step}）"
        )
    return stats


def main() -> int:
    setup_stdout_utf8()
    ap = argparse.ArgumentParser(
        description=(
            "step01：分销；头程 pay→base 与人民币定价折算一律使用 python/v2/config/fx_rates.json；"
            "补全 pay+base 后收口 calc_node"
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
        help="只处理指定订单号（与 --date-from/--date-to 可同时使用）；多个单号用逗号分隔，如 A001,B002",
    )
    ap.add_argument("--dry-run", action="store_true", help="不写库，仅预览前 3 条变更")
    ap.add_argument(
        "--through-step",
        type=int,
        default=0,
        metavar="N",
        help="0（默认）或 4=步骤1～4；1=仅步骤1；2=步骤1～2；3=步骤1～3",
    )
    ap.add_argument(
        "--calc-mode",
        choices=(CALC_MODE_STRICT, CALC_MODE_LOOSE),
        default=CALC_MODE_STRICT,
        help=(
            f"{CALC_MODE_STRICT}（默认）执行步骤 1.2：双 pay 已齐时用 fx_rates.json 回写 base；"
            f"{CALC_MODE_LOOSE} 跳过 1.2"
        ),
    )
    args = ap.parse_args()

    try:
        _last_phase_index(args.through_step)
    except ValueError as e:
        ap.error(str(e))

    _LOG.info("=" * 60)
    nos = _order_nos_from_arg(args.order_no)
    if nos:
        tail = "，…" if len(nos) > 8 else ""
        _LOG.info(f"订单号筛选：共 {len(nos)} 个 → {', '.join(nos[:8])}{tail}")
    _LOG.info(
        f"step01（through_step={args.through_step}，calc_mode={args.calc_mode}）："
        "分销 → 1.1 双 base 打标 → [strict 时 1.2 按 pay 回写 base] → 补运费 → 补税费 → 4 收口"
    )

    cfg = load_db_config()
    _LOG.info(f"连接：{cfg.host}:{cfg.port}  db={cfg.database}")
    conn = connect(cfg)
    try:
        stats = run_step01(
            conn,
            date_from=args.date_from,
            date_to=args.date_to,
            order_no=args.order_no,
            dry_run=args.dry_run,
            through_step=args.through_step,
            calc_mode=args.calc_mode,
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
