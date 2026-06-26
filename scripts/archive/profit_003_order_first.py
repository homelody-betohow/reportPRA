from __future__ import annotations

"""
从 sales_order_shipped 表中读取指定 import_batch 的数据，
按「SKU 主数据中的头程/关税（RMB/件）」分两步更新利润表（数据源可切换为 Excel 或 DB）：

  步骤 1 — first_leg_shipping_base
    p.first_leg_shipping_base = 单个头程(RMB/件) * warehouse_sku_qty / RMB_di_EUR
    calc_node → first_leg_shipping

  步骤 2 — first_leg_tax_base
    p.first_leg_tax_base = 单个关税(RMB/件) * warehouse_sku_qty / RMB_di_EUR
    若头程 base 已非 0，calc_node → first_leg；否则 → first_leg_tax

distribution_lev 非 0 的分销行（仅步骤 1 处理）：标记 calc_node=first_leg，不查主数据头程/关税。

注意：为了便于追溯数据，本脚本**不更新** sales_order_shipped 表，仅使用其字段做计算输入。

用法：
  cd d:\\py-project\\report
  python scripts\\archive\\profit_003_order_first.py
  python scripts\\archive\\profit_003_order_first.py --step 1
  python scripts\\archive\\profit_003_order_first.py --step 2
  python scripts\\archive\\profit_003_order_first.py --batch 20260616_203140
  python scripts\\archive\\profit_003_order_first.py --dry-run
"""

import argparse
import sys
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Literal

import pymysql.cursors

_REPORT_ROOT = Path(__file__).resolve().parents[2]
_PROJECT_ROOT = Path(__file__).resolve().parents[3]  # d:\py-project
_ARCHIVE_DIR = Path(__file__).resolve().parent
_DATA_IMPORT_DIR = _REPORT_ROOT / "scripts" / "dataImport"

for _p in (_PROJECT_ROOT, _REPORT_ROOT, _ARCHIVE_DIR, _DATA_IMPORT_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from batch_lock import read_import_batch_from_lock  # noqa: E402  # pyright: ignore[reportMissingImports]
from database.db_connection import DatabaseConfig, get_db_manager  # noqa: E402
from config import common as cfg_common  # noqa: E402

SOURCE_TABLE = "sales_order_shipped"
PROFIT_TABLE = "sales_order_sku_profit"
PRICING_TABLE = "product_sku_pricing"

CALC_NODE_DONE = "first_leg"
CALC_NODE_SHIPPING = "first_leg_shipping"
CALC_NODE_TAX = "first_leg_tax"
DISTRIBUTION_LEV = 1
_ZERO_DEC = Decimal("0")
_SIX_PLACES = Decimal("0.000001")

StepKind = Literal["shipping", "tax"]

# =========================================================
# 主数据来源切换（你只需要改这里）
#   pricing_schedule = "excel" -> 使用 report/config/common.py 的 BTH_ALL_SKU_DETAIL_PATH
#   pricing_schedule = "db"    -> 使用数据库表 product_sku_pricing
# =========================================================
pricing_schedule = "excel"  # 改成 "db" 即可切换到 DB
# pricing_schedule = "db"  # 改成 "excel" 即可切换到 Excel

# RMB兑换EUR
# 使用 report/config/common.py 中的换算常量：RMB -> EUR（除法）
RMB_TO_EUR_DIVISOR = Decimal(str(cfg_common.RMB_di_EUR))


def _log(level: str, msg: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] [{level}] {msg}")


def _dec(v: Any) -> Decimal | None:
    if v is None:
        return None
    if isinstance(v, Decimal):
        return v
    try:
        s = str(v).strip()
        if not s:
            return None
        return Decimal(s)
    except Exception:
        return None


def _int(v: Any) -> int:
    if v is None:
        return 0
    try:
        return int(v)
    except Exception:
        return 0


def _str(v: Any) -> str:
    return "" if v is None else str(v).strip()


def _is_distribution_row(row: dict[str, Any]) -> bool:
    """distribution_lev 非 0 时视为分销行，不查主数据头程/关税。"""
    lev = row.get("p_distribution_lev")
    if lev is None:
        lev = row.get("distribution_lev")
    return _int(lev) != 0


def _sku_lookup_key(warehouse_sku: str) -> str:
    """去空格、去掉 -NW 后缀作为主数据查表键。"""
    s = (warehouse_sku or "").strip()
    if s.upper().endswith("-NW"):
        s = s[:-3]
    return s


def _load_bth_pricing_from_excel(excel_path: Path) -> dict[str, dict[str, Decimal | None]]:
    """
    从「BTH全部SKU明细」Excel 的 sheet「基础数据维护」读取头程/关税（RMB/件）。

    返回：sku -> {
      "first_leg_eu_au_cny": ...,
      "first_leg_us_cny": ...,
      "first_leg_uk_cny": ...,
      "duty_eu_cny": ...,
      "duty_us_cny": ...,
      "duty_uk_cny": ...,
    }
    """
    try:
        import pandas as pd  # type: ignore
    except Exception as e:  # pragma: no cover
        raise RuntimeError(f"缺少 pandas 依赖，无法使用 Excel 模式：{e}") from e

    if not excel_path.is_file():
        raise RuntimeError(f"Excel 不存在或无法访问：{excel_path}")

    df = pd.read_excel(
        excel_path,
        sheet_name="基础数据维护",
        header=[0, 1],
        engine="openpyxl",
        dtype=object,
    )

    def _norm(x: Any) -> str:
        if x is None:
            return ""
        s = str(x).replace("\n", "").replace("\r", "").strip()
        return "" if s.lower() == "nan" else s

    cols_2: list[tuple[str, str]] = []
    raw_cols: list[Any] = list(df.columns)
    for c in raw_cols:
        if isinstance(c, tuple) and len(c) >= 2:
            cols_2.append((_norm(c[0]), _norm(c[1])))
        else:
            cols_2.append((_norm(c), ""))

    sku_labels = {"SKU", "产品编码"}
    sku_col_idx = next(
        (
            i
            for i, (l1, l2) in enumerate(cols_2)
            if _norm(l1).upper() in {s.upper() for s in sku_labels}
            or _norm(l2).upper() in {s.upper() for s in sku_labels}
        ),
        None,
    )
    if sku_col_idx is None:
        sample = [f"{i}:{l1}|{l2}" for i, (l1, l2) in enumerate(cols_2[:50])]
        raise RuntimeError(
            "Excel 表头未找到 SKU 列（基础数据维护）。"
            f"前50列表头示例：{sample}"
        )
    sku_col = raw_cols[sku_col_idx]

    def _find_col(l1_want: str, l2_want: str) -> Any | None:
        for i, (l1, l2) in enumerate(cols_2):
            if l1 == l1_want and l2 == l2_want:
                return raw_cols[i]
        for i, (l1, l2) in enumerate(cols_2):
            if l2 == l2_want and l1_want in l1:
                return raw_cols[i]
        return None

    want_cols: dict[str, Any] = {}
    want_specs: dict[str, tuple[str, str]] = {
        "first_leg_eu_au_cny": ("头程（RMB）", "EU/AU"),
        "first_leg_us_cny": ("头程（RMB）", "US"),
        "first_leg_uk_cny": ("头程（RMB）", "UK"),
        "duty_eu_cny": ("关税（含税）", "EU"),
        "duty_us_cny": ("关税（含税）", "US"),
        "duty_uk_cny": ("关税（含税）", "UK"),
    }
    missing: list[str] = []
    for k, (l1, l2) in want_specs.items():
        col = _find_col(l1, l2)
        if col is None:
            missing.append(f"{l1}|{l2}")
        else:
            want_cols[k] = col

    if missing:
        sample = [f"{i}:{l1}|{l2}" for i, (l1, l2) in enumerate(cols_2[:80])]
        raise RuntimeError(
            f"Excel 缺少关键列：{missing}。"
            f"前80列表头示例：{sample}"
        )

    out: dict[str, dict[str, Decimal | None]] = {}
    for _, row in df.iterrows():
        sku_raw = row.get(sku_col)
        sku = _sku_lookup_key("" if sku_raw is None else str(sku_raw))
        if not sku:
            continue

        d: dict[str, Decimal | None] = {}
        for k, col in want_cols.items():
            d[k] = _dec(row.get(col))
        out[sku] = d
    return out


def _unit_shipping_cny_from_excel(pricing_row: dict[str, Decimal | None], market: str) -> Decimal | None:
    if market == "US":
        return pricing_row.get("first_leg_us_cny")
    if market == "UK":
        return pricing_row.get("first_leg_uk_cny")
    if market in {"CA", "JP", "AU"}:
        return pricing_row.get("first_leg_eu_au_cny")
    return pricing_row.get("first_leg_eu_au_cny")


def _unit_tax_cny_from_excel(pricing_row: dict[str, Decimal | None], market: str) -> Decimal | None:
    if market == "US":
        return pricing_row.get("duty_us_cny")
    if market == "UK":
        return pricing_row.get("duty_uk_cny")
    if market in {"CA", "JP", "AU"}:
        return pricing_row.get("duty_eu_cny")
    return pricing_row.get("duty_eu_cny")


def _market_group(*, platform_site: str, shop_alias: str, warehouse_name: str) -> str:
    """
    站点分组（用于选择 product_sku_pricing 的哪一列）：
    - US: platform_site=US，但 shop_alias=DLZ-US 视为非 US
    - UK: platform_site=UK/GB
    - CA / JP / AU
    - 默认：EU
    """
    site = (platform_site or "").strip().upper()
    alias = (shop_alias or "").strip().upper()
    wh = (warehouse_name or "").strip()

    if site == "US" and alias == "DLZ-US":
        return "EU"
    if site == "US" and ("HY-DLZ-DE" in wh or "德国" in wh):
        return "EU"
    if site == "US":
        return "US"
    if site in {"UK", "GB"}:
        return "UK"
    if site == "CA":
        return "CA"
    if site == "JP":
        return "JP"
    if site == "AU":
        return "AU"
    return "EU"


def _unit_shipping_cny(pricing_row: dict[str, Any], market: str) -> Decimal | None:
    if market == "US":
        return _dec(pricing_row.get("first_leg_us_cny"))
    if market == "UK":
        return _dec(pricing_row.get("first_leg_uk_cny"))
    if market == "CA":
        return _dec(pricing_row.get("first_leg_ca_cny"))
    if market == "JP":
        return _dec(pricing_row.get("first_leg_jp_cny"))
    if market == "AU":
        return _dec(pricing_row.get("first_leg_eu_au_cny"))
    return _dec(pricing_row.get("first_leg_eu_au_cny"))


def _unit_tax_cny(pricing_row: dict[str, Any], market: str) -> Decimal | None:
    if market == "US":
        return _dec(pricing_row.get("duty_us_cny"))
    if market == "UK":
        return _dec(pricing_row.get("duty_uk_cny"))
    if market == "CA":
        return _dec(pricing_row.get("duty_ca_au_cny"))
    if market == "JP":
        return _dec(pricing_row.get("duty_jp_cny"))
    if market == "AU":
        return _dec(pricing_row.get("duty_ca_au_cny"))
    return _dec(pricing_row.get("duty_eu_cny"))


def _cny_to_eur_amount(cny: Decimal) -> Decimal:
    """RMB -> EUR：保留 6 位小数以匹配 decimal(18,6)。"""
    if cny == _ZERO_DEC:
        return _ZERO_DEC
    return (cny / RMB_TO_EUR_DIVISOR).quantize(_SIX_PLACES, rounding=ROUND_HALF_UP)


def _unit_cost_cny(
    pricing_row: dict[str, Any],
    market: str,
    *,
    step: StepKind,
    pricing_source: str,
) -> Decimal | None:
    if pricing_source == "excel":
        if step == "shipping":
            return _unit_shipping_cny_from_excel(pricing_row, market)  # type: ignore[arg-type]
        return _unit_tax_cny_from_excel(pricing_row, market)  # type: ignore[arg-type]
    if step == "shipping":
        return _unit_shipping_cny(pricing_row, market)
    return _unit_tax_cny(pricing_row, market)


def fetch_shipping_candidates(conn, import_batch: str) -> list[dict[str, Any]]:
    """步骤 1：待补 first_leg_shipping_base 的普通仓 + 待打标的分销行。"""
    sql = f"""
        SELECT
            s.`line_hash` AS line_hash,
            s.`import_batch` AS import_batch,
            s.`platform_site` AS platform_site,
            s.`shop_alias` AS shop_alias,
            s.`warehouse_name` AS warehouse_name,
            s.`warehouse_sku` AS warehouse_sku,
            s.`warehouse_sku_qty` AS warehouse_sku_qty,
            p.`first_leg_shipping_base` AS p_first_leg_shipping_base,
            p.`calc_node` AS p_calc_node,
            p.`distribution_lev` AS p_distribution_lev
        FROM `{SOURCE_TABLE}` AS s
        INNER JOIN `{PROFIT_TABLE}` AS p
          ON p.`line_hash` = s.`line_hash`
        WHERE s.`import_batch` = %s
          AND (
            (
              IFNULL(p.`distribution_lev`, 0) <> 0
              AND IFNULL(p.`calc_node`, '') <> %s
            )
            OR (
              IFNULL(p.`distribution_lev`, 0) = 0
              AND (
                p.`first_leg_shipping_base` IS NULL OR p.`first_leg_shipping_base` = 0
                OR IFNULL(p.`calc_node`, '') NOT IN (%s, %s)
              )
            )
          )
    """
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        _log("INFO", f"[步骤1] 查询待更新 first_leg_shipping_base：batch={import_batch}")
        cur.execute(
            sql,
            (
                import_batch,
                CALC_NODE_DONE,
                CALC_NODE_SHIPPING,
                CALC_NODE_DONE,
            ),
        )
        rows = cur.fetchall()
        _log("INFO", f"[步骤1] 候选行数：{len(rows)}")
        return rows
    finally:
        cur.close()


def fetch_tax_candidates(conn, import_batch: str) -> list[dict[str, Any]]:
    """步骤 2：待补 first_leg_tax_base 的普通仓（distribution_lev=0）。"""
    sql = f"""
        SELECT
            s.`line_hash` AS line_hash,
            s.`import_batch` AS import_batch,
            s.`platform_site` AS platform_site,
            s.`shop_alias` AS shop_alias,
            s.`warehouse_name` AS warehouse_name,
            s.`warehouse_sku` AS warehouse_sku,
            s.`warehouse_sku_qty` AS warehouse_sku_qty,
            p.`first_leg_shipping_base` AS p_first_leg_shipping_base,
            p.`first_leg_tax_base` AS p_first_leg_tax_base,
            p.`calc_node` AS p_calc_node
        FROM `{SOURCE_TABLE}` AS s
        INNER JOIN `{PROFIT_TABLE}` AS p
          ON p.`line_hash` = s.`line_hash`
        WHERE s.`import_batch` = %s
          AND IFNULL(p.`distribution_lev`, 0) = 0
          AND (
            p.`first_leg_tax_base` IS NULL OR p.`first_leg_tax_base` = 0
            OR IFNULL(p.`calc_node`, '') <> %s
          )
    """
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        _log("INFO", f"[步骤2] 查询待更新 first_leg_tax_base：batch={import_batch}")
        cur.execute(sql, (import_batch, CALC_NODE_DONE))
        rows = cur.fetchall()
        _log("INFO", f"[步骤2] 候选行数：{len(rows)}")
        return rows
    finally:
        cur.close()


def fetch_pricing(conn, product_skus: list[str]) -> dict[str, dict[str, Any]]:
    """批量读取 product_sku_pricing，返回 product_sku -> row(dict)。"""
    product_skus = [s for s in product_skus if s]
    if not product_skus:
        return {}

    cols = [
        "product_sku",
        "first_leg_eu_au_cny",
        "first_leg_us_cny",
        "first_leg_uk_cny",
        "first_leg_ca_cny",
        "first_leg_jp_cny",
        "duty_eu_cny",
        "duty_us_cny",
        "duty_uk_cny",
        "duty_ca_au_cny",
        "duty_jp_cny",
    ]
    cols_sql = ", ".join(f"`{c}`" for c in cols)

    out: dict[str, dict[str, Any]] = {}
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        chunk = 800
        for i in range(0, len(product_skus), chunk):
            part = product_skus[i : i + chunk]
            placeholders = ",".join(["%s"] * len(part))
            sql = f"SELECT {cols_sql} FROM `{PRICING_TABLE}` WHERE `product_sku` IN ({placeholders})"
            cur.execute(sql, part)
            for row in cur.fetchall():
                sku = _str(row.get("product_sku"))
                if sku:
                    out[sku] = row
        return out
    finally:
        cur.close()


def _collect_sku_keys(candidates: list[dict[str, Any]]) -> list[str]:
    return sorted(
        {
            _sku_lookup_key(_str(r.get("warehouse_sku")))
            for r in candidates
            if _str(r.get("warehouse_sku")) and not _is_distribution_row(r)
        }
    )


def _load_pricing_by_sku(
    conn,
    sku_keys: list[str],
    *,
    pricing_source: str,
    excel_file: Path | None,
) -> dict[str, dict[str, Any]]:
    if not sku_keys:
        return {}

    if pricing_source == "excel":
        excel_path = excel_file
        if excel_path is None:
            excel_path = Path(str(cfg_common.BTH_ALL_SKU_DETAIL_PATH))
        if excel_path is None:
            raise RuntimeError(
                "excel 模式需要 --excel-file 指定 BTH全部SKU明细-*.xlsx，"
                "或在 report/config/common.py 配置 BTH_ALL_SKU_DETAIL_PATH"
            )
        pricing_by_sku = _load_bth_pricing_from_excel(excel_path)
        _log("WARN", f"已加载 Excel 主数据：{excel_path.name}（SKU数={len(pricing_by_sku)}）")
        return pricing_by_sku

    pricing_by_sku = fetch_pricing(conn, sku_keys)
    _log("INFO", f"已加载 {len(pricing_by_sku)} 个 SKU 的主数据价格（{PRICING_TABLE}）")
    return pricing_by_sku


def build_shipping_updates(
    candidates: list[dict[str, Any]],
    pricing_by_sku: dict[str, dict[str, Any]],
    *,
    pricing_source: str,
) -> tuple[
    list[tuple[Decimal, str, str]],
    list[tuple[str, str]],
    dict[str, int],
]:
    """
    步骤 1 UPDATE 参数：
    - 普通仓：(first_leg_shipping_base, calc_node, line_hash)
    - 分销行：(calc_node, line_hash)
    """
    normal_updates: list[tuple[Decimal, str, str]] = []
    dist_updates: list[tuple[str, str]] = []
    stats = {
        "candidates": len(candidates),
        "distribution_warehouse": 0,
        "missing_sku": 0,
        "missing_pricing": 0,
        "missing_cost": 0,
        "updated": 0,
        "zero_cost_allowed": 0,
    }

    for r in candidates:
        line_hash = _str(r.get("line_hash"))
        wh_sku = _str(r.get("warehouse_sku"))
        warehouse_name = _str(r.get("warehouse_name"))
        if not line_hash or not wh_sku:
            stats["missing_sku"] += 1
            continue

        if _is_distribution_row(r):
            dist_updates.append((CALC_NODE_DONE, line_hash))
            stats["distribution_warehouse"] += 1
            continue

        qty = _int(r.get("warehouse_sku_qty"))
        if qty <= 0:
            stats["missing_cost"] += 1
            continue

        lookup = _sku_lookup_key(wh_sku)
        pricing = pricing_by_sku.get(lookup) or pricing_by_sku.get(wh_sku)
        if not pricing:
            stats["missing_pricing"] += 1
            continue

        market = _market_group(
            platform_site=_str(r.get("platform_site")),
            shop_alias=_str(r.get("shop_alias")),
            warehouse_name=warehouse_name,
        )
        unit_ship_cny = _unit_cost_cny(pricing, market, step="shipping", pricing_source=pricing_source)
        if unit_ship_cny is None:
            stats["missing_cost"] += 1
            continue

        ship_eur = _cny_to_eur_amount(unit_ship_cny * Decimal(qty))
        if ship_eur == _ZERO_DEC:
            stats["zero_cost_allowed"] += 1

        normal_updates.append((ship_eur, CALC_NODE_SHIPPING, line_hash))

    stats["updated"] = len(normal_updates) + len(dist_updates)
    return normal_updates, dist_updates, stats


def build_tax_updates(
    candidates: list[dict[str, Any]],
    pricing_by_sku: dict[str, dict[str, Any]],
    *,
    pricing_source: str,
) -> tuple[list[tuple[Decimal, str, str]], dict[str, int]]:
    """
    步骤 2 UPDATE 参数：(first_leg_tax_base, calc_node, line_hash)
    头程 base 已非 0 时 calc_node=first_leg，否则 first_leg_tax。
    """
    updates: list[tuple[Decimal, str, str]] = []
    stats = {
        "candidates": len(candidates),
        "missing_sku": 0,
        "missing_pricing": 0,
        "missing_cost": 0,
        "updated": 0,
        "zero_cost_allowed": 0,
        "done_node": 0,
        "tax_node": 0,
    }

    for r in candidates:
        line_hash = _str(r.get("line_hash"))
        wh_sku = _str(r.get("warehouse_sku"))
        warehouse_name = _str(r.get("warehouse_name"))
        if not line_hash or not wh_sku:
            stats["missing_sku"] += 1
            continue

        qty = _int(r.get("warehouse_sku_qty"))
        if qty <= 0:
            stats["missing_cost"] += 1
            continue

        lookup = _sku_lookup_key(wh_sku)
        pricing = pricing_by_sku.get(lookup) or pricing_by_sku.get(wh_sku)
        if not pricing:
            stats["missing_pricing"] += 1
            continue

        market = _market_group(
            platform_site=_str(r.get("platform_site")),
            shop_alias=_str(r.get("shop_alias")),
            warehouse_name=warehouse_name,
        )
        unit_tax_cny = _unit_cost_cny(pricing, market, step="tax", pricing_source=pricing_source)
        if unit_tax_cny is None:
            stats["missing_cost"] += 1
            continue

        tax_eur = _cny_to_eur_amount(unit_tax_cny * Decimal(qty))
        if tax_eur == _ZERO_DEC:
            stats["zero_cost_allowed"] += 1

        ship_base = _dec(r.get("p_first_leg_shipping_base"))
        if ship_base is not None and ship_base != _ZERO_DEC:
            calc_node = CALC_NODE_DONE
            stats["done_node"] += 1
        else:
            calc_node = CALC_NODE_TAX
            stats["tax_node"] += 1

        updates.append((tax_eur, calc_node, line_hash))

    stats["updated"] = len(updates)
    return updates, stats


def apply_shipping_updates(
    conn,
    normal_updates: list[tuple[Decimal, str, str]],
    dist_updates: list[tuple[str, str]],
) -> int:
    if not normal_updates and not dist_updates:
        return 0

    normal_sql = f"""
        UPDATE `{PROFIT_TABLE}`
        SET
            first_leg_shipping_base = %s,
            calc_node = %s
        WHERE line_hash = %s
    """
    dist_sql = f"""
        UPDATE `{PROFIT_TABLE}`
        SET
            calc_node = %s,
            distribution_lev = %s
        WHERE line_hash = %s
    """
    cur = conn.cursor()
    try:
        chunk = 500
        affected = 0
        for i in range(0, len(normal_updates), chunk):
            part = normal_updates[i : i + chunk]
            cur.executemany(normal_sql, part)
            affected += len(part)
        dist_params = [(node, DISTRIBUTION_LEV, line_hash) for node, line_hash in dist_updates]
        for i in range(0, len(dist_params), chunk):
            part = dist_params[i : i + chunk]
            cur.executemany(dist_sql, part)
            affected += len(part)
        return affected
    finally:
        cur.close()


def apply_tax_updates(conn, updates: list[tuple[Decimal, str, str]]) -> int:
    if not updates:
        return 0

    sql = f"""
        UPDATE `{PROFIT_TABLE}`
        SET
            first_leg_tax_base = %s,
            calc_node = %s
        WHERE line_hash = %s
    """
    cur = conn.cursor()
    try:
        chunk = 500
        affected = 0
        for i in range(0, len(updates), chunk):
            part = updates[i : i + chunk]
            cur.executemany(sql, part)
            affected += len(part)
        return affected
    finally:
        cur.close()


def finalize_done_node(conn, import_batch: str) -> int:
    """
    步骤 2 收尾：头程/关税 base 均已非 0，但 calc_node 尚未 first_leg 的行统一收口。
    """
    sql = f"""
        UPDATE `{PROFIT_TABLE}` AS p
        INNER JOIN `{SOURCE_TABLE}` AS s ON s.`line_hash` = p.`line_hash`
        SET p.`calc_node` = %s
        WHERE s.`import_batch` = %s
          AND IFNULL(p.`distribution_lev`, 0) = 0
          AND p.`first_leg_shipping_base` IS NOT NULL AND p.`first_leg_shipping_base` <> 0
          AND p.`first_leg_tax_base` IS NOT NULL AND p.`first_leg_tax_base` <> 0
          AND IFNULL(p.`calc_node`, '') <> %s
    """
    cur = conn.cursor()
    try:
        cur.execute(sql, (CALC_NODE_DONE, import_batch, CALC_NODE_DONE))
        return int(cur.rowcount or 0)
    finally:
        cur.close()


def run_shipping_step(
    conn,
    import_batch: str,
    *,
    pricing_source: str,
    excel_file: Path | None,
    dry_run: bool,
) -> int:
    candidates = fetch_shipping_candidates(conn, import_batch)
    if not candidates:
        _log("INFO", "[步骤1] 没有需要更新的行")
        return 0

    sku_keys = _collect_sku_keys(candidates)
    if sku_keys:
        pricing_by_sku = _load_pricing_by_sku(
            conn, sku_keys, pricing_source=pricing_source, excel_file=excel_file
        )
    else:
        pricing_by_sku = {}
        _log("INFO", "[步骤1] 无需加载主数据（候选行均为分销行）")

    normal_updates, dist_updates, stats = build_shipping_updates(
        candidates, pricing_by_sku, pricing_source=pricing_source
    )
    _log(
        "INFO",
        "[步骤1] 统计："
        f"候选={stats['candidates']} "
        f"可更新={stats['updated']} "
        f"分销行直标={stats['distribution_warehouse']} "
        f"缺SKU/ID={stats['missing_sku']} "
        f"缺主数据={stats['missing_pricing']} "
        f"缺费用/销量=0={stats['missing_cost']} "
        f"计算结果全0={stats['zero_cost_allowed']}",
    )

    if not normal_updates and not dist_updates:
        _log("WARN", "[步骤1] 计算后无可写入行")
        return 0

    if dry_run:
        _log("INFO", "[步骤1] dry-run：已跳过写库")
        return 0

    n = apply_shipping_updates(conn, normal_updates, dist_updates)
    _log("INFO", f"[步骤1] 写库完成：first_leg_shipping_base 更新 {n} 行")
    return n


def run_tax_step(
    conn,
    import_batch: str,
    *,
    pricing_source: str,
    excel_file: Path | None,
    dry_run: bool,
) -> int:
    candidates = fetch_tax_candidates(conn, import_batch)
    if not candidates:
        _log("INFO", "[步骤2] 没有需要更新的行")
        return 0

    sku_keys = _collect_sku_keys(candidates)
    pricing_by_sku = _load_pricing_by_sku(
        conn, sku_keys, pricing_source=pricing_source, excel_file=excel_file
    ) if sku_keys else {}

    updates, stats = build_tax_updates(candidates, pricing_by_sku, pricing_source=pricing_source)
    _log(
        "INFO",
        "[步骤2] 统计："
        f"候选={stats['candidates']} "
        f"可更新={stats['updated']} "
        f"缺SKU/ID={stats['missing_sku']} "
        f"缺主数据={stats['missing_pricing']} "
        f"缺费用/销量=0={stats['missing_cost']} "
        f"计算结果全0={stats['zero_cost_allowed']} "
        f"calc_node→{CALC_NODE_DONE}={stats['done_node']} "
        f"calc_node→{CALC_NODE_TAX}={stats['tax_node']}",
    )

    affected = 0
    if updates:
        if dry_run:
            _log("INFO", "[步骤2] dry-run：已跳过写库")
        else:
            affected += apply_tax_updates(conn, updates)
            _log("INFO", f"[步骤2] 写库完成：first_leg_tax_base 更新 {affected} 行")

    if not dry_run:
        n_done = finalize_done_node(conn, import_batch)
        if n_done:
            _log("INFO", f"[步骤2] 收口 calc_node={CALC_NODE_DONE!r}：{n_done} 行")
        affected += n_done

    return affected


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description=(
            "分两步更新 sales_order_sku_profit 的头程/关税 base："
            "步骤1=first_leg_shipping_base，步骤2=first_leg_tax_base"
        )
    )
    ap.add_argument(
        "--batch",
        type=str,
        default=None,
        metavar="BATCH",
        help="指定导入批次号（默认从 run_batch.lock 文件读取）",
    )
    ap.add_argument(
        "--step",
        type=int,
        choices=(0, 1, 2),
        default=0,
        help="执行步骤：0=两步都跑（默认），1=仅头程运费，2=仅头程关税",
    )
    ap.add_argument(
        "--dry-run",
        action="store_true",
        help="只计算不写库（仍会输出统计）",
    )
    ap.add_argument(
        "--pricing-source",
        choices=("db", "excel"),
        default=str(pricing_schedule).strip().lower() if str(pricing_schedule).strip().lower() in {"db", "excel"} else "db",
        help=(
            "头程/关税主数据来源：db=product_sku_pricing；excel=BTH全部SKU明细。"
            "默认值来自脚本顶部变量 pricing_schedule（建议直接改变量）。"
        ),
    )
    ap.add_argument(
        "--excel-file",
        type=Path,
        default=None,
        help="当 --pricing-source=excel 时，可手动指定 BTH全部SKU明细-*.xlsx；未指定则使用 report/config/common.py 的 BTH_ALL_SKU_DETAIL_PATH",
    )
    return ap.parse_args()


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    args = parse_args()

    if args.batch and args.batch.strip():
        import_batch = args.batch.strip()
        _log("INFO", f"使用命令行指定批次号：{import_batch}")
    else:
        import_batch = read_import_batch_from_lock()
        if import_batch:
            _log("INFO", f"从 run_batch.lock 读取批次号：{import_batch}")
        if not import_batch:
            _log("ERROR", "无法获取批次号，请使用 --batch 指定或确保 run_batch.lock 存在")
            return 1

    step_label = {0: "1+2", 1: "1", 2: "2"}[args.step]
    _log("INFO", f"任务：更新 {PROFIT_TABLE} 头程/关税（步骤 {step_label}）")
    _log("INFO", f"批次号：{import_batch} dry_run={bool(args.dry_run)}")
    _log("INFO", f"主数据来源：{args.pricing_source}（pricing_schedule={pricing_schedule!r}）")

    db = get_db_manager(DatabaseConfig())
    conn = db.get_connection()
    try:
        total = 0
        if args.step in (0, 1):
            total += run_shipping_step(
                conn,
                import_batch,
                pricing_source=args.pricing_source,
                excel_file=args.excel_file,
                dry_run=bool(args.dry_run),
            )
        if args.step in (0, 2):
            total += run_tax_step(
                conn,
                import_batch,
                pricing_source=args.pricing_source,
                excel_file=args.excel_file,
                dry_run=bool(args.dry_run),
            )

        if not args.dry_run and total > 0:
            conn.commit()
            _log("INFO", f"全部完成：共更新 {total} 行（按 line_hash）")
        elif args.dry_run:
            _log("INFO", "dry-run：全程未写库")
        else:
            _log("INFO", "无需写库")

        _log("INFO", "=" * 80)
        return 0
    except Exception as e:
        conn.rollback()
        _log("ERROR", f"任务失败：{e}")
        import traceback

        traceback.print_exc()
        return 2
    finally:
        conn.close()


if __name__ == "__main__":
    raise SystemExit(main())
