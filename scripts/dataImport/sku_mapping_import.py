from __future__ import annotations

"""
从订单行抽取 SKU 映射，写入 product_sku_mapping 表。

规则：
- 每个订单行产生 2 条记录：platform 维度 + warehouse 维度
- platform：seller_sku=platform_sku，partner_name=shop_alias，从 platform_shop 同步 market 信息
  - 一票一件组合：seller_sku 保留 '+'（如 E16011001+ESB30007），product_sku=warehouse_sku（如 E58012002）
  - 单品：mapping_type=single，product_sku=warehouse_sku
- warehouse：warehouse_sku=warehouse_sku，partner_name=warehouse_name，始终为 single
- 暂时仅导入 order_offline_type='一票一件' 的订单行到 product_sku_mapping
- warehouse_sku 以 AMZN.GR 开头（不区分大小写）的行不写入 product_sku_mapping
- 一票一件：每行独立映射，mapping_type 恒为 single
  - 普通：platform_sku = warehouse_sku（如 E58012002 = E58012002）
  - 组合：platform_sku 含 '+' 时，整组平台 SKU 映射到 1 个 warehouse_sku
    （如 E16011001+ESB30007 -> E58012002），不按 '+' 拆 component_info
"""

import json
import re
from datetime import datetime
from typing import Any, Callable

import pymysql.cursors

from import_common import insert_ignore_rows, row_subset_for_line_hash, stable_line_hash

MAPPING_TABLE = "product_sku_mapping"
PLATFORM_SHOP_TABLE = "platform_shop"
DEFAULT_SOURCE_TYPE = "Excel"
ALLOW_OFFLINE_TYPE = "一票一件"

# 店铺唯一业务维度；shop_hash = stable_line_hash(下列三列)
SHOP_HASH_KEYS: tuple[str, ...] = ("platform", "platform_site", "shop_name_en")

# line_hash 参与列（与 product_sku_mapping 表结构保持一致）
MAPPING_LINE_HASH_KEYS: tuple[str, ...] = (
    "partner_code",
    "partner_type",
    "partner_name",
    "shop_hash",
    "seller_sku",
    "warehouse_sku",
    "mapping_type",
    "product_sku",
)

MAPPING_INSERT_COLUMNS: tuple[str, ...] = (
    "line_hash",
    "partner_code",
    "partner_type",
    "partner_name",
    "shop_hash",
    "market_region",
    "market_code",
    "seller_sku",
    "warehouse_sku",
    "mapping_type",
    "product_sku",
    "component_info",
    "dev_owner",
    "ops_owner",
    "source_type",
    "is_active",
)

LogFn = Callable[[str, str], None]


def _default_log(level: str, msg: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] [{level}] {msg}")


def _is_blank_str(v: Any) -> bool:
    if v is None:
        return True
    return not str(v).strip()


def _strip_or_empty(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _strip_or_none(v: Any) -> str | None:
    s = _strip_or_empty(v)
    return s or None


def _is_allowed_mapping_row(d: dict[str, Any]) -> bool:
    """暂时仅导入 order_offline_type='一票一件' 的订单行。"""
    return _strip_or_empty(d.get("order_offline_type")) == ALLOW_OFFLINE_TYPE


def _is_amzn_gr_warehouse_sku(v: Any) -> bool:
    """Amazon FBA 仓 SKU 前缀（如 AMZN.GR.U02033010_...），不参与 product_sku_mapping。"""
    s = _strip_or_empty(v)
    return bool(s) and s.upper().startswith("AMZN.GR")


def _is_bundle_order(d: dict[str, Any]) -> bool:
    """
    判断是否为「一票多件多个」类组合品（需写 bundle + component_info）。
    一票一件即使 platform_sku 含 '+'，也按 single 映射到单个 warehouse_sku。
    """
    if _strip_or_empty(d.get("order_offline_type")) == ALLOW_OFFLINE_TYPE:
        return False
    platform_sku = _strip_or_empty(d.get("platform_sku"))
    if "+" in platform_sku:
        return True
    return _strip_or_empty(d.get("order_offline_type")) == "一票多件多个"


def _canonical_platform_sku(platform_sku: str) -> str:
    """规范平台 seller_sku：保留 '+' 组合形态，仅去掉尾部变体/逗号重复。"""
    raw = platform_sku.strip()
    if not raw:
        return ""
    if "," in raw:
        raw = raw.split(",")[0].strip()
    return _strip_variant_suffix(raw)


def _strip_variant_suffix(platform_sku: str) -> str:
    """去掉尾部变体标识（如 ' -1', ' -2'）。"""
    return re.sub(r"\s+-\d+$", "", platform_sku.strip())


def _normalize_bundle_platform_sku(
    platform_sku: str,
) -> tuple[str, int, list[str]]:
    """
    将 platform_sku 规范为「单个组合单元」及其组件列表。

    规则：
    - 含 '+' 时，逗号表示同一组合重复购买，不是额外组件
      例：'E02022001+E22042000,E02022001+E22042000'
        -> seller_sku='E02022001+E22042000', repeat=2, components=['E02022001','E22042000']
    - 不含 '+' 时，逗号仅表示同一 SKU 重复（如 'HL8041-HL8020EU,HL8041-HL8020EU'）
    """
    raw = platform_sku.strip()
    if not raw:
        return "", 1, []

    if "+" in raw:
        segments = [s.strip() for s in raw.split(",") if s.strip()]
        if not segments:
            segments = [raw]
        units = [_strip_variant_suffix(seg) for seg in segments]
        canonical = units[0]
        repeat = len(units) if units and all(u == canonical for u in units) else 1
        components = [p.strip() for p in canonical.split("+") if p.strip()]
        return canonical, repeat, components

    if "," in raw:
        segments = [s.strip() for s in raw.split(",") if s.strip()]
        canonical = _strip_variant_suffix(segments[0]) if segments else raw
        repeat = len(segments) if segments and all(
            _strip_variant_suffix(s) == canonical for s in segments
        ) else 1
        return canonical, repeat, [canonical]

    canonical = _strip_variant_suffix(raw)
    return canonical, 1, [canonical]


def _build_bundle_component_info(
    *,
    unit_components: list[str],
    bundle_repeat: int,
    warehouse_qty_map: dict[str, int],
) -> list[dict[str, Any]]:
    """
    生成单个 seller_sku（组合单元）对应的 component_info。

    优先用 platform_sku 解析出的组件结构（每单元 qty=1）；
    若仓库发货数量与「单元组件数 × 重复次数」一致，则反推每组件单单元用量。
    """
    if not unit_components:
        unit_components = sorted(warehouse_qty_map.keys())

    if not unit_components:
        return []

    expected_total = bundle_repeat * len(unit_components)
    actual_total = sum(warehouse_qty_map.values())
    per_unit_qty = 1

    if (
        bundle_repeat > 0
        and len(unit_components) > 0
        and actual_total > 0
        and actual_total % (bundle_repeat * len(unit_components)) == 0
    ):
        per_unit_qty = actual_total // (bundle_repeat * len(unit_components))
    elif actual_total > 0 and len(unit_components) == 1:
        per_unit_qty = max(1, actual_total // max(bundle_repeat, 1))

    component_info: list[dict[str, Any]] = []
    for sku in unit_components:
        wh_qty = warehouse_qty_map.get(sku)
        if wh_qty is not None and bundle_repeat > 0 and wh_qty % bundle_repeat == 0:
            qty = wh_qty // bundle_repeat
        else:
            qty = per_unit_qty
        component_info.append({"product_sku": sku, "qty": max(qty, 1)})
    return component_info


def compute_shop_hash(d: dict[str, Any]) -> str:
    """根据 platform + platform_site + shop_name_en 计算 shop_hash。"""
    shop_dict = {
        "platform": _strip_or_empty(d.get("platform")),
        "platform_site": _strip_or_empty(d.get("platform_site")),
        "shop_name_en": _strip_or_empty(d.get("shop_name_en")),
    }
    return stable_line_hash(row_subset_for_line_hash(shop_dict, SHOP_HASH_KEYS))


def _lookup_shop_market_info(
    conn, shop_hashes: set[str]
) -> dict[str, tuple[str | None, str | None]]:
    """按 shop_hash 批量查询 platform_shop 的 (market_region, market_code)。"""
    if not shop_hashes:
        return {}

    result: dict[str, tuple[str | None, str | None]] = {}
    items = list(shop_hashes)
    chunk = 500
    cur = conn.cursor(pymysql.cursors.Cursor)
    try:
        for i in range(0, len(items), chunk):
            part = items[i : i + chunk]
            placeholders = ",".join(["%s"] * len(part))
            sql = (
                f"SELECT `shop_hash`, `market_region`, `market_code` FROM `{PLATFORM_SHOP_TABLE}` "
                f"WHERE `shop_hash` IN ({placeholders})"
            )
            cur.execute(sql, part)
            for row in cur.fetchall():
                shop_hash_val = str(row[0]).strip() if row[0] else ""
                market_region_val = (
                    str(row[1]).strip() if row[1] and str(row[1]).strip() else None
                )
                market_code_val = (
                    str(row[2]).strip() if row[2] and str(row[2]).strip() else None
                )
                result[shop_hash_val] = (market_region_val, market_code_val)
    finally:
        cur.close()
    return result


def _build_platform_bundle_row(
    *,
    platform: str,
    shop_hash: str,
    shop_alias: str,
    platform_sku: str,
    component_info: list[dict[str, Any]],
    market_region: str | None,
    market_code: str | None,
    dev_owner: str | None,
    ops_owner: str | None,
    source_type: str,
) -> tuple[Any, ...]:
    """构建平台维度组合品（bundle）映射行。"""
    partner_name = shop_alias or None
    component_json = json.dumps(component_info, ensure_ascii=False)
    bundle_mapping = {
        "partner_code": platform,
        "partner_type": "platform",
        "partner_name": partner_name or "",
        "shop_hash": shop_hash,
        "seller_sku": platform_sku,
        "warehouse_sku": "",
        "mapping_type": "bundle",
        "product_sku": "",  # bundle 时 product_sku 固定空串
    }
    line_hash = stable_line_hash(
        row_subset_for_line_hash(bundle_mapping, MAPPING_LINE_HASH_KEYS)
    )
    return (
        line_hash,
        platform,
        "platform",
        partner_name,
        shop_hash,
        market_region,
        market_code,
        platform_sku,
        "",
        "bundle",
        "",
        component_json,
        dev_owner,
        ops_owner,
        source_type,
        1,
    )


def _build_platform_mapping_row(
    *,
    platform: str,
    shop_hash: str,
    shop_alias: str,
    platform_sku: str,
    product_sku: str,
    market_region: str | None,
    market_code: str | None,
    dev_owner: str | None,
    ops_owner: str | None,
    source_type: str,
) -> tuple[Any, ...]:
    partner_name = shop_alias or None
    platform_mapping = {
        "partner_code": platform,
        "partner_type": "platform",
        "partner_name": partner_name or "",
        "shop_hash": shop_hash,
        "seller_sku": platform_sku,
        "warehouse_sku": "",
        "mapping_type": "single",
        "product_sku": product_sku,
    }
    line_hash = stable_line_hash(
        row_subset_for_line_hash(platform_mapping, MAPPING_LINE_HASH_KEYS)
    )
    return (
        line_hash,
        platform,
        "platform",
        partner_name,
        shop_hash,
        market_region,
        market_code,
        platform_sku,
        "",
        "single",
        product_sku,
        None,
        dev_owner,
        ops_owner,
        source_type,
        1,
    )


def _build_warehouse_mapping_row(
    *,
    platform: str,
    warehouse_name: str,
    wh_sku_str: str,
    product_sku: str,
    dev_owner: str | None,
    ops_owner: str | None,
    source_type: str,
) -> tuple[Any, ...]:
    partner_name = warehouse_name or None
    warehouse_mapping = {
        "partner_code": platform,
        "partner_type": "warehouse",
        "partner_name": partner_name or "",
        "shop_hash": "",
        "seller_sku": "",
        "warehouse_sku": wh_sku_str,
        "mapping_type": "single",
        "product_sku": product_sku,
    }
    line_hash = stable_line_hash(
        row_subset_for_line_hash(warehouse_mapping, MAPPING_LINE_HASH_KEYS)
    )
    return (
        line_hash,
        platform,
        "warehouse",
        partner_name,
        "",
        None,
        None,
        "",
        wh_sku_str,
        "single",
        product_sku,
        None,
        dev_owner,
        ops_owner,
        source_type,
        1,
    )


def upsert_product_sku_mapping(
    conn,
    dicts: list[dict[str, Any]],
    *,
    source_type: str = DEFAULT_SOURCE_TYPE,
    log_fn: LogFn | None = None,
) -> int:
    """
    按平台和仓库两个维度插入 product_sku_mapping。

    当前仅处理 order_offline_type='一票一件'：
    - 每行产生 platform single + warehouse single
    - product_sku 一律取 warehouse_sku（组合平台 SKU 亦映射到单个仓库 SKU）

    返回本批实际新增行数。
    """
    log = log_fn or _default_log
    if not dicts:
        return 0

    eligible = [d for d in dicts if _is_allowed_mapping_row(d)]
    skipped = len(dicts) - len(eligible)
    if skipped:
        log("INFO", f"{MAPPING_TABLE}：仅导入 order_offline_type={ALLOW_OFFLINE_TYPE}；本批跳过 {skipped} 条")
    if not eligible:
        return 0

    # ========== 第一步：收集 shop_hash 并批量查询市场信息 ==========
    shop_hashes_to_query: set[str] = set()
    amzn_gr_skipped = 0
    for d in eligible:
        if _is_blank_str(d.get("warehouse_sku")):
            continue
        if _is_amzn_gr_warehouse_sku(d.get("warehouse_sku")):
            amzn_gr_skipped += 1
            continue
        platform = _strip_or_empty(d.get("platform"))
        platform_sku = _strip_or_empty(d.get("platform_sku"))
        if platform and platform_sku:
            shop_hashes_to_query.add(compute_shop_hash(d))

    shop_market_info = _lookup_shop_market_info(conn, shop_hashes_to_query)

    # ========== 第二步：逐行生成映射（一票一件：不按订单聚合 bundle） ==========
    seen: set[str] = set()
    rows: list[tuple[Any, ...]] = []

    for d in eligible:
        wh_sku = d.get("warehouse_sku")
        if _is_blank_str(wh_sku):
            continue
        if _is_amzn_gr_warehouse_sku(wh_sku):
            continue

        platform = _strip_or_empty(d.get("platform"))
        platform_sku_raw = _strip_or_empty(d.get("platform_sku"))
        wh_sku_str = str(wh_sku).strip()
        product_sku = wh_sku_str
        seller_sku = _canonical_platform_sku(platform_sku_raw) or platform_sku_raw
        shop_alias = _strip_or_empty(d.get("shop_alias"))
        warehouse_name = _strip_or_empty(d.get("warehouse_name"))
        shop_hash = compute_shop_hash(d)
        dev_owner = _strip_or_none(d.get("dev_owner"))
        ops_owner = _strip_or_none(d.get("platform_sku_owner"))

        if platform and seller_sku:
            market_region, market_code = shop_market_info.get(shop_hash, (None, None))
            platform_row = _build_platform_mapping_row(
                platform=platform,
                shop_hash=shop_hash,
                shop_alias=shop_alias,
                platform_sku=seller_sku,
                product_sku=product_sku,
                market_region=market_region,
                market_code=market_code,
                dev_owner=dev_owner,
                ops_owner=ops_owner,
                source_type=source_type,
            )
            if platform_row[0] not in seen:
                seen.add(platform_row[0])
                rows.append(platform_row)

        warehouse_row = _build_warehouse_mapping_row(
            platform=platform,
            warehouse_name=warehouse_name,
            wh_sku_str=wh_sku_str,
            product_sku=product_sku,
            dev_owner=dev_owner,
            ops_owner=ops_owner,
            source_type=source_type,
        )
        if warehouse_row[0] not in seen:
            seen.add(warehouse_row[0])
            rows.append(warehouse_row)

    if not rows:
        if amzn_gr_skipped:
            log(
                "INFO",
                f"{MAPPING_TABLE}：无有效候选（含 AMZN.GR 前缀跳过 {amzn_gr_skipped} 条）",
            )
        else:
            log("INFO", f"{MAPPING_TABLE}：无有效候选（warehouse_sku 全空或 platform_sku 全空）")
        return 0

    if amzn_gr_skipped:
        log("INFO", f"{MAPPING_TABLE}：跳过 AMZN.GR 前缀 warehouse_sku {amzn_gr_skipped} 条")
    log(
        "INFO",
        f"{MAPPING_TABLE}：批次去重后 {len(rows)} 条（{ALLOW_OFFLINE_TYPE}，INSERT IGNORE）",
    )
    n_new = insert_ignore_rows(
        conn,
        table=MAPPING_TABLE,
        columns=list(MAPPING_INSERT_COLUMNS),
        rows=rows,
    )
    log(
        "INFO",
        f"{MAPPING_TABLE}：新增 {n_new} 条（已存在 {len(rows) - n_new} 条跳过）",
    )
    return n_new
