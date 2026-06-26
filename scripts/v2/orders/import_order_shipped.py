from __future__ import annotations

"""
订单统计 Excel -> 表 sales_order_shipped；同批有效行会按店铺写入 platform_shop_config（shop_hash 唯一，INSERT IGNORE 无则新增、有则跳过）。

落库列见 _SHIPPED_MAP；line_hash 默认用「全部业务列 + source_type」。
自定义 line_hash：修改下方 LINE_HASH_KEYS，只列参与哈希的库字段名（须为 LINE_HASH_SOURCE_FIELDS 子集）。
  注意：改规则后历史行的 line_hash 与库内不一致，需按业务重导或清表；字段越少越容易「不同行撞同一 hash」。
  计算时机：行级 dict 在写入前已做类型转换；ref_no 空值会规范为 '' 后参与哈希。
  不参与：line_hash、id、created_at、updated_at、source 文件名。
  算法：excel_common.row_subset_for_line_hash + stable_line_hash（键排序 JSON + SHA-256）。

  日志风格对齐 warehouse-rent/import_provider_4px_detail.py：读取、行数、写入、提交。
  可选：环境变量 ORDER_IMPORT_VERBOSE=1 启动时额外打印全部 line_hash 字段名。

  补漏映射（从已落库发货表反写 product_sku_mapping，与导入时 _upsert_sku_mappings 规则一致）：
    python python/v2/orders/import_order_shipped.py --backfill-mapping-from-shipped
    python python/v2/orders/import_order_shipped.py --backfill-mapping-from-shipped --date-from 2025-01-01 --date-to 2026-12-31
    python python/v2/orders/import_order_shipped.py --backfill-mapping-from-shipped --backfill-chunk-size 5000
"""

import argparse
import os
import sys
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

_ORDERS_DIR = Path(__file__).resolve().parent
_V2_DIR = _ORDERS_DIR.parent
_WR_DIR = _V2_DIR / "warehouse-rent"
for _p in (_ORDERS_DIR, _V2_DIR, _WR_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

import pandas as pd
from logger import get_logger, setup_stdout_utf8

from db import connect, load_db_config
from excel_common import (
    cell_decimal,
    cell_dt,
    cell_int,
    cell_margin_rate,
    cell_str,
    cell_str_or_empty,
    default_order_excel_dir,
    row_subset_for_line_hash,
    stable_line_hash,
    upsert_rows,
)

TABLE = "sales_order_shipped"
MAPPING_TABLE = "product_sku_mapping"
PLATFORM_SHOP_TABLE = "platform_shop_config"
SOURCE_TYPE = "Excel"
_LOG = get_logger("ORDER-SHIPPED")

# 产品SKU映射表参与 line_hash 的列；与 docs/database/002_product_sku_mapping.sql 注释一致
# 改动需同步评估历史行哈希迁移成本
MAPPING_LINE_HASH_KEYS: tuple[str, ...] = (
    "platform",
    "platform_site",
    "shop_name_en",
    "warehouse_name",
    "warehouse_sku",
    "platform_sku",
)

MAPPING_INSERT_COLUMNS: tuple[str, ...] = (
    "line_hash",
    "product_sku",
    "platform",
    "platform_site",
    "shop_name_en",
    "warehouse_name",
    "warehouse_sku",
    "platform_sku",
    "dev_owner",
    "ops_owner",
    "source_type",
)

# 与 docs/database/005_platform_shop.sql 中 uk_psc_shop_hash 语义一致：同平台+站点+英文店名 -> 同一 shop_hash
SHOP_HASH_KEYS: tuple[str, ...] = ("platform", "platform_site", "shop_name_en")
PLATFORM_SHOP_INSERT_COLUMNS: tuple[str, ...] = (
    "shop_hash",
    "shop_name_en",
    "shop_name_cn",
    "platform",
    "platform_site",
    "currency",
    "fx_rate",
    "commission_type",
    "commission_rate",
    "vat_rate_type",
    "vat_rate",
    "ops_owner",
)

# Excel 表头在第 5 行 -> pandas header=4；中文表头 -> 库字段
_SHIPPED_MAP: list[tuple[str, str, str]] = [
    ("平台", "platform", "s64"),
    ("店铺英文名", "shop_name_en", "s128"),
    ("店铺别名", "shop_alias", "s128"),
    ("站点", "platform_site", "s64"),
    ("仓库", "warehouse_name", "s255"),
    ("创建时间", "order_created_at", "dt"),
    ("付款时间", "pay_time", "dt"),
    ("发货时间", "ship_time", "dt"),
    ("审核时间", "audit_time", "dt"),
    ("订单销售状态", "order_sales_status", "s64"),
    ("订单类型", "order_type", "s64"),
    ("订单号", "order_no", "s128"),
    ("参考号", "ref_no", "s128e"),
    ("原平台sku", "platform_sku_orig", "s255"),
    ("平台sku", "platform_sku", "s255"),
    ("仓库sku", "warehouse_sku", "s128"),
    ("仓库sku销量", "warehouse_sku_qty", "int"),
    ("产品名称", "product_name", "s512"),
    ("产品款式", "product_style", "s512"),
    ("品牌", "brand", "s128"),
    ("一级品类", "category_lv1", "s128"),
    ("二级品类", "category_lv2", "s128"),
    ("三级品类", "category_lv3", "s128"),
    ("产品销售状态", "product_sales_status", "s64"),
    ("仓库单号", "warehouse_order_no", "s128"),
    ("服务商单号", "provider_order_no", "s128"),
    ("订单下架类型", "order_offline_type", "s64"),
    ("是否发货", "is_shipped", "s16"),
    ("是否售后", "is_after_sale", "s16"),
    ("跟踪单号", "tracking_no", "s255"),
    ("交易号", "transaction_no", "s255"),
    ("运输方式", "shipping_method", "s255"),
    ("产品重量", "product_weight", "dec"),
    ("发货批次号", "ship_batch_no", "s128"),
    ("订单付款币种", "pay_currency", "s16"),
    ("汇率（转本位币汇率）", "fx_rate_to_base", "dec"),
    ("销售单价（付款币种）", "unit_price_pay", "dec"),
    ("订单总金额（付款币种）", "order_total_pay", "dec"),
    ("订单商品金额（付款币种）", "order_goods_pay", "dec"),
    ("平台运费（付款币种）", "platform_shipping_pay", "dec"),
    ("支付手续费（付款币种）", "payment_fee_pay", "dec"),
    ("平台手续费（付款币种）", "platform_fee_pay", "dec"),
    ("fba费用（付款币种）", "fba_fee_pay", "dec"),
    ("平台补贴费（付款币种）", "platform_subsidy_pay", "dec"),
    ("税费（付款币种）", "tax_pay", "dec"),
    ("其他费用（付款币种）", "other_fee_pay", "dec"),
    ("采购成本（付款币种）", "purchase_cost_pay", "dec"),
    ("采购运费（付款币种）", "purchase_shipping_pay", "dec"),
    ("采购税费（付款币种）", "purchase_tax_pay", "dec"),
    ("头程运费（付款币种）", "first_leg_shipping_pay", "dec"),
    ("头程税费（付款币种）", "first_leg_tax_pay", "dec"),
    ("包材费用（付款币种）", "packaging_fee_pay", "dec"),
    ("派送运费（付款币种）", "delivery_shipping_pay", "dec"),
    ("币种", "base_currency", "s16"),
    ("销售单价", "unit_price_base", "dec"),
    ("订单总金额", "order_total_base", "dec"),
    ("订单商品金额", "order_goods_base", "dec"),
    ("平台运费", "platform_shipping_base", "dec"),
    ("支付手续费", "payment_fee_base", "dec"),
    ("平台手续费", "platform_fee_base", "dec"),
    ("fba费用", "fba_fee_base", "dec"),
    ("平台补贴费", "platform_subsidy_base", "dec"),
    ("税费", "tax_base", "dec"),
    ("其他费用", "other_fee_base", "dec"),
    ("采购成本", "purchase_cost_base", "dec"),
    ("采购运费", "purchase_shipping_base", "dec"),
    ("采购税费", "purchase_tax_base", "dec"),
    ("头程运费", "first_leg_shipping_base", "dec"),
    ("头程税费", "first_leg_tax_base", "dec"),
    ("包材费用", "packaging_fee_base", "dec"),
    ("派送运费", "delivery_shipping_base", "dec"),
    ("总费用", "total_fee_base", "dec"),
    ("总成本", "total_cost_base", "dec"),
    ("毛利", "gross_profit_base", "dec"),
    ("毛利率", "gross_margin_rate", "margin"),
    ("ITEM_ID", "item_id", "s128"),
    ("ASIN", "asin", "s64"),
    ("原销售订单号", "orig_sales_order_no", "s128"),
    ("销售负责人", "sales_owner", "s128"),
    ("采购负责人", "purchase_owner", "s128"),
    ("开发负责人", "dev_owner", "s128"),
    ("附属销售员", "sub_sales", "s128"),
    ("平台sku负责人", "platform_sku_owner", "s128"),
    ("店铺负责人", "shop_owner", "s128"),
    ("买家ID", "buyer_id", "s128"),
    ("买家邮箱", "buyer_email", "s255"),
    ("电话", "phone", "s128"),
    ("买家姓名", "buyer_name", "s255"),
    ("收件人", "consignee", "s255"),
    ("国家", "country", "s128"),
    ("州/省", "state", "s128"),
    ("城市", "city", "s128"),
    ("联系地址", "address_line", "s1024"),
    ("邮编", "postal_code", "s64"),
    ("销售订单号", "sales_order_no", "s128"),
    ("客服备注", "cs_remark", "s2048"),
    ("系统备注", "system_remark", "s2048"),
    ("财务备注", "finance_remark", "s2048"),
    ("shopify支付方式", "shopify_pay_method", "s128"),
    ("shopify支付交易号", "shopify_pay_txn_no", "s255"),
    ("线下费用汇总", "offline_fee_summary", "dec"),
]

# 落库行里会出现的全部业务键（不含 line_hash）；用于校验 LINE_HASH_KEYS
LINE_HASH_SOURCE_FIELDS: tuple[str, ...] = tuple(c for _, c, _ in _SHIPPED_MAP) + ("source_type",)

# 参与 line_hash 的键, [请谨慎修改，否则历史行的 line_hash 与库内不一致，需按业务重导或清表]
# ("order_no", "ref_no", "warehouse_sku", "pay_time", "platform", "source_type")
LINE_HASH_KEYS: tuple[str, ...] = (
    "platform",
    "platform_site",
    "shop_name_en",
    "warehouse_name",
    "order_no",
    "ref_no",
    "warehouse_sku",
    "pay_time",
)

_src_set = frozenset(LINE_HASH_SOURCE_FIELDS)
_bad_keys = tuple(k for k in LINE_HASH_KEYS if k not in _src_set)
if _bad_keys:
    raise ValueError(f"LINE_HASH_KEYS 含未知列（非 LINE_HASH_SOURCE_FIELDS 子集）: {_bad_keys}")

_bad_shop_keys = tuple(k for k in SHOP_HASH_KEYS if k not in _src_set)
if _bad_shop_keys:
    raise ValueError(f"SHOP_HASH_KEYS 含未知列（非 LINE_HASH_SOURCE_FIELDS 子集）: {_bad_shop_keys}")


def _convert(v: Any, kind: str) -> Any:
    if kind == "dt":
        return cell_dt(v)
    if kind == "int":
        return cell_int(v)
    if kind == "dec":
        return cell_decimal(v)
    if kind == "margin":
        return cell_margin_rate(v)
    if kind.startswith("s") and kind.endswith("e"):
        n = int(kind[1:-1])
        return cell_str_or_empty(v, n)
    if kind.startswith("s"):
        n = int(kind[1:])
        return cell_str(v, n)
    return cell_str(v)


def _row_dict(series: pd.Series) -> dict[str, Any]:
    # 先占位全部库字段，避免 Excel 缺列（表头变更/旧模板）时 insert_cols 与 d 键不一致 KeyError
    out: dict[str, Any] = {col: None for _, col, _ in _SHIPPED_MAP}
    for zh, col, kind in _SHIPPED_MAP:
        if zh not in series.index:
            continue
        out[col] = _convert(series[zh], kind)
    out["source_type"] = SOURCE_TYPE
    return out


def _read_shipped_frame(xlsx: Path) -> pd.DataFrame:
    _LOG.warn(f"读取 Excel：{xlsx} sheet=0 header=第5行")
    df = pd.read_excel(xlsx, sheet_name=0, header=4, engine="openpyxl", dtype=object)
    df.columns = [("" if c is None else str(c)).replace("\n", " ").strip() for c in df.columns]
    df = df.dropna(how="all")
    _LOG.info(f"读取完成：行数={len(df)} 列数={len(df.columns)}")
    return df


def _insert_columns() -> list[str]:
    cols = ["line_hash"]
    for _, c, _ in _SHIPPED_MAP:
        cols.append(c)
    cols.append("source_type")
    return cols


def _is_empty_str(v: Any) -> bool:
    """None / 空串 / 只含空白 -> True。"""
    if v is None:
        return True
    if isinstance(v, str) and not v.strip():
        return True
    return False


def _lookup_platform_sku_orig_from_db(
    conn, keys: set[tuple[str, str]]
) -> dict[tuple[str, str], str]:
    """
    根据 (order_no, warehouse_sku) 二元组批量反查
    sales_order_shipped.platform_sku_orig（仅取非空）。

    返回 {(order_no, warehouse_sku): platform_sku_orig}；
    同一 (order_no, warehouse_sku) 理论上唯一，若多行取首个非空。
    分批查询，单次最多 500 个二元组，避免参数过多。
    """
    if not keys:
        return {}
    pairs = [(o, w) for o, w in keys if o and w]
    if not pairs:
        return {}
    result: dict[tuple[str, str], str] = {}
    chunk = 500
    cur = conn.cursor()
    try:
        for i in range(0, len(pairs), chunk):
            part = pairs[i : i + chunk]
            placeholders = ",".join(["(%s,%s)"] * len(part))
            sql = (
                f"SELECT `order_no`, `warehouse_sku`, `platform_sku_orig` FROM `{TABLE}` "
                f"WHERE (`order_no`, `warehouse_sku`) IN ({placeholders}) "
                f"AND `platform_sku_orig` IS NOT NULL AND `platform_sku_orig` <> ''"
            )
            params: list[str] = []
            for o, w in part:
                params.append(o)
                params.append(w)
            cur.execute(sql, params)
            for order_no, wh_sku, sku in cur.fetchall():
                if not order_no or not wh_sku or not sku:
                    continue
                key = (str(order_no).strip(), str(wh_sku).strip())
                if key[0] and key[1] and key not in result:
                    result[key] = str(sku)
    finally:
        cur.close()
    return result


def _fill_platform_sku_orig(conn, dicts: list[dict[str, Any]]) -> int:
    """
    回填规则：platform_sku_orig 为空 + orig_sales_order_no 有值 ->
              用 (orig_sales_order_no, warehouse_sku) 去匹配
              sales_order_shipped.(order_no, warehouse_sku)，
              取对应订单同仓库 SKU 行的 platform_sku_orig 赋值到当前行。

    查找顺序（节省一次 DB 往返）：
      1. 本批次（当前 Excel）已存在 (order_no, warehouse_sku) 与
         (orig_sales_order_no, warehouse_sku) 相同且 platform_sku_orig 非空的行；
      2. 剩余项再批量查 MySQL。
    返回实际回填的行数。
    """
    pending: list[tuple[int, tuple[str, str]]] = []
    for i, d in enumerate(dicts):
        if not _is_empty_str(d.get("platform_sku_orig")):
            continue
        orig = d.get("orig_sales_order_no")
        wh_sku = d.get("warehouse_sku")
        if _is_empty_str(orig) or _is_empty_str(wh_sku):
            continue
        key = (str(orig).strip(), str(wh_sku).strip())
        pending.append((i, key))
    if not pending:
        return 0

    needed: set[tuple[str, str]] = {k for _, k in pending}

    in_batch: dict[tuple[str, str], str] = {}
    for d in dicts:
        on = d.get("order_no")
        wh = d.get("warehouse_sku")
        sku = d.get("platform_sku_orig")
        if _is_empty_str(on) or _is_empty_str(wh) or _is_empty_str(sku):
            continue
        key = (str(on).strip(), str(wh).strip())
        if key in needed and key not in in_batch:
            in_batch[key] = sku

    remain = needed - set(in_batch.keys())
    db_map = _lookup_platform_sku_orig_from_db(conn, remain) if remain else {}

    filled = batch_hits = db_hits = 0
    for i, k in pending:
        if k in in_batch:
            dicts[i]["platform_sku_orig"] = in_batch[k]
            batch_hits += 1
            filled += 1
        elif k in db_map:
            dicts[i]["platform_sku_orig"] = db_map[k]
            db_hits += 1
            filled += 1
    _LOG.info(
        f"补全 platform_sku_orig：候选={len(pending)} 批次内命中={batch_hits} "
        f"数据库命中={db_hits} 未命中={len(pending) - filled}"
    )
    return filled


def _insert_ignore_rows(
    conn,
    *,
    table: str,
    columns: list[str],
    rows: list[tuple[Any, ...]],
    chunk_size: int = 300,
) -> int:
    """
    批量 INSERT IGNORE：遇到唯一键冲突时静默跳过，已有行不会被覆盖。
    用于「查无则新增」语义，避免 ON DUPLICATE KEY UPDATE 把已有非空字段被空值刷掉。
    返回 cursor.rowcount 累计值。
    """
    if not rows:
        return 0
    cols_sql = ", ".join(f"`{c}`" for c in columns)
    placeholders = ", ".join(["%s"] * len(columns))
    sql = f"INSERT IGNORE INTO `{table}` ({cols_sql}) VALUES ({placeholders})"
    affected = 0
    cur = conn.cursor()
    try:
        for i in range(0, len(rows), chunk_size):
            batch = rows[i : i + chunk_size]
            cur.executemany(sql, batch)
            if cur.rowcount and cur.rowcount > 0:
                affected += cur.rowcount
    finally:
        cur.close()
    return affected


def _upsert_sku_mappings(conn, dicts: list[dict[str, Any]]) -> int:
    """
    从订单行中抽取「店铺×仓库×SKU」映射关系，按 line_hash 在批次内去重后
    批量 INSERT IGNORE 到 product_sku_mapping 表（无则新增、有则保留）。

    规则：
      - warehouse_sku 必填；缺失则跳过该行的映射写入
      - product_sku 默认 = warehouse_sku（订单 Excel 不含独立 product_sku 列）
      - line_hash 参与列见 MAPPING_LINE_HASH_KEYS（不含 platform_sku_orig；原平台 SKU 仅存 sales_order_shipped）
      - ops_owner 取自订单 Excel 的「平台sku负责人」列（dict 中名为 platform_sku_owner）
    须在 _fill_platform_sku_orig 之后调用：保证写发货表前 Excel 行已补全 platform_sku_orig（映射表不落该列）。
    """
    if not dicts:
        return 0

    seen: set[str] = set()
    rows: list[tuple[Any, ...]] = []
    for d in dicts:
        wh_sku = d.get("warehouse_sku")
        if _is_empty_str(wh_sku):
            continue
        mapping = {
            "platform": d.get("platform"),
            "platform_site": d.get("platform_site"),
            "shop_name_en": d.get("shop_name_en"),
            "warehouse_name": d.get("warehouse_name"),
            "warehouse_sku": wh_sku,
            "platform_sku": d.get("platform_sku"),
        }
        h_in = row_subset_for_line_hash(mapping, MAPPING_LINE_HASH_KEYS)
        line_hash = stable_line_hash(h_in)
        if line_hash in seen:
            continue
        seen.add(line_hash)
        rows.append(
            (
                line_hash,
                wh_sku,  # product_sku 默认 = warehouse_sku
                mapping["platform"],
                mapping["platform_site"],
                mapping["shop_name_en"],
                mapping["warehouse_name"],
                wh_sku,
                mapping["platform_sku"],
                d.get("dev_owner"),
                d.get("platform_sku_owner"),  # 订单的「平台sku负责人」-> 映射表 ops_owner
                SOURCE_TYPE,
            )
        )

    if not rows:
        _LOG.info("产品SKU映射：无有效候选（warehouse_sku 全空）")
        return 0

    _LOG.info(
        f"产品SKU映射：批次去重后 {len(rows)} 条 -> {MAPPING_TABLE}（INSERT IGNORE，无则新增）"
    )
    n_new = _insert_ignore_rows(
        conn,
        table=MAPPING_TABLE,
        columns=list(MAPPING_INSERT_COLUMNS),
        rows=rows,
    )
    _LOG.info(f"产品SKU映射：新增 {n_new} 条（已存在 {len(rows) - n_new} 条跳过）")
    return n_new


def _insert_ignore_platform_shops(conn, dicts: list[dict[str, Any]]) -> int:
    """
    从订单行抽取店铺：platform、shop_name_en 均非空才写入；
    shop_hash = stable_line_hash(platform, platform_site, shop_name_en)，与 uk_psc_shop_hash 一致。
    批次内按 shop_hash 去重后 INSERT IGNORE（无则新增，已存在整行不覆盖）。

    表要求 NOT NULL 且 Excel 常缺的列：currency 用 base_currency 或 pay_currency 或空串；
    fx_rate 用 fx_rate_to_base，缺省为 0；佣金/VAT 相关缺省为空串与 0；shop_name_cn 用 shop_alias。
    """
    if not dicts:
        return 0

    zero_dec = Decimal("0")
    seen: set[str] = set()
    out_rows: list[tuple[Any, ...]] = []

    for d in dicts:
        if _is_empty_str(d.get("platform")) or _is_empty_str(d.get("shop_name_en")):
            continue
        h_in = row_subset_for_line_hash(d, SHOP_HASH_KEYS)
        shop_hash = stable_line_hash(h_in)
        if shop_hash in seen:
            continue
        seen.add(shop_hash)

        plat = str(d.get("platform")).strip()
        site = d.get("platform_site")
        site_s = "" if site is None else str(site).strip()
        shop_en = str(d.get("shop_name_en")).strip()
        shop_cn = d.get("shop_alias")
        shop_cn_s = "" if shop_cn is None else str(shop_cn).strip()

        cur = d.get("base_currency") or d.get("pay_currency")
        currency_s = "" if cur is None else str(cur).strip()

        fx = d.get("fx_rate_to_base")
        if fx is None:
            fx_v: Any = zero_dec
        else:
            fx_v = fx

        ops = d.get("shop_owner")
        ops_s = "" if ops is None else str(ops).strip()

        out_rows.append(
            (
                shop_hash,
                shop_en,
                shop_cn_s,
                plat,
                site_s,
                currency_s,
                fx_v,
                "",
                zero_dec,
                "",
                zero_dec,
                ops_s,
            )
        )

    if not out_rows:
        _LOG.info(f"平台店铺：无有效候选（缺 platform 或 shop_name_en）")
        return 0

    _LOG.info(
        f"平台店铺：批次去重后 {len(out_rows)} 条 -> {PLATFORM_SHOP_TABLE}（INSERT IGNORE，无则新增）"
    )
    n_new = _insert_ignore_rows(
        conn,
        table=PLATFORM_SHOP_TABLE,
        columns=list(PLATFORM_SHOP_INSERT_COLUMNS),
        rows=out_rows,
    )
    _LOG.info(f"平台店铺：新增 {n_new} 条（已存在 {len(out_rows) - n_new} 条跳过）")
    return n_new


def import_file(conn, xlsx: Path) -> tuple[int, int, int]:
    """
    Returns:
        (executemany 累计行数, 跳过行数, Excel 有效行数)
    """
    df = _read_shipped_frame(xlsx)
    insert_cols = _insert_columns()
    dicts: list[dict[str, Any]] = []
    skipped = 0
    for _, series in df.iterrows():
        d = _row_dict(series)
        order_no = d.get("order_no")
        wh_sku = d.get("warehouse_sku")
        if not order_no or not str(order_no).strip() or not wh_sku or not str(wh_sku).strip():
            skipped += 1
            continue
        if d.get("ref_no") is None:
            d["ref_no"] = ""
        dicts.append(d)
    n_excel = len(df)
    if not dicts:
        _LOG.warn(f"无有效行可写：Excel 行数={n_excel} 跳过={skipped}（缺订单号或仓库sku）")
        return 0, skipped, n_excel

    # 回填发货行 platform_sku_orig（写 sales_order_shipped；映射表不落该列）
    _fill_platform_sku_orig(conn, dicts)

    # 抽取 SKU 映射 -> product_sku_mapping（INSERT IGNORE）
    _upsert_sku_mappings(conn, dicts)

    # 店铺维度写入 platform_shop_config（shop_hash 唯一，INSERT IGNORE）
    _insert_ignore_platform_shops(conn, dicts)

    rows: list[tuple[Any, ...]] = []
    for d in dicts:
        h_in = row_subset_for_line_hash(d, LINE_HASH_KEYS)
        d["line_hash"] = stable_line_hash(h_in)
        rows.append(tuple(d[c] for c in insert_cols))
    _LOG.info(
        f"准备写入 MySQL：表={TABLE} 行数={len(rows)}（跳过={skipped}，line_hash 键数={len(LINE_HASH_KEYS)}）"
    )
    n = upsert_rows(conn, table=TABLE, columns=insert_cols, rows=rows)
    _LOG.info(f"MySQL executemany 完成：批次累计行数={n}")
    return n, skipped, n_excel


# 从发货表反补映射时 SELECT 的列（须含 _fill_platform_sku_orig 与 _upsert_sku_mappings 所需键）
_BACKFILL_MAPPING_SHIPPED_COLS: tuple[str, ...] = (
    "id",
    "order_no",
    "orig_sales_order_no",
    "platform",
    "platform_site",
    "shop_name_en",
    "warehouse_name",
    "warehouse_sku",
    "platform_sku",
    "platform_sku_orig",
    "dev_owner",
    "platform_sku_owner",
)


def _backfill_pay_time_clause(alias: str, date_from: date | None, date_to: date | None) -> tuple[str, list[Any]]:
    col = f"`{alias}`.`pay_time`"
    parts: list[str] = []
    params: list[Any] = []
    if date_from:
        parts.append(f"{col} >= %s")
        params.append(datetime(date_from.year, date_from.month, date_from.day, 0, 0, 0))
    if date_to:
        parts.append(f"{col} <= %s")
        params.append(datetime(date_to.year, date_to.month, date_to.day, 23, 59, 59))
    clause = (" AND " + " AND ".join(parts)) if parts else ""
    return clause, params


def backfill_product_sku_mapping_from_shipped(
    conn,
    *,
    date_from: date | None = None,
    date_to: date | None = None,
    chunk_size: int = 2000,
    commit_per_chunk: bool = True,
) -> tuple[int, int]:
    """
    根据已落库的 sales_order_shipped 行，按与 Excel 导入相同的规则补写 product_sku_mapping
    （先 _fill_platform_sku_orig，再 _upsert_sku_mappings，INSERT IGNORE）。

    适用于：历史数据、映射表被清空、或其它入口只写了发货表未写映射等「漏映射」场景。

    返回 (INSERT IGNORE 累计新增行数, 扫描的发货行数)。
    """
    date_clause, date_params = _backfill_pay_time_clause("s", date_from, date_to)
    cols_sql = ", ".join(f"s.`{c}`" for c in _BACKFILL_MAPPING_SHIPPED_COLS)
    base_sql = (
        f"SELECT {cols_sql} FROM `{TABLE}` s "
        "WHERE TRIM(COALESCE(s.`warehouse_sku`,'')) <> '' "
        "AND s.`order_no` IS NOT NULL AND TRIM(s.`order_no`) <> ''"
        f"{date_clause} "
        "AND s.`id` > %s "
        "ORDER BY s.`id` ASC "
        "LIMIT %s"
    )

    total_shipped = 0
    total_mapping_new = 0
    last_id = 0
    cur = conn.cursor(dictionary=True)
    try:
        while True:
            params = [*date_params, last_id, chunk_size]
            cur.execute(base_sql, params)
            batch = cur.fetchall() or []
            if not batch:
                break
            dicts: list[dict[str, Any]] = []
            for raw in batch:
                d = dict(raw)
                d.pop("id", None)
                dicts.append(d)
            total_shipped += len(dicts)
            last_id = int(raw["id"])  # type: ignore[arg-type]
            _fill_platform_sku_orig(conn, dicts)
            n_new = _upsert_sku_mappings(conn, dicts)
            total_mapping_new += n_new
            if commit_per_chunk:
                conn.commit()
            _LOG.info(
                f"[backfill-mapping] 进度 id<={last_id} 本批发货行={len(dicts)} "
                f"映射新增={n_new} 累计扫描={total_shipped} 映射累计新增={total_mapping_new}"
            )
    finally:
        cur.close()

    if not commit_per_chunk:
        conn.commit()
    _LOG.info(
        f"[backfill-mapping] 完成：扫描 sales_order_shipped={total_shipped} 行，"
        f"product_sku_mapping INSERT IGNORE 新增合计={total_mapping_new}"
    )
    return total_mapping_new, total_shipped


def main() -> int:
    setup_stdout_utf8()
    ap = argparse.ArgumentParser(description="导入 订单统计-*.xlsx -> sales_order_shipped")
    ap.add_argument(
        "--file",
        type=Path,
        default=None,
        help="指定单个 xlsx；默认导入目录下全部「订单统计」开头的文件",
    )
    ap.add_argument(
        "--dir",
        type=Path,
        default=None,
        help=f"Excel 目录，默认 {default_order_excel_dir()}",
    )
    ap.add_argument(
        "--backfill-mapping-from-shipped",
        action="store_true",
        help="不读 Excel：按 sales_order_shipped 已存在行补写 product_sku_mapping（INSERT IGNORE，与导入映射规则一致）",
    )
    ap.add_argument(
        "--date-from",
        type=date.fromisoformat,
        default=None,
        metavar="YYYY-MM-DD",
        help="与 --backfill-mapping-from-shipped 合用：仅 pay_time >= 当天 00:00:00",
    )
    ap.add_argument(
        "--date-to",
        type=date.fromisoformat,
        default=None,
        metavar="YYYY-MM-DD",
        help="与 --backfill-mapping-from-shipped 合用：仅 pay_time <= 当天 23:59:59",
    )
    ap.add_argument(
        "--backfill-chunk-size",
        type=int,
        default=2000,
        metavar="N",
        help="补写映射时每批从发货表拉取行数（默认 2000）",
    )
    args = ap.parse_args()

    if args.backfill_mapping_from_shipped:
        _LOG.info("任务：从 sales_order_shipped 补写 product_sku_mapping（漏映射修复）")
        cfg = load_db_config()
        _LOG.info(f"连接数据库：host={cfg.host} port={cfg.port} database={cfg.database} user={cfg.user}")
        conn = connect(cfg)
        try:
            backfill_product_sku_mapping_from_shipped(
                conn,
                date_from=args.date_from,
                date_to=args.date_to,
                chunk_size=max(1, int(args.backfill_chunk_size)),
            )
            conn.commit()
            _LOG.info("补写映射已提交")
            return 0
        except Exception:
            conn.rollback()
            _LOG.error("补写映射失败，已回滚")
            raise
        finally:
            conn.close()
            _LOG.info("数据库连接已关闭")

    base = args.dir or default_order_excel_dir()
    _LOG.info(f"任务：订单发货明细导入 -> {TABLE}")
    # _LOG.info(
    #     f"line_hash 参与键共 {len(LINE_HASH_KEYS)} 个（LINE_HASH_KEYS，默认可改为自定义子集）；"
    #     "说明见模块文档字符串；算法 row_subset_for_line_hash + stable_line_hash"
    # )
    if os.environ.get("ORDER_IMPORT_VERBOSE") == "1":
        _LOG.info("line_hash 键列表 LINE_HASH_KEYS: " + ", ".join(LINE_HASH_KEYS))
    if not base.is_dir():
        _LOG.error(f"目录不存在: {base}")
        return 2

    files: list[Path]
    if args.file:
        files = [args.file]
    else:
        files = sorted(base.glob("订单统计*.xlsx"))
        if not files:
            files = sorted(base.glob("*订单统计*.xlsx"))
    if not files:
        _LOG.error(f"未找到 xlsx: {base}")
        return 1

    _LOG.info(f"待处理文件数={len(files)} 目录={base}")

    cfg = load_db_config()
    _LOG.info(f"连接数据库：host={cfg.host} port={cfg.port} database={cfg.database} user={cfg.user}")
    conn = connect(cfg)
    total_upsert = 0
    total_skip = 0
    try:
        for p in files:
            if not p.is_file():
                _LOG.warn(f"跳过（非文件）: {p}")
                continue
            n, skipped, n_excel = import_file(conn, p)
            conn.commit()
            _LOG.info(f"已提交：{p.name} Excel行={n_excel} 写入UPSERT累计={n} 跳过={skipped}")
            total_upsert += n
            total_skip += skipped
        _LOG.info(f"全部完成：写入UPSERT累计行数={total_upsert} 总跳过={total_skip} 文件数={len(files)}")
        return 0
    except Exception:
        conn.rollback()
        _LOG.error("已回滚事务")
        raise
    finally:
        conn.close()
        _LOG.info("数据库连接已关闭")


if __name__ == "__main__":
    raise SystemExit(main())
