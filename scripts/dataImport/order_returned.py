from __future__ import annotations

"""
从鸿羽仓「二次上架明细」Excel 写入 MySQL 表 sales_order_returned。

路径：config.path_config.SECOND_RELISTING_PATH + MODE_PATTERN + 日期子目录
  每天 -> .../每天/鸿羽仓二次上架明细/6.9/鸿羽-二次上架明细-6.1-6.9.xls
  （日期子目录为 M.D，由 DATE_PATH 的 YYYY-MM-DD 转换，如 2026-06-09 -> 6.9）

工作表：ReturnOrders（否则第一个 sheet），表头第 1 行。
文件匹配：*二次上架明细-*.xls / *.xlsx

line_hash：LINE_HASH_KEYS 子集经 stable_line_hash（键排序 JSON + SHA-256）。

report_hash：未指定 --import-batch 时，默认取 run_batch.lock 的 import_batch。

check_lock：库内已存在且 check_lock=1 的行，重复导入时保留
platform / shop_name_en / shop_alias / platform_site / check_lock 五列不变，其余字段照常 UPSERT。

发货表回填（sales_order_shipped -> sales_order_returned，默认开启，--no-shipped-enrich 可关闭）：

匹配规则（按顺序尝试，命中即用）：
  1. orig_sales_order_no = sales_order_shipped.order_no
  2. orig_order_no = sales_order_shipped.provider_order_no
  3. orig_order_no = sales_order_shipped.order_no

多条发货记录时：优先 warehouse_sku 相同；否则取 ship_time 最新的一条。

写入策略：
  - 多数字段：仅当退件行对应列为空时写入
  - shop_alias：只要匹配到发货行且发货侧有值，一律覆盖（含 Excel「卖家店铺」已填的情况）

回填字段一览（退件列 <- 发货列）：
  platform           <- platform
  shop_name_en       <- shop_name_en
  shop_alias         <- shop_alias          （强制覆盖）
  platform_site      <- platform_site
  warehouse_name     <- warehouse_name
  platform_sku       <- platform_sku
  product_name       <- product_name
  platform_sku_owner <- platform_sku_owner
  orig_tracking_no   <- tracking_no
  sales_owner        <- sales_owner

市场字段回填（platform_shop -> sales_order_returned，导入全部文件提交后执行）：

关联键：platform + platform_site + shop_name_en（与 platform_shop.shop_hash 业务维度一致）。
仅更新 market_region 为空的行（NULL 或 TRIM 后为空字符串）。

用法：
  cd d:\\py-project\\report
  python scripts\\dataImport\\order_returned.py
  python scripts\\dataImport\\order_returned.py --date 2026-06-09
  python scripts\\dataImport\\order_returned.py --file "\\\\Betohow\\...\\鸿羽-二次上架明细-6.1-6.9.xls"
"""

import argparse
import re
import sys
import time
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

import pandas as pd
import pymysql.cursors

_REPORT_ROOT = Path(__file__).resolve().parents[2]
_DATA_IMPORT_DIR = Path(__file__).resolve().parent
for _p in (_REPORT_ROOT, _DATA_IMPORT_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from batch_lock import read_import_batch_from_lock  # noqa: E402
from config.path_config import DATE_PATH, MODE_PATTERN, SECOND_RELISTING_PATH  # noqa: E402
from database.db_connection import DatabaseConfig, get_db_manager  # noqa: E402
from scripts.console_log import (  # noqa: E402
    colorize,
    enable_windows_ansi,
    log as _log,
    log_success as _log_success,
)
from import_common import (  # noqa: E402
    cell_decimal,
    cell_dt,
    cell_str,
    cell_str_or_empty,
    row_subset_for_line_hash,
    stable_line_hash,
)

TABLE = "sales_order_returned"
SHIPPED_TABLE = "sales_order_shipped"
PLATFORM_SHOP_TABLE = "platform_shop"
_COLLATE = "utf8mb4_unicode_ci"
SOURCE_TYPE = "Excel"
Decimal0 = Decimal("0")
_WAREHOUSE_SKU_PRODUCT_PREFIX = "900008-"
_WH_BRACKET_RE = re.compile(r"^\s*\[([^\]]+)\]\s*(.*)$")

# 从发货表查询的列（用于匹配 warehouse_sku、按 ship_time 择优）
_SHIPPED_FETCH_COLS: tuple[str, ...] = (
    "id",
    "ship_time",
    "order_no",
    "provider_order_no",
    "ref_no",
    "platform",
    "shop_name_en",
    "shop_alias",
    "platform_site",
    "warehouse_name",
    "warehouse_sku",
    "platform_sku",
    "product_name",
    "platform_sku_owner",
    "sales_owner",
    "tracking_no",
)

# 退件列 <- sales_order_shipped 列（默认仅退件侧为空时写入；见 _SHIPPED_FILL_ALWAYS_COLS）
_SHIPPED_FILL_MAP: tuple[tuple[str, str, int | None], ...] = (
    ("platform", "platform", 64),
    ("shop_name_en", "shop_name_en", 128),
    ("shop_alias", "shop_alias", 128),
    ("platform_site", "platform_site", 64),
    ("warehouse_name", "warehouse_name", 255),
    ("platform_sku", "platform_sku", 255),
    ("product_name", "product_name", 512),
    ("platform_sku_owner", "platform_sku_owner", 128),
    ("orig_tracking_no", "tracking_no", 255),
    ("sales_owner", "sales_owner", 128),
)

# 匹配到发货行后始终覆盖（不因退件行已有值而跳过）
_SHIPPED_FILL_ALWAYS_COLS: frozenset[str] = frozenset({"shop_alias"})

# 逻辑键 -> Excel 列名（鸿羽 ReturnOrders 模板）
_HEADER: dict[str, str] = {
    "sku": "SKU",
    "return_doc": "退件号",
    "order_no": "订单号",
    "ref_no": "参考号",
    "order_ref": "订单参考号",
    "claim_no": "认领单号",
    "warehouse": "仓库",
    "status": "状态",
    "return_type": "退件类型",
    "label_svc": "Label服务",
    "track_no": "跟踪号",
    "qty": "数量",
    "disposition": "处理方式",
    "qty_received": "实收数量",
    "qty_good": "良品",
    "qty_bad": "不良品",
    "zone_transfer": "转运区",
    "reason": "退件原因",
    "note": "退件说明",
    "created_by": "创建人",
    "created_at": "创建时间",
    "audit_at": "审核时间",
    "done_at": "完成时间",
    "fee_rmb": "退件费用(RMB)",
    "shop_seller": "卖家店铺",
    "return_rec_client": "退货记录（客户端）",
    "return_rec_wh": "退货记录（仓储端）",
}

# check_lock=1 时 ON DUPLICATE KEY UPDATE 保留下列列（不写入新值）
_CHECK_LOCK_PRESERVE_COLS: frozenset[str] = frozenset(
    {
        "platform",
        "shop_name_en",
        "shop_alias",
        "platform_site",
        "check_lock",
    }
) 
# 注意：这里不可随意修改，否则历史行的 line_hash 与库内不一致，需按业务重导或清表
# 注意：这里不可随意修改，否则历史行的 line_hash 与库内不一致，需按业务重导或清表
LINE_HASH_KEYS: tuple[str, ...] = (
    "return_doc_no",
    "warehouse_sku",
    "return_tracking_no",
    "orig_order_no",
    "orig_ref_no",
)
# ================================================================================ 

_INSERT_COLUMNS: tuple[str, ...] = (
    "line_hash",
    "platform",
    "shop_name_en",
    "shop_alias",
    "platform_site",
    "warehouse_name",
    "warehouse_code",
    "return_doc_no",
    "receiving_no",
    "rma_no",
    "orig_order_no",
    "orig_ref_no",
    "orig_sales_order_no",
    "orig_tracking_no",
    "return_type",
    "return_status",
    "disposition",
    "platform_sku",
    "warehouse_sku",
    "product_sku",
    "product_name",
    "return_qty",
    "received_qty",
    "putaway_qty",
    "return_tracking_no",
    "carrier",
    "shipping_method",
    "apply_at",
    "buyer_ship_at",
    "received_at",
    "putaway_at",
    "inspected_at",
    "length_cm",
    "width_cm",
    "height_cm",
    "volume_m3",
    "weight_kg",
    "currency_code",
    "handling_fee",
    "declared_value",
    "sales_owner",
    "platform_sku_owner",
    "cs_remark",
    "warehouse_remark",
    "profit_calc_node",
    "source_type",
)


def _insert_columns(*, with_report_hash: bool = False) -> list[str]:
    cols = list(_INSERT_COLUMNS)
    if with_report_hash:
        cols.insert(cols.index("source_type"), "report_hash")
    return cols


def _is_blank_for_fill(v: Any) -> bool:
    if v is None:
        return True
    if isinstance(v, str) and not v.strip():
        return True
    return False


def _pick_shipped_row(
    candidates: list[dict[str, Any]],
    warehouse_sku: str | None,
) -> tuple[dict[str, Any] | None, str]:
    """优先按 warehouse_sku 命中，否则取 ship_time 最新的一条。"""
    if not candidates:
        return None, "none"
    dedup: dict[Any, dict[str, Any]] = {}
    for r in candidates:
        rid = r.get("id")
        dedup[rid if rid is not None else id(r)] = r
    uniq = list(dedup.values())

    if warehouse_sku:
        for r in uniq:
            if (r.get("warehouse_sku") or "") == warehouse_sku:
                return r, "sku"

    def _ts(r: dict[str, Any]) -> float:
        st = r.get("ship_time")
        if isinstance(st, datetime):
            return st.timestamp()
        if st is not None and hasattr(st, "timestamp"):
            try:
                return float(st.timestamp())
            except Exception:
                return 0.0
        return 0.0

    def _rid(r: dict[str, Any]) -> int:
        try:
            return int(r.get("id") or 0)
        except (TypeError, ValueError):
            return 0

    return max(uniq, key=lambda r: (_ts(r), _rid(r))), "latest"


def _fetch_shipped_by_field(
    conn: Any,
    values: frozenset[str],
    *,
    field: str,
    chunk_size: int = 200,
) -> tuple[dict[str, list[dict[str, Any]]], dict[str, Any]]:
    """按 sales_order_shipped 指定列批量查询，返回 field 值 -> 发货行列表。"""
    out: dict[str, list[dict[str, Any]]] = defaultdict(list)
    meta: dict[str, Any] = {"n_chunks": 0, "n_rows_fetched": 0, "elapsed_ms": 0.0}
    if not values:
        return out, meta

    cols_sql = ", ".join(f"`{c}`" for c in _SHIPPED_FETCH_COLS)
    wanted_list = sorted(values)
    cur = conn.cursor(pymysql.cursors.DictCursor)
    t0 = time.perf_counter()
    seen_ids: dict[str, set[Any]] = {}
    try:
        for i in range(0, len(wanted_list), chunk_size):
            chunk = wanted_list[i : i + chunk_size]
            if not chunk:
                break
            ph = ", ".join(["%s"] * len(chunk))
            sql = f"SELECT {cols_sql} FROM `{SHIPPED_TABLE}` WHERE `{field}` IN ({ph})"
            cur.execute(sql, tuple(chunk))
            batch = cur.fetchall() or []
            meta["n_rows_fetched"] += len(batch)
            meta["n_chunks"] += 1
            wanted_f = frozenset(chunk)
            for raw in batch:
                row = {k: raw.get(k) for k in _SHIPPED_FETCH_COLS}
                raw_key = row.get(field)
                if raw_key is None:
                    continue
                key = str(raw_key).strip()
                if not key or key not in wanted_f:
                    continue
                rid = row.get("id")
                bucket = seen_ids.setdefault(key, set())
                dedupe_key = rid if rid is not None else id(row)
                if dedupe_key in bucket:
                    continue
                bucket.add(dedupe_key)
                out[key].append(row)
    finally:
        cur.close()
    meta["elapsed_ms"] = (time.perf_counter() - t0) * 1000.0
    return out, meta


def _enrich_row_from_shipped(
    d: dict[str, Any],
    shipped: dict[str, Any] | None,
    cell_counts: dict[str, int],
) -> bool:
    if not shipped:
        return False
    changed = False
    for ret_col, ship_col, maxlen in _SHIPPED_FILL_MAP:
        val = cell_str(shipped.get(ship_col), maxlen)
        if val is None:
            continue
        force = ret_col in _SHIPPED_FILL_ALWAYS_COLS
        if not force and not _is_blank_for_fill(d.get(ret_col)):
            continue
        d[ret_col] = val
        cell_counts[ret_col] = cell_counts.get(ret_col, 0) + 1
        changed = True
    return changed


def _enrich_rows_from_shipped(conn: Any, rows: list[dict[str, Any]]) -> dict[str, Any]:
    """
    用 sales_order_shipped 回填退件行（匹配规则见模块文档）。
    shop_alias 匹配到发货行后强制覆盖，其余字段仅填空。
    """
    sales_order_keys = frozenset(
        str(d["orig_sales_order_no"]).strip()
        for d in rows
        if d.get("orig_sales_order_no") and str(d["orig_sales_order_no"]).strip()
    )
    orig_order_keys = frozenset(
        str(d["orig_order_no"]).strip()
        for d in rows
        if d.get("orig_order_no") and str(d["orig_order_no"]).strip()
    )
    _log(
        "INFO",
        f"【发货回填】开始：规则1 order_no<-orig_sales_order_no（{len(sales_order_keys)} 键），"
        f"规则2 provider_order_no<-orig_order_no（{len(orig_order_keys)} 键），"
        f"规则3 order_no<-orig_order_no（{len(orig_order_keys)} 键），退件行={len(rows)}",
    )

    # order_no 查询既要覆盖 orig_sales_order_no，也要覆盖 orig_order_no（规则3）
    order_no_keys = frozenset(set(sales_order_keys) | set(orig_order_keys))
    shipped_by_order_no, meta1 = _fetch_shipped_by_field(conn, order_no_keys, field="order_no")
    shipped_by_provider_no, meta2 = _fetch_shipped_by_field(
        conn, orig_order_keys, field="provider_order_no"
    )
    n_hit_rule1 = sum(1 for k in sales_order_keys if shipped_by_order_no.get(k))
    n_hit_rule2 = sum(1 for k in orig_order_keys if shipped_by_provider_no.get(k))
    n_hit_rule3 = sum(1 for k in orig_order_keys if shipped_by_order_no.get(k))

    _log(
        "INFO",
        f"【发货回填】查询完成：规则1 批次数={meta1['n_chunks']} 发货行={meta1['n_rows_fetched']} "
        f"命中={n_hit_rule1}/{len(sales_order_keys)} 耗时={meta1['elapsed_ms']:.1f}ms；"
        f"规则2 批次数={meta2['n_chunks']} 发货行={meta2['n_rows_fetched']} "
        f"命中={n_hit_rule2}/{len(orig_order_keys)} 耗时={meta2['elapsed_ms']:.1f}ms；"
        f"规则3 命中={n_hit_rule3}/{len(orig_order_keys)}（复用规则1的 order_no 查询结果）",
    )

    cell_counts: dict[str, int] = {}
    n_rows_touched = 0
    n_rows_no_ship = 0
    n_pick_sku = 0
    n_pick_latest = 0
    n_match_rule1 = 0
    n_match_rule2 = 0
    n_match_rule3 = 0

    for d in rows:
        candidates: list[dict[str, Any]] = []
        match_rule = ""

        sales_key = str(d.get("orig_sales_order_no") or "").strip()
        if sales_key:
            candidates = shipped_by_order_no.get(sales_key) or []
            if candidates:
                match_rule = "rule1"

        if not candidates:
            orig_key = str(d.get("orig_order_no") or "").strip()
            if orig_key:
                candidates = shipped_by_provider_no.get(orig_key) or []
                if candidates:
                    match_rule = "rule2"

        if not candidates:
            orig_key = str(d.get("orig_order_no") or "").strip()
            if orig_key:
                candidates = shipped_by_order_no.get(orig_key) or []
                if candidates:
                    match_rule = "rule3"

        if not candidates:
            n_rows_no_ship += 1
            continue

        ship, pick_reason = _pick_shipped_row(candidates, d.get("warehouse_sku"))
        if ship is None:
            n_rows_no_ship += 1
            continue

        if match_rule == "rule1":
            n_match_rule1 += 1
        elif match_rule == "rule2":
            n_match_rule2 += 1
        elif match_rule == "rule3":
            n_match_rule3 += 1
        if pick_reason == "sku":
            n_pick_sku += 1
        elif pick_reason == "latest":
            n_pick_latest += 1
        if _enrich_row_from_shipped(d, ship, cell_counts):
            n_rows_touched += 1

    col_summary = ", ".join(f"{k}={v}" for k, v in sorted(cell_counts.items()) if v > 0)
    _log(
        "INFO",
        f"【发货回填】完成：补全退件行={n_rows_touched}/{len(rows)}，"
        f"规则1命中={n_match_rule1}，规则2命中={n_match_rule2}，规则3命中={n_match_rule3}，"
        f"SKU命中={n_pick_sku}，最近发货={n_pick_latest}，无发货记录={n_rows_no_ship}"
        + (f"，字段写入：{col_summary}" if col_summary else ""),
    )
    return {
        "touched_rows": n_rows_touched,
        "keys_hit_rule1": n_hit_rule1,
        "keys_hit_rule2": n_hit_rule2,
        "keys_hit_rule3": n_hit_rule3,
        "match_rule1": n_match_rule1,
        "match_rule2": n_match_rule2,
        "match_rule3": n_match_rule3,
        "cell_counts": cell_counts,
    }


def _xc(series: pd.Series, logical: str) -> Any:
    zh = _HEADER[logical]
    return series.get(zh) if zh in series.index else None


def _excel_engine(path: Path) -> str:
    return "xlrd" if path.suffix.lower() == ".xls" else "openpyxl"


def _product_sku_from_warehouse_sku(wh: Any) -> str | None:
    s = cell_str(wh, 128)
    if not s:
        return None
    if s.startswith(_WAREHOUSE_SKU_PRODUCT_PREFIX):
        tail = s[len(_WAREHOUSE_SKU_PRODUCT_PREFIX) :].strip()
        return cell_str(tail, 128) if tail else None
    return None


def _parse_warehouse(raw: Any) -> tuple[str | None, str | None]:
    s = cell_str(raw, 255)
    if not s:
        return None, None
    m = _WH_BRACKET_RE.match(s)
    if not m:
        return s, None
    code = (m.group(1) or "").strip() or None
    tail = (m.group(2) or "").strip()
    return tail or code, code


def _combine_remark(*parts: Any, max_len: int = 2048) -> str | None:
    texts: list[str] = []
    for p in parts:
        s = cell_str(p, max_len)
        if s:
            texts.append(s)
    if not texts:
        return None
    out = " | ".join(texts)
    return out[:max_len] if len(out) > max_len else out


def _orig_order_and_ref(series: pd.Series) -> tuple[str | None, str]:
    order_no = cell_str(_xc(series, "order_no"), 128)
    ref = cell_str(_xc(series, "ref_no"), 128)
    order_ref = cell_str(_xc(series, "order_ref"), 128)
    rma_doc = cell_str(_xc(series, "return_doc"), 128)
    if order_no:
        return order_no, cell_str_or_empty(ref, 128)
    if ref:
        return ref, cell_str_or_empty(order_ref, 128)
    if order_ref:
        return order_ref, ""
    if rma_doc:
        return rma_doc, ""
    return None, ""


def _build_row(series: pd.Series) -> dict[str, Any] | None:
    orig_order_no, orig_ref_no = _orig_order_and_ref(series)
    sku = cell_str(_xc(series, "sku"), 128)
    if not sku or not orig_order_no:
        return None

    wh_name, wh_code = _parse_warehouse(_xc(series, "warehouse"))
    fee = cell_decimal(_xc(series, "fee_rmb"))
    bad = cell_str(_xc(series, "qty_bad"), 32)
    zone = cell_str(_xc(series, "zone_transfer"), 64)

    return {
        "platform": None,
        "shop_name_en": None,
        "shop_alias": cell_str(_xc(series, "shop_seller"), 128),
        "platform_site": None,
        "warehouse_name": wh_name,
        "warehouse_code": wh_code,
        "return_doc_no": cell_str(_xc(series, "return_doc"), 128),
        "receiving_no": None,
        "rma_no": cell_str(_xc(series, "claim_no"), 128),
        "orig_order_no": orig_order_no,
        "orig_ref_no": orig_ref_no,
        "orig_sales_order_no": cell_str(_xc(series, "order_ref"), 128),
        "orig_tracking_no": None,
        "return_type": cell_str(_xc(series, "return_type"), 64),
        "return_status": cell_str(_xc(series, "status"), 64),
        "disposition": cell_str(_xc(series, "disposition"), 64),
        "platform_sku": None,
        "warehouse_sku": sku,
        "product_sku": _product_sku_from_warehouse_sku(sku),
        "product_name": None,
        "return_qty": cell_decimal(_xc(series, "qty")) or Decimal0,
        "received_qty": cell_decimal(_xc(series, "qty_received")),
        "putaway_qty": cell_decimal(_xc(series, "qty_good")),
        "return_tracking_no": cell_str(_xc(series, "track_no"), 255),
        "carrier": None,
        "shipping_method": cell_str(_xc(series, "label_svc"), 128),
        "apply_at": cell_dt(_xc(series, "created_at")),
        "buyer_ship_at": None,
        "received_at": None,
        "putaway_at": cell_dt(_xc(series, "done_at")),
        "inspected_at": cell_dt(_xc(series, "audit_at")),
        "length_cm": None,
        "width_cm": None,
        "height_cm": None,
        "volume_m3": None,
        "weight_kg": None,
        "currency_code": "CNY" if fee is not None else None,
        "handling_fee": fee,
        "declared_value": None,
        "sales_owner": cell_str(_xc(series, "created_by"), 128),
        "platform_sku_owner": None,
        "cs_remark": _combine_remark(_xc(series, "reason"), _xc(series, "note")),
        "warehouse_remark": _combine_remark(
            f"不良品:{bad}" if bad is not None else None,
            f"转运区:{zone}" if zone is not None else None,
            _xc(series, "return_rec_client"),
            _xc(series, "return_rec_wh"),
        ),
        "profit_calc_node": None,
        "source_type": SOURCE_TYPE,
    }


def _collate_trim(expr: str) -> str:
    return f"TRIM({expr}) COLLATE {_COLLATE}"


def _collate_trim_ifnull(expr: str, default: str = "''") -> str:
    return f"TRIM(IFNULL({expr}, {default})) COLLATE {_COLLATE}"


def _shop_join_sql(*, returned_alias: str, shop_alias: str) -> str:
    return f"""
    {_collate_trim(f'`{returned_alias}`.`platform`')} = {_collate_trim(f'`{shop_alias}`.`platform`')}
    AND {_collate_trim_ifnull(f'`{returned_alias}`.`platform_site`')} = {_collate_trim_ifnull(f'`{shop_alias}`.`platform_site`')}
    AND {_collate_trim_ifnull(f'`{returned_alias}`.`shop_name_en`')} = {_collate_trim_ifnull(f'`{shop_alias}`.`shop_name_en`')}
  """.strip()


def _market_region_empty_sql(*, alias: str) -> str:
    return f"(`{alias}`.`market_region` IS NULL OR TRIM(`{alias}`.`market_region`) = '')"


def _fill_market_from_platform_shop(conn: Any) -> dict[str, int]:
    """
    根据 platform_shop 回填 sales_order_returned.market_region / market_code。
    仅处理 market_region 为空的行。
    """
    shop_join = _shop_join_sql(returned_alias="r", shop_alias="ps")
    empty_region = _market_region_empty_sql(alias="r")

    stats_sql = {
        "pending": f"SELECT COUNT(*) FROM `{TABLE}` AS r WHERE {empty_region}",
        "matched": (
            f"SELECT COUNT(*) FROM `{TABLE}` AS r "
            f"INNER JOIN `{PLATFORM_SHOP_TABLE}` AS ps ON {shop_join} "
            f"WHERE {empty_region}"
        ),
        "unmatched": (
            f"SELECT COUNT(*) FROM `{TABLE}` AS r "
            f"LEFT JOIN `{PLATFORM_SHOP_TABLE}` AS ps ON {shop_join} "
            f"WHERE {empty_region} AND ps.`id` IS NULL"
        ),
    }
    update_sql = f"""
        UPDATE `{TABLE}` AS r
        INNER JOIN `{PLATFORM_SHOP_TABLE}` AS ps ON {shop_join}
        SET
          r.`market_region` = ps.`market_region`,
          r.`market_code` = ps.`market_code`
        WHERE {empty_region}
    """

    cur = conn.cursor(pymysql.cursors.Cursor)
    stats: dict[str, int] = {}
    try:
        for key, sql in stats_sql.items():
            cur.execute(sql)
            row = cur.fetchone()
            stats[key] = int(row[0] or 0) if row else 0

        _log(
            "INFO",
            f"【市场回填】开始：market_region 为空 {stats['pending']} 条，"
            f"可匹配 platform_shop {stats['matched']} 条，无匹配 {stats['unmatched']} 条",
        )
        if stats["matched"] == 0:
            _log("INFO", "【市场回填】跳过：无可更新行")
            stats["updated"] = 0
            return stats

        cur.execute(update_sql)
        stats["updated"] = int(cur.rowcount or 0)
    finally:
        cur.close()

    _log(
        "INFO",
        f"【市场回填】完成：更新 {stats['updated']} 条"
        f"（待填 {stats['pending']}，匹配 {stats['matched']}，无店铺 {stats['unmatched']}）",
    )
    return stats


def _upsert_returned_rows(
    conn,
    *,
    columns: list[str],
    rows: list[tuple[Any, ...]],
    chunk_size: int = 300,
) -> int:
    """
    UPSERT sales_order_returned。
    若库内行 check_lock=1，则 platform/shop_name_en/shop_alias/platform_site/check_lock 保持原值。
    """
    if not rows:
        return 0

    cols_sql = ", ".join(f"`{c}`" for c in columns)
    placeholders = ", ".join(["%s"] * len(columns))
    update_parts: list[str] = []
    for col in columns:
        if col == "line_hash":
            continue
        if col in _CHECK_LOCK_PRESERVE_COLS:
            update_parts.append(
                f"`{col}`=IF(COALESCE(`check_lock`, 0) = 1, `{col}`, VALUES(`{col}`))"
            )
        else:
            update_parts.append(f"`{col}`=VALUES(`{col}`)")
    updates = ", ".join(update_parts)
    sql = (
        f"INSERT INTO `{TABLE}` ({cols_sql}) VALUES ({placeholders}) "
        f"ON DUPLICATE KEY UPDATE {updates}"
    )

    cur = conn.cursor()
    n = 0
    buf: list[tuple[Any, ...]] = []
    try:
        for row in rows:
            buf.append(row)
            if len(buf) >= chunk_size:
                cur.executemany(sql, buf)
                n += len(buf)
                buf.clear()
        if buf:
            cur.executemany(sql, buf)
            n += len(buf)
    finally:
        cur.close()
    return n


def _lookup_locked_line_hashes(conn: Any, line_hashes: set[str]) -> set[str]:
    """返回库内 check_lock=1 的 line_hash 集合。"""
    if not line_hashes:
        return set()
    locked: set[str] = set()
    items = sorted(line_hashes)
    chunk = 200
    cur = conn.cursor()
    try:
        for i in range(0, len(items), chunk):
            part = items[i : i + chunk]
            ph = ", ".join(["%s"] * len(part))
            sql = (
                f"SELECT `line_hash` FROM `{TABLE}` "
                f"WHERE `line_hash` IN ({ph}) AND COALESCE(`check_lock`, 0) = 1"
            )
            cur.execute(sql, tuple(part))
            for (line_hash,) in cur.fetchall() or []:
                locked.add(str(line_hash).strip())
    finally:
        cur.close()
    return locked


def _read_returned_frame(path: Path) -> pd.DataFrame:
    engine = _excel_engine(path)
    last: Exception | None = None
    df: pd.DataFrame | None = None
    sheet_used: str | int | None = None
    for sn in ("ReturnOrders", 0):
        try:
            df = pd.read_excel(path, sheet_name=sn, header=0, engine=engine, dtype=object)
            sheet_used = sn
            break
        except Exception as e:
            last = e
    if df is None:
        raise RuntimeError(f"无法读取工作表（需 ReturnOrders 或第一个 sheet）：{path}") from last

    _log("INFO", f"读取 Excel：{path.name} engine={engine} sheet={sheet_used!r}（表头第 1 行）")
    df.columns = [("" if c is None else str(c)).replace("\n", " ").strip() for c in df.columns]
    df = df.dropna(how="all")

    if _HEADER["sku"] not in df.columns:
        raise RuntimeError(f"Excel 缺少必需列「{_HEADER['sku']}」：{path}")

    lineage = {_HEADER[k] for k in ("order_no", "ref_no", "order_ref", "return_doc")}
    if not (lineage & set(df.columns)):
        raise RuntimeError(f"Excel 缺少原订单号解析列（至少其一）：{sorted(lineage)}")

    _log("INFO", f"读取完成：行数={len(df)} 列数={len(df.columns)}")
    return df


def import_file(
    conn,
    path: Path,
    *,
    enrich_from_shipped: bool = True,
    import_batch: str | None = None,
) -> tuple[int, int, int]:
    """返回 (UPSERT 行数, 跳过行数, Excel 总行数)。"""
    df = _read_returned_frame(path)
    dicts: list[dict[str, Any]] = []
    skipped = 0

    for _, series in df.iterrows():
        d = _build_row(series)
        if d is None:
            skipped += 1
            continue
        if d.get("orig_ref_no") is None:
            d["orig_ref_no"] = ""
        h_in = row_subset_for_line_hash(d, LINE_HASH_KEYS)
        d["line_hash"] = stable_line_hash(h_in)
        dicts.append(d)

    if not dicts:
        _log("WARN", f"无有效行：Excel 行数={len(df)} 跳过={skipped}")
        return 0, skipped, len(df)

    locked_hashes = _lookup_locked_line_hashes(conn, {d["line_hash"] for d in dicts})
    n_locked_partial = sum(1 for d in dicts if d["line_hash"] in locked_hashes)
    if n_locked_partial:
        _log(
            "INFO",
            f"check_lock=1：{n_locked_partial} 行将部分更新"
            f"（保留 platform/shop_name_en/shop_alias/platform_site/check_lock）",
        )

    if enrich_from_shipped:
        _enrich_rows_from_shipped(conn, dicts)
    else:
        _log("INFO", "已跳过发货表回填（--no-shipped-enrich）")

    rows: list[tuple[Any, ...]] = []
    insert_cols = _insert_columns(with_report_hash=import_batch is not None)
    for d in dicts:
        if import_batch is not None:
            d["report_hash"] = import_batch
        rows.append(tuple(d[c] for c in insert_cols))

    _log(
        "INFO",
        f"准备写入 {TABLE}：有效行={len(rows)} 跳过={skipped} "
        f"（含 check_lock 部分更新={n_locked_partial}） line_hash 键数={len(LINE_HASH_KEYS)}",
    )
    n = _upsert_returned_rows(conn, columns=insert_cols, rows=rows)
    return n, skipped, len(df)


def relisting_base_dir(mode: str | None = None) -> Path:
    return Path(SECOND_RELISTING_PATH.format(MODE_PATTERN=mode or MODE_PATTERN))


def date_to_relisting_folder(on_date: date) -> str:
    """二次上架目录日期格式：M.D，如 2026-06-09 -> 6.9。"""
    return f"{on_date.month}.{on_date.day}"


def date_path_to_relisting_folder(date_path: str) -> str:
    """将 path_config.DATE_PATH 转为二次上架子目录名。"""
    m = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", date_path.strip())
    if m:
        return f"{int(m.group(2))}.{int(m.group(3))}"
    return date_path


def default_date_dir(base: Path) -> Path:
    return base / date_path_to_relisting_folder(DATE_PATH)


def resolve_work_dir(base: Path, on_date: date | None) -> Path:
    if on_date is not None:
        return base / date_to_relisting_folder(on_date)
    return default_date_dir(base)


def discover_returned_files(directory: Path) -> list[Path]:
    files: list[Path] = []
    for pattern in ("*二次上架明细-*.xlsx", "*二次上架明细-*.xls"):
        files.extend(directory.glob(pattern))
    return sorted({p.resolve() for p in files if p.is_file() and not p.name.startswith("~$")})


def _resolve_import_batch(cli_batch: str | None) -> str | None:
    """report_hash 优先取命令行 --import-batch，否则读 run_batch.lock 的 import_batch。"""
    if cli_batch and cli_batch.strip():
        return cli_batch.strip()
    batch = read_import_batch_from_lock()
    if batch:
        _log("INFO", f"从 run_batch.lock 读取 report_hash：{batch}")
    return batch


def _warn_missing_shop_name_en(conn: Any) -> int:
    """
    导入完成后检查 sales_order_returned.shop_name_en 仍为空的行数。
    若存在空值，以醒目颜色提示需人工维护。
    """
    where_blank = "`shop_name_en` IS NULL OR TRIM(`shop_name_en`) = ''"
    with conn.cursor(pymysql.cursors.DictCursor) as cur:
        cur.execute(f"SELECT COUNT(*) AS cnt FROM `{TABLE}` WHERE {where_blank}")
        row = cur.fetchone() or {}
        missing = int(row.get("cnt") or 0)

    if missing <= 0:
        _log_success(f"检查完成：{TABLE}.shop_name_en 无空值")
        return 0

    with conn.cursor(pymysql.cursors.DictCursor) as cur:
        cur.execute(
            f"""
            SELECT `return_doc_no`, `warehouse_sku`, `orig_order_no`, `orig_sales_order_no`, `platform`
            FROM `{TABLE}`
            WHERE {where_blank}
            ORDER BY `id` DESC
            LIMIT 10
            """
        )
        samples = cur.fetchall() or []

    sep = colorize("=" * 72, "RED", "BOLD")
    title = colorize(
        f" 【重要】{TABLE} 中有 {missing} 条 shop_name_en 为空，需人工维护！ ",
        "BG_RED",
        "WHITE",
        "BOLD",
    )
    print(sep, flush=True)
    print(title, flush=True)
    print(sep, flush=True)
    print(colorize("  示例（最多 10 条，按 id 倒序）：", "YELLOW", "BOLD"), flush=True)
    for i, sample in enumerate(samples, 1):
        bits = ", ".join(
            f"{key}={sample.get(key)!r}"
            for key in ("return_doc_no", "warehouse_sku", "orig_order_no", "orig_sales_order_no")
            if sample.get(key)
        )
        print(colorize(f"    {i}. {bits or '(无关键字段)'}", "YELLOW"), flush=True)
    print(
        colorize(
            "  维护建议：先确认 sales_order_shipped 已导入；"
            "或运行 scripts/handle/upReturnedShop.py 辅助回填",
            "CYAN",
            "BOLD",
        ),
        flush=True,
    )
    print(sep, flush=True)
    return missing


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="二次上架明细 Excel -> sales_order_returned")
    ap.add_argument(
        "--date",
        type=date.fromisoformat,
        default=None,
        metavar="YYYY-MM-DD",
        help=f"覆盖日期子目录（默认由 DATE_PATH={DATE_PATH} 转为 M.D 格式）",
    )
    ap.add_argument(
        "--mode",
        choices=("每天", "每月"),
        default=None,
        help=f"路径模式，默认 path_config.MODE_PATTERN（{MODE_PATTERN}）",
    )
    ap.add_argument("--dir", type=Path, default=None, help="直接指定 Excel 目录")
    ap.add_argument("--file", type=Path, default=None, help="指定单个 xls/xlsx")
    ap.add_argument(
        "--no-shipped-enrich",
        action="store_true",
        help="不从 sales_order_shipped 回填平台/店铺等（跳过 order_no 与 provider_order_no 两条规则）",
    )
    ap.add_argument(
        "--import-batch",
        "--batch",
        dest="import_batch",
        default=None,
        metavar="BATCH",
        help="导入批次号写入 report_hash（默认读 run_batch.lock 的 import_batch）",
    )
    return ap.parse_args()


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    enable_windows_ansi()

    args = parse_args()
    mode = args.mode or MODE_PATTERN
    import_batch = _resolve_import_batch(args.import_batch)

    if args.file:
        files = [args.file.resolve()]
        work_dir = args.file.parent
    elif args.dir:
        work_dir = args.dir.resolve()
        if not work_dir.is_dir():
            _log("ERROR", f"目录不存在：{work_dir}")
            return 2
        files = discover_returned_files(work_dir)
    else:
        work_dir = resolve_work_dir(relisting_base_dir(mode), args.date)
        if not work_dir.is_dir():
            _log("ERROR", f"日期目录不存在：{work_dir}")
            return 2
        files = discover_returned_files(work_dir)

    if not files:
        _log("ERROR", f"未找到 *二次上架明细-*.xls(x) ：{work_dir}")
        return 1

    _log("INFO", f"任务：导入 -> {TABLE}")
    _log("INFO", f"模式={mode} 目录={work_dir} 文件数={len(files)}")

    db = get_db_manager(DatabaseConfig())
    conn = db.get_connection()
    total_upsert = 0
    total_skip = 0
    try:
        for fp in files:
            n, skipped, n_excel = import_file(
                conn,
                fp,
                enrich_from_shipped=not args.no_shipped_enrich,
                import_batch=import_batch,
            )
            conn.commit()
            _log("INFO", f"已提交：{fp.name} Excel行={n_excel} UPSERT={n} 跳过={skipped}")
            total_upsert += n
            total_skip += skipped
        market_stats = _fill_market_from_platform_shop(conn)
        conn.commit()
        _log(
            "INFO",
            f"全部完成：UPSERT累计={total_upsert} 总跳过={total_skip} "
            f"市场回填={market_stats.get('updated', 0)}",
        )
        _warn_missing_shop_name_en(conn)
        return 0
    except Exception:
        conn.rollback()
        _log("ERROR", "导入失败，已回滚")
        raise
    finally:
        conn.close()


if __name__ == "__main__":
    raise SystemExit(main())
