from __future__ import annotations

"""
鸿羽「二次上架/退件」明细 xls/xlsx -> 表 sales_order_returned。

默认文件：python/excel/daily/order/*二次上架*.xls（及 .xlsx）
工作表名：ReturnOrders（否则取第一个 sheet）
表头：第 1 行（header=0）

line_hash：LINE_HASH_KEYS 子集经 row_subset_for_line_hash + stable_line_hash；
改键后历史行 hash 会变，需按业务重导或清表。
product_sku：由 warehouse_sku 去掉前缀「900008-」后的后缀（须以前缀开头；否则为 NULL）。

发货表回填：默认用 orig_order_no 在 sales_order_shipped 上匹配
order_no / ref_no / sales_order_no / provider_order_no（任一相等即命中）。
同一订单多行时优先 warehouse_sku 与退件 Excel「SKU」列一致的发货行，否则取最近 ship_time；
仅当退件行对应字段为空时才写入（不覆盖 Excel 已有值）。
当退件 orig_order_no 与选中发货行的 provider_order_no 一致时，另用发货 warehouse_sku 同步退件 product_sku（见 _maybe_sync_product_sku_from_shipped_provider）。
过程日志：分阶段输出查询批次数、拉回行数、命中/未命中 orig、SKU 匹配与「最近发货」选用次数、各列补全计数与耗时。

表头映射：见 RETURNED_IMPORT_COLUMN_MAP（逻辑键 -> 当前模板 Excel 列名 -> 落库字段说明）。
鸿羽导出改版时只改映射中的「Excel 列名」字符串；导入时会校验必需列并打日志（含一行 JSON 快照）。
"""

import argparse
import json
import os
import re
import sys
import time
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from collections import defaultdict
from typing import Any

import mysql.connector

Decimal0 = Decimal("0")

SHIPPED_TABLE = "sales_order_shipped"

# 从发货表取出的列（用于回填 + 匹配 SKU / 排序）
_SHIPPED_FETCH_COLS: tuple[str, ...] = (
    "id",
    "ship_time",
    "order_no",
    "ref_no",
    "sales_order_no",
    "provider_order_no",
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

# 发货表回填：仅当退件行对应列为空时写入（_enrich_row_from_shipped）。
# 退件列 <- sales_order_shipped 列：
#   platform, shop_name_en, shop_alias, platform_site, warehouse_name,
#   platform_sku, product_name, platform_sku_owner, orig_tracking_no <- tracking_no, sales_owner
# 另：若本行 orig_order_no 与选中发货行的 provider_order_no 一致（即由服务商单号命中），
#     则用该发货行的 warehouse_sku 同步退件 product_sku（先按 900008- 前缀派生，否则用整段 warehouse_sku）。
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
    cell_str,
    cell_str_or_empty,
    default_order_excel_dir,
    row_subset_for_line_hash,
    stable_line_hash,
    upsert_rows,
)

TABLE = "sales_order_returned"
SOURCE_TYPE = "Excel"
_LOG = get_logger("ORDER-RETURNED")

# 鸿羽 ReturnOrders / 二次上架明细：逻辑键 -> 当前模板 Excel 列名 -> sales_order_returned 相关列说明
# 文件表头变更时：只改第二列字符串；第三列仅作文档/日志说明。
# required=True 的列必须在文件中出现，否则导入报错（尽早失败）。
RETURNED_IMPORT_COLUMN_MAP: tuple[tuple[str, str, str, bool], ...] = (
    ("sku", "SKU", "warehouse_sku", True),
    ("return_doc", "退件号", "return_doc_no", False),
    ("order_no", "订单号", "orig_order_no 优先来源", False),
    ("ref_no", "参考号", "orig_order_no / orig_ref_no", False),
    ("order_ref", "订单参考号", "orig_sales_order_no / orig_ref_no", False),
    ("claim_no", "认领单号", "rma_no", False),
    ("warehouse", "仓库", "warehouse_name / warehouse_code", False),
    ("status", "状态", "return_status", False),
    ("return_type", "退件类型", "return_type", False),
    ("label_svc", "Label服务", "shipping_method", False),
    ("track_no", "跟踪号", "return_tracking_no", False),
    ("qty", "数量", "return_qty", False),
    ("disposition", "处理方式", "disposition", False),
    ("qty_received", "实收数量", "received_qty", False),
    ("qty_good", "良品", "putaway_qty", False),
    ("qty_bad", "不良品", "warehouse_remark", False),
    ("zone_transfer", "转运区", "warehouse_remark", False),
    ("reason", "退件原因", "cs_remark", False),
    ("note", "退件说明", "cs_remark", False),
    ("created_by", "创建人", "sales_owner", False),
    ("created_at", "创建时间", "apply_at", False),
    ("audit_at", "审核时间", "inspected_at", False),
    ("done_at", "完成时间", "putaway_at", False),
    ("fee_rmb", "退件费用(RMB)", "handling_fee / currency_code", False),
    ("shop_seller", "卖家店铺", "shop_alias", False),
    ("return_rec_client", "退货记录（客户端）", "warehouse_remark", False),
    ("return_rec_wh", "退货记录（仓储端）", "warehouse_remark", False),
)

_logicals = [row[0] for row in RETURNED_IMPORT_COLUMN_MAP]
if len(_logicals) != len(frozenset(_logicals)):
    raise ValueError("RETURNED_IMPORT_COLUMN_MAP 存在重复的逻辑键")

RETURNED_HEADER_BY_LOGICAL: dict[str, str] = {logical: zh for logical, zh, _, _ in RETURNED_IMPORT_COLUMN_MAP}

# 解析 orig_order_no 时至少要有一列存在（不要求每行都有值）
_ORDER_LINEAGE_LOGICALS: tuple[str, ...] = ("order_no", "ref_no", "order_ref", "return_doc")


def _xc(series: pd.Series, logical: str) -> Any:
    """按逻辑键取 Excel 单元格（列名来自 RETURNED_HEADER_BY_LOGICAL）。"""
    zh = RETURNED_HEADER_BY_LOGICAL[logical]
    return series.get(zh) if zh in series.index else None


def _validate_and_log_excel_headers(path: Path, df: pd.DataFrame) -> None:
    cols = frozenset(df.columns)
    h = RETURNED_HEADER_BY_LOGICAL
    missing_required: list[str] = []
    for logical, zh, _note, req in RETURNED_IMPORT_COLUMN_MAP:
        if req and zh not in cols:
            missing_required.append(f"{logical!r}->{zh!r}")

    lineage_headers = {h[k] for k in _ORDER_LINEAGE_LOGICALS}
    if not (lineage_headers & cols):
        raise RuntimeError(
            f"Excel 缺少 orig_order_no 解析所需列（至少要有其一）: "
            f"{sorted(lineage_headers)}；文件={path}"
        )

    if missing_required:
        raise RuntimeError(
            f"Excel 缺少映射表中标记为必需的列：{missing_required}；文件={path}。"
            "若模板改版，请同步修改 RETURNED_IMPORT_COLUMN_MAP。"
        )

    rows_log: list[dict[str, Any]] = []
    missing_optional: list[str] = []
    for logical, zh, db_note, req in RETURNED_IMPORT_COLUMN_MAP:
        present = zh in cols
        rows_log.append(
            {
                "logical": logical,
                "excel_header": zh,
                "db_hint": db_note,
                "required": req,
                "present_in_file": present,
            }
        )
        if not present and not req:
            missing_optional.append(zh)

    mapped_headers = {row[1] for row in RETURNED_IMPORT_COLUMN_MAP}
    extra = [c for c in df.columns if c not in mapped_headers]

    _LOG.info(
        "ReturnOrders 表头映射（逻辑键 -> Excel列 -> 库字段说明；present=文件是否含该列）：\n"
        + "\n".join(
            f"  [{r['logical']}] {r['excel_header']!r} -> {r['db_hint']} "
            f"{'[必需]' if r['required'] else ''} "
            f"{'OK' if r['present_in_file'] else '缺失'}"
            for r in rows_log
        )
    )
    if missing_optional:
        _LOG.warn(
            f"以下映射列未出现在本文件中（对应库字段将多为 NULL）：共 {len(missing_optional)} 列；"
            f"示例：{missing_optional[:12]}{'...' if len(missing_optional) > 12 else ''}"
        )
    if extra:
        _LOG.info(
            f"文件中未出现在 RETURNED_IMPORT_COLUMN_MAP 的列（当前脚本未读取）："
            f"共 {len(extra)} 列；示例：{extra[:20]}{'...' if len(extra) > 20 else ''}"
        )

    snapshot = {
        "file": str(path.resolve()),
        "sheet_columns_n": len(df.columns),
        "column_map": rows_log,
        "excel_columns_not_in_map": extra,
    }
    _LOG.info("HEADER_MAP_JSON " + json.dumps(snapshot, ensure_ascii=False, separators=(",", ":")))

# 仅用于 line_hash 的「业务源字段」名（与 _build_row 返回键一致，勿随意改名）
LINE_HASH_KEYS: tuple[str, ...] = (
    "return_doc_no",
    "warehouse_sku",
    "return_tracking_no",
    "orig_order_no",
    "orig_ref_no",
)

# 仓库 SKU 前缀；product_sku = 去掉此前缀后的后缀（不含此前缀则 product_sku 为 NULL）
_WAREHOUSE_SKU_PRODUCT_PREFIX = "900008-"


def _product_sku_from_warehouse_sku(wh: Any) -> str | None:
    """product_sku：warehouse_sku 以「900008-」为前缀时去掉此前缀后的子串；否则 NULL。"""
    s = cell_str(wh, 128)
    if not s:
        return None
    p = _WAREHOUSE_SKU_PRODUCT_PREFIX
    if s.startswith(p):
        tail = s[len(p) :].strip()
        return cell_str(tail, 128) if tail else None
    return None


_WH_BRACKET_RE = re.compile(r"^\s*\[([^\]]+)\]\s*(.*)$")


def _excel_engine(path: Path) -> str:
    suf = path.suffix.lower()
    if suf == ".xls":
        return "xlrd"
    return "openpyxl"


def _parse_warehouse(raw: Any) -> tuple[str | None, str | None]:
    """如 '[DEHY]DEHY' -> ('DEHY','DEHY')；无方括号则整体为仓库名称。"""
    s = cell_str(raw, 255)
    if not s:
        return None, None
    m = _WH_BRACKET_RE.match(s)
    if not m:
        return s, None
    code = (m.group(1) or "").strip() or None
    tail = (m.group(2) or "").strip()
    name = tail or code
    return name, code


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
    """
    去重后优先按 warehouse_sku 命中，否则取 ship_time 最新的一条。
    返回 (发货行或 None, 选用原因: sku|latest|none)。
    """
    if not candidates:
        return None, "none"
    dedup: dict[Any, dict[str, Any]] = {}
    for r in candidates:
        rid = r.get("id")
        if rid is not None:
            dedup[rid] = r
        else:
            dedup[id(r)] = r
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


def _register_shipped_row(
    reg: dict[str, list[dict[str, Any]]],
    wanted: frozenset[str],
    row: dict[str, Any],
    seen_ids: dict[str, set[Any]],
) -> None:
    """同一发货行可能同时命中 order_no/ref_no 等多列，按 id 去重避免 reg[s] 膨胀。"""
    rid = row.get("id")
    for k in ("order_no", "ref_no", "sales_order_no", "provider_order_no"):
        v = row.get(k)
        if v is None:
            continue
        s = str(v).strip()
        if not s or s not in wanted:
            continue
        bucket = seen_ids.setdefault(s, set())
        dedupe_key = rid if rid is not None else id(row)
        if dedupe_key in bucket:
            continue
        bucket.add(dedupe_key)
        reg[s].append(row)


def _fetch_shipped_by_orig_orders(
    conn: mysql.connector.MySQLConnection,
    orig_orders: frozenset[str],
    *,
    chunk_size: int = 200,
) -> tuple[dict[str, list[dict[str, Any]]], dict[str, Any]]:
    """
    按 order_no / ref_no / sales_order_no / provider_order_no 命中 orig。
    返回 (orig -> 发货行列表, 统计元数据)。
    """
    out: dict[str, list[dict[str, Any]]] = defaultdict(list)
    meta: dict[str, Any] = {
        "n_chunks": 0,
        "n_rows_fetched": 0,
        "elapsed_ms": 0.0,
    }
    if not orig_orders:
        return out, meta

    cols_sql = ", ".join(f"`{c}`" for c in _SHIPPED_FETCH_COLS)
    wanted_list = sorted(orig_orders)
    cur = conn.cursor(dictionary=True)
    t0 = time.perf_counter()
    seen_ids: dict[str, set[Any]] = {}
    try:
        for i in range(0, len(wanted_list), chunk_size):
            chunk = wanted_list[i : i + chunk_size]
            if not chunk:
                break
            ph = ", ".join(["%s"] * len(chunk))
            sql = (
                f"SELECT {cols_sql} FROM `{SHIPPED_TABLE}` "
                f"WHERE `order_no` IN ({ph}) OR `ref_no` IN ({ph}) "
                f"OR `sales_order_no` IN ({ph}) OR `provider_order_no` IN ({ph})"
            )
            params = tuple(chunk) * 4
            cur.execute(sql, params)
            batch = cur.fetchall() or []
            meta["n_rows_fetched"] += len(batch)
            meta["n_chunks"] += 1
            wanted_f = frozenset(chunk)
            for raw in batch:
                row = {k: raw.get(k) for k in _SHIPPED_FETCH_COLS}
                _register_shipped_row(out, wanted_f, row, seen_ids)
    finally:
        cur.close()
    meta["elapsed_ms"] = (time.perf_counter() - t0) * 1000.0
    return out, meta


def _enrich_row_from_shipped(
    d: dict[str, Any],
    shipped: dict[str, Any] | None,
    cell_counts: dict[str, int],
) -> bool:
    """若有任意字段被写入返回 True；cell_counts 累计各退件列写入次数。"""
    if not shipped:
        return False
    changed = False
    for ret_col, ship_col, maxlen in _SHIPPED_FILL_MAP:
        if not _is_blank_for_fill(d.get(ret_col)):
            continue
        val = cell_str(shipped.get(ship_col), maxlen)
        if val is not None:
            d[ret_col] = val
            cell_counts[ret_col] = cell_counts.get(ret_col, 0) + 1
            changed = True
    return changed


def _maybe_sync_product_sku_from_shipped_provider(
    d: dict[str, Any],
    shipped: dict[str, Any],
    orig: str,
    cell_counts: dict[str, int],
) -> bool:
    """
    orig 与发货行 provider_order_no 一致时，用发货 warehouse_sku 更新退件 product_sku。
    优先 _product_sku_from_warehouse_sku；无前缀则退回整段 warehouse_sku（截断 128）。
    """
    prov = cell_str(shipped.get("provider_order_no"), 128)
    if not prov or prov != orig:
        return False
    whs = cell_str(shipped.get("warehouse_sku"), 128)
    if not whs:
        return False
    derived = _product_sku_from_warehouse_sku(whs)
    new_ps = derived if derived is not None else whs
    old_ps = d.get("product_sku")
    if old_ps == new_ps:
        return False
    d["product_sku"] = new_ps
    cell_counts["product_sku"] = cell_counts.get("product_sku", 0) + 1
    return True


def _enrich_rows_from_shipped(
    conn: mysql.connector.MySQLConnection,
    rows: list[dict[str, Any]],
) -> dict[str, Any]:
    """
    用 sales_order_shipped 回填 rows。
    返回统计 dict（便于日志与后续扩展）。
    """
    orig_set = frozenset(str(d["orig_order_no"]).strip() for d in rows if d.get("orig_order_no"))
    n_distinct = len(orig_set)

    _LOG.info(
        f"【发货回填】开始：表={SHIPPED_TABLE}，退件有效行={len(rows)}，"
        f"distinct orig_order_no={n_distinct}，匹配列=order_no/ref_no/sales_order_no/provider_order_no"
    )

    shipped_map, fetch_meta = _fetch_shipped_by_orig_orders(conn, orig_set)
    n_keys_hit = sum(1 for k in orig_set if shipped_map.get(k))
    miss_origs = sorted(k for k in orig_set if not shipped_map.get(k))
    hit_lists = [shipped_map[k] for k in orig_set if shipped_map.get(k)]
    n_candidates = sum(len(x) for x in hit_lists)
    avg_cand = (n_candidates / n_keys_hit) if n_keys_hit else 0.0

    _LOG.info(
        f"【发货回填】查询完成：SQL 批次数={fetch_meta['n_chunks']}，"
        f"拉回发货行={fetch_meta['n_rows_fetched']}（去重前），"
        f"耗时={fetch_meta['elapsed_ms']:.1f}ms"
    )
    _LOG.info(
        f"【发货回填】候选行：命中 orig 下发货行条数合计={n_candidates}（去重注册后），"
        f"平均 {avg_cand:.2f} 条/命中 orig"
    )
    _LOG.info(
        f"【发货回填】orig 命中：至少一条发货={n_keys_hit}/{n_distinct}，"
        f"发货表完全无记录={len(miss_origs)}"
    )
    if miss_origs:
        preview = miss_origs[:12]
        _LOG.warn(
            f"【发货回填】未命中 orig_order_no 示例（最多12个）: {preview}"
            f"{' ...' if len(miss_origs) > 12 else ''}"
        )

    cell_counts: dict[str, int] = {}
    n_rows_touched = 0
    n_rows_no_ship = 0
    n_pick_sku = 0
    n_pick_latest = 0
    n_product_sku_via_provider = 0

    t1 = time.perf_counter()
    for d in rows:
        orig = str(d.get("orig_order_no") or "").strip()
        if not orig:
            continue
        candidates = shipped_map.get(orig) or []
        ship, pick_reason = _pick_shipped_row(candidates, d.get("warehouse_sku"))
        if ship is None:
            n_rows_no_ship += 1
            continue
        if pick_reason == "sku":
            n_pick_sku += 1
        elif pick_reason == "latest":
            n_pick_latest += 1
        row_changed = False
        if _enrich_row_from_shipped(d, ship, cell_counts):
            row_changed = True
        if _maybe_sync_product_sku_from_shipped_provider(d, ship, orig, cell_counts):
            row_changed = True
            n_product_sku_via_provider += 1
        if row_changed:
            n_rows_touched += 1
    enrich_ms = (time.perf_counter() - t1) * 1000.0

    total_cells = sum(cell_counts.values())
    col_summary = ", ".join(f"{k}={v}" for k, v in sorted(cell_counts.items()) if v > 0)
    if not col_summary:
        col_summary = "（无，退件行已有值或发货侧为空）"

    _LOG.info(
        f"【发货回填】逐行选用：按 SKU 命中发货行={n_pick_sku}，"
        f"按最近 ship_time={n_pick_latest}，"
        f"无候选发货行（本批退件行）={n_rows_no_ship}"
    )
    _LOG.info(
        f"【发货回填】写入退件表字段：至少补到一列的退件行={n_rows_touched}/{len(rows)}，"
        f"其中 provider_order_no 命中后同步 product_sku={n_product_sku_via_provider}，"
        f"补全单元格合计={total_cells}，逐列计数: {col_summary}"
    )
    _LOG.info(f"【发货回填】逐行补全耗时={enrich_ms:.1f}ms，阶段结束")

    if os.environ.get("ORDER_IMPORT_VERBOSE") == "1" and miss_origs:
        _LOG.info("【发货回填】VERBOSE 未命中 orig 全集（每行一个）:\n" + "\n".join(f"  {x}" for x in miss_origs))

    return {
        "touched_rows": n_rows_touched,
        "keys_hit": n_keys_hit,
        "n_distinct_orig": n_distinct,
        "n_rows_no_shipped": n_rows_no_ship,
        "n_pick_by_sku": n_pick_sku,
        "n_pick_by_latest": n_pick_latest,
        "n_product_sku_via_provider_order_no": n_product_sku_via_provider,
        "n_shipped_rows_registered": n_candidates,
        "avg_shipped_rows_per_hit_orig": round(avg_cand, 4),
        "cell_counts": dict(cell_counts),
        "total_cells_filled": total_cells,
        "fetch_ms": fetch_meta["elapsed_ms"],
        "enrich_ms": enrich_ms,
        "n_rows_fetched": fetch_meta["n_rows_fetched"],
        "n_chunks": fetch_meta["n_chunks"],
        "miss_orig_count": len(miss_origs),
    }


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
    """订单号优先；否则参考号、订单参考号、退件号。orig_ref_no 对齐常见 ERP 参考列。"""
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
    if not sku:
        return None
    if not orig_order_no:
        return None

    wh_name, wh_code = _parse_warehouse(_xc(series, "warehouse"))
    return_doc = cell_str(_xc(series, "return_doc"), 128)
    track = cell_str(_xc(series, "track_no"), 255)

    qty = cell_decimal(_xc(series, "qty")) or Decimal0
    recv_q = cell_decimal(_xc(series, "qty_received"))
    good_q = cell_decimal(_xc(series, "qty_good"))
    fee = cell_decimal(_xc(series, "fee_rmb"))

    cs_parts = (_xc(series, "reason"), _xc(series, "note"))
    bad = cell_str(_xc(series, "qty_bad"), 32)
    zone = cell_str(_xc(series, "zone_transfer"), 64)
    wh_parts = (
        f"不良品:{bad}" if bad is not None else None,
        f"转运区:{zone}" if zone is not None else None,
        _xc(series, "return_rec_client"),
        _xc(series, "return_rec_wh"),
    )

    row: dict[str, Any] = {
        "platform": None,
        "shop_name_en": None,
        "shop_alias": cell_str(_xc(series, "shop_seller"), 128),
        "platform_site": None,
        "warehouse_name": wh_name,
        "warehouse_code": wh_code,
        "return_doc_no": return_doc,
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
        "return_qty": qty,
        "received_qty": recv_q,
        "putaway_qty": good_q,
        "return_tracking_no": track,
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
        "cs_remark": _combine_remark(*cs_parts),
        "warehouse_remark": _combine_remark(*wh_parts),
        "profit_calc_node": None,
        "source_type": SOURCE_TYPE,
    }
    return row


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
            continue
    if df is None:
        raise RuntimeError(f"无法读取工作表（需 ReturnOrders 或第一个 sheet）: {path}") from last
    _LOG.warn(f"读取 Excel：{path.name} engine={engine} sheet={sheet_used!r} header=第1行")
    df.columns = [("" if c is None else str(c)).replace("\n", " ").strip() for c in df.columns]
    df = df.dropna(how="all")
    _LOG.info(f"读取完成：行数={len(df)} 列数={len(df.columns)}")
    _validate_and_log_excel_headers(path, df)
    return df


def _insert_columns() -> list[str]:
    return [
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
    ]


def import_file(
    conn: mysql.connector.MySQLConnection,
    path: Path,
    *,
    enrich_from_shipped: bool = True,
) -> tuple[int, int, int]:
    df = _read_returned_frame(path)
    insert_cols = _insert_columns()
    row_dicts: list[dict[str, Any]] = []
    skipped = 0
    for _, series in df.iterrows():
        d = _build_row(series)
        if d is None:
            skipped += 1
            continue
        if d["return_qty"] is None:
            d["return_qty"] = Decimal0
        row_dicts.append(d)

    n_excel = len(df)
    if not row_dicts:
        _LOG.warn(f"无有效行可写：Excel 行数={n_excel} 跳过={skipped}（缺 SKU 或无法解析原订单号）")
        return 0, skipped, n_excel

    if enrich_from_shipped:
        _enrich_rows_from_shipped(conn, row_dicts)
        # _LOG.info(
        #     "SHIPPED_ENRICH_JSON "
        #     + json.dumps(st, ensure_ascii=False, separators=(",", ":"), default=str)
        # )
    else:
        _LOG.info("已跳过发货表回填（--no-shipped-enrich）")

    rows: list[tuple[Any, ...]] = []
    for d in row_dicts:
        h_in = row_subset_for_line_hash(d, LINE_HASH_KEYS)
        d["line_hash"] = stable_line_hash(h_in)
        rows.append(tuple(d[c] for c in insert_cols))
    _LOG.info(
        f"准备写入 MySQL：表={TABLE} 行数={len(rows)}（跳过={skipped}，line_hash 键数={len(LINE_HASH_KEYS)}）"
    )
    n = upsert_rows(conn, table=TABLE, columns=insert_cols, rows=rows)
    _LOG.info(f"MySQL executemany 完成：批次累计行数={n}")
    return n, skipped, n_excel


def main() -> int:
    setup_stdout_utf8()
    ap = argparse.ArgumentParser(description="导入 二次上架明细 xls/xlsx -> sales_order_returned")
    ap.add_argument("--file", type=Path, default=None, help="指定单个 xls/xlsx")
    ap.add_argument("--dir", type=Path, default=None, help=f"Excel 目录，默认 {default_order_excel_dir()}")
    ap.add_argument(
        "--no-shipped-enrich",
        action="store_true",
        help="不从 sales_order_shipped 按 orig_order_no 回填店铺/平台/SKU 等",
    )
    args = ap.parse_args()
    base = args.dir or default_order_excel_dir()
    _LOG.info(f"任务：退件/二次上架明细导入 -> {TABLE}")
    _LOG.info(
        f"line_hash 参与键共 {len(LINE_HASH_KEYS)} 个；"
        "说明见模块文档字符串；算法 row_subset_for_line_hash + stable_line_hash"
    )
    if os.environ.get("ORDER_IMPORT_VERBOSE") == "1":
        _LOG.info("LINE_HASH_KEYS: " + ", ".join(LINE_HASH_KEYS))
    if not base.is_dir():
        _LOG.error(f"目录不存在: {base}")
        return 2

    if args.file:
        files = [args.file]
    else:
        files = sorted(base.glob("*二次上架*.xls")) + sorted(base.glob("*二次上架*.xlsx"))
    if not files:
        _LOG.error(f"未找到 *二次上架*.xls / *二次上架*.xlsx: {base}")
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
            n, skipped, n_excel = import_file(conn, p, enrich_from_shipped=not args.no_shipped_enrich)
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
