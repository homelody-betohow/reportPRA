from __future__ import annotations

"""
根据 temu_order_item 更新 sales_order_shipped 的订单价格字段。

筛选：platform = semitemu 且 order_type = 销售订单。

关联键（优先）：line_hash（与 sales_order_shipped 一致，走 idx_temu_line_hash 索引）
回退关联：ref_no = order_no + platform_sku = sku_id（历史行 line_hash 为空时）

更新字段（付款币 <- temu_order_item；本位币 EUR 按 config/common.py 汇率折算）：
  pay_currency、unit_price_pay、order_goods_pay、order_total_pay、platform_shipping_pay、
  fx_rate_to_base、base_currency、unit_price_base、order_goods_base、order_total_base、platform_shipping_base。

默认仅处理 run_batch.lock 中 import_batch 对应的发货行；加 --all 可更新全表可匹配行。
首次部署或历史数据请执行：python scripts/dataImport/order_temu.py --backfill-line-hash

用法：
  cd d:\\py-project\\report
  python scripts\\dataImport\\order_temu.py
  python scripts\\dataImport\\order_temu.py --backfill-line-hash
  python scripts\\dataImport\\order_temu.py --batch 20260624_115249
  python scripts\\dataImport\\order_temu.py --all
  python scripts\\dataImport\\order_temu.py --dry-run
"""

import argparse
import sys
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

import pymysql.cursors

_REPORT_ROOT = Path(__file__).resolve().parents[2]
_DATA_IMPORT_DIR = Path(__file__).resolve().parent

for _p in (_REPORT_ROOT, _DATA_IMPORT_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from batch_lock import read_import_batch_from_lock  # noqa: E402
from config import common as fx_common  # noqa: E402
from database.db_connection import DatabaseConfig, get_db_manager  # noqa: E402

SHIPPED_TABLE = "sales_order_shipped"
TEMU_ORDER_ITEM_TABLE = "temu_order_item"
PLATFORM_SEMITEMU = "semitemu"
ORDER_TYPE_SALES = "销售订单"
BASE_CURRENCY = "EUR"
_HASH_CHUNK = 500
_KEY_CHUNK = 200
_UPDATE_CHUNK = 300
_QUANTIZE_PAY = Decimal("0.000001")
_QUANTIZE_FX = Decimal("0.00000001")

_SHIPPED_SELECT_COLS: tuple[str, ...] = (
    "id",
    "line_hash",
    "order_no",
    "ref_no",
    "platform_sku",
    "pay_currency",
    "fx_rate_to_base",
    "unit_price_pay",
    "order_goods_pay",
    "order_total_pay",
    "platform_shipping_pay",
    "base_currency",
    "unit_price_base",
    "order_goods_base",
    "order_total_base",
    "platform_shipping_base",
)

_TEMU_SELECT_COLS: tuple[str, ...] = (
    "line_hash",
    "order_no",
    "sku_id",
    "currency",
    "declared_price",
    "order_payment",
    "sales_revenue",
    "shipping_income",
)

# sales_order_shipped 付款币列 <- temu_order_item 列
_PAY_FIELD_MAP: tuple[tuple[str, str], ...] = (
    ("pay_currency", "currency"),
    ("unit_price_pay", "declared_price"),
    ("order_goods_pay", "order_payment"),
    ("order_total_pay", "sales_revenue"),
    ("platform_shipping_pay", "shipping_income"),
)

# 付款币 -> 本位币（同名后缀 _pay / _base）
_PAY_TO_BASE_FIELDS: tuple[tuple[str, str], ...] = (
    ("unit_price_pay", "unit_price_base"),
    ("order_goods_pay", "order_goods_base"),
    ("order_total_pay", "order_total_base"),
    ("platform_shipping_pay", "platform_shipping_base"),
)

# ISO / 符号 -> common.py 中的汇率变量名（1 单位外币 = rate EUR）
_FX_VAR_BY_CURRENCY: dict[str, str] = {
    "USD": "USD_to_EUR",
    "CAD": "CAD_to_EUR",
    "CZK": "kc_to_EUR",
    "PLN": "zl_to_EUR",
    "HUF": "Ft_to_EUR",
    "RON": "Lei_to_EUR",
    "SEK": "kr_to_EUR",
}


def _log(level: str, msg: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] [{level}] {msg}", flush=True)


def _log_unmatched_orders(unmatched: list[dict[str, str]]) -> None:
    if not unmatched:
        return

    order_nos = sorted(
        {str(x.get("order_no") or "").strip() for x in unmatched if str(x.get("order_no") or "").strip()}
    )
    _log("WARN", f"未匹配 temu_order_item：{len(unmatched)} 行，涉及订单号 {len(order_nos)} 个")
    if order_nos:
        _log("WARN", "未匹配订单号：" + "、".join(order_nos))

    for i, item in enumerate(unmatched, 1):
        _log(
            "WARN",
            f"  明细 {i}/{len(unmatched)}："
            f"order_no={item.get('order_no') or ''} "
            f"ref_no={item.get('ref_no') or ''} "
            f"platform_sku={item.get('platform_sku') or ''}",
        )


def _unmatched_item(row: dict[str, Any]) -> dict[str, str]:
    return {
        "order_no": str(row.get("order_no") or "").strip(),
        "ref_no": str(row.get("ref_no") or "").strip(),
        "platform_sku": str(row.get("platform_sku") or "").strip(),
    }


def _norm_key(ref_no: Any, platform_sku: Any) -> tuple[str, str]:
    return (str(ref_no or "").strip(), str(platform_sku or "").strip())


def _norm_hash(line_hash: Any) -> str:
    return str(line_hash or "").strip()


def _to_decimal(v: Any) -> Decimal:
    if v is None:
        return Decimal("0")
    if isinstance(v, Decimal):
        return v
    try:
        return Decimal(str(v))
    except (InvalidOperation, ValueError, TypeError):
        return Decimal("0")


def _str_equal(a: Any, b: Any) -> bool:
    return str(a or "").strip() == str(b or "").strip()


def _dec_equal(a: Any, b: Any) -> bool:
    return _to_decimal(a) == _to_decimal(b)


def _normalize_pay_currency(pay_currency: Any) -> str:
    c = str(pay_currency or "").strip().upper()
    aliases = {
        "€": "EUR",
        "RMB": "CNY",
        "KC": "CZK",
        "ZL": "PLN",
        "FT": "HUF",
        "LEI": "RON",
        "KR": "SEK",
    }
    return aliases.get(c, c)


def _fx_rate_to_base(pay_currency: Any) -> Decimal | None:
    """
    返回「1 单位付款币 = ? EUR」，与 import_temu_fee / A_报表 折算口径一致。
    CNY/RMB：除以 RMB_di_EUR；其余外币：乘以 common.py 中对应系数。
    """
    c = _normalize_pay_currency(pay_currency)
    if not c:
        return None
    if c == "EUR":
        return Decimal("1")
    if c == "CNY":
        return (Decimal("1") / Decimal(str(fx_common.RMB_di_EUR))).quantize(_QUANTIZE_FX)
    var_name = _FX_VAR_BY_CURRENCY.get(c)
    if not var_name:
        return None
    rate = getattr(fx_common, var_name, None)
    if rate is None:
        return None
    return Decimal(str(rate)).quantize(_QUANTIZE_FX)


def _pay_to_base(amount: Any, fx_rate: Decimal) -> Decimal:
    return (_to_decimal(amount) * fx_rate).quantize(_QUANTIZE_PAY)


def _build_target_row(temu: dict[str, Any]) -> dict[str, Any] | None:
    pay_currency = str(temu.get("currency") or "").strip()
    fx_rate = _fx_rate_to_base(pay_currency)
    if fx_rate is None:
        return None

    target: dict[str, Any] = {
        "pay_currency": pay_currency,
        "fx_rate_to_base": fx_rate,
        "base_currency": BASE_CURRENCY,
    }
    for shipped_col, temu_col in _PAY_FIELD_MAP:
        if shipped_col == "pay_currency":
            continue
        pay_val = temu.get(temu_col)
        target[shipped_col] = pay_val
        for pay_col, base_col in _PAY_TO_BASE_FIELDS:
            if pay_col == shipped_col:
                target[base_col] = _pay_to_base(pay_val, fx_rate)
                break
    return target


def _row_differs(shipped: dict[str, Any], target: dict[str, Any]) -> bool:
    if not _str_equal(shipped.get("pay_currency"), target.get("pay_currency")):
        return True
    if not _str_equal(shipped.get("base_currency"), target.get("base_currency")):
        return True
    if not _dec_equal(shipped.get("fx_rate_to_base"), target.get("fx_rate_to_base")):
        return True
    for shipped_col, _ in _PAY_FIELD_MAP:
        if shipped_col == "pay_currency":
            continue
        if not _dec_equal(shipped.get(shipped_col), target.get(shipped_col)):
            return True
    for _, base_col in _PAY_TO_BASE_FIELDS:
        if not _dec_equal(shipped.get(base_col), target.get(base_col)):
            return True
    return False


def _target_to_update_tuple(target: dict[str, Any], row_id: int) -> tuple[Any, ...]:
    return (
        target["pay_currency"],
        target["unit_price_pay"],
        target["order_goods_pay"],
        target["order_total_pay"],
        target["platform_shipping_pay"],
        target["fx_rate_to_base"],
        target["base_currency"],
        target["unit_price_base"],
        target["order_goods_base"],
        target["order_total_base"],
        target["platform_shipping_base"],
        row_id,
    )


def _temu_has_line_hash(conn) -> bool:
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        cur.execute(f"SHOW COLUMNS FROM `{TEMU_ORDER_ITEM_TABLE}` LIKE 'line_hash'")
        return cur.fetchone() is not None
    finally:
        cur.close()


def _shipped_scope_sql(*, extra: str = "") -> str:
    """semitemu + 销售订单 的公共 WHERE 片段（不含 import_batch）。"""
    base = (
        f"TRIM(`platform`) = %s AND TRIM(`order_type`) = %s"
    )
    if extra:
        return f"{base} AND {extra}"
    return base


def _shipped_scope_params() -> list[Any]:
    return [PLATFORM_SEMITEMU, ORDER_TYPE_SALES]


def fetch_shipped_rows(conn, *, import_batch: str | None) -> list[dict[str, Any]]:
    cols = ", ".join(f"`{c}`" for c in _SHIPPED_SELECT_COLS)
    sql = f"SELECT {cols} FROM `{SHIPPED_TABLE}` WHERE {_shipped_scope_sql()}"
    params = _shipped_scope_params()
    if import_batch:
        sql += " AND `import_batch` = %s"
        params.append(import_batch)

    _log("INFO", f"正在查询 semitemu / {ORDER_TYPE_SALES} 发货行…")
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        cur.execute(sql, params)
        rows = list(cur.fetchall())
    finally:
        cur.close()
    _log("INFO", f"发货行查询完成：{len(rows)} 条")
    return rows


def fetch_temu_by_line_hashes(
    conn,
    line_hashes: list[str],
) -> dict[str, dict[str, Any]]:
    unique = sorted({_norm_hash(h) for h in line_hashes if _norm_hash(h)})
    if not unique:
        return {}

    cols = ", ".join(f"`{c}`" for c in _TEMU_SELECT_COLS)
    result: dict[str, dict[str, Any]] = {}
    cur = conn.cursor(pymysql.cursors.DictCursor)
    total = len(unique)
    _log("INFO", f"按 line_hash 查询 temu_order_item：{total} 个…")

    try:
        for i in range(0, total, _HASH_CHUNK):
            part = unique[i : i + _HASH_CHUNK]
            placeholders = ",".join(["%s"] * len(part))
            sql = (
                f"SELECT {cols} FROM `{TEMU_ORDER_ITEM_TABLE}` "
                f"WHERE `line_hash` IN ({placeholders})"
            )
            cur.execute(sql, part)
            for row in cur.fetchall():
                lh = _norm_hash(row.get("line_hash"))
                if lh:
                    result[lh] = row
            done = min(i + _HASH_CHUNK, total)
            if done == total or done % 2000 == 0 or done <= _HASH_CHUNK:
                _log("INFO", f"line_hash 查询进度：{done}/{total}")
    finally:
        cur.close()

    _log("INFO", f"line_hash 命中：{len(result)} 条")
    return result


def fetch_temu_by_keys(
    conn,
    keys: list[tuple[str, str]],
) -> dict[tuple[str, str], dict[str, Any]]:
    unique_keys = sorted({k for k in keys if k[0] and k[1]})
    if not unique_keys:
        return {}

    cols = ", ".join(f"`{c}`" for c in _TEMU_SELECT_COLS)
    result: dict[tuple[str, str], dict[str, Any]] = {}
    cur = conn.cursor(pymysql.cursors.DictCursor)
    total = len(unique_keys)
    _log("INFO", f"按 order_no+sku_id 回退查询 temu_order_item：{total} 个…")

    try:
        for i in range(0, total, _KEY_CHUNK):
            part = unique_keys[i : i + _KEY_CHUNK]
            placeholders = ",".join(["(%s,%s)"] * len(part))
            sql = (
                f"SELECT {cols} FROM `{TEMU_ORDER_ITEM_TABLE}` "
                f"WHERE (`order_no`, `sku_id`) IN ({placeholders})"
            )
            params: list[str] = []
            for order_no, sku_id in part:
                params.extend([order_no, sku_id])
            cur.execute(sql, params)
            for row in cur.fetchall():
                key = _norm_key(row.get("order_no"), row.get("sku_id"))
                result[key] = row
    finally:
        cur.close()

    _log("INFO", f"order_no+sku_id 命中：{len(result)} 条")
    return result


def fetch_temu_maps(
    conn,
    shipped_rows: list[dict[str, Any]],
    *,
    use_line_hash: bool,
) -> tuple[dict[str, dict[str, Any]], dict[tuple[str, str], dict[str, Any]]]:
    by_hash: dict[str, dict[str, Any]] = {}
    if use_line_hash:
        by_hash = fetch_temu_by_line_hashes(
            conn, [_norm_hash(r.get("line_hash")) for r in shipped_rows]
        )

    fallback_keys: list[tuple[str, str]] = []
    for row in shipped_rows:
        lh = _norm_hash(row.get("line_hash"))
        if use_line_hash and lh and lh in by_hash:
            continue
        key = _norm_key(row.get("ref_no"), row.get("platform_sku"))
        if key[0] and key[1]:
            fallback_keys.append(key)

    by_key = fetch_temu_by_keys(conn, fallback_keys) if fallback_keys else {}
    return by_hash, by_key


def resolve_temu_row(
    shipped: dict[str, Any],
    by_hash: dict[str, dict[str, Any]],
    by_key: dict[tuple[str, str], dict[str, Any]],
) -> dict[str, Any] | None:
    lh = _norm_hash(shipped.get("line_hash"))
    if lh:
        hit = by_hash.get(lh)
        if hit is not None:
            return hit

    key = _norm_key(shipped.get("ref_no"), shipped.get("platform_sku"))
    if not key[0] or not key[1]:
        return None
    return by_key.get(key)


def backfill_temu_line_hash(
    conn,
    *,
    import_batch: str | None,
    dry_run: bool,
) -> int:
    """
    将 sales_order_shipped.line_hash 回填到 temu_order_item（仅填空值行）。
    按主键 id 分批 UPDATE，避免大表 JOIN 锁表。
    """
    cols = "`id`, `line_hash`, `ref_no`, `platform_sku`"
    line_hash_clause = "`line_hash` IS NOT NULL AND TRIM(`line_hash`) <> ''"
    sql = (
        f"SELECT {cols} FROM `{SHIPPED_TABLE}` "
        f"WHERE {_shipped_scope_sql(extra=line_hash_clause)}"
    )
    params: list[Any] = _shipped_scope_params()
    if import_batch:
        sql += " AND `import_batch` = %s"
        params.append(import_batch)

    _log("INFO", "正在扫描需回填 line_hash 的发货行…")
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        cur.execute(sql, params)
        shipped_rows = list(cur.fetchall())
    finally:
        cur.close()

    if not shipped_rows:
        _log("INFO", "无发货行可回填 line_hash")
        return 0

    by_key = fetch_temu_by_keys(
        conn,
        [_norm_key(r.get("ref_no"), r.get("platform_sku")) for r in shipped_rows],
    )
    pending: list[tuple[str, str, str]] = []
    for row in shipped_rows:
        key = _norm_key(row.get("ref_no"), row.get("platform_sku"))
        temu = by_key.get(key)
        if temu is None:
            continue
        lh = _norm_hash(row.get("line_hash"))
        if not lh:
            continue
        if _norm_hash(temu.get("line_hash")) == lh:
            continue
        pending.append((lh, key[0], key[1]))

    if not pending:
        _log("INFO", "temu_order_item.line_hash 均已与发货表一致，无需回填")
        return 0

    _log("INFO", f"待回填 line_hash：{len(pending)} 条")
    if dry_run:
        _log("INFO", f"[dry-run] 将回填 {len(pending)} 条 line_hash")
        return len(pending)

    update_sql = (
        f"UPDATE `{TEMU_ORDER_ITEM_TABLE}` SET `line_hash`=%s "
        f"WHERE `order_no`=%s AND `sku_id`=%s "
        f"AND (`line_hash` IS NULL OR TRIM(`line_hash`) = '')"
    )
    cur = conn.cursor()
    affected = 0
    try:
        for i in range(0, len(pending), _UPDATE_CHUNK):
            batch = pending[i : i + _UPDATE_CHUNK]
            cur.executemany(update_sql, batch)
            affected += cur.rowcount
            done = min(i + _UPDATE_CHUNK, len(pending))
            _log("INFO", f"line_hash 回填进度：{done}/{len(pending)}")
    finally:
        cur.close()

    return max(affected, 0)


def build_updates(
    shipped_rows: list[dict[str, Any]],
    by_hash: dict[str, dict[str, Any]],
    by_key: dict[tuple[str, str], dict[str, Any]],
) -> tuple[list[tuple[Any, ...]], dict[str, Any]]:
    updates: list[tuple[Any, ...]] = []
    stats: dict[str, Any] = {
        "scope_total": len(shipped_rows),
        "matched_by_hash": 0,
        "matched_by_key": 0,
        "matched": 0,
        "already_ok": 0,
        "unmatched": 0,
        "fx_missing": 0,
        "pending_update": 0,
        "unmatched_orders": [],
        "fx_missing_orders": [],
    }

    for row in shipped_rows:
        lh = _norm_hash(row.get("line_hash"))
        temu = resolve_temu_row(row, by_hash, by_key)
        if temu is None:
            stats["unmatched"] += 1
            stats["unmatched_orders"].append(_unmatched_item(row))
            continue

        stats["matched"] += 1
        if lh and _norm_hash(temu.get("line_hash")) == lh:
            stats["matched_by_hash"] += 1
        else:
            stats["matched_by_key"] += 1

        target = _build_target_row(temu)
        if target is None:
            stats["fx_missing"] += 1
            item = _unmatched_item(row)
            item["pay_currency"] = str(temu.get("currency") or row.get("pay_currency") or "").strip()
            stats["fx_missing_orders"].append(item)
            continue

        if not _row_differs(row, target):
            stats["already_ok"] += 1
            continue

        stats["pending_update"] += 1
        updates.append(_target_to_update_tuple(target, row["id"]))

    return updates, stats


def apply_updates(
    conn,
    updates: list[tuple[Any, ...]],
    *,
    dry_run: bool,
) -> int:
    if not updates:
        _log("INFO", "没有需要更新的行")
        return 0

    if dry_run:
        _log("INFO", f"[dry-run] 将更新 {len(updates)} 行")
        return len(updates)

    sql = (
        f"UPDATE `{SHIPPED_TABLE}` SET "
        f"`pay_currency`=%s, "
        f"`unit_price_pay`=%s, "
        f"`order_goods_pay`=%s, "
        f"`order_total_pay`=%s, "
        f"`platform_shipping_pay`=%s, "
        f"`fx_rate_to_base`=%s, "
        f"`base_currency`=%s, "
        f"`unit_price_base`=%s, "
        f"`order_goods_base`=%s, "
        f"`order_total_base`=%s, "
        f"`platform_shipping_base`=%s "
        f"WHERE `id`=%s"
    )
    cur = conn.cursor()
    affected = 0
    total = len(updates)
    _log("INFO", f"开始写入：共 {total} 行，每批 {_UPDATE_CHUNK} 行…")

    try:
        for i in range(0, total, _UPDATE_CHUNK):
            batch = updates[i : i + _UPDATE_CHUNK]
            cur.executemany(sql, batch)
            affected += cur.rowcount
            done = min(i + _UPDATE_CHUNK, total)
            _log("INFO", f"写入进度：{done}/{total}（本批 rowcount={cur.rowcount}）")
    finally:
        cur.close()

    return max(affected, 0)


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="根据 temu_order_item 更新 sales_order_shipped 的订单价格字段"
    )
    ap.add_argument(
        "--import-batch",
        "--batch",
        dest="import_batch",
        type=str,
        default=None,
        metavar="BATCH",
        help="指定 import_batch（默认从 run_batch.lock 读取；与 --all 互斥）",
    )
    ap.add_argument(
        "--all",
        action="store_true",
        help="更新全表可匹配行，不按 import_batch 过滤",
    )
    ap.add_argument(
        "--backfill-line-hash",
        action="store_true",
        help="先将 sales_order_shipped.line_hash 回填到 temu_order_item（仅填空值）",
    )
    ap.add_argument(
        "--dry-run",
        action="store_true",
        help="仅统计待更新行数，不执行 UPDATE",
    )
    return ap.parse_args()


def resolve_import_batch(args: argparse.Namespace) -> str | None:
    if args.all:
        return None
    if args.import_batch:
        return args.import_batch.strip()
    batch = read_import_batch_from_lock()
    return batch.strip() if batch else None


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    args = parse_args()
    if args.all and args.import_batch:
        _log("ERROR", "--all 与 --import-batch/--batch 不能同时使用")
        return 1

    import_batch = resolve_import_batch(args)
    if not args.all and not import_batch:
        _log(
            "ERROR",
            "无法获取批次号，请使用 --batch 指定、--all 全表更新，或确保 run_batch.lock 存在",
        )
        return 1

    scope_desc = "全表" if import_batch is None else f"批次 {import_batch}"
    _log("INFO", f"任务：{TEMU_ORDER_ITEM_TABLE} -> {SHIPPED_TABLE} 价格字段")
    _log("INFO", f"范围：{scope_desc} (platform={PLATFORM_SEMITEMU}, order_type={ORDER_TYPE_SALES})")
    _log(
        "INFO",
        f"本位币={BASE_CURRENCY}；汇率来源 config/common.py "
        f"(RMB_di_EUR={fx_common.RMB_di_EUR}, USD_to_EUR={fx_common.USD_to_EUR})",
    )
    if args.dry_run:
        _log("INFO", "模式：dry-run（不写入数据库）")

    _log("INFO", "正在连接数据库…")
    db = get_db_manager(DatabaseConfig())
    conn = db.get_connection()
    _log("INFO", "数据库连接就绪")

    try:
        use_line_hash = _temu_has_line_hash(conn)
        if use_line_hash:
            _log("INFO", "temu_order_item 已含 line_hash 列，优先按 line_hash 关联")
        else:
            _log("WARN", "temu_order_item 尚无 line_hash 列，仅按 order_no+sku_id 关联（请执行迁移 SQL）")

        if args.backfill_line_hash:
            if not use_line_hash:
                _log("ERROR", "无法回填：temu_order_item 缺少 line_hash 列，请先执行 alter_temu_order_item_add_line_hash.sql")
                return 1
            n_backfill = backfill_temu_line_hash(
                conn,
                import_batch=import_batch,
                dry_run=args.dry_run,
            )
            _log("INFO", f"line_hash 回填：{'预计' if args.dry_run else '实际'} {n_backfill} 条")

        shipped_rows = fetch_shipped_rows(conn, import_batch=import_batch)
        by_hash, by_key = fetch_temu_maps(
            conn, shipped_rows, use_line_hash=use_line_hash
        )
        updates, stats = build_updates(shipped_rows, by_hash, by_key)

        _log("INFO", f"范围内 {PLATFORM_SEMITEMU}/{ORDER_TYPE_SALES} 发货行：{stats['scope_total']} 条")
        _log("INFO", f"可匹配 temu_order_item：{stats['matched']} 条（line_hash={stats['matched_by_hash']}，回退键={stats['matched_by_key']}）")
        _log("INFO", f"价格字段已正确：{stats['already_ok']} 条")
        _log("INFO", f"无法匹配订单明细：{stats['unmatched']} 条")
        _log_unmatched_orders(stats.get("unmatched_orders") or [])
        if stats.get("fx_missing"):
            _log("WARN", f"付款币种无汇率配置，跳过：{stats['fx_missing']} 条")
            fx_orders = stats.get("fx_missing_orders") or []
            fx_order_nos = sorted(
                {x.get("order_no") or "" for x in fx_orders if x.get("order_no")}
            )
            if fx_order_nos:
                _log("WARN", "无汇率订单号：" + "、".join(fx_order_nos))
        _log("INFO", f"待更新：{stats['pending_update']} 条")

        n_updated = apply_updates(conn, updates, dry_run=args.dry_run)

        if args.dry_run:
            conn.rollback()
            _log("INFO", f"dry-run 完成：预计更新 {n_updated} 条")
        else:
            conn.commit()
            _log("INFO", f"更新完成：实际更新 {n_updated} 条")

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
