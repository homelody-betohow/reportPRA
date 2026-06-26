from __future__ import annotations

"""
根据 temu_order_item 表更新 sales_order_sku_profit 的订单价格字段。

关联键（优先）：line_hash（走 idx_temu_line_hash / uk_sosp_line_hash）
回退关联：ref_no = order_no + platform_sku = sku_id

更新字段：order_total_pay、order_goods_base（来自 temu sales_revenue / order_payment）。
仅更新 temu_order_item 中 file_name 不为空的数据（RPA/手工导入的订单详情）。
sales_order_shipped.order_type =「重发订单」的行跳过，不更新价格。
执行结束后校验：platform=semitemu 且非重发订单的 order_total_base 须 > 0，否则红色醒目提醒。

性能：先按批次缩小 profit 范围，再分批 IN 查询 temu，最后按主键 id 分批 UPDATE，
避免大表 TRIM/COLLATE JOIN。

用法：
  cd d:\\py-project\\report
  python scripts\\archive\\profit_002_order_price.py
  python scripts\\archive\\profit_002_order_price.py --batch 20260616_203140
  python scripts\\archive\\profit_002_order_price.py --all
  python scripts\\archive\\profit_002_order_price.py --dry-run
"""

import argparse
import sys
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

import pymysql.cursors

_REPORT_ROOT = Path(__file__).resolve().parents[2]
_ARCHIVE_DIR = Path(__file__).resolve().parent
_SCRIPTS_DIR = _ARCHIVE_DIR.parent
_DATA_IMPORT_DIR = _REPORT_ROOT / "scripts" / "dataImport"

for _p in (_REPORT_ROOT, _SCRIPTS_DIR, _ARCHIVE_DIR, _DATA_IMPORT_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from batch_lock import read_import_batch_from_lock  # noqa: E402  # pyright: ignore[reportMissingImports]
from console_style import init_console, print_banner  # noqa: E402
from database.db_connection import DatabaseConfig, get_db_manager  # noqa: E402

PROFIT_TABLE = "sales_order_sku_profit"
SHIPPED_TABLE = "sales_order_shipped"
TEMU_ORDER_ITEM_TABLE = "temu_order_item"
PLATFORM_SEMITEMU = "semitemu"
RESEND_ORDER_TYPE = "重发订单"

_HASH_CHUNK = 500
_KEY_CHUNK = 200
_UPDATE_CHUNK = 300
_CHECK_SAMPLE_LIMIT = 20

_PROFIT_SELECT_COLS: tuple[str, ...] = (
    "id",
    "line_hash",
    "ref_no",
    "platform_sku",
    "order_total_pay",
    "order_goods_base",
)

_TEMU_SELECT_COLS: tuple[str, ...] = (
    "line_hash",
    "order_no",
    "sku_id",
    "sales_revenue",
    "order_payment",
    "file_name",
)


def _log(level: str, msg: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] [{level}] {msg}", flush=True)


def _norm_hash(line_hash: Any) -> str:
    return str(line_hash or "").strip()


def _norm_key(ref_no: Any, platform_sku: Any) -> tuple[str, str]:
    return (str(ref_no or "").strip(), str(platform_sku or "").strip())


def _to_decimal(v: Any) -> Decimal:
    if v is None:
        return Decimal("0")
    if isinstance(v, Decimal):
        return v
    try:
        return Decimal(str(v))
    except (InvalidOperation, ValueError, TypeError):
        return Decimal("0")


def _dec_equal(a: Any, b: Any) -> bool:
    return _to_decimal(a) == _to_decimal(b)


def _has_file_name(temu: dict[str, Any]) -> bool:
    return bool(str(temu.get("file_name") or "").strip())


def _temu_has_line_hash(conn) -> bool:
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        cur.execute(f"SHOW COLUMNS FROM `{TEMU_ORDER_ITEM_TABLE}` LIKE 'line_hash'")
        return cur.fetchone() is not None
    finally:
        cur.close()


def fetch_profit_rows(conn, *, import_batch: str | None) -> list[dict[str, Any]]:
    """查询范围内 semitemu 利润行（批次模式经 shipped 过滤并排除重发）。"""
    cols = ", ".join(f"p.`{c}`" for c in _PROFIT_SELECT_COLS)
    params: list[Any] = [PLATFORM_SEMITEMU]

    if import_batch:
        sql = f"""
            SELECT {cols}
            FROM `{PROFIT_TABLE}` AS p
            INNER JOIN `{SHIPPED_TABLE}` AS s ON s.`line_hash` = p.`line_hash`
            WHERE s.`import_batch` = %s
              AND p.`platform` = %s
              AND IFNULL(s.`order_type`, '') <> %s
        """
        params = [import_batch, PLATFORM_SEMITEMU, RESEND_ORDER_TYPE]
        scope = f"批次 {import_batch}"
    else:
        sql = f"""
            SELECT {", ".join(f"`{c}`" for c in _PROFIT_SELECT_COLS)}
            FROM `{PROFIT_TABLE}`
            WHERE `platform` = %s
              AND IFNULL(`order_type`, '') <> %s
        """
        params.append(RESEND_ORDER_TYPE)
        scope = "全表"

    _log("INFO", f"正在查询 {scope} 内 {PLATFORM_SEMITEMU} 利润行…")
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        cur.execute(sql, params)
        rows = list(cur.fetchall())
    finally:
        cur.close()
    _log("INFO", f"利润行查询完成：{len(rows)} 条")
    return rows


def fetch_temu_by_line_hashes(conn, line_hashes: list[str]) -> dict[str, dict[str, Any]]:
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
                if lh and _has_file_name(row):
                    result[lh] = row
            done = min(i + _HASH_CHUNK, total)
            if done == total or done % 2000 == 0 or done <= _HASH_CHUNK:
                _log("INFO", f"line_hash 查询进度：{done}/{total}")
    finally:
        cur.close()

    _log("INFO", f"line_hash 命中（file_name 非空）：{len(result)} 条")
    return result


def fetch_temu_by_keys(conn, keys: list[tuple[str, str]]) -> dict[tuple[str, str], dict[str, Any]]:
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
                if not _has_file_name(row):
                    continue
                key = _norm_key(row.get("order_no"), row.get("sku_id"))
                result[key] = row
    finally:
        cur.close()

    _log("INFO", f"order_no+sku_id 命中（file_name 非空）：{len(result)} 条")
    return result


def fetch_temu_maps(
    conn,
    profit_rows: list[dict[str, Any]],
    *,
    use_line_hash: bool,
) -> tuple[dict[str, dict[str, Any]], dict[tuple[str, str], dict[str, Any]]]:
    by_hash: dict[str, dict[str, Any]] = {}
    if use_line_hash:
        by_hash = fetch_temu_by_line_hashes(
            conn, [_norm_hash(r.get("line_hash")) for r in profit_rows]
        )

    fallback_keys: list[tuple[str, str]] = []
    for row in profit_rows:
        lh = _norm_hash(row.get("line_hash"))
        if use_line_hash and lh and lh in by_hash:
            continue
        key = _norm_key(row.get("ref_no"), row.get("platform_sku"))
        if key[0] and key[1]:
            fallback_keys.append(key)

    by_key = fetch_temu_by_keys(conn, fallback_keys) if fallback_keys else {}
    return by_hash, by_key


def resolve_temu_row(
    profit: dict[str, Any],
    by_hash: dict[str, dict[str, Any]],
    by_key: dict[tuple[str, str], dict[str, Any]],
) -> dict[str, Any] | None:
    lh = _norm_hash(profit.get("line_hash"))
    if lh:
        hit = by_hash.get(lh)
        if hit is not None:
            return hit

    key = _norm_key(profit.get("ref_no"), profit.get("platform_sku"))
    if not key[0] or not key[1]:
        return None
    return by_key.get(key)


def _price_differs(profit: dict[str, Any], temu: dict[str, Any]) -> bool:
    return not _dec_equal(profit.get("order_total_pay"), temu.get("sales_revenue")) or not _dec_equal(
        profit.get("order_goods_base"), temu.get("order_payment")
    )


def build_updates(
    profit_rows: list[dict[str, Any]],
    by_hash: dict[str, dict[str, Any]],
    by_key: dict[tuple[str, str], dict[str, Any]],
) -> tuple[list[tuple[Any, ...]], dict[str, int]]:
    updates: list[tuple[Any, ...]] = []
    stats = {
        "scope_total": len(profit_rows),
        "matched_by_hash": 0,
        "matched_by_key": 0,
        "matched": 0,
        "already_ok": 0,
        "unmatched": 0,
        "pending_update": 0,
    }

    for row in profit_rows:
        lh = _norm_hash(row.get("line_hash"))
        temu = resolve_temu_row(row, by_hash, by_key)
        if temu is None:
            stats["unmatched"] += 1
            continue

        stats["matched"] += 1
        if lh and _norm_hash(temu.get("line_hash")) == lh:
            stats["matched_by_hash"] += 1
        else:
            stats["matched_by_key"] += 1

        if not _price_differs(row, temu):
            stats["already_ok"] += 1
            continue

        stats["pending_update"] += 1
        updates.append(
            (
                temu.get("sales_revenue"),
                temu.get("order_payment"),
                row["id"],
            )
        )

    return updates, stats


def apply_updates(conn, updates: list[tuple[Any, ...]], *, dry_run: bool) -> int:
    if not updates:
        return 0

    if dry_run:
        _log("INFO", f"[dry-run] 将更新 {len(updates)} 行")
        return len(updates)

    sql = (
        f"UPDATE `{PROFIT_TABLE}` SET "
        f"`order_total_pay`=%s, `order_goods_base`=%s "
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


def fetch_invalid_order_total_base(
    conn, *, import_batch: str | None
) -> list[dict[str, Any]]:
    """查询 platform=semitemu、非重发订单且 order_total_base<=0 的利润行。"""
    if import_batch:
        sql = f"""
            SELECT
                p.`id`,
                p.`line_hash`,
                p.`ref_no`,
                p.`platform_sku`,
                p.`order_type`,
                p.`order_total_base`
            FROM `{PROFIT_TABLE}` AS p
            INNER JOIN `{SHIPPED_TABLE}` AS s ON s.`line_hash` = p.`line_hash`
            WHERE s.`import_batch` = %s
              AND p.`platform` = %s
              AND IFNULL(p.`order_type`, '') <> %s
              AND IFNULL(p.`order_total_base`, 0) <= 0
            ORDER BY p.`id`
        """
        params: list[Any] = [import_batch, PLATFORM_SEMITEMU, RESEND_ORDER_TYPE]
    else:
        sql = f"""
            SELECT
                `id`,
                `line_hash`,
                `ref_no`,
                `platform_sku`,
                `order_type`,
                `order_total_base`
            FROM `{PROFIT_TABLE}`
            WHERE `platform` = %s
              AND IFNULL(`order_type`, '') <> %s
              AND IFNULL(`order_total_base`, 0) <= 0
            ORDER BY `id`
        """
        params = [PLATFORM_SEMITEMU, RESEND_ORDER_TYPE]

    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        cur.execute(sql, params)
        return list(cur.fetchall())
    finally:
        cur.close()


def _log_order_total_base_alert(
    invalid_rows: list[dict[str, Any]],
    *,
    scope_desc: str,
    use_color: bool,
) -> None:
    count = len(invalid_rows)
    headline = (
        f"⚠ 数据异常：{scope_desc} 内 platform={PLATFORM_SEMITEMU}、"
        f"order_type≠「{RESEND_ORDER_TYPE}」但 order_total_base<=0 共 {count} 条"
    )
    body_lines: list[str] = []
    for row in invalid_rows[:_CHECK_SAMPLE_LIMIT]:
        body_lines.append(
            f"  id={row.get('id')} ref_no={row.get('ref_no')} "
            f"sku={row.get('platform_sku')} order_type={row.get('order_type')} "
            f"order_total_base={row.get('order_total_base')}"
        )
    if count > len(body_lines):
        body_lines.append(f"  … 另有 {count - len(body_lines)} 条未列出")

    print_banner(
        headline,
        kind="alert",
        use_color=use_color,
        body_lines=body_lines,
        footer="请检查发货源数据或 profit_001 写入过滤逻辑。",
    )


def _log_order_total_base_ok(*, scope_desc: str, use_color: bool) -> None:
    headline = (
        f"✓ 校验通过：{scope_desc} 内 platform={PLATFORM_SEMITEMU}、"
        f"非「{RESEND_ORDER_TYPE}」订单 order_total_base 均 > 0"
    )
    print_banner(headline, kind="ok", use_color=use_color)


def check_order_total_base(
    conn,
    *,
    import_batch: str | None,
    scope_desc: str,
    use_color: bool,
) -> bool:
    """执行后校验 order_total_base；存在异常行时输出醒目提醒。返回是否存在异常。"""
    _log("INFO", f"校验 {scope_desc} 内非重发 semitemu 订单 order_total_base…")
    invalid_rows = fetch_invalid_order_total_base(conn, import_batch=import_batch)
    if not invalid_rows:
        _log_order_total_base_ok(scope_desc=scope_desc, use_color=use_color)
        return False

    _log_order_total_base_alert(invalid_rows, scope_desc=scope_desc, use_color=use_color)
    return True


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="根据 temu_order_item 更新 sales_order_sku_profit 的订单价格字段"
    )
    ap.add_argument(
        "--batch",
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
        "--dry-run",
        action="store_true",
        help="仅统计待更新行数，不执行 UPDATE",
    )
    return ap.parse_args()


def resolve_import_batch(args: argparse.Namespace) -> str | None:
    if args.all:
        return None
    if args.batch:
        return args.batch.strip()
    batch = read_import_batch_from_lock()
    return batch.strip() if batch else None


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    use_color = init_console()

    args = parse_args()
    if args.all and args.batch:
        _log("ERROR", "--all 与 --batch 不能同时使用")
        return 1

    import_batch = resolve_import_batch(args)
    if not args.all and not import_batch:
        _log(
            "ERROR",
            "无法获取批次号，请使用 --batch 指定、--all 全表更新，或确保 run_batch.lock 存在",
        )
        return 1

    scope_desc = "全表" if import_batch is None else f"批次 {import_batch}"
    _log("INFO", f"任务：{TEMU_ORDER_ITEM_TABLE} -> {PROFIT_TABLE}.order_*_pay")
    _log("INFO", f"范围：{scope_desc} (platform={PLATFORM_SEMITEMU})")
    if args.dry_run:
        _log("INFO", "模式：dry-run（不写入数据库）")

    db = get_db_manager(DatabaseConfig())
    conn = db.get_connection()

    try:
        use_line_hash = _temu_has_line_hash(conn)
        if use_line_hash:
            _log("INFO", "temu_order_item 已含 line_hash 列，优先按 line_hash 关联")
        else:
            _log("WARN", "temu_order_item 尚无 line_hash 列，仅按 ref_no+platform_sku 关联")

        profit_rows = fetch_profit_rows(conn, import_batch=import_batch)
        if not profit_rows:
            _log("INFO", "范围内无利润行，结束")
            return 0

        by_hash, by_key = fetch_temu_maps(conn, profit_rows, use_line_hash=use_line_hash)
        updates, stats = build_updates(profit_rows, by_hash, by_key)

        _log("INFO", f"范围内 {PLATFORM_SEMITEMU} 利润行：{stats['scope_total']} 条")
        _log(
            "INFO",
            f"可匹配 temu_order_item (file_name非空)：{stats['matched']} 条"
            f"（line_hash={stats['matched_by_hash']}，回退键={stats['matched_by_key']}）",
        )
        _log("INFO", f"价格字段已正确：{stats['already_ok']} 条")
        _log("INFO", f"无法匹配订单明细：{stats['unmatched']} 条")
        _log("INFO", f"待更新：{stats['pending_update']} 条")

        n_updated = apply_updates(conn, updates, dry_run=args.dry_run)

        if args.dry_run:
            conn.rollback()
            _log("INFO", f"dry-run 完成：预计更新 {n_updated} 条")
        else:
            conn.commit()
            _log("INFO", f"更新完成：实际更新 {n_updated} 条")

        check_order_total_base(
            conn,
            import_batch=import_batch,
            scope_desc=scope_desc,
            use_color=use_color,
        )

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
