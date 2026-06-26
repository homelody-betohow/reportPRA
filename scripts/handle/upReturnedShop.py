from __future__ import annotations

r"""
根据 platform_shop 补全 sales_order_returned 中 platform 为空的退件行。

查询条件：
  platform IS NULL AND TRIM(shop_name_en) <> ''

按 shop_name_en 关联 platform_shop，仅当退件行 platform 为空时才处理该行；
platform 已有值（含手工修复）的行整行跳过，不更新任何字段。
其余字段仅当退件行对应列为空时写入。
成功写入的行会将 check_lock 置为 1，标记已完成店铺对接补全。
同一 shop_name_en 对应多条店铺时：先按 shop_alias 收窄；仍多条则优先 shop_status=1；
仍无法唯一确定则取第一条（按 platform_shop.id 排序）直接更新。

用法：
  cd d:\py-project\report
  python scripts\handle\upReturnedShop.py
  python scripts\handle\upReturnedShop.py --dry-run
  python scripts\handle\upReturnedShop.py --limit 100
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

import pymysql.cursors

_REPORT_ROOT = Path(__file__).resolve().parents[2]
_HANDLE_DIR = Path(__file__).resolve().parent
for _p in (_REPORT_ROOT, _HANDLE_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from database.db_connection import DatabaseConfig, get_db_manager  # noqa: E402
from scripts.dataImport.import_common import cell_str  # noqa: E402

RETURNED_TABLE = "sales_order_returned"
PLATFORM_SHOP_TABLE = "platform_shop"

# 退件列 <- platform_shop 列（仅退件侧为空时写入；platform 已有值则整行跳过）
_FILL_MAP: tuple[tuple[str, str, int | None], ...] = (
    ("platform", "platform", 64),
    ("shop_alias", "shop_alias", 128),
    ("platform_site", "platform_site", 64),
    ("sales_owner", "ops_owner", 128),
)

_FETCH_RETURNED_SQL = f"""
SELECT `id`, `shop_name_en`, `shop_alias`, `platform`, `platform_site`, `sales_owner`, `orig_order_no`
FROM `{RETURNED_TABLE}`
WHERE (`platform` IS NULL OR TRIM(`platform`) = '') AND TRIM(`shop_name_en`) <> ''
ORDER BY `orig_order_no`
"""

_PLATFORM_SHOP_COLS: tuple[str, ...] = (
    "id",
    "shop_name_en",
    "shop_alias",
    "shop_name_cn",
    "platform",
    "platform_site",
    "ops_owner",
    "shop_status",
)


def _is_blank(v: Any) -> bool:
    if v is None:
        return True
    return not str(v).strip()


def _norm_shop_name(v: Any) -> str:
    return "" if v is None else str(v).strip()


def _fetch_pending_rows(conn, *, limit: int | None) -> list[dict[str, Any]]:
    sql = _FETCH_RETURNED_SQL
    params: tuple[Any, ...] = ()
    if limit is not None and limit > 0:
        sql += " LIMIT %s"
        params = (limit,)
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        cur.execute(sql, params)
        return list(cur.fetchall() or [])
    finally:
        cur.close()


def _lookup_platform_shops(conn, shop_names: set[str]) -> dict[str, list[dict[str, Any]]]:
    """shop_name_en（TRIM 后）-> platform_shop 行列表。"""
    out: dict[str, list[dict[str, Any]]] = {k: [] for k in shop_names}
    if not shop_names:
        return out

    cols_sql = ", ".join(f"`{c}`" for c in _PLATFORM_SHOP_COLS)
    items = sorted(shop_names)
    chunk = 200
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        for i in range(0, len(items), chunk):
            part = items[i : i + chunk]
            ph = ", ".join(["%s"] * len(part))
            sql = (
                f"SELECT {cols_sql} FROM `{PLATFORM_SHOP_TABLE}` "
                f"WHERE TRIM(`shop_name_en`) IN ({ph})"
            )
            cur.execute(sql, tuple(part))
            for row in cur.fetchall() or []:
                key = _norm_shop_name(row.get("shop_name_en"))
                if key in out:
                    out[key].append(row)
    finally:
        cur.close()
    return out


def _pick_shop_row(
    returned: dict[str, Any],
    candidates: list[dict[str, Any]],
) -> tuple[dict[str, Any] | None, str]:
    if not candidates:
        return None, "no_shop"

    if len(candidates) == 1:
        return candidates[0], "unique"

    ret_alias = _norm_shop_name(returned.get("shop_alias"))
    if ret_alias:
        alias_matched = [
            c
            for c in candidates
            if ret_alias
            in {
                _norm_shop_name(c.get("shop_alias")),
                _norm_shop_name(c.get("shop_name_cn")),
            }
        ]
        if len(alias_matched) == 1:
            return alias_matched[0], "alias"
        if len(alias_matched) > 1:
            candidates = alias_matched

    enabled = [c for c in candidates if int(c.get("shop_status") or 0) == 1]
    if len(enabled) == 1:
        return enabled[0], "shop_status"
    if len(enabled) > 1:
        candidates = enabled

    if len(candidates) == 1:
        return candidates[0], "fallback"

    ordered = sorted(candidates, key=lambda c: int(c.get("id") or 0))
    return ordered[0], "first"


def _build_updates(
    returned: dict[str, Any],
    shop: dict[str, Any],
) -> dict[str, Any]:
    if not _is_blank(returned.get("platform")):
        return {}
    updates: dict[str, Any] = {}
    for ret_col, shop_col, maxlen in _FILL_MAP:
        if not _is_blank(returned.get(ret_col)):
            continue
        val = cell_str(shop.get(shop_col), maxlen)
        if val is not None:
            updates[ret_col] = val
    return updates


def _apply_updates(
    conn,
    pending: list[dict[str, Any]],
    *,
    dry_run: bool,
) -> dict[str, Any]:
    shop_names = {_norm_shop_name(r.get("shop_name_en")) for r in pending}
    shop_names.discard("")
    shop_index = _lookup_platform_shops(conn, shop_names)

    stats: dict[str, Any] = {
        "total": len(pending),
        "updated": 0,
        "skipped_no_shop": 0,
        "picked_first": 0,
        "skipped_no_change": 0,
        "skipped_has_platform": 0,
        "skipped_race_platform": 0,
        "pick_reasons": {},
        "field_counts": {},
        "ambiguous_samples": [],
        "no_shop_samples": [],
    }

    update_rows: list[tuple[int, dict[str, Any]]] = []

    for row in pending:
        if not _is_blank(row.get("platform")):
            stats["skipped_has_platform"] += 1
            continue

        shop_en = _norm_shop_name(row.get("shop_name_en"))
        candidates = shop_index.get(shop_en) or []
        shop, reason = _pick_shop_row(row, candidates)
        stats["pick_reasons"][reason] = stats["pick_reasons"].get(reason, 0) + 1

        if shop is None:
            stats["skipped_no_shop"] += 1
            if len(stats["no_shop_samples"]) < 20:
                stats["no_shop_samples"].append(
                    {
                        "id": row.get("id"),
                        "shop_name_en": shop_en,
                        "orig_order_no": row.get("orig_order_no"),
                    }
                )
            continue

        if reason == "first":
            stats["picked_first"] += 1
            if len(stats["ambiguous_samples"]) < 20:
                stats["ambiguous_samples"].append(
                    {
                        "id": row.get("id"),
                        "shop_name_en": shop_en,
                        "shop_alias": row.get("shop_alias"),
                        "orig_order_no": row.get("orig_order_no"),
                        "candidate_count": len(candidates),
                        "picked_shop_id": shop.get("id"),
                    }
                )

        updates = _build_updates(row, shop)
        if not updates:
            stats["skipped_no_change"] += 1
            continue

        updates["check_lock"] = 1
        update_rows.append((int(row["id"]), updates))

    stats["pending_updates"] = update_rows
    if dry_run or not update_rows:
        stats["updated"] = len(update_rows)
        for _, updates in update_rows:
            for col in updates:
                stats["field_counts"][col] = stats["field_counts"].get(col, 0) + 1
        return stats

    cur = conn.cursor()
    try:
        for rid, updates in update_rows:
            set_sql = ", ".join(f"`{col}`=%s" for col in updates)
            sql = (
                f"UPDATE `{RETURNED_TABLE}` SET {set_sql} "
                f"WHERE `id`=%s AND (`platform` IS NULL OR TRIM(`platform`) = '')"
            )
            cur.execute(sql, tuple(updates[c] for c in updates) + (rid,))
            if cur.rowcount == 0:
                stats["skipped_race_platform"] += 1
                continue
            stats["updated"] += 1
            for col in updates:
                stats["field_counts"][col] = stats["field_counts"].get(col, 0) + 1
    finally:
        cur.close()

    return stats


def _print_stats(stats: dict[str, Any], *, dry_run: bool) -> None:
    print("\n" + "=" * 80)
    print("处理结果" + ("（预览模式，未写入数据库）" if dry_run else ""))
    print("=" * 80)
    print(f"待处理行数: {stats['total']}")
    print(f"将更新行数: {stats['updated']}")
    print(f"无匹配店铺: {stats['skipped_no_shop']}")
    if stats.get("picked_first"):
        print(f"店铺不唯一取首条: {stats['picked_first']}")
    print(f"platform 已有值跳过: {stats['skipped_has_platform']}")
    if stats.get("skipped_race_platform"):
        print(f"写入时 platform 已被占用跳过: {stats['skipped_race_platform']}")
    print(f"无需变更: {stats['skipped_no_change']}")

    if stats.get("pick_reasons"):
        reasons = ", ".join(f"{k}={v}" for k, v in sorted(stats["pick_reasons"].items()))
        print(f"店铺匹配方式: {reasons}")

    if stats.get("field_counts"):
        fields = ", ".join(f"{k}={v}" for k, v in sorted(stats["field_counts"].items()))
        print(f"字段写入计数: {fields}")
        if stats["field_counts"].get("check_lock"):
            print(f"check_lock=1 行数: {stats['field_counts']['check_lock']}")

    if stats.get("no_shop_samples"):
        print("\n未匹配店铺样本（最多 20 条）:")
        for s in stats["no_shop_samples"]:
            print(
                f"  id={s['id']} shop_name_en={s['shop_name_en']!r} "
                f"orig_order_no={s['orig_order_no']!r}"
            )

    if stats.get("ambiguous_samples"):
        print("\n店铺不唯一取首条样本（最多 20 条）:")
        for s in stats["ambiguous_samples"]:
            print(
                f"  id={s['id']} shop_name_en={s['shop_name_en']!r} "
                f"shop_alias={s['shop_alias']!r} candidates={s['candidate_count']} "
                f"picked_shop_id={s['picked_shop_id']} "
                f"orig_order_no={s['orig_order_no']!r}"
            )

    pending = stats.get("pending_updates") or []
    if dry_run and pending:
        print("\n更新预览（前 5 条）:")
        for rid, updates in pending[:5]:
            cols = ", ".join(f"{k}={v!r}" for k, v in updates.items())
            print(f"  id={rid}: {cols}")
        if len(pending) > 5:
            print(f"  ... 还有 {len(pending) - 5} 条")


def main() -> None:
    ap = argparse.ArgumentParser(
        description="按 platform_shop 补全 sales_order_returned 的 platform 等字段"
    )
    ap.add_argument(
        "--dry-run",
        action="store_true",
        help="仅统计与预览，不写入数据库",
    )
    ap.add_argument(
        "--limit",
        type=int,
        default=None,
        help="最多处理多少条退件行（调试用）",
    )
    args = ap.parse_args()

    print("=" * 80)
    print("sales_order_returned <- platform_shop 店铺补全")
    print("=" * 80)
    print(f"开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    if args.dry_run:
        print("模式: 预览（--dry-run）")
    if args.limit:
        print(f"限制行数: {args.limit}")

    db_manager = get_db_manager()
    conn = db_manager.get_connection()
    try:
        pending = _fetch_pending_rows(conn, limit=args.limit)
        print(f"查询到待补全行数: {len(pending)}")
        if not pending:
            print("没有需要处理的记录")
            return

        stats = _apply_updates(conn, pending, dry_run=args.dry_run)
        if not args.dry_run:
            conn.commit()
        _print_stats(stats, dry_run=args.dry_run)
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

    print(f"\n结束时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")


if __name__ == "__main__":
    main()
