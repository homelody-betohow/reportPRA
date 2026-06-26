from __future__ import annotations

"""
根据 platform_shop 表更新 sales_order_sku_profit 的 market_region、market_code；
并将 warehouse_name 含「分销」的行的 distribution_lev 设为 1。

关联键：platform + platform_site + shop_name_en（与 platform_shop.shop_hash 业务维度一致）。

默认仅处理本批 import_batch 对应的利润行（经 sales_order_shipped.line_hash 关联）；
加 --all 可更新全表可匹配行。

用法：
  cd d:\\py-project\\report
  python scripts\\archive\\profit_002_order_market.py
  python scripts\\archive\\profit_002_order_market.py --batch 20260616_203140
  python scripts\\archive\\profit_002_order_market.py --all
  python scripts\\archive\\profit_002_order_market.py --dry-run
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

import pymysql.cursors

_REPORT_ROOT = Path(__file__).resolve().parents[2]
_ARCHIVE_DIR = Path(__file__).resolve().parent
_DATA_IMPORT_DIR = _REPORT_ROOT / "scripts" / "dataImport"

for _p in (_REPORT_ROOT, _ARCHIVE_DIR, _DATA_IMPORT_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from batch_lock import read_import_batch_from_lock  # noqa: E402
from database.db_connection import DatabaseConfig, get_db_manager  # noqa: E402

PROFIT_TABLE = "sales_order_sku_profit"
SHIPPED_TABLE = "sales_order_shipped"
PLATFORM_SHOP_TABLE = "platform_shop"
_COLLATE = "utf8mb4_unicode_ci"

_MARKET_COLUMNS = ("market_region", "market_code")
_DISTRIBUTION_WH_MARK = "分销"
_DISTRIBUTION_LEV = 1


def _collate_trim(expr: str) -> str:
    return f"TRIM({expr}) COLLATE {_COLLATE}"


def _collate_trim_ifnull(expr: str, default: str = "''") -> str:
    return f"TRIM(IFNULL({expr}, {default})) COLLATE {_COLLATE}"


def _log(level: str, msg: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] [{level}] {msg}")


def _ensure_required_columns(conn) -> None:
    """确认 profit 表已存在 market_region / market_code / distribution_lev 列。"""
    cur = conn.cursor(pymysql.cursors.Cursor)
    try:
        cur.execute(f"SHOW COLUMNS FROM `{PROFIT_TABLE}`")
        existing = {str(row[0]).strip() for row in cur.fetchall() or []}
    finally:
        cur.close()

    required = (*_MARKET_COLUMNS, "distribution_lev")
    missing = [c for c in required if c not in existing]
    if missing:
        raise RuntimeError(
            f"表 `{PROFIT_TABLE}` 缺少列 {missing}，请先执行 ALTER TABLE 增加对应列"
        )


def _shop_join_sql(*, profit_alias: str, shop_alias: str) -> str:
    return f"""
    {_collate_trim(f'`{profit_alias}`.`platform`')} = {_collate_trim(f'`{shop_alias}`.`platform`')}
    AND {_collate_trim_ifnull(f'`{profit_alias}`.`platform_site`')} = {_collate_trim_ifnull(f'`{shop_alias}`.`platform_site`')}
    AND {_collate_trim_ifnull(f'`{profit_alias}`.`shop_name_en`')} = {_collate_trim_ifnull(f'`{shop_alias}`.`shop_name_en`')}
  """.strip()


def _batch_filter_sql(*, profit_alias: str, shipped_alias: str) -> str:
    return (
        f"`{shipped_alias}`.`import_batch` = %s "
        f"AND `{shipped_alias}`.`line_hash` COLLATE {_COLLATE} = `{profit_alias}`.`line_hash` COLLATE {_COLLATE}"
    )


def _value_changed_sql(*, profit_alias: str, shop_alias: str) -> str:
    return (
        f"IFNULL(`{profit_alias}`.`market_region`, '') <> IFNULL(`{shop_alias}`.`market_region`, '') "
        f"OR IFNULL(`{profit_alias}`.`market_code`, '') <> IFNULL(`{shop_alias}`.`market_code`, '')"
    )


def _build_count_sql(*, import_batch: str | None) -> str:
    shop_join = _shop_join_sql(profit_alias="p", shop_alias="ps")
    changed = _value_changed_sql(profit_alias="p", shop_alias="ps")

    if import_batch:
        batch_filter = _batch_filter_sql(profit_alias="p", shipped_alias="s")
        return f"""
            SELECT COUNT(*) AS cnt
            FROM `{PROFIT_TABLE}` AS p
            INNER JOIN `{SHIPPED_TABLE}` AS s ON {batch_filter}
            INNER JOIN `{PLATFORM_SHOP_TABLE}` AS ps ON {shop_join}
            WHERE {changed}
        """

    return f"""
        SELECT COUNT(*) AS cnt
        FROM `{PROFIT_TABLE}` AS p
        INNER JOIN `{PLATFORM_SHOP_TABLE}` AS ps ON {shop_join}
        WHERE {changed}
    """


def _build_update_sql(*, import_batch: str | None) -> str:
    shop_join = _shop_join_sql(profit_alias="p", shop_alias="ps")
    changed = _value_changed_sql(profit_alias="p", shop_alias="ps")

    if import_batch:
        batch_filter = _batch_filter_sql(profit_alias="p", shipped_alias="s")
        return f"""
            UPDATE `{PROFIT_TABLE}` AS p
            INNER JOIN `{SHIPPED_TABLE}` AS s ON {batch_filter}
            INNER JOIN `{PLATFORM_SHOP_TABLE}` AS ps ON {shop_join}
            SET
              p.`market_region` = ps.`market_region`,
              p.`market_code` = ps.`market_code`
            WHERE {changed}
        """

    return f"""
        UPDATE `{PROFIT_TABLE}` AS p
        INNER JOIN `{PLATFORM_SHOP_TABLE}` AS ps ON {shop_join}
        SET
          p.`market_region` = ps.`market_region`,
          p.`market_code` = ps.`market_code`
        WHERE {changed}
    """


def _build_stats_sql(*, import_batch: str | None) -> dict[str, str]:
    shop_join = _shop_join_sql(profit_alias="p", shop_alias="ps")

    if import_batch:
        batch_filter = _batch_filter_sql(profit_alias="p", shipped_alias="s")
        scope_from = (
            f"`{PROFIT_TABLE}` AS p "
            f"INNER JOIN `{SHIPPED_TABLE}` AS s ON {batch_filter}"
        )
    else:
        scope_from = f"`{PROFIT_TABLE}` AS p"

    return {
        "scope_total": f"SELECT COUNT(*) FROM {scope_from}",
        "matched": (
            f"SELECT COUNT(*) FROM {scope_from} "
            f"INNER JOIN `{PLATFORM_SHOP_TABLE}` AS ps ON {shop_join}"
        ),
        "already_ok": (
            f"SELECT COUNT(*) FROM {scope_from} "
            f"INNER JOIN `{PLATFORM_SHOP_TABLE}` AS ps ON {shop_join} "
            f"WHERE NOT ({_value_changed_sql(profit_alias='p', shop_alias='ps')})"
        ),
        "unmatched": (
            f"SELECT COUNT(*) FROM {scope_from} "
            f"LEFT JOIN `{PLATFORM_SHOP_TABLE}` AS ps ON {shop_join} "
            f"WHERE ps.`id` IS NULL"
        ),
    }


def _fetch_scalar(conn, sql: str, params: tuple[Any, ...] = ()) -> int:
    cur = conn.cursor(pymysql.cursors.Cursor)
    try:
        cur.execute(sql, params)
        row = cur.fetchone()
        return int(row[0] or 0) if row else 0
    finally:
        cur.close()


def collect_stats(conn, *, import_batch: str | None) -> dict[str, int]:
    stats_sql = _build_stats_sql(import_batch=import_batch)
    params: tuple[Any, ...] = (import_batch,) if import_batch else ()

    scope_total = _fetch_scalar(conn, stats_sql["scope_total"], params)
    matched = _fetch_scalar(conn, stats_sql["matched"], params)
    already_ok = _fetch_scalar(conn, stats_sql["already_ok"], params)
    unmatched = _fetch_scalar(conn, stats_sql["unmatched"], params)
    pending = _fetch_scalar(conn, _build_count_sql(import_batch=import_batch), params)

    return {
        "scope_total": scope_total,
        "matched": matched,
        "already_ok": already_ok,
        "unmatched": unmatched,
        "pending_update": pending,
    }


def _distribution_pending_filter_sql(*, profit_alias: str) -> str:
    return (
        f"`{profit_alias}`.`warehouse_name` LIKE %s "
        f"AND IFNULL(`{profit_alias}`.`distribution_lev`, 0) <> {_DISTRIBUTION_LEV}"
    )


def _distribution_params(import_batch: str | None) -> tuple[Any, ...]:
    like = f"%{_DISTRIBUTION_WH_MARK}%"
    if import_batch:
        return (import_batch, like)
    return (like,)


def _build_distribution_count_sql(*, import_batch: str | None) -> str:
    pending = _distribution_pending_filter_sql(profit_alias="p")

    if import_batch:
        batch_filter = _batch_filter_sql(profit_alias="p", shipped_alias="s")
        return f"""
            SELECT COUNT(*) AS cnt
            FROM `{PROFIT_TABLE}` AS p
            INNER JOIN `{SHIPPED_TABLE}` AS s ON {batch_filter}
            WHERE {pending}
        """

    return f"""
        SELECT COUNT(*) AS cnt
        FROM `{PROFIT_TABLE}` AS p
        WHERE {pending}
    """


def _build_distribution_update_sql(*, import_batch: str | None) -> str:
    pending = _distribution_pending_filter_sql(profit_alias="p")

    if import_batch:
        batch_filter = _batch_filter_sql(profit_alias="p", shipped_alias="s")
        return f"""
            UPDATE `{PROFIT_TABLE}` AS p
            INNER JOIN `{SHIPPED_TABLE}` AS s ON {batch_filter}
            SET p.`distribution_lev` = {_DISTRIBUTION_LEV}
            WHERE {pending}
        """

    return f"""
        UPDATE `{PROFIT_TABLE}` AS p
        SET p.`distribution_lev` = {_DISTRIBUTION_LEV}
        WHERE {pending}
    """


def update_distribution_lev(
    conn,
    *,
    import_batch: str | None,
    dry_run: bool,
) -> int:
    params = _distribution_params(import_batch)
    pending = _fetch_scalar(conn, _build_distribution_count_sql(import_batch=import_batch), params)

    if pending <= 0:
        _log("INFO", "没有需要标记分销等级的行")
        return 0

    if dry_run:
        _log("INFO", f"[dry-run] 将标记分销等级 {pending} 行")
        return pending

    sql = _build_distribution_update_sql(import_batch=import_batch)
    cur = conn.cursor(pymysql.cursors.Cursor)
    try:
        cur.execute(sql, params)
        return int(cur.rowcount or 0)
    finally:
        cur.close()


def update_market_fields(
    conn,
    *,
    import_batch: str | None,
    dry_run: bool,
) -> int:
    params: tuple[Any, ...] = (import_batch,) if import_batch else ()
    pending = _fetch_scalar(conn, _build_count_sql(import_batch=import_batch), params)

    if pending <= 0:
        _log("INFO", "没有需要更新的行")
        return 0

    if dry_run:
        _log("INFO", f"[dry-run] 将更新 {pending} 行")
        return pending

    sql = _build_update_sql(import_batch=import_batch)
    cur = conn.cursor(pymysql.cursors.Cursor)
    try:
        cur.execute(sql, params)
        return int(cur.rowcount or 0)
    finally:
        cur.close()


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description=(
            "根据 platform_shop 更新 sales_order_sku_profit 的 market_region、market_code；"
            "并将 warehouse_name 含「分销」的行的 distribution_lev 设为 1"
        )
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
    _log("INFO", f"任务1：{PLATFORM_SHOP_TABLE} -> {PROFIT_TABLE}.market_*")
    _log("INFO", f"任务2：warehouse_name 含「{_DISTRIBUTION_WH_MARK}」-> {PROFIT_TABLE}.distribution_lev={_DISTRIBUTION_LEV}")
    _log("INFO", f"范围：{scope_desc}")
    if args.dry_run:
        _log("INFO", "模式：dry-run（不写入数据库）")

    db = get_db_manager(DatabaseConfig())
    conn = db.get_connection()

    try:
        _ensure_required_columns(conn)

        stats = collect_stats(conn, import_batch=import_batch)
        _log("INFO", f"范围内利润行：{stats['scope_total']} 条")
        _log("INFO", f"可匹配 platform_shop：{stats['matched']} 条")
        _log("INFO", f"市场字段已正确：{stats['already_ok']} 条")
        _log("INFO", f"无法匹配店铺：{stats['unmatched']} 条")
        _log("INFO", f"待更新市场字段：{stats['pending_update']} 条")

        n_market_updated = update_market_fields(
            conn,
            import_batch=import_batch,
            dry_run=args.dry_run,
        )

        dist_params = _distribution_params(import_batch)
        n_dist_pending = _fetch_scalar(
            conn, _build_distribution_count_sql(import_batch=import_batch), dist_params
        )
        _log("INFO", f"待标记分销等级：{n_dist_pending} 条")

        n_dist_updated = update_distribution_lev(
            conn,
            import_batch=import_batch,
            dry_run=args.dry_run,
        )

        if args.dry_run:
            conn.rollback()
            _log("INFO", f"dry-run 完成：预计更新市场字段 {n_market_updated} 条，分销等级 {n_dist_updated} 条")
        else:
            conn.commit()
            _log("INFO", f"更新完成：市场字段 {n_market_updated} 条，分销等级 {n_dist_updated} 条")

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
