from __future__ import annotations

"""
SKU 清洗：根据 warehouse_sku 清洗规则回填 product_sku_mapping.product_sku。

背景：
- 订单导入 (import_order_shipped) 时 product_sku 默认 = warehouse_sku；但仓库 SKU 在
  AMZN 历史数据里掺杂前缀/后缀（如 AMZN.GR.xxx、xxx#FBA、xxx_FBAUS-UEHPSW…）。
- 本脚本复用 v1 的清洗规则（python/v1/J_AMZ_仓租/月报/J1_计算_AMZ仓租.py 的 extract_values），
  仅对 product_sku 进行覆写；warehouse_sku 与 line_hash 保持原样（line_hash 不参与 product_sku）。

清洗规则（与 v1 等价，大小写都识别 amzn.gr.）：
  1) warehouse_sku 含 'amzn.gr.'（不区分大小写）：
       去掉 'amzn.gr.' 前缀，再依次按 '-' / '_' 各取首段
       例：'AMZN.GR.U02033010_FBAUS-UEHPSW2956BRX-VG' -> 'U02033010'
  2) 否则：
       依次按 '#' / 'BCFBAFL' / 'FBFBAFL' 各取首段
       例：'E51033005#FBA' -> 'E51033005'
       例：'E51033005BCFBAFL' -> 'E51033005'
  3) 清洗后为空字符串 -> None（保留原 product_sku 不动）

用法（在 python/ 目录下）：
  python v2/products/run_sku_clean.py                       # 实跑：清洗全表
  python v2/products/run_sku_clean.py --dry-run             # 不写库，只统计 + 打印前若干差异行
  python v2/products/run_sku_clean.py --dry-run --limit 50  # 调试：只看前 50 条差异
  python v2/products/run_sku_clean.py --sku AMZN.GR.U02033010_FBAUS-UEHPSW2956BRX-VG
                                                            # 单 SKU 试算，不连库（演示清洗结果）
"""

import argparse
import re
import sys
from pathlib import Path
from typing import Any, Iterable

_PROD_DIR = Path(__file__).resolve().parent
_V2_DIR = _PROD_DIR.parent
_ORDERS_DIR = _V2_DIR / "orders"
_WR_DIR = _V2_DIR / "warehouse-rent"
for _p in (_PROD_DIR, _V2_DIR, _ORDERS_DIR, _WR_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from db import connect, load_db_config  # noqa: E402
from logger import get_logger, setup_stdout_utf8  # noqa: E402

_LOG = get_logger("SKU-CLEAN")

MAPPING_TABLE = "product_sku_mapping"

# 不区分大小写匹配 amzn.gr.（v1 仅小写，本脚本兼容 AMZN.GR. 等大写写法）
_AMZN_GR_RE = re.compile(r"amzn\.gr\.", re.IGNORECASE)


def clean_warehouse_sku(s: Any) -> str | None:
    """
    将仓库 SKU 清洗成可用于 product_sku 的"干净 SKU"。

    与 v1 (J1_计算_AMZ仓租.py / extract_values) 等价，并修正大小写敏感问题：
      - 'amzn.gr.' 前缀（不区分大小写）：取 'amzn.gr.' 之后，再按 '-' / '_' 各取首段
      - 否则：依次按 '#' / 'BCFBAFL' / 'FBFBAFL' 各取首段
    返回 None 表示输入空或清洗结果为空（调用方应保留原值或跳过）。
    """
    if s is None:
        return None
    text = str(s).strip()
    if not text:
        return None

    if _AMZN_GR_RE.search(text):
        # 按 amzn.gr.（不区分大小写）切分，取最后一段
        tail = _AMZN_GR_RE.split(text)[-1]
        cleaned = tail.split("-", 1)[0].split("_", 1)[0].strip()
    else:
        cleaned = (
            text.split("#", 1)[0]
            .split("BCFBAFL", 1)[0]
            .split("FBFBAFL", 1)[0]
            .strip()
        )

    return cleaned or None


# ============================================================
# DB 操作
# ============================================================
def _fetch_mapping_rows(conn) -> list[tuple[int, str, str | None]]:
    """
    读取 product_sku_mapping 全部行：(id, warehouse_sku, product_sku)。
    表规模约万级，一次读完即可（如未来超大可再加分页）。
    """
    cur = conn.cursor()
    cur.execute(
        f"SELECT `id`, `warehouse_sku`, `product_sku` "
        f"FROM `{MAPPING_TABLE}` "
        f"WHERE `warehouse_sku` IS NOT NULL AND `warehouse_sku` <> '' "
        f"ORDER BY `id`"
    )
    rows = [(int(r[0]), str(r[1]), (None if r[2] is None else str(r[2]))) for r in cur.fetchall()]
    cur.close()
    return rows


def _update_product_sku_by_id(conn, pairs: Iterable[tuple[int, str]], chunk: int = 500) -> int:
    """
    按主键 id 批量更新 product_sku（仅这一列），避免触发 ON UPDATE 之外的副作用。
    """
    items = list(pairs)
    if not items:
        return 0
    sql = f"UPDATE `{MAPPING_TABLE}` SET `product_sku`=%s WHERE `id`=%s"
    cur = conn.cursor()
    affected = 0
    try:
        for i in range(0, len(items), chunk):
            batch = items[i : i + chunk]
            # mysql.connector 的 executemany 参数顺序需与 SQL 占位符一致
            cur.executemany(sql, [(new_sku, _id) for (_id, new_sku) in batch])
            rc = cur.rowcount
            if rc is not None and rc > 0:
                affected += rc
    finally:
        cur.close()
    return affected


# ============================================================
# 主流程
# ============================================================
def run_sku_clean(
    conn,
    *,
    dry_run: bool = False,
    limit: int | None = None,
) -> dict[str, int]:
    """
    扫描 product_sku_mapping，对 warehouse_sku 应用清洗规则得到 cleaned，
    与现有 product_sku 比对，差异行触发 UPDATE。

    Returns:
        {"scanned": ..., "diff": ..., "updated": ..., "skipped_empty_clean": ...}
    """
    rows = _fetch_mapping_rows(conn)
    _LOG.info(f"读取 {MAPPING_TABLE}：{len(rows)} 行（含 warehouse_sku 非空）")

    diffs: list[tuple[int, str, str | None, str]] = []  # (id, warehouse_sku, old_product_sku, new_product_sku)
    skipped_empty_clean = 0
    for _id, wh_sku, prod_sku in rows:
        cleaned = clean_warehouse_sku(wh_sku)
        if cleaned is None:
            skipped_empty_clean += 1
            continue
        # 旧值视空字符串与 None 等价
        old_norm = prod_sku.strip() if isinstance(prod_sku, str) else prod_sku
        if not old_norm:
            old_norm = None
        if cleaned != old_norm:
            diffs.append((_id, wh_sku, prod_sku, cleaned))

    _LOG.info(
        f"扫描完成：总行={len(rows)} 需更新={len(diffs)} 清洗为空跳过={skipped_empty_clean}"
    )

    # 打印若干样例，便于人工核对
    sample_n = min(len(diffs), 10 if not dry_run else (limit or 20))
    if sample_n > 0:
        _LOG.info(f"差异样例（共 {len(diffs)} 条，展示前 {sample_n} 条）：")
        _LOG.info(f"  {'id':>8}  {'warehouse_sku':<48}  {'old_product_sku':<24} -> new_product_sku")
        for _id, wh_sku, old, new_sku in diffs[:sample_n]:
            _LOG.info(f"  {_id:>8}  {wh_sku[:48]:<48}  {(old or '<NULL>')[:24]:<24} -> {new_sku}")

    if dry_run:
        _LOG.warn("--dry-run：不写库")
        return {
            "scanned": len(rows),
            "diff": len(diffs),
            "updated": 0,
            "skipped_empty_clean": skipped_empty_clean,
        }

    # 实跑（如果指定 --limit，则只更新前 N 条）
    to_update = diffs if limit is None else diffs[:limit]
    if not to_update:
        _LOG.info("无需更新，任务结束")
        return {
            "scanned": len(rows),
            "diff": len(diffs),
            "updated": 0,
            "skipped_empty_clean": skipped_empty_clean,
        }

    _LOG.info(f"准备 UPDATE：{len(to_update)} 行（按 id 批量）")
    affected = _update_product_sku_by_id(conn, ((_id, new_sku) for (_id, _, _, new_sku) in to_update))
    _LOG.info(f"UPDATE 完成：受影响 {affected} 行")
    return {
        "scanned": len(rows),
        "diff": len(diffs),
        "updated": affected,
        "skipped_empty_clean": skipped_empty_clean,
    }


def main() -> int:
    setup_stdout_utf8()
    ap = argparse.ArgumentParser(
        description="按 warehouse_sku 清洗规则回填 product_sku_mapping.product_sku"
    )
    ap.add_argument("--dry-run", action="store_true", help="不写库，只统计 + 打印差异样例")
    ap.add_argument(
        "--limit",
        type=int,
        default=None,
        metavar="N",
        help="最多更新前 N 条差异（调试用，配合 --dry-run 时控制样例打印数）",
    )
    ap.add_argument(
        "--sku",
        default=None,
        metavar="WAREHOUSE_SKU",
        help="单 SKU 试算模式：不连库，直接打印清洗结果（用于验证规则）",
    )
    args = ap.parse_args()

    # === 单 SKU 试算（离线，无需数据库）===
    if args.sku is not None:
        cleaned = clean_warehouse_sku(args.sku)
        _LOG.info(f"输入 warehouse_sku = {args.sku!r}")
        _LOG.info(f"清洗后 product_sku = {cleaned!r}")
        return 0

    _LOG.info("=" * 60)
    _LOG.info(f"任务：清洗 {MAPPING_TABLE}.product_sku（依据 warehouse_sku）")
    cfg = load_db_config()
    _LOG.info(f"连接数据库：host={cfg.host} port={cfg.port} database={cfg.database} user={cfg.user}")
    conn = connect(cfg)
    try:
        stats = run_sku_clean(conn, dry_run=args.dry_run, limit=args.limit)
        if args.dry_run:
            conn.rollback()
            _LOG.info(
                f"已回滚（dry-run）：扫描={stats['scanned']} 差异={stats['diff']} "
                f"清洗为空跳过={stats['skipped_empty_clean']}"
            )
        else:
            conn.commit()
            _LOG.info(
                f"已提交事务：扫描={stats['scanned']} 差异={stats['diff']} "
                f"更新={stats['updated']} 清洗为空跳过={stats['skipped_empty_clean']}"
            )
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
