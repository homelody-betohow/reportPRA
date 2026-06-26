from __future__ import annotations

"""
从 amz_transaction 表汇总 Amazon FBA 派送相关费用，更新 sales_order_sku_profit.delivery_shipping_base。

业务口径（按你的确认）：
- amz_transaction 中若同 line_hash 有重复行，导入阶段已做金额累计；此处再按 (order_no + platform_sku) 汇总。
- delivery_shipping_base 作为“派送运费（本位币）”，以正数成本写入（因此对 fba_fees 取绝对值口径）。

关联键：
- profit.order_no      <-> amz_transaction.amazon_order_id
- profit.platform_sku  <-> amz_transaction.platform_sku（由 seller_sku 清洗得到，用于对齐订单统计口径）

币种转换：
- 若 amz_transaction.currency = 'EUR'：直接使用金额（EUR）
- 否则优先使用 amz_transaction.fx_rate_cny（转人民币汇率）：EUR = amount * fx_rate_cny / RMB_di_EUR
- 若 fx_rate_cny 缺失：回退用 config/common.py 中的固定汇率（USD/PLN/CZK/HUF/CAD/SEK/RON 等）

用法：
  cd d:\\py-project\\report
  python scripts\\archive\\profit_004_order_delivery_amz.py
  python scripts\\archive\\profit_004_order_delivery_amz.py --batch 20260616_203140
  python scripts\\archive\\profit_004_order_delivery_amz.py --all
  python scripts\\archive\\profit_004_order_delivery_amz.py --dry-run
"""

import argparse
import sys
from datetime import datetime
from decimal import Decimal
from pathlib import Path

_REPORT_ROOT = Path(__file__).resolve().parents[2]
_ARCHIVE_DIR = Path(__file__).resolve().parent
_DATA_IMPORT_DIR = _REPORT_ROOT / "scripts" / "dataImport"
for _p in (_REPORT_ROOT, _ARCHIVE_DIR, _DATA_IMPORT_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from batch_lock import read_import_batch_from_lock  # noqa: E402  # pyright: ignore[reportMissingImports]
from config import common as cfg_common  # noqa: E402
from database.db_connection import DatabaseConfig, get_db_manager  # noqa: E402

PROFIT_TABLE = "sales_order_sku_profit"
AMZ_TXN_TABLE = "amz_transaction"


def _log(level: str, msg: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] [{level}] {msg}")


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="从 amz_transaction 更新利润表派送运费（Amazon）")
    ap.add_argument("--batch", "--import-batch", dest="batch", default=None, help="指定 import_batch（默认从锁文件读取）")
    ap.add_argument("--all", action="store_true", help="不按批次过滤，更新全表 Amazon 可匹配行（谨慎使用）")
    ap.add_argument("--dry-run", action="store_true", help="仅统计不写库")
    return ap.parse_args()


def _build_update_sql(*, where_batch: bool) -> tuple[str, dict[str, object]]:
    """
    返回 (SQL, params)。

    说明：
    - 以 (amazon_order_id, platform_sku) 聚合出 delivery_shipping_base（EUR）
    - 用 UPDATE JOIN 回填到 profit.delivery_shipping_base
    """
    RMB_DI_EUR = Decimal(str(cfg_common.RMB_di_EUR))

    # 固定汇率回退（把“外币 -> EUR”当作乘法系数）
    USD_TO_EUR = Decimal(str(cfg_common.USD_to_EUR))
    PLN_TO_EUR = Decimal(str(getattr(cfg_common, "zl_to_EUR", 0)))
    CZK_TO_EUR = Decimal(str(getattr(cfg_common, "kc_to_EUR", 0)))
    HUF_TO_EUR = Decimal(str(getattr(cfg_common, "Ft_to_EUR", 0)))
    CAD_TO_EUR = Decimal(str(getattr(cfg_common, "CAD_to_EUR", 0)))
    SEK_TO_EUR = Decimal(str(getattr(cfg_common, "kr_to_EUR", 0)))
    RON_TO_EUR = Decimal(str(getattr(cfg_common, "Lei_to_EUR", 0)))

    # 用 CASE 做“金额绝对值 + 币种换算”
    # 注意：fx_rate_cny 为“转人民币汇率”，因此 非EUR 时优先：amount * fx_rate_cny / RMB_di_EUR
    #       若 fx_rate_cny 为空，再用固定汇率回退；都没有则按 0 处理。
    amount_abs = "CASE WHEN t.fba_fees < 0 THEN -t.fba_fees ELSE t.fba_fees END"

    eur_expr = f"""
        CASE
            WHEN t.currency = 'EUR' THEN {amount_abs}
            WHEN t.fx_rate_cny IS NOT NULL THEN ({amount_abs} * t.fx_rate_cny / {RMB_DI_EUR})
            WHEN t.currency IN ('CNY','RMB') THEN ({amount_abs} / {RMB_DI_EUR})
            WHEN t.currency = 'USD' THEN ({amount_abs} * {USD_TO_EUR})
            WHEN t.currency = 'PLN' THEN ({amount_abs} * {PLN_TO_EUR})
            WHEN t.currency = 'CZK' THEN ({amount_abs} * {CZK_TO_EUR})
            WHEN t.currency = 'HUF' THEN ({amount_abs} * {HUF_TO_EUR})
            WHEN t.currency = 'CAD' THEN ({amount_abs} * {CAD_TO_EUR})
            WHEN t.currency = 'SEK' THEN ({amount_abs} * {SEK_TO_EUR})
            WHEN t.currency = 'RON' THEN ({amount_abs} * {RON_TO_EUR})
            ELSE 0
        END
    """.strip()

    # 仅取 fulfillment_channel=FBA 且 amazon_order_id 有值的行（按业务要求）。
    # 仍要求 fba_fees 非空且非 0，避免写入无意义的 0。
    subquery_where = """
        t.amazon_order_id IS NOT NULL
        AND t.amazon_order_id <> ''
        AND t.platform_sku IS NOT NULL
        AND t.platform_sku <> ''
        AND t.fulfillment_channel = 'FBA'
        AND t.fba_fees IS NOT NULL
        AND t.fba_fees <> 0
    """.strip()

    # 双键匹配策略（最大化匹配覆盖）：
    # 1) 优先用 seller_sku（原始值）精确匹配 profit.platform_sku
    # 2) 如果匹配不上，再用 platform_sku（清洗后）匹配 profit 清洗后的 SKU
    #
    # 订单号规范化：取前三段（112-xxxx-xxxx）
    norm_order_no = "SUBSTRING_INDEX(p.order_no, '-', 3)"

    # profit.platform_sku 清洗表达式（与 amz_transaction.platform_sku 生成规则一致）
    norm_profit_sku = (
        "CASE "
        "WHEN LOCATE('amzn.gr.', LOWER(IFNULL(p.platform_sku,''))) > 0 "
        "THEN SUBSTRING_INDEX(SUBSTRING_INDEX(SUBSTRING_INDEX(LOWER(IFNULL(p.platform_sku,'')),'amzn.gr.',-1),'-',1),'_',1) "
        "ELSE SUBSTRING_INDEX(SUBSTRING_INDEX(LOWER(IFNULL(p.platform_sku,'')),'#',1),'BCFBAFL',1) "
        "END"
    )

    # 先从 profit 表筛出“需要匹配”的行，再去 amz_transaction 聚合匹配，最后只更新这些行
    # 过滤条件（按你的要求）：
    #   platform='amazon' AND delivery_shipping_base=0 AND shipping_method='FBA'
    sql = f"""
        UPDATE {PROFIT_TABLE} p
        JOIN (
            SELECT
                p2.line_hash AS line_hash,
                SUM({eur_expr}) AS delivery_base_eur
            FROM {PROFIT_TABLE} p2
            JOIN {AMZ_TXN_TABLE} t
              ON t.amazon_order_id = SUBSTRING_INDEX(p2.order_no, '-', 3)
             AND (
                 LOWER(IFNULL(t.seller_sku,'')) = LOWER(IFNULL(p2.platform_sku,''))
                 OR (
                    CASE
                        WHEN LOCATE('amzn.gr.', LOWER(IFNULL(t.platform_sku,''))) > 0
                        THEN SUBSTRING_INDEX(
                                SUBSTRING_INDEX(
                                    SUBSTRING_INDEX(LOWER(IFNULL(t.platform_sku,'')),'amzn.gr.',-1),
                                    '-', 1
                                ),
                                '_', 1
                             )
                        ELSE SUBSTRING_INDEX(
                                SUBSTRING_INDEX(LOWER(IFNULL(t.platform_sku,'')),'#',1),
                                'BCFBAFL', 1
                             )
                    END
                 ) = (
                    CASE
                        WHEN LOCATE('amzn.gr.', LOWER(IFNULL(p2.platform_sku,''))) > 0
                        THEN SUBSTRING_INDEX(
                                SUBSTRING_INDEX(
                                    SUBSTRING_INDEX(LOWER(IFNULL(p2.platform_sku,'')),'amzn.gr.',-1),
                                    '-', 1
                                ),
                                '_', 1
                             )
                        ELSE SUBSTRING_INDEX(
                                SUBSTRING_INDEX(LOWER(IFNULL(p2.platform_sku,'')),'#',1),
                                'BCFBAFL', 1
                             )
                    END
                 )
             )
            WHERE {subquery_where}
              AND p2.platform = 'amazon'
              AND p2.delivery_shipping_base = 0
              AND p2.shipping_method = 'FBA'
            GROUP BY p2.line_hash
        ) x
          ON x.line_hash = p.line_hash
        SET
            p.delivery_shipping_base = x.delivery_base_eur,
            p.calc_node = 'delivery_flag'
        WHERE p.platform = 'amazon'
          AND p.delivery_shipping_base = 0
          AND p.shipping_method = 'FBA'
    """.rstrip()

    params: dict[str, object] = {}
    if where_batch:
        sql += "\n  AND p.report_hash = %(batch)s"
    return sql, params


def main() -> int:
    args = parse_args()

    batch = args.batch
    if not args.all:
        if not batch:
            batch = read_import_batch_from_lock()
        if not batch:
            _log("ERROR", "未指定 --batch，且无法从锁文件读取 import_batch。")
            return 2

    where_batch = not args.all
    sql, params = _build_update_sql(where_batch=where_batch)
    if where_batch:
        params["batch"] = batch

    # SKU 清洗表达式（用于 dry-run 计数查询）
    def _sku_clean_expr(sku_col: str) -> str:
        return (
            f"CASE "
            f"WHEN LOCATE('amzn.gr.', LOWER(IFNULL({sku_col},''))) > 0 "
            f"THEN SUBSTRING_INDEX(SUBSTRING_INDEX(SUBSTRING_INDEX(LOWER(IFNULL({sku_col},'')),'amzn.gr.',-1),'-',1),'_',1) "
            f"ELSE SUBSTRING_INDEX(SUBSTRING_INDEX(LOWER(IFNULL({sku_col},'')),'#',1),'BCFBAFL',1) "
            f"END"
        )

    db = get_db_manager(DatabaseConfig())
    conn = db.get_connection()
    try:
        with conn.cursor() as cur:
            if args.dry_run:
                # dry-run：用 SELECT 统计可匹配行数（不 UPDATE）
                count_sql = f"""
                    SELECT COUNT(1) AS n
                    FROM {PROFIT_TABLE} p
                    JOIN {AMZ_TXN_TABLE} t
                      ON t.amazon_order_id = SUBSTRING_INDEX(p.order_no, '-', 3)
                     AND (
                         LOWER(IFNULL(t.seller_sku,'')) = LOWER(IFNULL(p.platform_sku,''))
                         OR {_sku_clean_expr('t.platform_sku')} = {_sku_clean_expr('p.platform_sku')}
                     )
                    WHERE p.platform = 'amazon'
                      AND p.delivery_shipping_base = 0
                      AND p.shipping_method = 'FBA'
                      AND t.amazon_order_id IS NOT NULL
                      AND t.amazon_order_id <> ''
                      AND t.platform_sku IS NOT NULL
                      AND t.platform_sku <> ''
                      AND t.fulfillment_channel = 'FBA'
                      AND t.fba_fees IS NOT NULL
                      AND t.fba_fees <> 0
                """.rstrip()
                if where_batch:
                    count_sql += " AND p.report_hash = %(batch)s"
                cur.execute(count_sql, params)
                n = cur.fetchone()
                _log("INFO", f"dry-run：可更新行数={n.get('n') if isinstance(n, dict) else n}")
                _log("INFO", "dry-run：未写库")
                return 0

            _log("INFO", f"开始更新：{PROFIT_TABLE}.delivery_shipping_base（Amazon）")
            if where_batch:
                _log("INFO", f"范围：report_hash={batch}")
            else:
                _log("WARN", "范围：--all 全表更新（请确认这真是你想要的）")

            cur.execute(sql, params)
            affected = cur.rowcount
            conn.commit()
            _log("INFO", f"完成：影响行数={affected}")
            return 0
    except Exception:
        conn.rollback()
        _log("ERROR", "更新失败，已回滚")
        raise
    finally:
        conn.close()


if __name__ == "__main__":
    raise SystemExit(main())

