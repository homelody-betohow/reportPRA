from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

try:
    from dotenv import load_dotenv
except Exception:  # pragma: no cover
    load_dotenv = None  # type: ignore[assignment]

# Same import strategy as other scripts in this folder.
_THIS_DIR = Path(__file__).resolve().parent
_V2_DIR = _THIS_DIR.parent
sys.path.insert(0, str(_V2_DIR))

from db import connect, load_db_config  # noqa: E402
from logger import get_logger, setup_stdout_utf8  # noqa: E402


_logger = get_logger("SUMMARY")


def upsert_daily_from_details(cur, provider: str) -> int:
    """
    从明细表聚合并 upsert 到日汇总表（按 日 + 仓 + 币种 + SKU）。
    """
    sql = """
        INSERT INTO warehouse_rent_daily (
          provider, charge_date, warehouse_code, warehouse_name, currency, sku,
          amount_total, discount_total, qty_total, volume_total_m3, line_count
        )
        SELECT
          provider,
          charge_date,
          warehouse_code,
          warehouse_name,
          currency,
          COALESCE(sku, '') AS sku,
          ROUND(SUM(amount), 6) AS amount_total,
          ROUND(SUM(COALESCE(discount_amount, 0)), 6) AS discount_total,
          ROUND(SUM(COALESCE(qty, 0)), 6) AS qty_total,
          ROUND(SUM(COALESCE(volume_m3, 0)), 6) AS volume_total_m3,
          COUNT(*) AS line_count
        FROM warehouse_rent_detail
        WHERE provider = %s
        GROUP BY provider, charge_date, warehouse_code, warehouse_name, currency, COALESCE(sku, '')
        ON DUPLICATE KEY UPDATE
          amount_total = VALUES(amount_total),
          discount_total = VALUES(discount_total),
          qty_total = VALUES(qty_total),
          volume_total_m3 = VALUES(volume_total_m3),
          line_count = VALUES(line_count)
    """
    cur.execute(sql, (provider,))
    return int(cur.rowcount or 0)


def run(provider: str | None = None) -> None:
    cfg = load_db_config()
    conn = connect(cfg)
    try:
        cur = conn.cursor()
        if provider:
            _logger.info(f"开始汇总 provider={provider} -> warehouse_rent_daily")
            upsert_daily_from_details(cur, provider)
        else:
            for p in ("HY", "4PX", "AMZ_FBA"):
                _logger.info(f"开始汇总 provider={p} -> warehouse_rent_daily")
                upsert_daily_from_details(cur, p)
        conn.commit()
        _logger.info("汇总完成（已提交）")
    except Exception:
        _logger.error("发生异常，准备回滚")
        conn.rollback()
        raise
    finally:
        _logger.info("关闭数据库连接")
        conn.close()


def main() -> None:
    setup_stdout_utf8()
    if load_dotenv:
        load_dotenv()

    parser = argparse.ArgumentParser(description="从 warehouse_rent_detail 聚合写入 warehouse_rent_daily（按日+SKU）")
    parser.add_argument("--provider", default="", help="可选：只汇总某个 provider（HY/4PX），留空则都汇总")
    args = parser.parse_args()
    run(provider=args.provider.strip() or None)


if __name__ == "__main__":
    main()

