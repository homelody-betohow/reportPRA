from __future__ import annotations

import argparse
import sys
from pathlib import Path

try:
    from dotenv import load_dotenv
except Exception:  # pragma: no cover
    load_dotenv = None  # type: ignore[assignment]

# NOTE:
# This script lives under `python/v2/warehouse-rent/` (hyphen in folder name),
# so we can't use package-relative imports reliably.
# We import:
# - `python/v2/db.py` via sys.path
# - sibling import modules via sys.path (this dir)
_THIS_DIR = Path(__file__).resolve().parent
# `.../python/v2/warehouse-rent/` -> parent is `.../python/v2/`
_V2_DIR = _THIS_DIR.parent
sys.path.insert(0, str(_V2_DIR))
sys.path.insert(0, str(_THIS_DIR))

from db import connect, load_db_config  # noqa: E402
from summary_detail_to_daily import upsert_daily_from_details  # noqa: E402
from import_provider_amazon_fba_detail import import_amazon_fba_file  # noqa: E402
from import_provider_4px_detail import import_4px_file  # noqa: E402
from import_provider_hy_detail import import_hy_file  # noqa: E402
from logger import get_logger, setup_stdout_utf8  # noqa: E402

log = get_logger("ALL").info


def run_import(daily_dir: Path) -> None:
    cfg = load_db_config()
    conn = connect(cfg)
    try:
        cur = conn.cursor()

        # Import files
        hy_files = sorted(daily_dir.glob("鸿羽*仓-仓租明细*.xlsx"))
        px_files = sorted(daily_dir.glob("4PX*仓-仓租明细*.xlsx"))
        fba_files = sorted(daily_dir.glob("FBA仓租明细*.xlsx"))

        counts = {"HY": {"files": 0, "rows": 0}, "4PX": {"files": 0, "rows": 0}, "AMZ_FBA": {"files": 0, "rows": 0}}

        log(f"扫描目录：{daily_dir}")
        log(f"发现 HY 文件数：{len(hy_files)}；4PX 文件数：{len(px_files)}；AMZ_FBA 文件数：{len(fba_files)}")

        for fp in hy_files:
            log(f"开始导入 HY 仓租明细：{fp.name}")
            counts["HY"]["files"] += 1
            counts["HY"]["rows"] += import_hy_file(cur, fp)
            log(f"完成导入 HY 仓租明细：{fp.name}（累计写入/尝试写入行数：{counts['HY']['rows']}）")

        for fp in px_files:
            log(f"开始导入 4PX 仓租明细：{fp.name}")
            counts["4PX"]["files"] += 1
            counts["4PX"]["rows"] += import_4px_file(cur, fp)
            log(f"完成导入 4PX 仓租明细：{fp.name}（累计写入/尝试写入行数：{counts['4PX']['rows']}）")

        for fp in fba_files:
            log(f"开始导入 AMZ_FBA 仓租明细：{fp.name}")
            counts["AMZ_FBA"]["files"] += 1
            counts["AMZ_FBA"]["rows"] += import_amazon_fba_file(cur, fp)
            log(f"完成导入 AMZ_FBA 仓租明细：{fp.name}（累计写入/尝试写入行数：{counts['AMZ_FBA']['rows']}）")

        # Refresh daily aggregates per provider
        log("开始汇总写入 warehouse_rent_daily（HY）")
        upsert_daily_from_details(cur, "HY")
        log("开始汇总写入 warehouse_rent_daily（4PX）")
        upsert_daily_from_details(cur, "4PX")
        log("开始汇总写入 warehouse_rent_daily（AMZ_FBA）")
        upsert_daily_from_details(cur, "AMZ_FBA")

        conn.commit()
        log(f"导入完成（已提交）：{counts}")
    except Exception:
        get_logger("ALL").error("发生异常，准备回滚")
        conn.rollback()
        raise
    finally:
        log("关闭数据库连接")
        conn.close()


def main() -> None:
    setup_stdout_utf8()
    if load_dotenv:
        load_dotenv()

    parser = argparse.ArgumentParser(description="从 python/excel/daily 导入 HY/4PX 仓租到 MySQL，并生成按日+SKU 汇总表")
    parser.add_argument(
        "--daily-dir",
        default=r"python/excel/daily",
        help="daily Excel 目录",
    )
    args = parser.parse_args()

    daily_dir = Path(args.daily_dir)
    if not daily_dir.exists():
        raise FileNotFoundError(f"找不到目录：{daily_dir}")

    run_import(daily_dir=daily_dir)


if __name__ == "__main__":
    main()

