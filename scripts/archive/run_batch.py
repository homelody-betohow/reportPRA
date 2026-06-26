from __future__ import annotations

"""
按 import_batch 顺序执行 archive 利润归档流水线。

批次号默认从 dataImport/run_batch.lock 读取（须先完成 Excel 导入）；
也可用 --batch 手动指定。

顺序：
  1. profit_001_order_sku.py      发货 -> 利润表初始化
  2. profit_002_order_market.py   回填市场区域/编码
  3. profit_002_order_price.py    覆盖 Temu 价格（RPA 明细）
  4. profit_003_order_first.py    重算头程/关税
  5. profit_004_order_delivery_amz.py 更新派送运费（Amazon FBA）

用法：
  cd d:\\path\\to\\rpa-task
  python scripts\\archive\\run_batch.py
  python scripts\\archive\\run_batch.py --batch 20260616_203140
  python scripts\\archive\\run_batch.py --dry-run
  python scripts\\archive\\run_batch.py --pricing-source excel
"""

import argparse
import subprocess
import sys
from pathlib import Path

_ARCHIVE_DIR = Path(__file__).resolve().parent
_SCRIPTS_DIR = _ARCHIVE_DIR.parent
_REPORT_ROOT = _ARCHIVE_DIR.parents[1]
_DATA_IMPORT_DIR = _REPORT_ROOT / "scripts" / "dataImport"

for _p in (_REPORT_ROOT, _SCRIPTS_DIR, _DATA_IMPORT_DIR, _ARCHIVE_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from batch_lock import read_import_batch_from_lock  # noqa: E402
from console_style import (  # noqa: E402
    BG_BLUE,
    BG_CYAN,
    BG_MAGENTA,
    BG_YELLOW,
    FG_BLACK,
    FG_WHITE,
    STYLE_FAIL,
    STYLE_OK,
    init_console,
    log_level,
    paint_bg_fg,
)

_STEPS: tuple[tuple[str, str], ...] = (
    ("profit_001_order_sku.py", "利润表初始化"),
    ("profit_002_order_market.py", "市场信息回填"),
    ("profit_002_order_price.py", "Temu 价格覆盖"),
    ("profit_003_order_first.py", "头程/关税重算"),
    ("profit_004_order_delivery_amz.py", "Amazon FBA 派送费"),
)

_DRY_RUN_SCRIPTS = frozenset(
    {
        "profit_002_order_market.py",
        "profit_002_order_price.py",
        "profit_003_order_first.py",
        "profit_004_order_delivery_amz.py",
    }
)

_STEP_STYLES: dict[str, tuple[str, str]] = {
    "profit_001_order_sku.py": (BG_BLUE, FG_WHITE),
    "profit_002_order_market.py": (BG_YELLOW, FG_BLACK),
    "profit_002_order_price.py": (BG_MAGENTA, FG_WHITE),
    "profit_003_order_first.py": (BG_CYAN, FG_BLACK),
    "profit_004_order_delivery_amz.py": (BG_YELLOW, FG_BLACK),
}


def _log(level: str, msg: str, *, use_color: bool) -> None:
    log_level(level, msg, use_color=use_color)


def _log_step_banner(
    script: str,
    label: str,
    step_i: int,
    total: int,
    *,
    phase: str,
    detail: str = "",
    use_color: bool,
) -> None:
    if phase == "start":
        title = f" ▶ 步骤 {step_i}/{total}：{label}（{script}） "
        bg, fg = _STEP_STYLES.get(script, (BG_CYAN, FG_BLACK))
    elif phase == "ok":
        title = f" ✓ 完成 {step_i}/{total}：{label}（{script}） "
        bg, fg = STYLE_OK
    else:
        title = f" ✗ 失败 {step_i}/{total}：{label}（{script}） "
        bg, fg = STYLE_FAIL

    if detail:
        title = f"{title} {detail} "

    width = 72
    pad = max(0, width - len(title.encode("gbk", errors="ignore")))
    banner = paint_bg_fg(title + (" " * pad), bg, fg, use_color=use_color)
    print(banner, flush=True)


def resolve_import_batch(override: str | None, *, use_color: bool) -> str | None:
    """从命令行或 dataImport/run_batch.lock 解析 import_batch（不自动生成新批次）。"""
    if override and override.strip():
        return override.strip()

    batch = read_import_batch_from_lock()
    if batch:
        _log(
            "INFO",
            f"从 dataImport/run_batch.lock 读取批次：import_batch={batch}",
            use_color=use_color,
        )
        return batch.strip()

    return None


def build_child_argv(
    script: str,
    import_batch: str,
    *,
    dry_run: bool,
    pricing_source: str | None,
    excel_file: Path | None,
    first_leg_step: int,
) -> list[str]:
    argv = [
        sys.executable,
        str(_ARCHIVE_DIR / script),
        "--batch",
        import_batch,
    ]
    if dry_run and script in _DRY_RUN_SCRIPTS:
        argv.append("--dry-run")
    if script == "profit_003_order_first.py":
        if first_leg_step != 0:
            argv.extend(["--step", str(first_leg_step)])
        if pricing_source is not None:
            argv.extend(["--pricing-source", pricing_source])
        if excel_file is not None:
            argv.extend(["--excel-file", str(excel_file)])
    return argv


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="按批次顺序执行 archive 利润归档流水线")
    ap.add_argument(
        "--import-batch",
        "--batch",
        dest="import_batch",
        default=None,
        metavar="BATCH",
        help="导入批次号（默认从 dataImport/run_batch.lock 读取）",
    )
    ap.add_argument(
        "--dry-run",
        action="store_true",
        help="传给支持 dry-run 的子脚本（002_market、002_price、003、004_amz）；001 仍会写库",
    )
    ap.add_argument(
        "--step",
        type=int,
        choices=(0, 1, 2),
        default=0,
        help="传给 profit_003：0=头程两步都跑（默认），1=仅头程运费，2=仅头程关税",
    )
    ap.add_argument(
        "--pricing-source",
        choices=("db", "excel"),
        default=None,
        help="传给 profit_003：头程/关税主数据来源",
    )
    ap.add_argument(
        "--excel-file",
        type=Path,
        default=None,
        help="传给 profit_003：Excel 模式下的 BTH全部SKU明细 路径",
    )
    ap.add_argument(
        "--continue-on-error",
        action="store_true",
        help="某步失败后继续执行后续脚本（默认遇错即停）",
    )
    ap.add_argument(
        "--no-color",
        action="store_true",
        help="禁用终端彩色/背景色日志",
    )
    return ap.parse_args()


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    args = parse_args()
    use_color = init_console(no_color=args.no_color)

    import_batch = resolve_import_batch(args.import_batch, use_color=use_color)
    if not import_batch:
        _log(
            "ERROR",
            "无法获取批次号：请先执行 dataImport/run_batch.py，或使用 --batch 指定",
            use_color=use_color,
        )
        return 1

    _log("INFO", f"本批 import_batch={import_batch}", use_color=use_color)
    if args.dry_run:
        _log(
            "WARN",
            "dry-run 模式：001 仍会 UPSERT 利润表；002/003/004 仅统计不写库",
            use_color=use_color,
        )

    exit_code = 0
    failed = 0
    for step_i, (script, label) in enumerate(_STEPS, 1):
        argv = build_child_argv(
            script,
            import_batch,
            dry_run=args.dry_run,
            pricing_source=args.pricing_source,
            excel_file=args.excel_file,
            first_leg_step=args.step,
        )
        _log_step_banner(
            script, label, step_i, len(_STEPS), phase="start", use_color=use_color
        )
        _log("INFO", f"执行：{' '.join(argv)}", use_color=use_color)
        rc = subprocess.run(argv, cwd=_REPORT_ROOT).returncode
        if rc != 0:
            _log_step_banner(
                script,
                label,
                step_i,
                len(_STEPS),
                phase="fail",
                detail=f"退出码={rc}",
                use_color=use_color,
            )
            _log("ERROR", f"{script} 失败，退出码={rc}", use_color=use_color)
            failed += 1
            if not args.continue_on_error:
                exit_code = rc
                return exit_code
        else:
            _log_step_banner(
                script, label, step_i, len(_STEPS), phase="ok", use_color=use_color
            )

    if failed:
        _log(
            "ERROR",
            f"完成但有 {failed} 个步骤失败，import_batch={import_batch}",
            use_color=use_color,
        )
        exit_code = 1
    else:
        _log("INFO", f"全部完成，import_batch={import_batch}", use_color=use_color)
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
