from __future__ import annotations

"""
检查 run_batch.py 各步骤所需的 Excel 源文件是否已就绪。

与 run_batch.py 使用相同的路径规则（config.path_config + --date / --mode）。
order_temu.py 不读取 Excel，仅提示为数据库步骤。

用法：
  cd <项目根目录>
  python scripts\\dataImport\\check_list.py
  python scripts\\dataImport\\check_list.py --date 2026-06-09
  python scripts\\dataImport\\check_list.py --mode 每月
"""

import argparse
import os
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

_REPORT_ROOT = Path(__file__).resolve().parents[2]
_DATA_IMPORT_DIR = Path(__file__).resolve().parent
_JOBS_DIR = Path(__file__).resolve().parents[1] / "jobs"
for _p in (_REPORT_ROOT, _DATA_IMPORT_DIR, _JOBS_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from config.path_config import DATE_PATH, MODE_PATTERN  # noqa: E402
from run_batch import _STEPS  # noqa: E402

from amz_transaction import (  # noqa: E402
    discover_transaction_files,
    resolve_work_dir as txn_resolve_work_dir,
    transaction_base_dir,
)
from order_refund import (  # noqa: E402
    discover_refund_files,
    erp_base_dir,
    resolve_work_dir as refund_resolve_work_dir,
)
from order_returned import (  # noqa: E402
    discover_returned_files,
    relisting_base_dir,
    resolve_work_dir as returned_resolve_work_dir,
)
from order_shipped import (  # noqa: E402
    discover_shipped_files,
    resolve_work_dir as shipped_resolve_work_dir,
)
from notify_email import send_notify  # noqa: E402

_RESET = "\033[0m"
_LEVEL_FG = {
    "INFO": "\033[36m",
    "WARN": "\033[33m",
    "OK": "\033[32m",
    "SKIP": "\033[90m",
}


def _enable_windows_ansi() -> None:
    if os.name != "nt":
        return
    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32
        for handle_id in (-11, -12):
            handle = kernel32.GetStdHandle(handle_id)
            mode = ctypes.c_uint32()
            if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
                kernel32.SetConsoleMode(handle, mode.value | 0x0004)
    except Exception:
        pass


def _supports_color(*, no_color: bool) -> bool:
    if no_color or os.getenv("NO_COLOR"):
        return False
    if not hasattr(sys.stdout, "isatty") or not sys.stdout.isatty():
        return False
    return True


def _log(level: str, msg: str, *, use_color: bool) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] [{level}] {msg}"
    if use_color and level in _LEVEL_FG:
        line = f"\033[1m{_LEVEL_FG[level]}{line}{_RESET}"
    print(line, flush=True)


@dataclass(frozen=True)
class StepCheckResult:
    script: str
    label: str
    needs_excel: bool
    work_dir: Path | None
    glob_hint: str
    files: tuple[Path, ...]
    dir_exists: bool
    warnings: tuple[str, ...]
    note: str = ""


def _check_excel_step(
    script: str,
    label: str,
    work_dir: Path,
    glob_hint: str,
    discover,
) -> StepCheckResult:
    warnings: list[str] = []
    dir_exists = work_dir.is_dir()
    if not dir_exists:
        warnings.append(f"目录不存在：{work_dir}")

    files: tuple[Path, ...] = ()
    if dir_exists:
        files = tuple(discover(work_dir))
        if not files:
            warnings.append(f"未找到匹配文件（{glob_hint}）：{work_dir}")

    return StepCheckResult(
        script=script,
        label=label,
        needs_excel=True,
        work_dir=work_dir,
        glob_hint=glob_hint,
        files=files,
        dir_exists=dir_exists,
        warnings=tuple(warnings),
    )


def build_checks(mode: str, on_date: date | None) -> list[StepCheckResult]:
    erp_base = erp_base_dir(mode)

    relisting_base = relisting_base_dir(mode)
    relisting_dir = returned_resolve_work_dir(relisting_base, on_date)

    txn_base = transaction_base_dir(mode)

    return [
        _check_excel_step(
            "order_shipped.py",
            "订单发货",
            shipped_resolve_work_dir(erp_base, mode, on_date),
            "订单统计*.xlsx / *订单统计*.xlsx",
            discover_shipped_files,
        ),
        _check_excel_step(
            "order_refund.py",
            "RMA 退款",
            refund_resolve_work_dir(erp_base, mode, on_date),
            "RMA*.xlsx",
            discover_refund_files,
        ),
        _check_excel_step(
            "order_returned.py",
            "二次上架退件",
            relisting_dir,
            "*二次上架明细-*.xls / *.xlsx",
            discover_returned_files,
        ),
        StepCheckResult(
            script="order_temu.py",
            label="更新TEMU订单费用",
            needs_excel=False,
            work_dir=None,
            glob_hint="",
            files=(),
            dir_exists=True,
            warnings=(),
            note="无需 Excel；依赖 order_shipped 写入的 temu_order_item 与数据库",
        ),
        _check_excel_step(
            "amz_transaction.py",
            "Amazon交易明细",
            txn_resolve_work_dir(txn_base, mode, on_date),
            "transaction交易明细*.xlsx",
            discover_transaction_files,
        ),
    ]


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="检查 run_batch.py 各步骤所需的 Excel 源文件是否存在"
    )
    ap.add_argument(
        "--date",
        type=date.fromisoformat,
        default=None,
        metavar="YYYY-MM-DD",
        help=f"覆盖日期子目录（发货/退款/退件默认 DATE_PATH={DATE_PATH}；交易明细默认当天）",
    )
    ap.add_argument(
        "--mode",
        choices=("每天", "每月"),
        default=None,
        help=f"路径模式（每天/每月），默认 path_config.MODE_PATTERN（{MODE_PATTERN}）",
    )
    ap.add_argument(
        "--no-color",
        action="store_true",
        help="禁用终端彩色日志",
    )
    ap.add_argument(
        "--no-email",
        action="store_true",
        help="检查失败时不发送 warning 邮件",
    )
    return ap.parse_args()


def _missing_steps(results: list[StepCheckResult]) -> list[StepCheckResult]:
    return [r for r in results if r.needs_excel and not r.files]


def _send_missing_files_warning(
    *,
    results: list[StepCheckResult],
    mode: str,
    on_date: date | None,
) -> None:
    missing = _missing_steps(results)
    if not missing:
        return

    date_label = on_date.isoformat() if on_date else f"发货/退款/退件={DATE_PATH}；交易明细=当天"
    lines = [
        f"有 {len(missing)} 个 Excel 步骤缺少源文件，run_batch 执行时可能失败。",
        "",
        "缺少文件的步骤：",
    ]
    for item in missing:
        lines.append(f"- {item.label}（{item.script}）")
        assert item.work_dir is not None
        lines.append(f"  目录：{item.work_dir}")
        lines.append(f"  匹配规则：{item.glob_hint}")
        for w in item.warnings:
            lines.append(f"  {w}")

    details: list[tuple[str, str]] = [
        ("模式", mode),
        ("日期", date_label),
        ("缺少步骤数", str(len(missing))),
    ]
    for i, item in enumerate(missing, 1):
        assert item.work_dir is not None
        details.append((f"步骤 {i}", f"{item.label} · {item.work_dir}"))

    send_notify(
        category="warning",
        subject="Excel 源文件检查失败",
        body="\n".join(lines),
        details=details,
        subtitle="rpa-task · check_list",
        log_prefix="[check_list]",
    )


def _print_summary(
    results: list[StepCheckResult],
    *,
    use_color: bool,
    mode: str,
    on_date: date | None,
    send_email: bool,
) -> int:
    missing = 0
    _log("INFO", f"共 {len(_STEPS)} 个 run_batch 步骤，其中 {sum(1 for r in results if r.needs_excel)} 个需要 Excel", use_color=use_color)

    for step_i, result in enumerate(results, 1):
        title = f"步骤 {step_i}/{len(results)}：{result.label}（{result.script}）"
        _log("INFO", title, use_color=use_color)

        if not result.needs_excel:
            _log("SKIP", result.note or "无需 Excel", use_color=use_color)
            continue

        assert result.work_dir is not None
        _log("INFO", f"目录：{result.work_dir}", use_color=use_color)
        _log("INFO", f"匹配规则：{result.glob_hint}", use_color=use_color)

        if result.files:
            for f in result.files:
                _log("OK", f"已找到：{f.name}", use_color=use_color)
        else:
            missing += 1
            for w in result.warnings:
                _log("WARN", w, use_color=use_color)

    print(flush=True)
    if missing:
        _log(
            "WARN",
            f"检查完成：有 {missing} 个 Excel 步骤缺少源文件，run_batch 执行时可能失败",
            use_color=use_color,
        )
        if send_email:
            _send_missing_files_warning(results=results, mode=mode, on_date=on_date)
        return 1

    _log("OK", "检查完成：所需 Excel 文件均已就绪", use_color=use_color)
    return 0


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    _enable_windows_ansi()
    args = parse_args()
    use_color = _supports_color(no_color=args.no_color)
    mode = args.mode or MODE_PATTERN

    if args.date is not None:
        _log("INFO", f"使用指定日期：{args.date.isoformat()}，模式：{mode}", use_color=use_color)
    else:
        _log(
            "INFO",
            f"默认日期：发货/退款/退件 DATE_PATH={DATE_PATH}；交易明细=当天；模式：{mode}",
            use_color=use_color,
        )

    results = build_checks(mode, args.date)
    return _print_summary(
        results,
        use_color=use_color,
        mode=mode,
        on_date=args.date,
        send_email=not args.no_email,
    )


if __name__ == "__main__":
    raise SystemExit(main())
