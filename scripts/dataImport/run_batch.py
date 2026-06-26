from __future__ import annotations

"""
按时间生成 import_batch，顺序执行 dataImport 导入流水线。

当日重复执行时复用 `run_batch.lock` 中的 import_batch；跨日自动清除旧锁并生成新批次。
执行结束后写入/更新锁文件。

顺序：
  1. order_shipped.py   -> sales_order_shipped（写入 import_batch）
  2. order_refund.py    -> sales_order_refund（写入 report_hash）
  3. order_returned.py  -> sales_order_returned（写入 report_hash）
  4. order_temu.py      -> sales_order_shipped（更新Temu订单费用）

用法：
  cd d:\\py-project\\report
  python scripts\\dataImport\\run_batch.py
  python scripts\\dataImport\\run_batch.py --date 2026-06-09
  python scripts\\dataImport\\run_batch.py --import-batch 20260616_120000
  python scripts\\dataImport\\run_batch.py --mode 每天
"""

import argparse
import json
import os
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any

_DATA_IMPORT_DIR = Path(__file__).resolve().parent
_REPORT_ROOT = _DATA_IMPORT_DIR.parents[1]
_LOCK_FILE = _DATA_IMPORT_DIR / "run_batch.lock"

_STEPS: tuple[tuple[str, str], ...] = (
    ("order_shipped.py", "订单发货"),
    ("order_refund.py", "RMA 退款"),
    ("order_returned.py", "二次上架退件"),
    ("order_temu.py", "更新TEMU订单费用"),
    ("amz_transaction.py", "Amazon交易明细"),
)

# 各步骤横幅背景色（ANSI）：(背景, 前景)
_STEP_STYLES: dict[str, tuple[str, str]] = {
    "order_shipped.py": ("\033[44m", "\033[97m"),   # 蓝底白字
    "order_refund.py": ("\033[43m", "\033[30m"),    # 黄底黑字
    "order_returned.py": ("\033[45m", "\033[97m"),  # 洋红底白字
}
_STYLE_OK = ("\033[42m", "\033[30m")    # 绿底黑字
_STYLE_FAIL = ("\033[41m", "\033[97m")  # 红底白字
_RESET = "\033[0m"
_LEVEL_FG = {
    "INFO": "\033[36m",
    "WARN": "\033[33m",
    "ERROR": "\033[31m",
}


def _enable_windows_ansi() -> None:
    if os.name != "nt":
        return
    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32
        for handle_id in (-11, -12):  # stdout, stderr
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


def _paint(text: str, bg: str, fg: str, *, use_color: bool) -> str:
    if not use_color:
        return text
    return f"{bg}{fg}{text}{_RESET}"


def _log(level: str, msg: str, *, use_color: bool) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] [{level}] {msg}"
    if use_color and level in _LEVEL_FG:
        line = f"\033[1m{_LEVEL_FG[level]}{line}{_RESET}"
    print(line, flush=True)


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
        bg, fg = _STEP_STYLES.get(script, ("\033[46m", "\033[30m"))
    elif phase == "ok":
        title = f" ✓ 完成 {step_i}/{total}：{label}（{script}） "
        bg, fg = _STYLE_OK
    else:
        title = f" ✗ 失败 {step_i}/{total}：{label}（{script}） "
        bg, fg = _STYLE_FAIL

    if detail:
        title = f"{title} {detail} "

    width = 72
    pad = max(0, width - len(title.encode("gbk", errors="ignore")))
    banner = _paint(title + (" " * pad), bg, fg, use_color=use_color)
    print(banner, flush=True)


def make_import_batch(override: str | None) -> str:
    if override and override.strip():
        return override.strip()
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _read_lock() -> dict[str, Any] | None:
    if not _LOCK_FILE.is_file():
        return None
    try:
        data = json.loads(_LOCK_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    return data if isinstance(data, dict) else None


def _clear_lock(*, use_color: bool, reason: str) -> None:
    if not _LOCK_FILE.is_file():
        return
    try:
        _LOCK_FILE.unlink()
        _log("INFO", f"已删除锁文件（{reason}）：{_LOCK_FILE}", use_color=use_color)
    except OSError as e:
        _log("WARN", f"删除锁文件失败：{_LOCK_FILE} ({e})", use_color=use_color)


def resolve_import_batch(override: str | None, *, use_color: bool) -> str:
    """未显式指定批次时：当日锁有效则复用，否则生成新批次。"""
    if override and override.strip():
        return override.strip()

    today = date.today().isoformat()
    lock = _read_lock()
    if lock is not None:
        lock_date = str(lock.get("run_date", "")).strip()
        lock_batch = str(lock.get("import_batch", "")).strip()
        if lock_date == today and lock_batch:
            _log(
                "INFO",
                f"复用当日锁文件批次：import_batch={lock_batch}（{_LOCK_FILE.name}）",
                use_color=use_color,
            )
            return lock_batch
        if lock_date and lock_date != today:
            _clear_lock(use_color=use_color, reason=f"执行日期已变更 {lock_date} -> {today}")
        else:
            _clear_lock(use_color=use_color, reason="锁文件内容无效")

    batch = make_import_batch(None)
    _log("INFO", f"生成新批次：import_batch={batch}", use_color=use_color)
    return batch


def write_lock(import_batch: str, *, use_color: bool) -> None:
    payload = {
        "run_date": date.today().isoformat(),
        "import_batch": import_batch,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    try:
        _LOCK_FILE.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        _log("INFO", f"已写入锁文件：{_LOCK_FILE}", use_color=use_color)
    except OSError as e:
        _log("WARN", f"写入锁文件失败：{_LOCK_FILE} ({e})", use_color=use_color)


def build_child_argv(
    script: str,
    import_batch: str,
    *,
    date_arg: date | None,
    mode: str | None,
    no_shipped_enrich: bool,
) -> list[str]:
    argv = [
        sys.executable,
        str(_DATA_IMPORT_DIR / script),
        "--import-batch",
        import_batch,
    ]
    if date_arg is not None:
        argv.extend(["--date", date_arg.isoformat()])
    if mode is not None:
        argv.extend(["--mode", mode])
    if script == "order_returned.py" and no_shipped_enrich:
        argv.append("--no-shipped-enrich")
    return argv


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="按批次顺序执行 dataImport 导入流水线")
    ap.add_argument(
        "--import-batch",
        "--batch",
        dest="import_batch",
        default=None,
        metavar="BATCH",
        help="导入批次号（显式指定时忽略锁文件；默认当日复用 run_batch.lock）",
    )
    ap.add_argument(
        "--date",
        type=date.fromisoformat,
        default=None,
        metavar="YYYY-MM-DD",
        help="覆盖各脚本的日期子目录",
    )
    ap.add_argument(
        "--mode",
        choices=("每天", "每月"),
        default=None,
        help="路径模式（每天/每月）",
    )
    ap.add_argument(
        "--no-shipped-enrich",
        action="store_true",
        help="传给 order_returned.py：不从 sales_order_shipped 回填平台/店铺等",
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

    _enable_windows_ansi()
    args = parse_args()
    use_color = _supports_color(no_color=args.no_color)
    import_batch = resolve_import_batch(args.import_batch, use_color=use_color)
    _log("INFO", f"本批 import_batch={import_batch}", use_color=use_color)

    exit_code = 0
    failed = 0
    try:
        for step_i, (script, label) in enumerate(_STEPS, 1):
            argv = build_child_argv(
                script,
                import_batch,
                date_arg=args.date,
                mode=args.mode,
                no_shipped_enrich=args.no_shipped_enrich,
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
    finally:
        write_lock(import_batch, use_color=use_color)


if __name__ == "__main__":
    raise SystemExit(main())
