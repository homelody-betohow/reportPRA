from __future__ import annotations

"""
终端 ANSI 彩色输出公用方法。

供 report/scripts 下各脚本复用：Windows 控制台启用 ANSI、检测是否着色、
横幅/日志/文本染色等。

用法：
  from console_style import init_console, paint, print_banner, log_level

  use_color = init_console()
  print_banner("校验通过", kind="ok", use_color=use_color)
  print_banner("数据异常", kind="alert", use_color=use_color, body_lines=["行1", "行2"])
"""

import os
import sys
from datetime import datetime
from typing import Literal, Sequence

RESET = "\033[0m"
BOLD = "\033[1m"

FG_BLACK = "\033[30m"
FG_WHITE = "\033[97m"
FG_RED = "\033[91m"
FG_GREEN = "\033[32m"
FG_YELLOW = "\033[33m"
FG_CYAN = "\033[36m"

BG_RED = "\033[41m"
BG_GREEN = "\033[42m"
BG_YELLOW = "\033[43m"
BG_BLUE = "\033[44m"
BG_MAGENTA = "\033[45m"
BG_CYAN = "\033[46m"

# 预设：背景 + 前景
STYLE_OK = (BG_GREEN, FG_BLACK)
STYLE_FAIL = (BG_RED, FG_WHITE)
STYLE_ALERT = (BG_RED, FG_WHITE)

BORDER_OK = f"{BOLD}{FG_GREEN}"
BORDER_ALERT = f"{BOLD}{FG_RED}"

LEVEL_FG = {
    "INFO": FG_CYAN,
    "WARN": FG_YELLOW,
    "ERROR": FG_RED,
}

BannerKind = Literal["ok", "alert", "fail", "info"]


def enable_windows_ansi() -> None:
    """在 Windows 控制台启用 ANSI 转义序列。"""
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


def supports_color(*, no_color: bool = False) -> bool:
    """当前 stdout 是否适合输出 ANSI 颜色。"""
    if no_color or os.getenv("NO_COLOR"):
        return False
    if not hasattr(sys.stdout, "isatty") or not sys.stdout.isatty():
        return False
    return True


def init_console(*, no_color: bool = False) -> bool:
    """启用 Windows ANSI 并返回是否使用颜色。"""
    enable_windows_ansi()
    return supports_color(no_color=no_color)


def paint(text: str, *styles: str, use_color: bool = True) -> str:
    """为文本套上 ANSI 样式；styles 可为常量字符串或命名样式元组。"""
    if not use_color or not styles:
        return text
    codes: list[str] = []
    for style in styles:
        if isinstance(style, tuple):
            codes.extend(style)
        else:
            codes.append(style)
    return f"{''.join(codes)}{text}{RESET}"


def paint_bg_fg(text: str, bg: str, fg: str, *, use_color: bool = True) -> str:
    """按背景色 + 前景色染色（兼容 run_batch 横幅写法）。"""
    return paint(text, bg, fg, use_color=use_color)


def print_line(text: str, *styles: str, use_color: bool = True) -> None:
    """打印一行染色文本。"""
    print(paint(text, *styles, use_color=use_color), flush=True)


def log_level(level: str, msg: str, *, use_color: bool = True) -> None:
    """带时间戳的彩色日志行。"""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] [{level}] {msg}"
    if use_color and level in LEVEL_FG:
        line = paint(line, BOLD, LEVEL_FG[level], use_color=True)
    print(line, flush=True)


def _banner_styles(kind: BannerKind) -> tuple[str, tuple[str, ...], tuple[str, ...]]:
    if kind == "ok":
        return BORDER_OK, STYLE_OK, (FG_BLACK,)
    if kind in ("alert", "fail"):
        return BORDER_ALERT, STYLE_FAIL, (FG_WHITE, BOLD)
    return f"{BOLD}{FG_CYAN}", (BG_CYAN, FG_BLACK), (FG_BLACK,)


def print_banner(
    headline: str,
    *,
    kind: BannerKind = "info",
    use_color: bool = True,
    width: int = 80,
    body_lines: Sequence[str] | None = None,
    footer: str | None = None,
) -> None:
    """
    打印带边框的醒目横幅。

    kind:
      ok    — 绿底标题 + 绿色边框（校验通过等）
      alert / fail — 红底标题 + 红色边框（异常提醒）
      info  — 青底标题
    """
    border_style, headline_style, body_style = _banner_styles(kind)
    border = "=" * width

    print_line(border, border_style, use_color=use_color)
    print_line(headline, *headline_style, BOLD, use_color=use_color)
    print_line(border, border_style, use_color=use_color)

    if body_lines:
        for line in body_lines:
            print_line(line, *body_style, use_color=use_color)

    if footer:
        print_line(footer, *headline_style, use_color=use_color)
        print_line(border, border_style, use_color=use_color)
