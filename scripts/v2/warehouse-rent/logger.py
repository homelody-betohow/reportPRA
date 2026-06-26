from __future__ import annotations

import os
import sys
from dataclasses import dataclass
from datetime import datetime
from typing import Literal

Level = Literal["DEBUG", "INFO", "WARN", "ERROR"]


@dataclass(frozen=True)
class Logger:
    prefix: str = ""
    use_color: bool = True

    def debug(self, msg: str) -> None:
        self._log("DEBUG", msg)

    def info(self, msg: str) -> None:
        self._log("INFO", msg)

    def warn(self, msg: str) -> None:
        self._log("WARN", msg)

    def error(self, msg: str) -> None:
        self._log("ERROR", msg)

    def _log(self, level: Level, msg: str) -> None:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        pfx = f" {self.prefix}" if self.prefix else ""
        line = f"[{ts}] [{level}]{pfx} {msg}"
        print(self._style(level, line))

    def _style(self, level: Level, text: str) -> str:
        if not self.use_color:
            return text
        # ANSI styles: bold + colors
        reset = "\033[0m"
        bold = "\033[1m"
        color = {
            "DEBUG": "\033[90m",  # gray
            "INFO": "\033[36m",  # cyan
            "WARN": "\033[33m",  # yellow
            "ERROR": "\033[31m",  # red
        }.get(level, "")
        return f"{bold}{color}{text}{reset}"


def setup_stdout_utf8() -> None:
    """Best-effort: avoid Windows console mojibake when printing Chinese."""
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass


def _supports_color() -> bool:
    if os.getenv("NO_COLOR"):
        return False
    if not hasattr(sys.stdout, "isatty") or not sys.stdout.isatty():
        return False
    # Many terminals support ANSI. Windows CMD/PowerShell on modern Win10 usually does.
    return True


def get_logger(prefix: str = "") -> Logger:
    return Logger(prefix=prefix, use_color=_supports_color())

