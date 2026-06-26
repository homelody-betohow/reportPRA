"""
report 项目根目录引导：把上级 py-project 加入 sys.path，以便 import A_报表。
各脚本在入口最前面调用：bootstrap(__file__)
"""
import sys
from pathlib import Path

_REPORT_DIR_NAME = "report"
_A_BAO_DIR_NAME = "A_报表"


def find_report_root(start: Path | str | None = None) -> Path:
    """定位 report 目录（项目根）。"""
    start_path = Path(start).resolve() if start else Path.cwd()
    for candidate in [start_path, *start_path.parents]:
        if candidate.name == _REPORT_DIR_NAME and (candidate / "docker-compose.yml").is_file():
            return candidate
        if (candidate / "docker-compose.yml").is_file() and (candidate / "database").is_dir():
            return candidate
    raise RuntimeError(
        f"未找到 report 项目根（需含 docker-compose.yml），当前起点：{start_path}"
    )


def find_py_project_root(report_root: Path) -> Path:
    """report 的上级目录，内含 A_报表 与 ensure_project_root.py。"""
    parent = report_root.parent
    if (parent / _A_BAO_DIR_NAME).is_dir():
        return parent
    raise RuntimeError(
        f"未找到 {_A_BAO_DIR_NAME}，请确认目录结构：py-project/report 与 py-project/A_报表 并列"
    )


def bootstrap(script_file: str | None = None) -> Path:
    """
    1. 将 py-project 加入 sys.path（可 import A_报表）
    2. 将 report 加入 sys.path（可 import config / database）
    返回 report 根路径。
    """
    if script_file:
        report_root = find_report_root(Path(script_file))
    else:
        report_root = find_report_root()

    py_root = find_py_project_root(report_root)

    for p in (str(py_root), str(report_root)):
        if p not in sys.path:
            sys.path.insert(0, p)

    return report_root


def report_root() -> Path:
    """当前 report 根目录。"""
    return find_report_root()
