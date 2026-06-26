"""
项目根目录引导：将项目根加入 sys.path，以便 import config / database。
各脚本在入口最前面调用：bootstrap(__file__)
"""
import sys
from pathlib import Path

_KNOWN_ROOT_NAMES = frozenset({"rpa-task", "rpaReport", "report"})


def find_project_root(start: Path | str | None = None) -> Path:
    """定位项目根目录（含 docker-compose.yml 与 database/）。"""
    start_path = Path(start).resolve() if start else Path.cwd()
    for candidate in [start_path, *start_path.parents]:
        if (
            candidate.name in _KNOWN_ROOT_NAMES
            and (candidate / "docker-compose.yml").is_file()
        ):
            return candidate
        if (candidate / "docker-compose.yml").is_file() and (candidate / "database").is_dir():
            return candidate
    raise RuntimeError(
        f"未找到项目根（需含 docker-compose.yml 与 database/），当前起点：{start_path}"
    )


def bootstrap(script_file: str | None = None) -> Path:
    """
    将项目根加入 sys.path（可 import config / database）。
    返回项目根路径。
    """
    root = find_project_root(Path(script_file) if script_file else None)
    p = str(root)
    if p not in sys.path:
        sys.path.insert(0, p)
    return root


def report_root() -> Path:
    """当前项目根目录（兼容旧名）。"""
    return find_project_root()


# 兼容旧调用
find_report_root = find_project_root
