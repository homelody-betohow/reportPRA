"""
路径配置：report 为项目根，报表脚本与数据仍在 ../A_报表。
"""
from pathlib import Path

from bootstrap import find_report_root, find_py_project_root

_REPORT_ROOT = find_report_root(__file__)
_PY_PROJECT = find_py_project_root(_REPORT_ROOT)
A_BAO_ROOT = _PY_PROJECT / "A_报表"

# 桌面（换电脑时只改这里）
DESKTOP_ROOT = Path(r"C:\Users\BTH-windows\Desktop")

# 网络共享盘
BTH_ALL_SKU_DETAIL_PATH = Path(
    r"\\Betohow\数据报表\数据库\BTH全部SKU明细-v2026.06.02.xlsx"
)

# RPA / ERP 查询结果目录
RPA_ORDER_QUERY_ROOT = Path(r"\\Betohow\数据报表\RPA\报表-无站点-订单查询")
RPA_RELISTING_QUERY_ROOT = Path(r"\\Betohow\数据报表\RPA\二次上架-数据查询")
