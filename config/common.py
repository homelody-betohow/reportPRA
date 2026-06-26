

# 汇率 别的币种 转 EUR
RMB_di_EUR = 7.3  # 人民币转欧元(除法)
USD_to_EUR = 0.860  # 美元
kc_to_EUR = 0.04133  # 捷克克朗
zl_to_EUR = 0.237  # 波兰兹罗提
Ft_to_EUR = 0.002611  # 匈牙利福林
CAD_to_EUR = 0.6179  # 加拿大元
kr_to_EUR = 0.0934  # 瑞典克朗
Lei_to_EUR = 0.196  # 罗马尼亚列伊

RATE_SHIP_FEE = 1.05  # 运费费率
SKU_NW_DISCOUNT = 0.8  # NW尾缀SKU 折扣

# =========================================================
# Excel 主数据路径
# - 主要用于「BTH全部SKU明细」：头程（RMB）/ 关税（含税）等字段
# - 优先：自动在共享目录中查找最新版本 BTH全部SKU明细-*.xlsx
# - 回退：config/path_config.py 中的 EXCEL_MAPPING_PATH
# =========================================================
from pathlib import Path


def _pick_latest_bth_excel() -> str | None:
    """
    在共享目录中自动选择最新的 BTH全部SKU明细-*.xlsx。
    选择策略：
    - 先按文件名排序（一般版本号/日期在文件名里，排序越靠后越新）
    - 若文件名无法保证，则改成按 mtime 选最新也很容易（目前先用文件名排序，稳定且快）
    """
    net_dir = Path(r"\\Betohow\数据报表\数据库")
    try:
        if not net_dir.is_dir():
            return None
        files = sorted(net_dir.glob("BTH全部SKU明细-*.xlsx"))
        files = [p for p in files if p.is_file() and not p.name.startswith("~$")]
        if not files:
            return None
        return str(files[-1])
    except Exception:
        return None


_auto = _pick_latest_bth_excel()
if _auto:
    BTH_ALL_SKU_DETAIL_PATH = _auto
else:
    try:
        from .path_config import EXCEL_MAPPING_PATH as BTH_ALL_SKU_DETAIL_PATH  # type: ignore
    except Exception:
        # 最后兜底：一个通用文件名（需要权限访问共享目录） 
        # BTH_ALL_SKU_DETAIL_PATH = r"\\Betohow\数据报表\数据库\BTH全部SKU明细.xlsx"
        BTH_ALL_SKU_DETAIL_PATH = r"\\Betohow\数据报表\数据库\BTH全部SKU明细-v2026.06.02.xlsx"