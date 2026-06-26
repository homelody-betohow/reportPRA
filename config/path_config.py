from datetime import datetime, timedelta
from pathlib import Path

# Excel 映射表路径
EXCEL_MAPPING_PATH = r"\\Betohow\数据报表\数据库\BTH全部SKU明细-v2026.06.02.xlsx"
BTH_ALL_SKU_DETAIL_PATH = EXCEL_MAPPING_PATH

MODE_RUN = "日报"


if MODE_RUN == "日报":
    # 3天前的日期
    DATE_PATH = datetime.now() - timedelta(days=3)
    DATE_PATH = DATE_PATH.strftime("%Y-%m-%d")
else:
    # 上个月1号到上个月的第一天
    DATE_PATH = datetime.now().replace(day=1)
    DATE_PATH = DATE_PATH.strftime("%Y-%m")


# 模式：每天 / 每月
MODE_PATTERN = "每天"
# ERP-订单统计表路径 = sales_order_shipped
ERP_ORDER_STA_PATH = r"\\Betohow\数据报表\报表自动化下载\其它报表\{MODE_PATTERN}\ERP订单、RMA下载"
# 二次上架表路径  = sales_order_returned
SECOND_RELISTING_PATH = r"\\Betohow\数据报表\报表自动化下载\其它报表\{MODE_PATTERN}\鸿羽仓二次上架明细"
# 交易明细表路径 = amz_transaction
TRANSACTION_PATH = r"\\Betohow\数据报表\报表自动化下载\其它报表\{MODE_PATTERN}\transaction交易明细"
# 亚马逊利润报表 = amz_seller_sku_profit_snapshot
AMAZON_PROFIT_PATH = r"\\Betohow\数据报表\报表自动化下载\其它报表\{MODE_PATTERN}\亚马逊利润报表"

# 桌面（换电脑时只改这里）
DESKTOP_ROOT = Path(r"C:\Users\BTH-windows\Desktop")

# RPA / ERP 查询结果目录
RPA_ORDER_QUERY_ROOT = Path(r"\\Betohow\数据报表\RPA\报表-无站点-订单查询")
RPA_RELISTING_QUERY_ROOT = Path(r"\\Betohow\数据报表\RPA\二次上架-数据查询")
