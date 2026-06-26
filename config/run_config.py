"""
交互式配置向导
在报表生成前统一配置所有参数
"""
import json
import sys
from pathlib import Path
from datetime import datetime, timedelta
from typing import Dict, Any


def clear_screen():
    """清屏"""
    import os
    os.system('cls' if os.name == 'nt' else 'clear')


def print_header(title: str):
    """打印标题"""
    width = 60
    print("\n" + "=" * width)
    print(f"{title:^{width}}")
    print("=" * width + "\n")


def select_option(prompt: str, options: list, default: int = 0) -> int:
    """
    选择选项
    
    Args:
        prompt: 提示文字
        options: 选项列表
        default: 默认选项索引
    
    Returns:
        选中的选项索引
    """
    print(prompt)
    for i, option in enumerate(options, 1):
        marker = "✓" if i - 1 == default else " "
        print(f"  [{marker}] {i}. {option}")
    
    while True:
        choice = input(f"\n请输入选项编号 (1-{len(options)}) [默认: {default + 1}]: ").strip()
        if not choice:
            return default
        if choice.isdigit() and 1 <= int(choice) <= len(options):
            return int(choice) - 1
        print(f"❌ 无效输入，请输入 1-{len(options)} 之间的数字")


def input_text(prompt: str, default: str = "") -> str:
    """
    输入文本
    
    Args:
        prompt: 提示文字
        default: 默认值
    
    Returns:
        输入的文本
    """
    if default:
        value = input(f"{prompt} [默认: {default}]: ").strip()
        return value if value else default
    else:
        while True:
            value = input(f"{prompt}: ").strip()
            if value:
                return value
            print("❌ 此项不能为空，请重新输入")


def confirm(prompt: str, default: bool = True) -> bool:
    """
    确认选项
    
    Args:
        prompt: 提示文字
        default: 默认值
    
    Returns:
        True/False
    """
    default_text = "Y/n" if default else "y/N"
    while True:
        value = input(f"{prompt} [{default_text}]: ").strip().lower()
        if not value:
            return default
        if value in ('y', 'yes', '是'):
            return True
        if value in ('n', 'no', '否'):
            return False
        print("❌ 请输入 y(是) 或 n(否)")


def calculate_date_range(report_type: str) -> Dict[str, Any]:
    """
    根据报表类型计算日期范围
    
    Args:
        report_type: 报表类型（日报/月报）
    
    Returns:
        日期配置字典
    """
    today = datetime.now()
    
    if report_type == '日报':
        # 日报：统计 3 天前的数据，区间为当月 1 号至该日
        anchor = today - timedelta(days=3)
        start_date = anchor.replace(day=1)
        end_date = anchor
        shared_date = f"{anchor.month}.1-{anchor.month}.{anchor.day}"
        ku_cun_date = f"{anchor.year}.{anchor.month}.{anchor.day}"
        test_sheet_name = f"{anchor.year}.{anchor.month}"
        test_start = f"{anchor.year}-{anchor.month}-1"
        test_end = f"{anchor.year}-{anchor.month}-{anchor.day}"
        
        # transaction 文件日期：+3 天
        transaction_anchor = anchor + timedelta(days=3)
        transaction_date = f"{anchor.month}.1-{transaction_anchor.month}.{transaction_anchor.day}"
    else:
        # 月报：上一个自然月 1 号至月末
        first_this_month = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        last_day_prev = first_this_month - timedelta(days=1)
        start_date = last_day_prev.replace(day=1)
        end_date = last_day_prev
        shared_date = f"{last_day_prev.month}.1-{last_day_prev.month}.{last_day_prev.day}"
        ku_cun_date = f"{last_day_prev.year}.{last_day_prev.month}.{last_day_prev.day}"
        test_sheet_name = f"{last_day_prev.year}.{last_day_prev.month}"
        test_start = f"{last_day_prev.year}-{last_day_prev.month}-1"
        test_end = f"{last_day_prev.year}-{last_day_prev.month}-{last_day_prev.day}"
        
        # transaction 文件日期：每月 5 号
        transaction_date = f"{last_day_prev.month}.1-{today.month}.5"
    
    return {
        "report_type": report_type,
        "start_date": start_date.strftime("%Y-%m-%d"),
        "end_date": end_date.strftime("%Y-%m-%d"),
        "shared_date": shared_date,
        "ku_cun_date": ku_cun_date,
        "test_sheet_name": test_sheet_name,
        "test_start_date": test_start,
        "test_end_date": test_end,
        "transaction_date": transaction_date,
        "current_year_month": f"{today.year}.{today.month}"
    }


def run_config_wizard() -> Dict[str, Any]:
    """运行配置向导"""
    clear_screen()
    print_header("报表系统配置向导")
    
    print("欢迎使用报表系统！")
    print("在开始生成报表前，请先完成以下配置。\n")
    
    config = {}
    
    # 1. 选择报表类型
    print_header("步骤 1/5：选择报表类型")
    report_type_idx = select_option(
        "请选择要生成的报表类型：",
        ["日报（统计 3 天前的数据）", "月报（统计上个月的数据）"],
        default=0
    )
    report_type = "日报" if report_type_idx == 0 else "月报"
    
    # 计算日期范围
    date_config = calculate_date_range(report_type)
    config.update(date_config)
    
    print(f"\n✓ 已选择：{report_type}")
    print(f"  统计时间段：{date_config['shared_date']}")
    print(f"  测评时间：{date_config['test_sheet_name']}")
    
    input("\n按回车继续...")
    
    # 2. 目标拆解表
    print_header("步骤 2/5：目标拆解表配置")
    default_month_goal = f"26.{datetime.now().month}月目标拆解及跟进.xlsx"
    month_goal_excel = input_text(
        "请输入目标拆解表文件名（用于 Amazon 销售负责人映射）",
        default=default_month_goal
    )
    config['month_goal_excel'] = month_goal_excel
    
    print(f"\n✓ 目标拆解表：{month_goal_excel}")
    input("\n按回车继续...")
    
    # 3. 数据库配置
    print_header("步骤 3/5：数据库映射配置")
    print("数据库映射可以大幅提升查询速度（比 Excel 快 10-100 倍）。")
    print("如果尚未配置数据库，可以选择继续使用 Excel 映射。\n")
    
    enable_db_mapping = confirm("是否启用数据库映射？", default=False)
    config['enable_db_mapping'] = enable_db_mapping
    
    if enable_db_mapping:
        print("\n将使用数据库进行映射查询。")
        print("请确保已完成以下步骤：")
        print("  1. 已安装 MySQL 并创建数据库 'report_system'")
        print("  2. 已执行 schema.sql 创建表结构")
        print("  3. 已运行数据迁移脚本导入数据")
        print("  4. 已配置 report/config/db_config.json")
        
        if not confirm("\n是否已完成上述步骤？", default=False):
            print("\n⚠ 将回退到 Excel 映射模式")
            config['enable_db_mapping'] = False
    else:
        print("\n✓ 将使用 Excel 映射（兼容模式）")
    
    input("\n按回车继续...")
    
    # 4. 自动化配置
    print_header("步骤 4/5：自动化功能配置")
    print("以下功能需要额外的环境配置，如果尚未配置可以跳过。\n")
    
    auto_rpa = confirm("是否启用自动 RPA 查询（ERP/订单管理）？", default=False)
    config['auto_rpa'] = auto_rpa
    
    if auto_rpa:
        print("\n⚠ 自动 RPA 功能尚未实现，将暂时保持手动查询模式")
        config['auto_rpa'] = False
    
    auto_open_files = confirm("是否在检查点自动打开文件？", default=True)
    config['auto_open_files'] = auto_open_files
    
    input("\n按回车继续...")
    
    # 5. 报表模块选择
    print_header("步骤 5/5：选择要生成的报表模块")
    print("可以选择生成部分或全部报表模块。\n")
    
    all_modules = [
        ("B", "订单统计（sale_resend）"),
        ("C", "退款"),
        ("D", "广告"),
        ("E", "秒杀（仅 AMZ）"),
        ("F", "测评"),
        ("G", "二次上架"),
        ("H", "AMZ 利润报表 + OTTO 客户经理费"),
        ("K", "仓租映射产品信息"),
        ("M", "毛利 + 销售负责人")
    ]
    
    print("0. 全部模块（推荐）")
    for i, (code, name) in enumerate(all_modules, 1):
        print(f"{i}. {name}")
    
    choice = input("\n请输入要生成的模块编号（多个用逗号分隔，0=全部）[默认: 0]: ").strip()
    
    if not choice or choice == '0':
        selected_modules = [code for code, _ in all_modules]
    else:
        indices = [int(x.strip()) - 1 for x in choice.split(',') if x.strip().isdigit()]
        selected_modules = [all_modules[i][0] for i in indices if 0 <= i < len(all_modules)]
    
    config['selected_modules'] = selected_modules
    print(f"\n✓ 已选择 {len(selected_modules)} 个模块")
    
    input("\n按回车继续...")
    
    # 6. 确认配置
    clear_screen()
    print_header("配置确认")
    print("请确认以下配置信息：\n")
    print(f"报表类型：{config['report_type']}")
    print(f"统计时间段：{config['shared_date']}")
    print(f"目标拆解表：{config['month_goal_excel']}")
    print(f"数据库映射：{'是' if config['enable_db_mapping'] else '否（使用 Excel）'}")
    print(f"自动 RPA：{'是' if config['auto_rpa'] else '否'}")
    print(f"自动打开文件：{'是' if config['auto_open_files'] else '否'}")
    print(f"生成模块：{', '.join(config['selected_modules']) if config['selected_modules'] else '全部'}")
    
    if not confirm("\n确认配置并保存？", default=True):
        print("\n❌ 已取消配置")
        return None
    
    # 保存配置
    config_path = Path(__file__).parent / "config.json"
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    
    print(f"\n✅ 配置已保存到：{config_path}")
    print("\n下一步：运行 master_runner.py 开始生成报表")
    
    return config


def load_config() -> Dict[str, Any]:
    """加载配置文件"""
    config_path = Path(__file__).parent / "config.json"
    if not config_path.exists():
        print("❌ 配置文件不存在，请先运行配置向导")
        return None
    
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)


if __name__ == "__main__":
    config = run_config_wizard()
    if config:
        print("\n" + "=" * 60)
        print("配置完成！")
        print("=" * 60)
