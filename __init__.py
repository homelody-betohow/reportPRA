"""
报表系统重构模块

提供数据库映射、配置管理、自动化运行等功能。

模块结构：
    - config/      配置管理
    - database/    数据库连接和映射服务
    - automation/  自动化工具（待实现）
    - runners/     自动化运行器（待实现）
    - utils/       工具类（待实现）

使用示例：
    from report.database.mapping_service import MappingService
    from report.config.run_config import run_config_wizard
    
    # 运行配置向导
    config = run_config_wizard()
    
    # 使用映射服务
    service = MappingService()
    mapping = service.map_sku_to_product_id(['ABC-001', 'DEF-002'])
"""

__version__ = '1.0.0'
__author__ = 'Report System Team'
__date__ = '2026-06-03'
