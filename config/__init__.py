"""
配置管理模块

提供交互式配置向导和配置文件管理。

使用示例：
    from report.config import run_config_wizard, load_config
    
    # 运行配置向导
    config = run_config_wizard()
    
    # 加载已保存的配置
    config = load_config()
"""

from .run_config import run_config_wizard, load_config

__all__ = [
    'run_config_wizard',
    'load_config'
]
