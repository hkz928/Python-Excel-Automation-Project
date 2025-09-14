# -*- coding: utf-8 -*-
"""
@Time    ：2025/9/14 上午11:33
@Author  ：庄洪奎（ARTHUR)
@FileName：config.py
@Software：PyCharm
@Subions ：zhk0459
"""
"""
项目配置文件
包含Excel操作的默认配置和常量定义
"""

import os
from pathlib import Path

# 项目根目录
PROJECT_ROOT = Path(__file__).parent

# 示例文件目录
EXAMPLES_DIR = PROJECT_ROOT / "examples"

# 输出文件目录
OUTPUT_DIR = PROJECT_ROOT / "output"

# 确保输出目录存在
OUTPUT_DIR.mkdir(exist_ok=True)

# Excel 应用程序默认配置
EXCEL_CONFIG = {
    "visible": False,           # 是否显示Excel界面
    "display_alerts": False,   # 是否显示警告对话框
    "screen_updating": False,  # 是否启用屏幕更新
    "calculation": "automatic", # 计算模式: automatic, manual, semiautomatic
    "enable_events": True,     # 是否启用事件
}

# 文件操作配置
FILE_CONFIG = {
    "backup_enabled": True,    # 是否创建备份
    "auto_save": True,        # 是否自动保存
    "save_format": "xlsx",    # 默认保存格式
    "encoding": "utf-8",      # 文件编码
}

# 日志配置
LOG_CONFIG = {
    "level": "INFO",
    "format": "{time:YYYY-MM-DD HH:mm:ss} | {level} | {name}:{function}:{line} | {message}",
    "rotation": "10 MB",
    "retention": "30 days",
    "log_file": PROJECT_ROOT / "logs" / "excel_automation.log"
}

# 确保日志目录存在
LOG_CONFIG["log_file"].parent.mkdir(exist_ok=True)

# 性能配置
PERFORMANCE_CONFIG = {
    "batch_size": 1000,        # 批量操作大小
    "timeout": 300,            # 操作超时时间（秒）
    "max_retries": 3,          # 最大重试次数
    "retry_delay": 1,          # 重试延迟（秒）
}

# 图表默认配置
CHART_CONFIG = {
    "default_type": "xlColumnClustered",  # 默认图表类型
    "width": 400,              # 默认宽度
    "height": 300,             # 默认高度
    "left": 100,               # 默认左边距
    "top": 50,                 # 默认上边距
}

# 数据透视表默认配置
PIVOT_CONFIG = {
    "version": 5,              # 透视表版本
    "source_type": 1,          # 数据源类型 (xlDatabase)
    "table_style": "TableStyleMedium2",  # 表格样式
}

# 打印配置
PRINT_CONFIG = {
    "orientation": 1,          # 页面方向 (1=纵向, 2=横向)
    "paper_size": 9,           # 纸张大小 (A4)
    "left_margin": 72,         # 左边距（磅）
    "right_margin": 72,        # 右边距（磅）
    "top_margin": 72,          # 上边距（磅）
    "bottom_margin": 72,       # 下边距（磅）
    "zoom": 100,               # 缩放比例
}

# 格式配置
FORMAT_CONFIG = {
    "default_font": "微软雅黑",
    "default_font_size": 11,
    "header_font_size": 14,
    "title_font_size": 16,
    "number_format": "#,##0.00",
    "date_format": "yyyy-mm-dd",
    "percentage_format": "0.00%",
}

# 错误处理配置
ERROR_CONFIG = {
    "continue_on_error": False,  # 遇到错误是否继续
    "log_errors": True,          # 是否记录错误日志
    "show_error_dialog": False,  # 是否显示错误对话框
}

# 安全配置
SECURITY_CONFIG = {
    "enable_macros": False,      # 是否启用宏
    "trust_vba_project": False,  # 是否信任VBA项目
    "disable_external_data": False,  # 是否禁用外部数据
}

# 获取配置函数
def get_config(config_name: str, key: str = None):
    """
    获取配置值
    
    Args:
        config_name: 配置名称
        key: 配置键名（可选）
    
    Returns:
        配置值
    """
    config_map = {
        "excel": EXCEL_CONFIG,
        "file": FILE_CONFIG,
        "log": LOG_CONFIG,
        "performance": PERFORMANCE_CONFIG,
        "chart": CHART_CONFIG,
        "pivot": PIVOT_CONFIG,
        "print": PRINT_CONFIG,
        "format": FORMAT_CONFIG,
        "error": ERROR_CONFIG,
        "security": SECURITY_CONFIG,
    }
    
    config = config_map.get(config_name)
    if config is None:
        raise ValueError(f"未知的配置名称: {config_name}")
    
    if key is None:
        return config
    
    return config.get(key)

# 更新配置函数
def update_config(config_name: str, key: str, value):
    """
    更新配置值
    
    Args:
        config_name: 配置名称
        key: 配置键名
        value: 新的配置值
    """
    config_map = {
        "excel": EXCEL_CONFIG,
        "file": FILE_CONFIG,
        "log": LOG_CONFIG,
        "performance": PERFORMANCE_CONFIG,
        "chart": CHART_CONFIG,
        "pivot": PIVOT_CONFIG,
        "print": PRINT_CONFIG,
        "format": FORMAT_CONFIG,
        "error": ERROR_CONFIG,
        "security": SECURITY_CONFIG,
    }
    
    config = config_map.get(config_name)
    if config is None:
        raise ValueError(f"未知的配置名称: {config_name}")
    
    config[key] = value