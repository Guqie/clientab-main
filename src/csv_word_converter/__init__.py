#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CSV-Word转换工具包

这是一个专业的CSV到Word文档转换工具，支持多种模板和自定义配置。

主要功能:
- CSV数据读取和处理
- 多种Word文档模板支持
- 图片下载和嵌入
- 自定义样式和格式
- 批量处理能力

使用示例:
    >>> from csv_word_converter import convert_csv_to_word
    >>> result = convert_csv_to_word(
    ...     csv_file="data.csv",
    ...     template_type="guoziwei"
    ... )
    >>> print(f"生成的文档: {result}")

作者: AI Development Team
版本: 1.0.0
许可证: MIT
"""

import logging
from typing import Any, Dict, List, Optional

# 版本信息
__version__ = "1.0.0"
__author__ = "AI Development Team"
__email__ = "dev@example.com"
__license__ = "MIT"

# 配置日志
logger = logging.getLogger(__name__)


# 公共API列表
__all__ = [
    # 版本信息
    "__version__",
    # 核心功能
    "convert_csv_to_word",
    "csv_to_word_universal",
    "validate_csv_file", 
    "get_available_templates",
    # 核心类（按需导入）
    "UniversalDocumentGenerator",
    # "DocumentTemplate",
    # "ConfigBasedTemplate",
    # "EnhancedImageDownloader",
]


# 便捷函数定义
def convert_csv_to_word(
    csv_file: str, template_type: str = "default", output_dir: Optional[str] = None, **kwargs
) -> str:
    """
    便捷的CSV到Word转换函数

    参数:
        csv_file: CSV文件路径
        template_type: 模板类型，默认为"default"
        output_dir: 输出目录，默认为None（自动生成）
        **kwargs: 其他参数传递给核心转换函数

    返回:
        str: 生成的Word文档路径

    异常:
        FileNotFoundError: CSV文件不存在
        ValueError: 模板类型不支持
        RuntimeError: 转换过程中出现错误
    """
    # 延迟导入核心函数以避免循环依赖
    from .core import csv_to_word_universal
    try:
        # csv_to_word_universal函数只接受csv_file, template_type, config_path参数
        # 忽略output_dir参数，因为该函数内部会自动生成输出路径
        return csv_to_word_universal(csv_file=csv_file, template_type=template_type, **kwargs)
    except Exception as e:
        logger.error(f"CSV转Word失败: {e}")
        raise


def get_available_templates() -> List[str]:
    """
    获取可用的模板列表

    返回:
        List[str]: 可用模板名称列表
    """
    try:
        # 延迟导入以避免循环依赖
        from .core import TemplateFactory
        factory = TemplateFactory()
        return factory.get_available_templates()
    except Exception:
        # 如果配置文件不存在或有问题，返回默认列表
        logger.warning("无法从TemplateFactory加载模板列表，返回默认值。")
        return ["guoziwei"]


def validate_csv_file(csv_file: str) -> Dict[str, Any]:
    """
    验证CSV文件的有效性

    参数:
        csv_file: CSV文件路径

    返回:
        Dict[str, Any]: 验证结果，包含is_valid、row_count、column_count等信息

    异常:
        FileNotFoundError: 文件不存在
    """
    import os
    import pandas as pd

    if not os.path.exists(csv_file):
        raise FileNotFoundError(f"CSV文件不存在: {csv_file}")

    try:
        # 读取CSV文件进行验证
        df = pd.read_csv(csv_file)

        return {
            "is_valid": True,
            "file_path": csv_file,
            "file_size": os.path.getsize(csv_file),
            "row_count": len(df),
            "column_count": len(df.columns),
            "columns": df.columns.tolist(),
            "has_header": True,  # 假设有表头
            "encoding": "utf-8",  # 默认编码
        }

    except Exception as e:
        logger.error(f"CSV文件验证失败: {e}")
        return {
            "is_valid": False,
            "error": str(e),
            "file_path": csv_file,
        }


# 包级别配置
def configure_logging(level: str = "INFO") -> None:
    """
    配置日志系统

    参数:
        level: 日志级别，默认为"INFO"
    """
    logging.basicConfig(
        level=getattr(logging, level.upper()),
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )


# 延迟导入核心函数和类，避免循环依赖
def __getattr__(name):
    """动态导入模块属性"""
    if name == "csv_to_word_universal":
        from .core import csv_to_word_universal
        return csv_to_word_universal
    elif name == "UniversalDocumentGenerator":
        from .core import UniversalDocumentGenerator
        return UniversalDocumentGenerator
    raise AttributeError(f"module '{__name__}' has no attribute '{name}'")

# 初始化日志
configure_logging("INFO")

logger.info(f"CSV-Word转换工具包 v{__version__} 已加载")