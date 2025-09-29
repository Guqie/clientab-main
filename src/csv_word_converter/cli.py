#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CSV-Word转换工具 - 命令行接口

提供命令行方式调用CSV到Word转换功能，支持单文件转换、批量处理和多种输出格式。

使用示例:
    # 单文件转换
    csv2word input.csv --template guoziwei --output ./reports/
    csv-word-convert data.csv -t default -o output.docx --verbose
    
    # 批量处理
    csv2word --batch-dir ./data/ --output-dir ./reports/ --format pdf
    
    # 异步处理
    csv2word input.csv --async --progress
"""

import argparse
import asyncio
import logging
import os
import sys
from pathlib import Path
from typing import List, Optional

from . import (
    __version__,
    configure_logging,
    convert_csv_to_word,
    get_available_templates,
    validate_csv_file,
)
from .async_converter import AsyncConverter
from .batch_processor import BatchProcessor, BatchConfig
from .output_formats import OutputFormatFactory


def setup_argument_parser() -> argparse.ArgumentParser:
    """
    设置命令行参数解析器

    返回:
        argparse.ArgumentParser: 配置好的参数解析器
    """
    parser = argparse.ArgumentParser(
        prog="csv2word",
        description="CSV到Word文档转换工具，支持单文件转换、批量处理和多种输出格式",
        epilog=f"版本: {__version__} | 更多信息请访问项目主页",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    # 位置参数
    parser.add_argument(
        "csv_file",
        nargs="?",  # 使CSV文件参数可选
        help="输入的CSV文件路径（批量处理时可省略）",
        type=str,
    )

    # 批量处理参数
    parser.add_argument(
        "--batch-dir",
        dest="batch_dir",
        help="批量处理：指定包含CSV文件的目录",
        type=str,
    )

    parser.add_argument(
        "--batch-pattern",
        dest="batch_pattern",
        default="*.csv",
        help="批量处理：CSV文件匹配模式 (默认: *.csv)",
        type=str,
    )

    parser.add_argument(
        "--max-workers",
        dest="max_workers",
        type=int,
        default=4,
        help="批量处理：最大并发工作线程数 (默认: 4)",
    )

    # 输出格式参数
    parser.add_argument(
        "--format",
        dest="output_format",
        choices=["docx", "pdf", "html", "markdown", "excel", "json"],
        default="docx",
        help="输出格式 (默认: docx)",
    )

    # 异步处理参数
    parser.add_argument(
        "--async",
        dest="use_async",
        action="store_true",
        help="启用异步处理模式",
    )

    parser.add_argument(
        "--progress",
        dest="show_progress",
        action="store_true",
        help="显示处理进度",
    )

    # 可选参数
    parser.add_argument(
        "-t",
        "--template",
        dest="template_type",
        default="default",
        choices=get_available_templates(),
        help="Word文档模板类型 (默认: default)",
    )

    parser.add_argument(
        "-o",
        "--output",
        dest="output_path",
        help="输出文件路径或目录 (默认: 自动生成)",
        type=str,
    )

    parser.add_argument(
        "--output-dir",
        dest="output_dir",
        help="输出目录 (默认: ./outputs/)",
        type=str,
        default="./outputs/",
    )

    # 处理选项
    parser.add_argument(
        "--no-images",
        dest="download_images",
        action="store_false",
        default=True,
        help="禁用图片下载",
    )

    parser.add_argument(
        "--image-timeout",
        dest="image_timeout",
        type=int,
        default=30,
        help="图片下载超时时间(秒) (默认: 30)",
    )

    parser.add_argument(
        "--max-retries",
        dest="max_retries",
        type=int,
        default=3,
        help="图片下载最大重试次数 (默认: 3)",
    )

    # 输出控制
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="详细输出模式",
    )

    parser.add_argument(
        "-q",
        "--quiet",
        action="store_true",
        help="静默模式，只输出错误信息",
    )

    parser.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        default="INFO",
        help="日志级别 (默认: INFO)",
    )

    # 验证选项
    parser.add_argument(
        "--validate-only",
        action="store_true",
        help="仅验证CSV文件，不进行转换",
    )

    parser.add_argument(
        "--list-templates",
        action="store_true",
        help="列出所有可用模板",
    )

    # 版本信息
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )

    # 服务器端口（用于Heroku等云平台部署）
    parser.add_argument(
        "--port",
        dest="port",
        type=int,
        help="服务器端口号（用于云平台部署，如Heroku）",
    )

    return parser


def validate_arguments(args: argparse.Namespace) -> bool:
    """
    验证命令行参数的有效性

    参数:
        args: 解析后的命令行参数

    返回:
        bool: 参数是否有效
    """
    # 如果是列出模板或版本信息，跳过CSV文件检查
    if args.list_templates:
        return True
    
    # 批量处理模式验证
    if args.batch_dir:
        if not os.path.exists(args.batch_dir):
            print(f"错误: 批量处理目录不存在: {args.batch_dir}", file=sys.stderr)
            return False
        if not os.path.isdir(args.batch_dir):
            print(f"错误: 批量处理路径不是目录: {args.batch_dir}", file=sys.stderr)
            return False
        # 批量处理模式下不需要单个CSV文件
        return True
    
    # 单文件处理模式验证
    if not args.csv_file:
        print("错误: 需要提供CSV文件路径或使用 --batch-dir 进行批量处理", file=sys.stderr)
        return False
    
    # 检查CSV文件是否存在
    if not os.path.exists(args.csv_file):
        print(f"错误: CSV文件不存在: {args.csv_file}", file=sys.stderr)
        return False

    # 检查输出目录
    if args.output_dir:
        output_dir = Path(args.output_dir)
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"错误: 无法创建输出目录 {args.output_dir}: {e}", file=sys.stderr)
            return False

    # 检查冲突参数
    if args.verbose and args.quiet:
        print("错误: --verbose 和 --quiet 参数不能同时使用", file=sys.stderr)
        return False

    # 检查批量处理参数
    if args.max_workers < 1:
        print("错误: --max-workers 必须大于0", file=sys.stderr)
        return False

    return True


def setup_logging_from_args(args: argparse.Namespace) -> None:
    """
    根据命令行参数设置日志配置

    参数:
        args: 解析后的命令行参数
    """
    if args.quiet:
        log_level = "ERROR"
    elif args.verbose:
        log_level = "DEBUG"
    else:
        log_level = args.log_level

    configure_logging(log_level)


def list_available_templates() -> None:
    """列出所有可用的模板"""
    templates = get_available_templates()
    print("可用的Word文档模板:")
    for i, template in enumerate(templates, 1):
        print(f"  {i}. {template}")


async def process_single_file_async(
    csv_file: str,
    template_type: str,
    output_format: str,
    output_dir: str,
    show_progress: bool = False
) -> Optional[str]:
    """
    异步处理单个CSV文件
    
    参数:
        csv_file: CSV文件路径
        template_type: 模板类型
        output_format: 输出格式
        output_dir: 输出目录
        show_progress: 是否显示进度
        
    返回:
        Optional[str]: 输出文件路径，失败时返回None
    """
    converter = AsyncConverter()
    
    def progress_callback(progress: float, message: str):
        if show_progress:
            print(f"进度: {progress:.1%} - {message}")
    
    try:
        result = await converter.convert_single_async(
            csv_file=csv_file,
            template_type=template_type,
            output_format=output_format,
            output_dir=output_dir,
            progress_callback=progress_callback
        )
        return result
    except Exception as e:
        print(f"异步转换失败: {e}")
        return None


def process_batch_files(
    batch_dir: str,
    pattern: str,
    template_type: str,
    output_format: str,
    output_dir: str,
    max_workers: int,
    show_progress: bool = False
) -> bool:
    """
    批量处理CSV文件
    
    参数:
        batch_dir: 批量处理目录
        pattern: 文件匹配模式
        template_type: 模板类型
        output_format: 输出格式
        output_dir: 输出目录
        max_workers: 最大工作线程数
        show_progress: 是否显示进度
        
    返回:
        bool: 处理是否成功
    """
    try:
        # 配置批量处理器
        config = BatchConfig(
            input_directory=batch_dir,
            output_directory=output_dir,
            file_pattern=pattern,
            template_type=template_type,
            output_format=output_format,
            max_workers=max_workers,
            enable_progress=show_progress
        )
        
        processor = BatchProcessor(config)
        
        def progress_callback(completed: int, total: int, current_file: str):
            if show_progress:
                progress = completed / total if total > 0 else 0
                print(f"批量处理进度: {progress:.1%} ({completed}/{total}) - 当前: {current_file}")
        
        # 执行批量处理
        result = processor.process_batch(progress_callback=progress_callback)
        
        # 显示结果统计
        print(f"\n批量处理完成:")
        print(f"  成功: {result.successful_count}")
        print(f"  失败: {result.failed_count}")
        print(f"  总计: {result.total_count}")
        print(f"  耗时: {result.total_time:.2f}秒")
        
        if result.failed_files:
            print(f"\n失败的文件:")
            for file_path, error in result.failed_files.items():
                print(f"  - {file_path}: {error}")
        
        return result.failed_count == 0
        
    except Exception as e:
        print(f"批量处理失败: {e}")
        return False


def validate_csv_and_report(csv_file: str) -> bool:
    """
    验证CSV文件并输出报告

    参数:
        csv_file: CSV文件路径

    返回:
        bool: 验证是否成功
    """
    try:
        result = validate_csv_file(csv_file)

        if result["is_valid"]:
            print(f"✓ CSV文件验证通过: {csv_file}")
            print(f"  - 文件大小: {result['file_size']} 字节")
            print(f"  - 行数: {result['row_count']}")
            print(f"  - 列数: {result['column_count']}")
            print(f"  - 列名: {', '.join(result['columns'])}")
            return True
        else:
            print(f"✗ CSV文件验证失败: {result.get('error', '未知错误')}")
            return False

    except Exception as e:
        print(f"✗ CSV文件验证出错: {e}")
        return False


def main() -> int:
    """
    主函数 - 命令行入口点

    返回:
        int: 退出代码 (0=成功, 1=失败)
    """
    parser = setup_argument_parser()
    args = parser.parse_args()

    # 设置日志
    setup_logging_from_args(args)
    logger = logging.getLogger(__name__)

    try:
        # 处理特殊命令
        if args.list_templates:
            list_available_templates()
            return 0

        # 如果指定了端口参数，启动Web服务器模式
        if hasattr(args, 'port') and args.port:
            logger.info(f"启动Web服务器模式，端口: {args.port}")
            from .web_server import start_web_server
            start_web_server(port=args.port)
            return 0

        # 验证参数
        if not validate_arguments(args):
            return 1

        # 仅验证模式
        if args.validate_only:
            if args.batch_dir:
                print("批量模式下不支持仅验证选项")
                return 1
            success = validate_csv_and_report(args.csv_file)
            return 0 if success else 1

        # 批量处理模式
        if args.batch_dir:
            logger.info(f"开始批量处理: {args.batch_dir}")
            success = process_batch_files(
                batch_dir=args.batch_dir,
                pattern=args.batch_pattern,
                template_type=args.template_type,
                output_format=args.output_format,
                output_dir=args.output_dir or "outputs",
                max_workers=args.max_workers,
                show_progress=args.progress
            )
            return 0 if success else 1

        # 单文件处理模式
        if args.use_async:
            # 异步处理单个文件
            logger.info(f"开始异步转换CSV文件: {args.csv_file}")
            result_path = asyncio.run(process_single_file_async(
                csv_file=args.csv_file,
                template_type=args.template_type,
                output_format=args.output_format,
                output_dir=args.output_dir or "outputs",
                show_progress=args.progress
            ))
        else:
            # 同步处理单个文件
            logger.info(f"开始转换CSV文件: {args.csv_file}")
            logger.info(f"使用模板: {args.template_type}")

            # 准备转换参数
            convert_kwargs = {
                "csv_file": args.csv_file,
                "template_type": args.template_type,
            }

            # 添加可选参数
            if args.output_path:
                # 如果指定了完整的输出文件路径，提取目录部分作为output_dir
                output_dir = os.path.dirname(args.output_path)
                if output_dir:
                    convert_kwargs["output_dir"] = output_dir

            # 执行转换
            result_path = convert_csv_to_word(**convert_kwargs)

            # 如果需要转换为其他格式
            if args.output_format != "docx" and result_path:
                from .output_formats import convert_to_format
                # 读取生成的docx文件数据
                import pandas as pd
                csv_data = pd.read_csv(args.csv_file).to_dict('records')
                
                # 构建输出路径
                output_dir = Path(args.output_dir or "outputs")
                output_dir.mkdir(exist_ok=True)
                output_filename = Path(result_path).stem + f".{args.output_format}"
                output_path = output_dir / output_filename
                
                # 使用asyncio运行异步函数
                result_path = asyncio.run(convert_to_format(
                    csv_data=csv_data,
                    output_path=output_path,
                    title=f"转换报告 - {args.template_type}",
                    source=args.csv_file
                ))

        # 检查结果
        if result_path and os.path.exists(result_path):
            print(f"✓ 转换成功! 输出文件: {result_path}")

            # 显示文件信息
            file_size = os.path.getsize(result_path)
            print(f"  - 文件大小: {file_size} 字节")

            return 0
        else:
            print("✗ 转换失败: 未生成输出文件")
            return 1

    except KeyboardInterrupt:
        print("\n用户中断操作")
        return 1

    except Exception as e:
        logger.error(f"转换过程中出现错误: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())