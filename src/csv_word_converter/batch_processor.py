#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量处理管理器模块

提供批量处理多个CSV文件的能力，支持目录扫描、文件过滤、进度监控等功能。
集成异步转换器实现高效的并发处理。

主要功能:
- 批量文件发现和过滤
- 智能输出路径生成
- 实时进度监控和报告
- 错误统计和重试机制
- 处理结果汇总和导出
"""

import asyncio
import logging
import time
from pathlib import Path
from typing import Dict, List, Optional, Callable, Pattern, Union, Tuple
from dataclasses import dataclass, field
from datetime import datetime
import re
import json

from .async_converter import AsyncConverter, ConversionTask, TaskStatus


@dataclass
class BatchConfig:
    """批量处理配置类"""
    input_dir: Path                              # 输入目录
    output_dir: Path                             # 输出目录
    template_type: str = "default"               # 默认模板类型
    file_pattern: str = "*.csv"                  # 文件匹配模式
    recursive: bool = True                       # 是否递归搜索子目录
    max_concurrent: int = 4                      # 最大并发数
    max_retries: int = 3                         # 最大重试次数
    timeout: float = 300.0                       # 单任务超时时间
    overwrite: bool = False                      # 是否覆盖已存在文件
    exclude_patterns: List[str] = field(default_factory=list)  # 排除模式
    include_patterns: List[str] = field(default_factory=list)  # 包含模式

    def __post_init__(self):
        """在初始化后执行，确保路径是Path对象"""
        if isinstance(self.input_dir, str):
            self.input_dir = Path(self.input_dir)
        if isinstance(self.output_dir, str):
            self.output_dir = Path(self.output_dir)


@dataclass
class BatchResult:
    """批量处理结果类"""
    batch_id: str                                # 批次ID
    start_time: datetime                         # 开始时间
    end_time: Optional[datetime] = None          # 结束时间
    total_files: int = 0                         # 总文件数
    processed_files: int = 0                     # 已处理文件数
    successful_files: int = 0                    # 成功文件数
    failed_files: int = 0                        # 失败文件数
    skipped_files: int = 0                       # 跳过文件数
    tasks: Dict[str, ConversionTask] = field(default_factory=dict)  # 任务详情
    errors: List[str] = field(default_factory=list)  # 错误列表
    
    @property
    def success_rate(self) -> float:
        """计算成功率"""
        if self.processed_files == 0:
            return 0.0
        return (self.successful_files / self.processed_files) * 100
    
    @property
    def duration(self) -> Optional[float]:
        """计算总耗时（秒）"""
        if self.start_time and self.end_time:
            return (self.end_time - self.start_time).total_seconds()
        return None


class BatchProcessor:
    """批量处理管理器
    
    负责管理多个CSV文件的批量转换处理，提供文件发现、任务调度、
    进度监控、结果统计等功能。
    """

    def __init__(self, config: BatchConfig):
        """
        初始化批量处理器

        Args:
            config: 批量处理配置
        """
        self.config = config
        self.logger = logging.getLogger(__name__)
        
        # 创建异步转换器
        self.converter = AsyncConverter(
            max_workers=config.max_concurrent,
            max_concurrent_tasks=config.max_concurrent,
            default_timeout=config.timeout
        )
        
        # 进度回调
        self.progress_callbacks: List[Callable[[BatchResult], None]] = []
        
        # 当前批次结果
        self.current_result: Optional[BatchResult] = None

    def add_progress_callback(self, callback: Callable[[BatchResult], None]) -> None:
        """
        添加进度回调函数

        Args:
            callback: 回调函数，接收BatchResult参数
        """
        self.progress_callbacks.append(callback)

    def _notify_progress(self) -> None:
        """通知所有进度回调函数"""
        if self.current_result:
            for callback in self.progress_callbacks:
                try:
                    callback(self.current_result)
                except Exception as e:
                    self.logger.warning(f"进度回调函数执行失败: {e}")

    def _compile_patterns(self, patterns: List[str]) -> List[Pattern]:
        """
        编译正则表达式模式

        Args:
            patterns: 模式字符串列表

        Returns:
            List[Pattern]: 编译后的正则表达式列表
        """
        compiled = []
        for pattern in patterns:
            try:
                # 将glob模式转换为正则表达式
                regex_pattern = pattern.replace('*', '.*').replace('?', '.')
                compiled.append(re.compile(regex_pattern, re.IGNORECASE))
            except re.error as e:
                self.logger.warning(f"无效的模式 '{pattern}': {e}")
        return compiled

    def _should_include_file(self, file_path: Path) -> bool:
        """
        判断文件是否应该被包含在处理中

        Args:
            file_path: 文件路径

        Returns:
            bool: 是否包含该文件
        """
        file_name = file_path.name
        
        # 检查排除模式
        if self.config.exclude_patterns:
            exclude_patterns = self._compile_patterns(self.config.exclude_patterns)
            for pattern in exclude_patterns:
                if pattern.match(file_name):
                    return False
        
        # 检查包含模式
        if self.config.include_patterns:
            include_patterns = self._compile_patterns(self.config.include_patterns)
            for pattern in include_patterns:
                if pattern.match(file_name):
                    return True
            return False  # 如果有包含模式但都不匹配，则排除
        
        return True

    def discover_files(self) -> List[Path]:
        """
        发现需要处理的CSV文件

        Returns:
            List[Path]: CSV文件路径列表
        """
        self.logger.info(f"开始扫描目录: {self.config.input_dir}")
        
        if not self.config.input_dir.exists():
            raise FileNotFoundError(f"输入目录不存在: {self.config.input_dir}")
        
        if not self.config.input_dir.is_dir():
            raise NotADirectoryError(f"输入路径不是目录: {self.config.input_dir}")

        files = []
        
        # 根据配置选择搜索方式
        if self.config.recursive:
            pattern_path = self.config.input_dir / "**" / self.config.file_pattern
            found_files = list(self.config.input_dir.glob("**/" + self.config.file_pattern))
        else:
            found_files = list(self.config.input_dir.glob(self.config.file_pattern))

        # 过滤文件
        for file_path in found_files:
            if file_path.is_file() and self._should_include_file(file_path):
                files.append(file_path)

        self.logger.info(f"发现 {len(files)} 个CSV文件")
        return files

    def _generate_output_path(self, input_file: Path) -> Path:
        """
        生成输出文件路径

        Args:
            input_file: 输入文件路径

        Returns:
            Path: 输出文件路径
        """
        # 保持相对目录结构
        relative_path = input_file.relative_to(self.config.input_dir)
        
        # 更改文件扩展名为.docx
        output_name = relative_path.stem + ".docx"
        output_path = self.config.output_dir / relative_path.parent / output_name
        
        return output_path

    def _should_skip_file(self, input_file: Path, output_file: Path) -> bool:
        """
        判断是否应该跳过文件处理

        Args:
            input_file: 输入文件路径
            output_file: 输出文件路径

        Returns:
            bool: 是否跳过
        """
        # 如果输出文件已存在且不允许覆盖
        if output_file.exists() and not self.config.overwrite:
            # 检查文件修改时间
            if output_file.stat().st_mtime > input_file.stat().st_mtime:
                return True
        
        return False

    async def process_batch(self, 
                          files: Optional[List[Path]] = None,
                          batch_id: Optional[str] = None) -> BatchResult:
        """
        执行批量处理

        Args:
            files: 要处理的文件列表（可选，默认自动发现）
            batch_id: 批次ID（可选，自动生成）

        Returns:
            BatchResult: 批量处理结果
        """
        # 生成批次ID
        if batch_id is None:
            batch_id = f"batch_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

        # 初始化结果对象
        self.current_result = BatchResult(
            batch_id=batch_id,
            start_time=datetime.now()
        )

        try:
            # 发现文件
            if files is None:
                files = self.discover_files()
            
            self.current_result.total_files = len(files)
            self.logger.info(f"开始批量处理 {len(files)} 个文件")

            # 添加任务到转换器
            task_mapping = {}  # 映射任务ID到文件路径
            
            for input_file in files:
                output_file = self._generate_output_path(input_file)
                
                # 检查是否跳过
                if self._should_skip_file(input_file, output_file):
                    self.logger.info(f"跳过文件（已存在且较新）: {input_file}")
                    self.current_result.skipped_files += 1
                    continue

                # 添加转换任务
                try:
                    task_id = await self.converter.add_task(
                        csv_file=input_file,
                        output_path=output_file,
                        template_type=self.config.template_type,
                        max_retries=self.config.max_retries
                    )
                    task_mapping[task_id] = input_file
                    
                except Exception as e:
                    error_msg = f"添加任务失败 {input_file}: {e}"
                    self.logger.error(error_msg)
                    self.current_result.errors.append(error_msg)
                    self.current_result.failed_files += 1

            # 设置进度回调
            def task_progress_callback(task: ConversionTask):
                # 更新批次结果
                if task.task_id in self.current_result.tasks:
                    old_task = self.current_result.tasks[task.task_id]
                    if old_task.status != task.status:
                        # 状态发生变化，更新统计
                        if task.status == TaskStatus.COMPLETED:
                            self.current_result.successful_files += 1
                        elif task.status == TaskStatus.FAILED:
                            self.current_result.failed_files += 1
                
                self.current_result.tasks[task.task_id] = task
                self.current_result.processed_files = len([
                    t for t in self.current_result.tasks.values() 
                    if t.is_finished
                ])
                
                # 通知外部回调
                self._notify_progress()

            self.converter.add_progress_callback(task_progress_callback)

            # 执行所有任务
            if task_mapping:
                self.logger.info(f"开始并发处理 {len(task_mapping)} 个任务")
                completed_tasks = await self.converter.process_all_tasks()
                
                # 更新最终结果
                for task_id, task in completed_tasks.items():
                    self.current_result.tasks[task_id] = task
                    
                    if task.status == TaskStatus.COMPLETED:
                        self.current_result.successful_files += 1
                    elif task.status == TaskStatus.FAILED:
                        self.current_result.failed_files += 1
                        if task.error_message:
                            self.current_result.errors.append(
                                f"{task_mapping.get(task_id, 'Unknown')}: {task.error_message}"
                            )

            # 完成处理
            self.current_result.end_time = datetime.now()
            self.current_result.processed_files = len([
                t for t in self.current_result.tasks.values() 
                if t.is_finished
            ])

            self.logger.info(f"批量处理完成: {self.current_result.successful_files}/{self.current_result.total_files} 成功")
            self._notify_progress()

            return self.current_result

        except Exception as e:
            self.logger.error(f"批量处理失败: {e}")
            if self.current_result:
                self.current_result.end_time = datetime.now()
                self.current_result.errors.append(f"批量处理异常: {e}")
            raise

    def export_report(self, 
                     result: BatchResult, 
                     output_path: Path,
                     format: str = "json") -> None:
        """
        导出处理报告

        Args:
            result: 批量处理结果
            output_path: 输出文件路径
            format: 报告格式 ("json", "csv", "txt")
        """
        try:
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            if format.lower() == "json":
                self._export_json_report(result, output_path)
            elif format.lower() == "csv":
                self._export_csv_report(result, output_path)
            elif format.lower() == "txt":
                self._export_txt_report(result, output_path)
            else:
                raise ValueError(f"不支持的报告格式: {format}")
                
            self.logger.info(f"报告已导出: {output_path}")
            
        except Exception as e:
            self.logger.error(f"导出报告失败: {e}")
            raise

    def _export_json_report(self, result: BatchResult, output_path: Path) -> None:
        """导出JSON格式报告"""
        report_data = {
            "batch_id": result.batch_id,
            "start_time": result.start_time.isoformat(),
            "end_time": result.end_time.isoformat() if result.end_time else None,
            "duration": result.duration,
            "summary": {
                "total_files": result.total_files,
                "processed_files": result.processed_files,
                "successful_files": result.successful_files,
                "failed_files": result.failed_files,
                "skipped_files": result.skipped_files,
                "success_rate": result.success_rate
            },
            "tasks": {
                task_id: {
                    "csv_file": str(task.csv_file),
                    "output_path": str(task.output_path),
                    "template_type": task.template_type,
                    "status": task.status.value,
                    "progress": task.progress,
                    "duration": task.duration,
                    "error_message": task.error_message,
                    "retry_count": task.retry_count
                }
                for task_id, task in result.tasks.items()
            },
            "errors": result.errors
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(report_data, f, indent=2, ensure_ascii=False)

    def _export_csv_report(self, result: BatchResult, output_path: Path) -> None:
        """导出CSV格式报告"""
        import csv
        
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            
            # 写入标题行
            writer.writerow([
                "Task ID", "CSV File", "Output Path", "Template", 
                "Status", "Progress", "Duration", "Error", "Retries"
            ])
            
            # 写入任务数据
            for task_id, task in result.tasks.items():
                writer.writerow([
                    task_id,
                    str(task.csv_file),
                    str(task.output_path),
                    task.template_type,
                    task.status.value,
                    f"{task.progress:.1f}%",
                    f"{task.duration:.2f}s" if task.duration else "N/A",
                    task.error_message or "",
                    task.retry_count
                ])

    def _export_txt_report(self, result: BatchResult, output_path: Path) -> None:
        """导出文本格式报告"""
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(f"批量处理报告\n")
            f.write(f"=" * 50 + "\n\n")
            
            f.write(f"批次ID: {result.batch_id}\n")
            f.write(f"开始时间: {result.start_time}\n")
            f.write(f"结束时间: {result.end_time}\n")
            f.write(f"总耗时: {result.duration:.2f}秒\n\n" if result.duration else "总耗时: N/A\n\n")
            
            f.write(f"处理统计:\n")
            f.write(f"  总文件数: {result.total_files}\n")
            f.write(f"  已处理: {result.processed_files}\n")
            f.write(f"  成功: {result.successful_files}\n")
            f.write(f"  失败: {result.failed_files}\n")
            f.write(f"  跳过: {result.skipped_files}\n")
            f.write(f"  成功率: {result.success_rate:.1f}%\n\n")
            
            if result.errors:
                f.write(f"错误列表:\n")
                for i, error in enumerate(result.errors, 1):
                    f.write(f"  {i}. {error}\n")

    async def shutdown(self) -> None:
        """关闭批量处理器"""
        self.logger.info("正在关闭批量处理器...")
        await self.converter.shutdown()
        self.logger.info("批量处理器已关闭")