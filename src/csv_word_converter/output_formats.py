#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
输出格式扩展模块

提供多种输出格式的支持，包括PDF、HTML、Markdown、Excel、JSON等。
通过统一的接口和工厂模式实现格式转换的可扩展性。

支持的输出格式:
- Word (.docx) - 默认格式
- PDF (.pdf) - 通过python-docx2pdf或reportlab
- HTML (.html) - 结构化网页格式
- Markdown (.md) - 轻量级标记语言
- Excel (.xlsx) - 表格格式
- JSON (.json) - 结构化数据格式
"""

import json
import logging
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Dict, Any, List, Optional, Union
from dataclasses import dataclass
import pandas as pd
from datetime import datetime

try:
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

try:
    import docx2pdf
    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False


@dataclass
class FormatConfig:
    """输出格式配置类"""
    format_type: str                             # 格式类型
    output_path: Path                            # 输出路径
    template_data: Dict[str, Any]                # 模板数据
    options: Dict[str, Any] = None               # 格式特定选项
    
    def __post_init__(self):
        if self.options is None:
            self.options = {}


class OutputFormatter(ABC):
    """输出格式化器抽象基类
    
    定义所有输出格式化器的通用接口，子类需要实现具体的格式转换逻辑。
    """

    def __init__(self, config: FormatConfig):
        """
        初始化格式化器

        Args:
            config: 格式配置
        """
        self.config = config
        self.logger = logging.getLogger(self.__class__.__name__)

    @abstractmethod
    async def format_output(self) -> Path:
        """
        执行格式转换

        Returns:
            Path: 输出文件路径
        """
        pass

    @abstractmethod
    def validate_config(self) -> bool:
        """
        验证配置是否有效

        Returns:
            bool: 配置是否有效
        """
        pass

    def _ensure_output_dir(self) -> None:
        """确保输出目录存在"""
        self.config.output_path.parent.mkdir(parents=True, exist_ok=True)


class WordFormatter(OutputFormatter):
    """Word格式化器（默认格式）"""

    def validate_config(self) -> bool:
        """验证Word格式配置"""
        return self.config.output_path.suffix.lower() == '.docx'

    async def format_output(self) -> Path:
        """
        生成Word文档（通过现有转换器）

        Returns:
            Path: Word文档路径
        """
        from .core import csv_to_word_universal
        import tempfile
        import pandas as pd
        
        self._ensure_output_dir()
        
        try:
            # 从配置中获取数据
            data = self.config.template_data.get('data', [])
            if not data:
                raise ValueError("没有提供数据进行转换")
            
            # 将数据转换为DataFrame并保存为临时CSV文件
            df = pd.DataFrame(data)
            
            # 创建临时CSV文件
            with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as temp_csv:
                df.to_csv(temp_csv.name, index=False, encoding='utf-8')
                temp_csv_path = temp_csv.name
            
            # 使用现有的CSV到Word转换功能，使用第一个可用模板
            from .core import TemplateFactory
            factory = TemplateFactory()
            available_templates = factory.get_available_templates()
            
            if not available_templates:
                raise RuntimeError("没有可用的Word模板")
            
            # 使用第一个可用模板
            template_type = available_templates[0]
            self.logger.info(f"使用模板类型: {template_type}")
            
            result_path = csv_to_word_universal(
                csv_file=temp_csv_path,
                template_type=template_type
            )
            
            # 将生成的文档移动到目标位置
            if result_path and Path(result_path).exists():
                import shutil
                shutil.move(result_path, self.config.output_path)
                self.logger.info(f"Word文档已生成: {self.config.output_path}")
            else:
                raise RuntimeError("Word文档生成失败")
            
            # 清理临时文件
            Path(temp_csv_path).unlink(missing_ok=True)
            
            return self.config.output_path
            
        except Exception as e:
            self.logger.error(f"Word文档生成失败: {str(e)}")
            raise


class PDFFormatter(OutputFormatter):
    """PDF格式化器"""

    def validate_config(self) -> bool:
        """验证PDF格式配置"""
        if not self.config.output_path.suffix.lower() == '.pdf':
            return False
        
        # 检查依赖库
        if not REPORTLAB_AVAILABLE and not DOCX2PDF_AVAILABLE:
            self.logger.error("PDF格式需要安装 reportlab 或 docx2pdf 库")
            return False
        
        return True

    async def format_output(self) -> Path:
        """
        生成PDF文档

        Returns:
            Path: PDF文档路径
        """
        self._ensure_output_dir()
        
        # 优先使用reportlab直接生成
        if REPORTLAB_AVAILABLE:
            return await self._generate_pdf_with_reportlab()
        elif DOCX2PDF_AVAILABLE:
            return await self._convert_docx_to_pdf()
        else:
            raise RuntimeError("无可用的PDF生成库")

    async def _generate_pdf_with_reportlab(self) -> Path:
        """使用reportlab直接生成PDF"""
        doc = SimpleDocTemplate(
            str(self.config.output_path),
            pagesize=self.config.options.get('pagesize', A4),
            rightMargin=72, leftMargin=72,
            topMargin=72, bottomMargin=18
        )
        
        # 构建内容
        story = []
        styles = getSampleStyleSheet()
        
        # 标题
        title = self.config.template_data.get('title', '转换报告')
        story.append(Paragraph(title, styles['Title']))
        story.append(Spacer(1, 12))
        
        # 基本信息
        info_data = [
            ['生成时间', datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
            ['数据来源', str(self.config.template_data.get('source', 'CSV文件'))],
        ]
        
        info_table = Table(info_data, colWidths=[2*inch, 4*inch])
        info_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 14),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        story.append(info_table)
        story.append(Spacer(1, 12))
        
        # 数据表格
        if 'data' in self.config.template_data:
            data = self.config.template_data['data']
            if isinstance(data, list) and data:
                # 创建表格
                table_data = []
                if isinstance(data[0], dict):
                    # 字典列表格式
                    headers = list(data[0].keys())
                    table_data.append(headers)
                    for row in data:
                        table_data.append([str(row.get(h, '')) for h in headers])
                else:
                    # 直接使用数据
                    table_data = data
                
                data_table = Table(table_data)
                data_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 12),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
                
                story.append(data_table)
        
        # 构建PDF
        doc.build(story)
        self.logger.info(f"PDF文档已生成: {self.config.output_path}")
        return self.config.output_path

    async def _convert_docx_to_pdf(self) -> Path:
        """通过docx2pdf转换Word文档为PDF"""
        # 假设已有Word文档，进行转换
        docx_path = self.config.output_path.with_suffix('.docx')
        if not docx_path.exists():
            raise FileNotFoundError(f"Word文档不存在: {docx_path}")
        
        docx2pdf.convert(str(docx_path), str(self.config.output_path))
        self.logger.info(f"PDF文档已转换: {self.config.output_path}")
        return self.config.output_path


class HTMLFormatter(OutputFormatter):
    """HTML格式化器"""

    def validate_config(self) -> bool:
        """验证HTML格式配置"""
        return self.config.output_path.suffix.lower() == '.html'

    async def format_output(self) -> Path:
        """
        生成HTML文档

        Returns:
            Path: HTML文档路径
        """
        self._ensure_output_dir()
        
        # 构建HTML内容
        html_content = self._build_html_content()
        
        # 写入文件
        with open(self.config.output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        self.logger.info(f"HTML文档已生成: {self.config.output_path}")
        return self.config.output_path

    def _build_html_content(self) -> str:
        """构建HTML内容"""
        title = self.config.template_data.get('title', '转换报告')
        
        html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <style>
        body {{
            font-family: 'Microsoft YaHei', Arial, sans-serif;
            line-height: 1.6;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }}
        h1 {{
            color: #333;
            border-bottom: 3px solid #007acc;
            padding-bottom: 10px;
        }}
        .info-section {{
            background-color: #f8f9fa;
            padding: 15px;
            border-radius: 5px;
            margin: 20px 0;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }}
        th, td {{
            border: 1px solid #ddd;
            padding: 12px;
            text-align: left;
        }}
        th {{
            background-color: #007acc;
            color: white;
            font-weight: bold;
        }}
        tr:nth-child(even) {{
            background-color: #f2f2f2;
        }}
        tr:hover {{
            background-color: #e8f4f8;
        }}
        .footer {{
            margin-top: 30px;
            text-align: center;
            color: #666;
            font-size: 0.9em;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>{title}</h1>
        
        <div class="info-section">
            <h3>基本信息</h3>
            <p><strong>生成时间:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            <p><strong>数据来源:</strong> {self.config.template_data.get('source', 'CSV文件')}</p>
        </div>
"""
        
        # 添加数据表格
        if 'data' in self.config.template_data:
            data = self.config.template_data['data']
            if isinstance(data, list) and data:
                html += "\n        <h3>数据内容</h3>\n        <table>\n"
                
                if isinstance(data[0], dict):
                    # 字典列表格式
                    headers = list(data[0].keys())
                    html += "            <tr>\n"
                    for header in headers:
                        html += f"                <th>{header}</th>\n"
                    html += "            </tr>\n"
                    
                    for row in data:
                        html += "            <tr>\n"
                        for header in headers:
                            value = row.get(header, '')
                            html += f"                <td>{value}</td>\n"
                        html += "            </tr>\n"
                else:
                    # 假设第一行是标题
                    if data:
                        html += "            <tr>\n"
                        for cell in data[0]:
                            html += f"                <th>{cell}</th>\n"
                        html += "            </tr>\n"
                        
                        for row in data[1:]:
                            html += "            <tr>\n"
                            for cell in row:
                                html += f"                <td>{cell}</td>\n"
                            html += "            </tr>\n"
                
                html += "        </table>\n"
        
        html += """
        <div class="footer">
            <p>此报告由CSV转换工具自动生成</p>
        </div>
    </div>
</body>
</html>"""
        
        return html


class MarkdownFormatter(OutputFormatter):
    """Markdown格式化器"""

    def validate_config(self) -> bool:
        """验证Markdown格式配置"""
        return self.config.output_path.suffix.lower() == '.md'

    async def format_output(self) -> Path:
        """
        生成Markdown文档

        Returns:
            Path: Markdown文档路径
        """
        self._ensure_output_dir()
        
        # 构建Markdown内容
        md_content = self._build_markdown_content()
        
        # 写入文件
        with open(self.config.output_path, 'w', encoding='utf-8') as f:
            f.write(md_content)
        
        self.logger.info(f"Markdown文档已生成: {self.config.output_path}")
        return self.config.output_path

    def _build_markdown_content(self) -> str:
        """构建Markdown内容"""
        title = self.config.template_data.get('title', '转换报告')
        
        md = f"# {title}\n\n"
        
        # 基本信息
        md += "## 基本信息\n\n"
        md += f"- **生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        md += f"- **数据来源**: {self.config.template_data.get('source', 'CSV文件')}\n\n"
        
        # 数据表格
        if 'data' in self.config.template_data:
            data = self.config.template_data['data']
            if isinstance(data, list) and data:
                md += "## 数据内容\n\n"
                
                if isinstance(data[0], dict):
                    # 字典列表格式
                    headers = list(data[0].keys())
                    
                    # 表头
                    md += "| " + " | ".join(headers) + " |\n"
                    md += "| " + " | ".join(["---"] * len(headers)) + " |\n"
                    
                    # 数据行
                    for row in data:
                        values = [str(row.get(h, '')) for h in headers]
                        md += "| " + " | ".join(values) + " |\n"
                else:
                    # 假设第一行是标题
                    if data:
                        # 表头
                        md += "| " + " | ".join(str(cell) for cell in data[0]) + " |\n"
                        md += "| " + " | ".join(["---"] * len(data[0])) + " |\n"
                        
                        # 数据行
                        for row in data[1:]:
                            md += "| " + " | ".join(str(cell) for cell in row) + " |\n"
                
                md += "\n"
        
        # 页脚
        md += "---\n\n"
        md += "*此报告由CSV转换工具自动生成*\n"
        
        return md


class ExcelFormatter(OutputFormatter):
    """Excel格式化器"""

    def validate_config(self) -> bool:
        """验证Excel格式配置"""
        return self.config.output_path.suffix.lower() in ['.xlsx', '.xls']

    async def format_output(self) -> Path:
        """
        生成Excel文档

        Returns:
            Path: Excel文档路径
        """
        self._ensure_output_dir()
        
        # 创建Excel工作簿
        with pd.ExcelWriter(self.config.output_path, engine='openpyxl') as writer:
            # 基本信息工作表
            info_data = {
                '项目': ['生成时间', '数据来源', '格式版本'],
                '值': [
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    str(self.config.template_data.get('source', 'CSV文件')),
                    '1.0'
                ]
            }
            info_df = pd.DataFrame(info_data)
            info_df.to_excel(writer, sheet_name='基本信息', index=False)
            
            # 数据工作表
            if 'data' in self.config.template_data:
                data = self.config.template_data['data']
                if isinstance(data, list) and data:
                    if isinstance(data[0], dict):
                        # 字典列表格式
                        data_df = pd.DataFrame(data)
                    else:
                        # 二维列表格式
                        if data:
                            headers = data[0] if data else []
                            rows = data[1:] if len(data) > 1 else []
                            data_df = pd.DataFrame(rows, columns=headers)
                        else:
                            data_df = pd.DataFrame()
                    
                    data_df.to_excel(writer, sheet_name='数据内容', index=False)
        
        self.logger.info(f"Excel文档已生成: {self.config.output_path}")
        return self.config.output_path


class JSONFormatter(OutputFormatter):
    """JSON格式化器"""

    def validate_config(self) -> bool:
        """验证JSON格式配置"""
        return self.config.output_path.suffix.lower() == '.json'

    async def format_output(self) -> Path:
        """
        生成JSON文档

        Returns:
            Path: JSON文档路径
        """
        self._ensure_output_dir()
        
        # 构建JSON数据
        json_data = {
            'metadata': {
                'title': self.config.template_data.get('title', '转换报告'),
                'generated_at': datetime.now().isoformat(),
                'source': str(self.config.template_data.get('source', 'CSV文件')),
                'format_version': '1.0'
            },
            'data': self.config.template_data.get('data', [])
        }
        
        # 写入JSON文件
        with open(self.config.output_path, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, indent=2, ensure_ascii=False)
        
        self.logger.info(f"JSON文档已生成: {self.config.output_path}")
        return self.config.output_path


class OutputFormatFactory:
    """输出格式工厂类
    
    负责根据文件扩展名或格式类型创建相应的格式化器实例。
    """

    # 格式映射表
    FORMAT_MAPPING = {
        '.docx': WordFormatter,
        '.pdf': PDFFormatter,
        '.html': HTMLFormatter,
        '.md': MarkdownFormatter,
        '.xlsx': ExcelFormatter,
        '.xls': ExcelFormatter,
        '.json': JSONFormatter,
    }

    @classmethod
    def create_formatter(cls, config: FormatConfig) -> OutputFormatter:
        """
        创建格式化器实例

        Args:
            config: 格式配置

        Returns:
            OutputFormatter: 格式化器实例

        Raises:
            ValueError: 不支持的输出格式
        """
        # 根据文件扩展名确定格式
        file_extension = config.output_path.suffix.lower()
        
        if file_extension not in cls.FORMAT_MAPPING:
            raise ValueError(f"不支持的输出格式: {file_extension}")
        
        formatter_class = cls.FORMAT_MAPPING[file_extension]
        formatter = formatter_class(config)
        
        # 验证配置
        if not formatter.validate_config():
            raise ValueError(f"格式配置验证失败: {config.format_type}")
        
        return formatter

    @classmethod
    def get_supported_formats(cls) -> List[str]:
        """
        获取支持的输出格式列表

        Returns:
            List[str]: 支持的格式扩展名列表
        """
        return list(cls.FORMAT_MAPPING.keys())

    @classmethod
    def is_format_supported(cls, file_extension: str) -> bool:
        """
        检查格式是否支持

        Args:
            file_extension: 文件扩展名

        Returns:
            bool: 是否支持该格式
        """
        return file_extension.lower() in cls.FORMAT_MAPPING


async def convert_to_format(csv_data: List[Dict[str, Any]], 
                          output_path: Path,
                          title: str = "转换报告",
                          source: str = "CSV文件") -> Path:
    """
    便捷函数：将CSV数据转换为指定格式

    Args:
        csv_data: CSV数据（字典列表）
        output_path: 输出文件路径
        title: 报告标题
        source: 数据源描述

    Returns:
        Path: 输出文件路径
    """
    config = FormatConfig(
        format_type=output_path.suffix.lower(),
        output_path=output_path,
        template_data={
            'title': title,
            'source': source,
            'data': csv_data
        }
    )
    
    formatter = OutputFormatFactory.create_formatter(config)
    return await formatter.format_output()