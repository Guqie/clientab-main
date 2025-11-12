#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CSV-Word转换工具 - MCP服务器

这是一个基于Model Context Protocol (MCP)的服务器实现，
为CSV到Word转换工具提供标准化的工具接口，供其他项目集成使用。

支持的工具:
- convert_csv_to_word: 单文件转换
- batch_convert_csv: 批量转换
- get_available_templates: 获取可用模板
- validate_csv_file: 验证CSV文件
- get_conversion_status: 获取转换状态

作者: AI Development Team
版本: 1.0.0
"""

import asyncio
import json
import logging
import os
import sys
import tempfile
import traceback
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence
import uuid
from datetime import datetime

# MCP相关导入
from mcp.server import Server
from mcp.server.models import InitializationOptions
from mcp.server.stdio import stdio_server
from mcp.types import (
    Resource,
    Tool,
    TextContent,
    ImageContent,
    EmbeddedResource,
    LoggingLevel
)
import mcp.server.stdio
import mcp.types as types

# 添加项目根目录到Python路径
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root / "src"))

try:
    from csv_word_converter import (
        convert_csv_to_word,
        get_available_templates,
        validate_csv_file,
        __version__
    )
    from csv_word_converter.core import TemplateFactory
    from csv_word_converter.async_converter import AsyncConverter
    from csv_word_converter.batch_processor import BatchProcessor, BatchConfig
except ImportError as e:
    logging.error(f"导入核心模块失败: {e}")
    sys.exit(1)

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# 全局状态管理
conversion_jobs = {}  # 存储转换任务状态
temp_directories = {}  # 存储临时目录映射

class CSVWordMCPServer:
    """
    CSV-Word转换工具的MCP服务器实现
    """
    
    def __init__(self):
        """
        初始化MCP服务器
        """
        self.server = Server("csv-word-converter")
        self.setup_tools()
        self.setup_resources()
        
    def setup_tools(self):
        """
        设置MCP工具定义
        """
        
        @self.server.list_tools()
        async def handle_list_tools() -> List[Tool]:
            """
            列出所有可用的工具
            
            Returns:
                List[Tool]: 工具列表
            """
            return [
                Tool(
                    name="convert_csv_to_word",
                    description="将CSV文件转换为Word文档",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "csv_file_path": {
                                "type": "string",
                                "description": "CSV文件的绝对路径"
                            },
                            "template_type": {
                                "type": "string",
                                "description": "模板类型（如：guoziwei, new_energy）",
                                "default": "guoziwei"
                            },
                            "output_dir": {
                                "type": "string",
                                "description": "输出目录路径（可选）"
                            },
                            "output_format": {
                                "type": "string",
                                "description": "输出格式",
                                "enum": ["docx", "pdf"],
                                "default": "docx"
                            }
                        },
                        "required": ["csv_file_path"]
                    }
                ),
                Tool(
                    name="batch_convert_csv",
                    description="批量转换多个CSV文件",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "csv_files": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "CSV文件路径列表"
                            },
                            "template_type": {
                                "type": "string",
                                "description": "模板类型",
                                "default": "guoziwei"
                            },
                            "output_dir": {
                                "type": "string",
                                "description": "输出目录路径"
                            },
                            "max_workers": {
                                "type": "integer",
                                "description": "最大并发工作线程数",
                                "default": 4
                            }
                        },
                        "required": ["csv_files"]
                    }
                ),
                Tool(
                    name="get_available_templates",
                    description="获取所有可用的文档模板",
                    inputSchema={
                        "type": "object",
                        "properties": {},
                        "additionalProperties": False
                    }
                ),
                Tool(
                    name="validate_csv_file",
                    description="验证CSV文件的格式和内容",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "csv_file_path": {
                                "type": "string",
                                "description": "CSV文件的绝对路径"
                            }
                        },
                        "required": ["csv_file_path"]
                    }
                ),
                Tool(
                    name="get_conversion_status",
                    description="获取转换任务的状态",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "job_id": {
                                "type": "string",
                                "description": "转换任务ID"
                            }
                        },
                        "required": ["job_id"]
                    }
                ),
                Tool(
                    name="create_temp_workspace",
                    description="创建临时工作空间用于文件处理",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "workspace_name": {
                                "type": "string",
                                "description": "工作空间名称（可选）"
                            }
                        }
                    }
                ),
                Tool(
                    name="cleanup_temp_workspace",
                    description="清理临时工作空间",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "workspace_id": {
                                "type": "string",
                                "description": "工作空间ID"
                            }
                        },
                        "required": ["workspace_id"]
                    }
                )
            ]
        
        @self.server.call_tool()
        async def handle_call_tool(name: str, arguments: Dict[str, Any]) -> List[types.TextContent]:
            """
            处理工具调用
            
            Args:
                name (str): 工具名称
                arguments (Dict[str, Any]): 工具参数
                
            Returns:
                List[types.TextContent]: 工具执行结果
            """
            try:
                if name == "convert_csv_to_word":
                    return await self._convert_csv_to_word(arguments)
                elif name == "batch_convert_csv":
                    return await self._batch_convert_csv(arguments)
                elif name == "get_available_templates":
                    return await self._get_available_templates(arguments)
                elif name == "validate_csv_file":
                    return await self._validate_csv_file(arguments)
                elif name == "get_conversion_status":
                    return await self._get_conversion_status(arguments)
                elif name == "create_temp_workspace":
                    return await self._create_temp_workspace(arguments)
                elif name == "cleanup_temp_workspace":
                    return await self._cleanup_temp_workspace(arguments)
                else:
                    return [types.TextContent(
                        type="text",
                        text=f"未知工具: {name}"
                    )]
            except Exception as e:
                logger.error(f"工具调用失败 {name}: {e}")
                logger.error(traceback.format_exc())
                return [types.TextContent(
                    type="text",
                    text=f"工具执行失败: {str(e)}"
                )]
    
    def setup_resources(self):
        """
        设置MCP资源定义
        """
        
        @self.server.list_resources()
        async def handle_list_resources() -> List[Resource]:
            """
            列出所有可用资源
            
            Returns:
                List[Resource]: 资源列表
            """
            return [
                Resource(
                    uri="csv-word://templates",
                    name="可用模板列表",
                    description="获取所有可用的Word文档模板",
                    mimeType="application/json"
                ),
                Resource(
                    uri="csv-word://config",
                    name="配置信息",
                    description="获取转换工具的配置信息",
                    mimeType="application/json"
                ),
                Resource(
                    uri="csv-word://status",
                    name="服务状态",
                    description="获取MCP服务器的运行状态",
                    mimeType="application/json"
                )
            ]
        
        @self.server.read_resource()
        async def handle_read_resource(uri: str) -> str:
            """
            读取资源内容
            
            Args:
                uri (str): 资源URI
                
            Returns:
                str: 资源内容
            """
            if uri == "csv-word://templates":
                templates = get_available_templates()
                return json.dumps({
                    "templates": templates,
                    "count": len(templates),
                    "timestamp": datetime.now().isoformat()
                }, ensure_ascii=False, indent=2)
            
            elif uri == "csv-word://config":
                return json.dumps({
                    "version": __version__,
                    "project_root": str(project_root),
                    "supported_formats": ["docx", "pdf"],
                    "max_file_size": "100MB",
                    "timestamp": datetime.now().isoformat()
                }, ensure_ascii=False, indent=2)
            
            elif uri == "csv-word://status":
                return json.dumps({
                    "status": "running",
                    "active_jobs": len(conversion_jobs),
                    "temp_workspaces": len(temp_directories),
                    "uptime": "N/A",  # 可以添加启动时间跟踪
                    "timestamp": datetime.now().isoformat()
                }, ensure_ascii=False, indent=2)
            
            else:
                raise ValueError(f"未知资源URI: {uri}")
    
    async def _convert_csv_to_word(self, arguments: Dict[str, Any]) -> List[types.TextContent]:
        """
        执行CSV到Word转换
        
        Args:
            arguments (Dict[str, Any]): 转换参数
            
        Returns:
            List[types.TextContent]: 转换结果
        """
        csv_file_path = arguments["csv_file_path"]
        template_type = arguments.get("template_type", "guoziwei")
        output_dir = arguments.get("output_dir")
        output_format = arguments.get("output_format", "docx")
        
        # 验证文件存在
        if not os.path.exists(csv_file_path):
            return [types.TextContent(
                type="text",
                text=f"错误: CSV文件不存在: {csv_file_path}"
            )]
        
        # 创建任务ID
        job_id = str(uuid.uuid4())
        
        try:
            # 记录任务开始
            conversion_jobs[job_id] = {
                "status": "processing",
                "start_time": datetime.now().isoformat(),
                "csv_file": csv_file_path,
                "template_type": template_type,
                "output_format": output_format
            }
            
            # 执行转换
            result_path = convert_csv_to_word(
                csv_file=csv_file_path,
                template_type=template_type,
                output_dir=output_dir
            )
            
            # 更新任务状态
            conversion_jobs[job_id].update({
                "status": "completed",
                "end_time": datetime.now().isoformat(),
                "output_path": result_path
            })
            
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "success": True,
                    "job_id": job_id,
                    "output_path": result_path,
                    "message": f"转换成功: {os.path.basename(result_path)}"
                }, ensure_ascii=False, indent=2)
            )]
            
        except Exception as e:
            # 更新任务状态为失败
            conversion_jobs[job_id].update({
                "status": "failed",
                "end_time": datetime.now().isoformat(),
                "error": str(e)
            })
            
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "success": False,
                    "job_id": job_id,
                    "error": str(e),
                    "message": f"转换失败: {e}"
                }, ensure_ascii=False, indent=2)
            )]
    
    async def _batch_convert_csv(self, arguments: Dict[str, Any]) -> List[types.TextContent]:
        """
        执行批量CSV转换
        
        Args:
            arguments (Dict[str, Any]): 批量转换参数
            
        Returns:
            List[types.TextContent]: 批量转换结果
        """
        csv_files = arguments["csv_files"]
        template_type = arguments.get("template_type", "guoziwei")
        output_dir = arguments.get("output_dir")
        max_workers = arguments.get("max_workers", 4)
        
        # 创建批量任务ID
        job_id = str(uuid.uuid4())
        
        try:
            # 记录批量任务开始
            conversion_jobs[job_id] = {
                "status": "processing",
                "start_time": datetime.now().isoformat(),
                "type": "batch",
                "csv_files": csv_files,
                "template_type": template_type,
                "total_files": len(csv_files),
                "completed_files": 0,
                "failed_files": 0,
                "results": []
            }
            
            # 执行批量转换
            results = []
            for i, csv_file in enumerate(csv_files):
                try:
                    if not os.path.exists(csv_file):
                        result = {
                            "file": csv_file,
                            "success": False,
                            "error": "文件不存在"
                        }
                    else:
                        output_path = convert_csv_to_word(
                            csv_file=csv_file,
                            template_type=template_type,
                            output_dir=output_dir
                        )
                        result = {
                            "file": csv_file,
                            "success": True,
                            "output_path": output_path
                        }
                        conversion_jobs[job_id]["completed_files"] += 1
                    
                except Exception as e:
                    result = {
                        "file": csv_file,
                        "success": False,
                        "error": str(e)
                    }
                    conversion_jobs[job_id]["failed_files"] += 1
                
                results.append(result)
                conversion_jobs[job_id]["results"] = results
            
            # 更新任务状态
            conversion_jobs[job_id].update({
                "status": "completed",
                "end_time": datetime.now().isoformat()
            })
            
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "success": True,
                    "job_id": job_id,
                    "total_files": len(csv_files),
                    "completed_files": conversion_jobs[job_id]["completed_files"],
                    "failed_files": conversion_jobs[job_id]["failed_files"],
                    "results": results
                }, ensure_ascii=False, indent=2)
            )]
            
        except Exception as e:
            conversion_jobs[job_id].update({
                "status": "failed",
                "end_time": datetime.now().isoformat(),
                "error": str(e)
            })
            
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "success": False,
                    "job_id": job_id,
                    "error": str(e)
                }, ensure_ascii=False, indent=2)
            )]
    
    async def _get_available_templates(self, arguments: Dict[str, Any]) -> List[types.TextContent]:
        """
        获取可用模板列表
        
        Args:
            arguments (Dict[str, Any]): 参数（未使用）
            
        Returns:
            List[types.TextContent]: 模板列表
        """
        try:
            templates = get_available_templates()
            
            # 加载模板详细信息
            import yaml
            config_path = project_root / "templates_config.yaml"
            template_details = {}
            
            if config_path.exists():
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = yaml.safe_load(f)
                    template_details = config.get('templates', {})
            
            # 构建详细的模板信息
            detailed_templates = []
            for template_name in templates:
                template_info = {
                    "name": template_name,
                    "display_name": template_details.get(template_name, {}).get('name', template_name),
                    "description": template_details.get(template_name, {}).get('description', ''),
                    "available": True
                }
                detailed_templates.append(template_info)
            
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "success": True,
                    "templates": detailed_templates,
                    "count": len(detailed_templates)
                }, ensure_ascii=False, indent=2)
            )]
            
        except Exception as e:
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "success": False,
                    "error": str(e)
                }, ensure_ascii=False, indent=2)
            )]
    
    async def _validate_csv_file(self, arguments: Dict[str, Any]) -> List[types.TextContent]:
        """
        验证CSV文件
        
        Args:
            arguments (Dict[str, Any]): 验证参数
            
        Returns:
            List[types.TextContent]: 验证结果
        """
        csv_file_path = arguments["csv_file_path"]
        
        try:
            validation_result = validate_csv_file(csv_file_path)
            
            return [types.TextContent(
                type="text",
                text=json.dumps(validation_result, ensure_ascii=False, indent=2)
            )]
            
        except Exception as e:
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "valid": False,
                    "error": str(e),
                    "message": f"验证失败: {e}"
                }, ensure_ascii=False, indent=2)
            )]
    
    async def _get_conversion_status(self, arguments: Dict[str, Any]) -> List[types.TextContent]:
        """
        获取转换任务状态
        
        Args:
            arguments (Dict[str, Any]): 状态查询参数
            
        Returns:
            List[types.TextContent]: 任务状态
        """
        job_id = arguments["job_id"]
        
        if job_id in conversion_jobs:
            job_info = conversion_jobs[job_id]
            return [types.TextContent(
                type="text",
                text=json.dumps(job_info, ensure_ascii=False, indent=2)
            )]
        else:
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "error": f"未找到任务ID: {job_id}"
                }, ensure_ascii=False, indent=2)
            )]
    
    async def _create_temp_workspace(self, arguments: Dict[str, Any]) -> List[types.TextContent]:
        """
        创建临时工作空间
        
        Args:
            arguments (Dict[str, Any]): 工作空间参数
            
        Returns:
            List[types.TextContent]: 工作空间信息
        """
        workspace_name = arguments.get("workspace_name", "csv_word_workspace")
        
        try:
            # 创建临时目录
            temp_dir = tempfile.mkdtemp(prefix=f"{workspace_name}_")
            workspace_id = str(uuid.uuid4())
            
            # 记录工作空间
            temp_directories[workspace_id] = {
                "path": temp_dir,
                "name": workspace_name,
                "created_time": datetime.now().isoformat()
            }
            
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "success": True,
                    "workspace_id": workspace_id,
                    "workspace_path": temp_dir,
                    "message": f"临时工作空间已创建: {temp_dir}"
                }, ensure_ascii=False, indent=2)
            )]
            
        except Exception as e:
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "success": False,
                    "error": str(e)
                }, ensure_ascii=False, indent=2)
            )]
    
    async def _cleanup_temp_workspace(self, arguments: Dict[str, Any]) -> List[types.TextContent]:
        """
        清理临时工作空间
        
        Args:
            arguments (Dict[str, Any]): 清理参数
            
        Returns:
            List[types.TextContent]: 清理结果
        """
        workspace_id = arguments["workspace_id"]
        
        if workspace_id not in temp_directories:
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "success": False,
                    "error": f"未找到工作空间ID: {workspace_id}"
                }, ensure_ascii=False, indent=2)
            )]
        
        try:
            workspace_info = temp_directories[workspace_id]
            temp_dir = workspace_info["path"]
            
            # 清理目录
            import shutil
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            
            # 移除记录
            del temp_directories[workspace_id]
            
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "success": True,
                    "workspace_id": workspace_id,
                    "message": f"工作空间已清理: {temp_dir}"
                }, ensure_ascii=False, indent=2)
            )]
            
        except Exception as e:
            return [types.TextContent(
                type="text",
                text=json.dumps({
                    "success": False,
                    "error": str(e)
                }, ensure_ascii=False, indent=2)
            )]

async def main():
    """
    主函数：启动MCP服务器
    """
    logger.info(f"启动CSV-Word转换工具MCP服务器 v{__version__}")
    
    # 创建服务器实例
    mcp_server = CSVWordMCPServer()
    
    # 运行服务器
    async with stdio_server() as (read_stream, write_stream):
        await mcp_server.server.run(
            read_stream,
            write_stream,
            InitializationOptions(
                server_name="csv-word-converter",
                server_version=__version__,
                capabilities=mcp_server.server.get_capabilities(
                    notification_options=None,
                    experimental_capabilities=None,
                ),
            ),
        )

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("MCP服务器已停止")
    except Exception as e:
        logger.error(f"MCP服务器运行出错: {e}")
        logger.error(traceback.format_exc())
        sys.exit(1)