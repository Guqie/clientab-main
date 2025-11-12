#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CSV-Word转换工具 - Streamlit前端

这是一个基于Streamlit的Web前端，为CSV到Word转换工具提供友好的用户界面。

功能特性:
- 文件拖拽上传
- 模板选择和配置
- 实时预览和转换
- 批量处理支持
- 结果下载和管理

作者: AI Development Team
版本: 1.0.0
"""

import streamlit as st
import pandas as pd
import os
import sys
import tempfile
import shutil
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Any
import yaml
import logging
from datetime import datetime
import traceback

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
except ImportError as e:
    st.error(f"导入核心模块失败: {e}")
    st.stop()

# 配置页面
st.set_page_config(
    page_title="CSV-Word转换工具",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义CSS样式
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .upload-section {
        border: 2px dashed #cccccc;
        border-radius: 10px;
        padding: 2rem;
        text-align: center;
        margin: 1rem 0;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .error-box {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def initialize_session_state():
    """
    初始化Streamlit会话状态
    """
    if 'uploaded_files' not in st.session_state:
        st.session_state.uploaded_files = []
    if 'conversion_results' not in st.session_state:
        st.session_state.conversion_results = []
    if 'temp_dir' not in st.session_state:
        st.session_state.temp_dir = tempfile.mkdtemp(prefix="csv_word_")
    if 'processing' not in st.session_state:
        st.session_state.processing = False
    if 'save_to_outputs' not in st.session_state:
        st.session_state.save_to_outputs = False
    if 'zip_ready' not in st.session_state:
        st.session_state.zip_ready = None

def cleanup_temp_files():
    """
    清理临时文件
    """
    try:
        if 'temp_dir' in st.session_state and os.path.exists(st.session_state.temp_dir):
            shutil.rmtree(st.session_state.temp_dir)
            st.session_state.temp_dir = tempfile.mkdtemp(prefix="csv_word_")
    except Exception as e:
        st.warning(f"清理临时文件时出错: {e}")

def load_template_config() -> Dict[str, Any]:
    """
    加载模板配置
    
    Returns:
        Dict[str, Any]: 模板配置字典
    """
    try:
        config_path = project_root / "templates_config.yaml"
        if config_path.exists():
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        else:
            st.warning("未找到模板配置文件，使用默认配置")
            return {"templates": {}}
    except Exception as e:
        st.error(f"加载模板配置失败: {e}")
        return {"templates": {}}

def validate_uploaded_file(uploaded_file) -> Dict[str, Any]:
    """
    验证上传的CSV文件
    
    Args:
        uploaded_file: Streamlit上传的文件对象
        
    Returns:
        Dict[str, Any]: 验证结果，统一使用'valid'键名
    """
    try:
        # 保存临时文件
        temp_path = os.path.join(st.session_state.temp_dir, uploaded_file.name)
        with open(temp_path, 'wb') as f:
            f.write(uploaded_file.getbuffer())
        
        # 验证文件
        validation_result = validate_csv_file(temp_path)
        
        # 统一键名：将is_valid转换为valid
        if 'is_valid' in validation_result:
            validation_result['valid'] = validation_result.pop('is_valid')
        elif 'valid' not in validation_result:
            # 防御性编程：如果既没有is_valid也没有valid，默认为False
            validation_result['valid'] = False
            
        validation_result['temp_path'] = temp_path
        
        return validation_result
    except Exception as e:
        return {
            'valid': False,
            'error': str(e),
            'message': f"文件验证失败: {e}"
        }


def check_required_columns(file_path: str, required: Optional[List[str]] = None) -> Dict[str, Any]:
    """
    校验CSV是否包含必需列

    参数:
        file_path (str): CSV文件路径
        required (Optional[List[str]]): 必需列名列表，默认['title','content']

    返回:
        Dict[str, Any]: 包含valid与missing等信息的校验结果
    """
    required_cols = required or ['title', 'content']
    try:
        df = pd.read_csv(file_path, encoding='utf-8')
    except UnicodeDecodeError:
        df = pd.read_csv(file_path, encoding='gbk')
    cols = set(df.columns)
    missing = [c for c in required_cols if c not in cols]
    return {
        'valid': len(missing) == 0,
        'missing': missing,
        'columns': list(cols)
    }

def preview_csv_data(file_path: str, max_rows: int = 10) -> pd.DataFrame:
    """
    预览CSV数据
    
    Args:
        file_path (str): CSV文件路径
        max_rows (int): 最大显示行数
        
    Returns:
        pd.DataFrame: 预览数据
    """
    try:
        df = pd.read_csv(file_path, encoding='utf-8')
        return df.head(max_rows)
    except UnicodeDecodeError:
        try:
            df = pd.read_csv(file_path, encoding='gbk')
            return df.head(max_rows)
        except Exception as e:
            st.error(f"读取CSV文件失败: {e}")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"预览CSV数据失败: {e}")
        return pd.DataFrame()

def convert_single_file(file_path: str, template_type: str, output_format: str = "docx") -> Dict[str, Any]:
    """
    转换单个文件
    
    Args:
        file_path (str): CSV文件路径
        template_type (str): 模板类型
        output_format (str): 输出格式
        
    Returns:
        Dict[str, Any]: 转换结果
    """
    try:
        # 设置输出目录
        output_dir = (project_root / "outputs").as_posix() if st.session_state.save_to_outputs else st.session_state.temp_dir

        # 执行转换（生成docx）
        docx_path = convert_csv_to_word(
            csv_file=file_path,
            template_type=template_type,
            output_dir=output_dir
        )

        final_path = docx_path

        # 如需其他格式（例如pdf），进行二次转换
        if output_format and output_format != "docx":
            import pandas as pd
            from csv_word_converter.output_formats import convert_to_format
            csv_data = pd.read_csv(file_path).to_dict('records')
            filename = Path(docx_path).stem + f".{output_format}"
            out_dir = Path(output_dir)
            out_dir.mkdir(exist_ok=True)
            target_path = out_dir / filename
            final_path = asyncio.run(convert_to_format(
                csv_data=csv_data,
                output_path=target_path,
                title=f"转换报告 - {template_type}",
                source=file_path
            ))

        return {
            'success': True,
            'output_path': str(final_path),
            'message': f"转换成功: {os.path.basename(str(final_path))}"
        }
    except Exception as e:
        return {
            'success': False,
            'error': str(e),
            'message': f"转换失败: {e}"
        }

def create_download_link(file_path: str, link_text: str) -> str:
    """
    创建文件下载链接
    
    Args:
        file_path (str): 文件路径
        link_text (str): 链接文本
        
    Returns:
        str: 下载链接HTML
    """
    try:
        with open(file_path, 'rb') as f:
            file_data = f.read()
        
        import base64
        b64_data = base64.b64encode(file_data).decode()
        file_name = os.path.basename(file_path)
        
        return f'<a href="data:application/octet-stream;base64,{b64_data}" download="{file_name}">{link_text}</a>'
    except Exception as e:
        return f"下载链接生成失败: {e}"

def main():
    """
    主应用函数
    """
    # 初始化会话状态
    initialize_session_state()
    
    # 页面标题
    st.markdown('<h1 class="main-header">📄 CSV-Word转换工具</h1>', unsafe_allow_html=True)
    st.markdown(f"<p style='text-align: center; color: #666;'>版本 {__version__} | 专业的CSV到Word文档转换解决方案</p>", unsafe_allow_html=True)
    
    # 侧边栏配置
    with st.sidebar:
        st.header("⚙️ 转换配置")
        
        # 加载模板配置
        template_config = load_template_config()
        available_templates = list(template_config.get('templates', {}).keys())
        
        if not available_templates:
            st.error("未找到可用模板")
            return
        
        # 模板选择
        selected_template = st.selectbox(
            "选择文档模板",
            available_templates,
            help="选择适合的Word文档模板"
        )
        
        # 显示模板信息
        if selected_template in template_config['templates']:
            template_info = template_config['templates'][selected_template]
            st.info(f"**{template_info.get('name', selected_template)}**")
        
        # 输出格式选择
        output_format = st.selectbox(
            "输出格式",
            ["docx", "pdf"],
            help="选择输出文档格式"
        )

        # 目标输出目录选择
        st.session_state.save_to_outputs = st.checkbox("保存到项目 outputs 目录", value=False, help="默认保存在临时目录，可勾选保存到项目outputs")
        
        # 高级选项
        st.subheader("🔧 高级选项")
        
        batch_mode = st.checkbox("批量处理模式", help="同时处理多个CSV文件")
        show_preview = st.checkbox("显示数据预览", value=True, help="上传后显示CSV数据预览")
        auto_cleanup = st.checkbox("自动清理临时文件", value=True, help="转换完成后自动清理临时文件")
        
        # 清理按钮
        if st.button("🗑️ 清理临时文件"):
            cleanup_temp_files()
            st.success("临时文件已清理")
    
    # 主内容区域
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📁 文件上传")
        
        # 文件上传区域
        if batch_mode:
            uploaded_files = st.file_uploader(
                "选择CSV文件（支持多文件）",
                type=['csv'],
                accept_multiple_files=True,
                help="拖拽CSV文件到此区域或点击选择文件"
            )
        else:
            uploaded_files = st.file_uploader(
                "选择CSV文件",
                type=['csv'],
                accept_multiple_files=False,
                help="拖拽CSV文件到此区域或点击选择文件"
            )
            if uploaded_files:
                uploaded_files = [uploaded_files]
        
        # 处理上传的文件
        if uploaded_files:
            st.session_state.uploaded_files = uploaded_files
            
            # 验证文件
            validation_results = []
            for uploaded_file in uploaded_files:
                result = validate_uploaded_file(uploaded_file)
                validation_results.append({
                    'file': uploaded_file,
                    'result': result
                })
            
            # 显示验证结果
            valid_files = []
            for item in validation_results:
                file_name = item['file'].name
                result = item['result']
                
                # 防御性编程：确保result是字典且包含valid键
                if not isinstance(result, dict):
                    st.error(f"❌ {file_name} - 验证结果格式错误")
                    continue
                    
                is_valid = result.get('valid', False)  # 使用get方法避免KeyError
                
                if is_valid:
                    st.success(f"✅ {file_name} - 验证通过")
                    valid_files.append(item)
                    
                    # 显示数据预览
                    if show_preview and 'temp_path' in result:
                        with st.expander(f"📊 预览 {file_name}"):
                            preview_df = preview_csv_data(result['temp_path'])
                            if not preview_df.empty:
                                st.dataframe(preview_df, use_container_width=True)
                                st.info(f"数据形状: {preview_df.shape[0]} 行 × {preview_df.shape[1]} 列")
                                # 必需列校验
                                column_check = check_required_columns(result['temp_path'])
                                if column_check['valid']:
                                    st.success("必需列校验通过: 需包含 'title' 与 'content'")
                                else:
                                    st.warning(f"缺少必需列: {', '.join(column_check['missing'])}。转换可能失败或内容不完整。")
                            else:
                                st.warning("无法预览数据")
                else:
                    error_message = result.get('message', result.get('error', '验证失败'))
                    st.error(f"❌ {file_name} - {error_message}")
            
            # 转换按钮
            if valid_files and not st.session_state.processing:
                if st.button("🚀 开始转换", type="primary"):
                    st.session_state.processing = True
                    
                    # 创建进度条
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    conversion_results = []
                    
                    for i, item in enumerate(valid_files):
                        file_name = item['file'].name
                        temp_path = item['result']['temp_path']
                        
                        status_text.text(f"正在转换: {file_name}")
                        
                        # 执行转换
                        result = convert_single_file(temp_path, selected_template, output_format)
                        result['file_name'] = file_name
                        conversion_results.append(result)
                        
                        # 更新进度
                        progress_bar.progress((i + 1) / len(valid_files))
                    
                    # 保存结果
                    st.session_state.conversion_results = conversion_results
                    st.session_state.processing = False
                    
                    status_text.text("转换完成！")
                    st.success("所有文件转换完成！")

                    # 若为批量成功，提供ZIP打包下载
                    success_paths = [r['output_path'] for r in conversion_results if r.get('success')]
                    if len(success_paths) > 1:
                        zip_name = f"converted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
                        zip_path = Path(st.session_state.temp_dir) / zip_name
                        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for p in success_paths:
                                zf.write(p, arcname=os.path.basename(p))
                        st.session_state.zip_ready = str(zip_path)
                        st.info("已生成批量下载包")
            
            elif st.session_state.processing:
                st.info("⏳ 正在处理中，请稍候...")
    
    with col2:
        st.header("📋 转换结果")
        
        if st.session_state.conversion_results:
            for result in st.session_state.conversion_results:
                file_name = result['file_name']
                
                if result['success']:
                    st.success(f"✅ {file_name}")
                    
                    # 提供下载链接
                    if os.path.exists(result['output_path']):
                        with open(result['output_path'], 'rb') as f:
                            st.download_button(
                                label=f"📥 下载 {os.path.basename(result['output_path'])}",
                                data=f.read(),
                                file_name=os.path.basename(result['output_path']),
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                else:
                    st.error(f"❌ {file_name}")
                    st.error(result.get('message', '转换失败'))
        else:
            st.info("暂无转换结果")

        # 批量ZIP下载
        if st.session_state.zip_ready and os.path.exists(st.session_state.zip_ready):
            with open(st.session_state.zip_ready, 'rb') as f:
                st.download_button(
                    label=f"📦 下载批量结果 {os.path.basename(st.session_state.zip_ready)}",
                    data=f.read(),
                    file_name=os.path.basename(st.session_state.zip_ready),
                    mime="application/zip"
                )
    
    # 页脚信息
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: #666; font-size: 0.9rem;'>
        <p>CSV-Word转换工具 | 支持多种模板和格式 | 
        <a href='#' style='color: #1f77b4;'>使用文档</a> | 
        <a href='#' style='color: #1f77b4;'>问题反馈</a></p>
    </div>
    """, unsafe_allow_html=True)
    
    # 自动清理（如果启用）
    if auto_cleanup and st.session_state.conversion_results:
        # 在页面刷新时清理
        pass

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        st.error(f"应用运行出错: {e}")
        st.error("详细错误信息:")
        st.code(traceback.format_exc())