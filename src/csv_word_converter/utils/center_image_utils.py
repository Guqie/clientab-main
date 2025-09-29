#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图片描述段落居中处理工具

该模块提供图片描述段落的居中处理功能，用于文档后处理阶段。
"""

from docx.enum.text import WD_ALIGN_PARAGRAPH


def center_image_description_paragraphs(doc):
    """
    将文档中的图片描述段落居中对齐
    
    该函数遍历文档中的所有段落，识别包含图片描述的段落并将其设置为居中对齐。
    通常用于文档的后处理阶段，提升文档的视觉效果。
    
    Args:
        doc: python-docx Document对象，要处理的Word文档
        
    Returns:
        None: 直接修改传入的文档对象
        
    Note:
        - 该函数会识别包含"图"、"Figure"、"图片"等关键词的段落
        - 仅对符合条件的段落进行居中处理，不影响其他段落格式
        - 处理过程中会保留段落的其他格式属性
    """
    if not doc:
        return
        
    try:
        # 定义图片描述的关键词
        image_keywords = ["图", "Figure", "图片", "Fig.", "图表", "插图"]
        
        # 遍历文档中的所有段落
        for paragraph in doc.paragraphs:
            if not paragraph.text.strip():
                continue
                
            # 检查段落是否包含图片描述关键词
            text = paragraph.text.strip()
            is_image_description = any(keyword in text for keyword in image_keywords)
            
            # 如果是图片描述段落，设置为居中对齐
            if is_image_description:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
        # 同时处理表格中的段落
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if not paragraph.text.strip():
                            continue
                            
                        text = paragraph.text.strip()
                        is_image_description = any(keyword in text for keyword in image_keywords)
                        
                        if is_image_description:
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            
    except Exception as e:
        print(f"处理图片描述段落居中时出错: {e}")