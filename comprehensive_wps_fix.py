#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
综合WPS兼容性修复脚本
整合所有发现的问题并提供完整解决方案
"""

import logging
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.enum.text import WD_COLOR_INDEX
import os
from pathlib import Path

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)

class WPSCompatibilityFixer:
    """
    WPS兼容性修复器
    """
    
    def __init__(self, input_doc_path, output_doc_path):
        self.input_doc_path = input_doc_path
        self.output_doc_path = output_doc_path
        self.doc = None
        self.document_xml = None
        self.fixes_applied = []
    
    def load_document(self):
        """
        加载文档
        """
        logger.info(f"加载文档: {self.input_doc_path}")
        try:
            self.doc = Document(self.input_doc_path)
            self.document_xml = self.doc._element
            return True
        except Exception as e:
            logger.error(f"加载文档失败: {e}")
            return False
    
    def fix_ampersand_encoding(self):
        """
        修复半角&符号编码问题
        """
        logger.info("\n--- 修复半角&符号编码问题 ---")
        fixed_count = 0
        
        # 修复所有文本元素
        for elem in self.document_xml.iter():
            if elem.tag.endswith('}t') and elem.text and '&' in elem.text:
                original_text = elem.text
                fixed_text = original_text.replace('&', '＆')
                elem.text = fixed_text
                fixed_count += 1
                logger.info(f"修复文本: '{original_text}' -> '{fixed_text}'")
        
        # 修复超链接属性
        for hyperlink in self.document_xml.iter():
            if hyperlink.tag.endswith('}hyperlink'):
                # 修复锚点
                anchor = hyperlink.get(qn('w:anchor'))
                if anchor and '&' in anchor:
                    fixed_anchor = anchor.replace('&', '＆')
                    hyperlink.set(qn('w:anchor'), fixed_anchor)
                    fixed_count += 1
                    logger.info(f"修复锚点: '{anchor}' -> '{fixed_anchor}'")
                
                # 修复工具提示
                tooltip = hyperlink.get(qn('w:tooltip'))
                if tooltip and '&' in tooltip:
                    fixed_tooltip = tooltip.replace('&', '＆')
                    hyperlink.set(qn('w:tooltip'), fixed_tooltip)
                    fixed_count += 1
                    logger.info(f"修复工具提示: '{tooltip}' -> '{fixed_tooltip}'")
        
        self.fixes_applied.append(f"修复半角&符号: {fixed_count} 处")
        return fixed_count
    
    def optimize_hyperlink_structure(self):
        """
        优化超链接结构，提高WPS兼容性
        """
        logger.info("\n--- 优化超链接结构 ---")
        optimized_count = 0
        
        for hyperlink in self.document_xml.iter():
            if hyperlink.tag.endswith('}hyperlink'):
                # 简化工具提示
                tooltip = hyperlink.get(qn('w:tooltip'))
                if tooltip and len(tooltip) > 50:
                    simplified_tooltip = "返回目录"
                    hyperlink.set(qn('w:tooltip'), simplified_tooltip)
                    optimized_count += 1
                    logger.info(f"简化工具提示: '{tooltip}' -> '{simplified_tooltip}'")
                
                # 确保锚点格式正确
                anchor = hyperlink.get(qn('w:anchor'))
                if anchor:
                    # 移除可能的空格和特殊字符
                    clean_anchor = anchor.strip().replace(' ', '')
                    if clean_anchor != anchor:
                        hyperlink.set(qn('w:anchor'), clean_anchor)
                        optimized_count += 1
                        logger.info(f"清理锚点: '{anchor}' -> '{clean_anchor}'")
        
        self.fixes_applied.append(f"优化超链接结构: {optimized_count} 处")
        return optimized_count
    
    def ensure_bookmark_integrity(self):
        """
        确保书签完整性
        """
        logger.info("\n--- 确保书签完整性 ---")
        
        # 收集所有书签开始和结束标记
        bookmark_starts = {}
        bookmark_ends = {}
        
        for elem in self.document_xml.iter():
            if elem.tag.endswith('}bookmarkStart'):
                bookmark_id = elem.get(qn('w:id'))
                name = elem.get(qn('w:name'))
                bookmark_starts[bookmark_id] = name
            elif elem.tag.endswith('}bookmarkEnd'):
                bookmark_id = elem.get(qn('w:id'))
                bookmark_ends[bookmark_id] = True
        
        # 检查不匹配的书签
        missing_ends = []
        for bookmark_id, name in bookmark_starts.items():
            if bookmark_id not in bookmark_ends:
                missing_ends.append((bookmark_id, name))
                logger.warning(f"书签 '{name}' (ID: {bookmark_id}) 缺少结束标记")
        
        # 报告书签状态
        logger.info(f"书签开始标记: {len(bookmark_starts)} 个")
        logger.info(f"书签结束标记: {len(bookmark_ends)} 个")
        logger.info(f"不匹配的书签: {len(missing_ends)} 个")
        
        self.fixes_applied.append(f"书签完整性检查: {len(bookmark_starts)} 个书签")
        return len(missing_ends) == 0
    
    def add_wps_compatibility_metadata(self):
        """
        添加WPS兼容性元数据
        """
        logger.info("\n--- 添加WPS兼容性元数据 ---")
        
        # 获取文档核心属性
        core_props = self.doc.core_properties
        
        # 添加兼容性说明
        if not core_props.comments:
            core_props.comments = "已优化WPS兼容性 - 修复超链接跳转问题"
        else:
            if "WPS兼容性" not in core_props.comments:
                core_props.comments += " | 已优化WPS兼容性"
        
        # 更新修改者信息
        core_props.last_modified_by = "WPS兼容性修复工具"
        
        self.fixes_applied.append("添加WPS兼容性元数据")
        logger.info("已添加WPS兼容性元数据")
    
    def validate_fixes(self):
        """
        验证修复结果
        """
        logger.info("\n--- 验证修复结果 ---")
        
        issues_found = []
        
        # 检查是否还有半角&符号
        ampersand_count = 0
        for elem in self.document_xml.iter():
            if elem.tag.endswith('}t') and elem.text and '&' in elem.text:
                ampersand_count += 1
                issues_found.append(f"仍存在半角&符号: {elem.text}")
        
        # 检查超链接完整性
        hyperlink_issues = 0
        for hyperlink in self.document_xml.iter():
            if hyperlink.tag.endswith('}hyperlink'):
                anchor = hyperlink.get(qn('w:anchor'))
                if not anchor:
                    hyperlink_issues += 1
                    issues_found.append("发现无锚点的超链接")
        
        # 报告验证结果
        if not issues_found:
            logger.info("✓ 所有修复验证通过")
            return True
        else:
            logger.warning(f"发现 {len(issues_found)} 个问题:")
            for issue in issues_found[:5]:  # 只显示前5个
                logger.warning(f"  - {issue}")
            return False
    
    def save_document(self):
        """
        保存修复后的文档
        """
        logger.info(f"\n保存修复后的文档: {self.output_doc_path}")
        try:
            # 确保输出目录存在
            output_dir = Path(self.output_doc_path).parent
            output_dir.mkdir(parents=True, exist_ok=True)
            
            self.doc.save(self.output_doc_path)
            logger.info("✓ 文档保存成功")
            return True
        except Exception as e:
            logger.error(f"保存文档失败: {e}")
            return False
    
    def generate_report(self):
        """
        生成修复报告
        """
        logger.info("\n" + "="*60)
        logger.info("WPS兼容性修复报告")
        logger.info("="*60)
        
        logger.info(f"输入文档: {self.input_doc_path}")
        logger.info(f"输出文档: {self.output_doc_path}")
        logger.info(f"文档大小: {os.path.getsize(self.output_doc_path) / 1024:.1f} KB")
        
        logger.info("\n应用的修复:")
        for fix in self.fixes_applied:
            logger.info(f"  ✓ {fix}")
        
        logger.info("\nWPS使用建议:")
        logger.info("1. 使用WPS Office 2019或更新版本")
        logger.info("2. 确保启用了超链接功能")
        logger.info("3. 如果仍有问题，尝试另存为.docx格式")
        logger.info("4. 检查WPS的兼容性设置")
        
        logger.info("\n" + "="*60)
    
    def run_comprehensive_fix(self):
        """
        运行综合修复流程
        """
        logger.info("=== 开始WPS综合兼容性修复 ===")
        
        # 1. 加载文档
        if not self.load_document():
            return False
        
        # 2. 应用各种修复
        self.fix_ampersand_encoding()
        self.optimize_hyperlink_structure()
        self.ensure_bookmark_integrity()
        self.add_wps_compatibility_metadata()
        
        # 3. 验证修复结果
        validation_passed = self.validate_fixes()
        
        # 4. 保存文档
        if self.save_document():
            # 5. 生成报告
            self.generate_report()
            return validation_passed
        else:
            return False

def main():
    """
    主函数
    """
    # 文档路径配置
    input_doc = r"D:\桌面\clientab-main\temp-data\1758118990_466357_technology_fixed.docx"
    output_doc = r"D:\桌面\clientab-main\temp-data\1758118990_466357_technology_wps_final.docx"
    
    # 创建修复器实例
    fixer = WPSCompatibilityFixer(input_doc, output_doc)
    
    # 运行综合修复
    success = fixer.run_comprehensive_fix()
    
    if success:
        logger.info("\n🎉 WPS兼容性修复完成！")
        logger.info(f"请在WPS中打开文档测试: {output_doc}")
    else:
        logger.error("\n❌ 修复过程中出现问题，请检查日志")

if __name__ == "__main__":
    main()