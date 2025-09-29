#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
详细检查房地产周刊文档中的图片对齐情况
"""

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

def detailed_check():
    """详细检查所有图片段落的对齐情况"""
    doc = Document('temp-data/1758619325_386932_guoziwei.docx')
    print('详细检查所有图片段落:')
    
    image_paragraph_count = 0
    non_center_paragraphs = []
    
    for i, paragraph in enumerate(doc.paragraphs):
        has_image = False
        for run in paragraph.runs:
            if run.element.xpath('.//a:blip'):
                has_image = True
                break
        
        if has_image:
            image_paragraph_count += 1
            alignment = paragraph.alignment
            
            if alignment != WD_ALIGN_PARAGRAPH.CENTER:
                non_center_paragraphs.append({
                    'index': i,
                    'image_num': image_paragraph_count,
                    'alignment': alignment,
                    'text': paragraph.text[:50]
                })
    
    print(f'总图片段落数: {image_paragraph_count}')
    print(f'非居中对齐的图片段落数: {len(non_center_paragraphs)}')
    
    if non_center_paragraphs:
        print('非居中对齐的图片段落详情:')
        for p in non_center_paragraphs:
            print(f'  图片#{p["image_num"]} (段落{p["index"]}): 对齐={p["alignment"]}, 文本={repr(p["text"])}')
    else:
        print('✅ 所有图片段落都是居中对齐的!')
    
    return len(non_center_paragraphs) == 0

if __name__ == '__main__':
    detailed_check()