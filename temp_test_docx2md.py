"""docx 转 markdown 转换器"""
from docx import Document
from docx.shared import Pt
from pathlib import Path
import re


def docx_to_markdown(docx_path: str, output_path: str = None) -> str:
    doc = Document(docx_path)
    lines = []

    def get_para_style_name(para):
        style = para.style
        if style is None:
            return None
        return style.name

    def is_heading_style(style_name):
        if style_name is None:
            return False
        return "Heading" in style_name or "标题" in style_name

    def get_heading_level(style_name):
        if style_name is None:
            return 0
        if "1" in style_name:
            return 1
        if "2" in style_name:
            return 2
        if "3" in style_name:
            return 3
        if "4" in style_name:
            return 4
        if "标题 1" in style_name or "heading 1" in style_name.lower():
            return 1
        if "标题 2" in style_name or "heading 2" in style_name.lower():
            return 2
        if "标题 3" in style_name or "heading 3" in style_name.lower():
            return 3
        return 0

    def run_to_text(runs):
        text = ""
        for run in runs:
            t = run.text or ""
            if run.bold:
                t = f"**{t}**"
            if run.italic:
                t = f"*{t}*"
            if run.underline:
                t = f"<u>{t}</u>"
            text += t
        return text

    for para in doc.paragraphs:
        style_name = get_para_style_name(para)
        text = run_to_text(para.runs).strip()

        if not text:
            lines.append("")
            continue

        if is_heading_style(style_name):
            level = get_heading_level(style_name)
            if level > 0:
                lines.append(f"{'#' * level} {text}")
            else:
                lines.append(f"## {text}")
        else:
            lines.append(text)

    result = "\n".join(lines)

    if output_path:
        Path(output_path).write_text(result, encoding="utf-8")
        print(f"已保存到: {output_path}")

    return result


if __name__ == "__main__":
    import sys
    docx_path = r"d:\桌面\clientab-main\temp-data\markdown格式转换\中车-战新产业与未来产业月报.docx"
    md_path = docx_path.replace(".docx", ".md")
    content = docx_to_markdown(docx_path, md_path)
    print(f"转换完成！共 {len(content)} 字符")
    print("预览前 2000 字符：")
    print("=" * 60)
    print(content[:2000])
