# -*- coding: utf-8 -*-
"""
临时端到端验证脚本：
1) 寻找或生成最新 new_energy 模板的 Word 文档；
2) 扫描段落，定位包含图片的段落，记录图片宽度与前后相邻文本；
3) 生成 temp-data/_e2e_image_report.txt 报告文件，用于人工核查图片是否按占位符顺序插入、对齐与宽度是否符合预期。

运行方式：放在项目根目录，使用虚拟环境的 python 执行：
python _tmp_make_report.py
"""
import os
import glob
from typing import List, Tuple, Optional

from docx import Document

try:
    # 优先使用项目通用入口，避免直接依赖其他脚本
    from universal_csv_to_word import csv_to_word_universal
except Exception:
    csv_to_word_universal = None  # 若不可用，仅做已有文档的静态分析


def _find_or_build_docx() -> str:
    """定位最新的 new_energy 文档，若不存在且具备生成能力则触发生成。

    返回：
        生成或发现的 .docx 文件的相对路径。
    """
    candidates = [
        p for p in glob.glob(os.path.join('temp-data', '*new_energy*.docx'))
        if not os.path.basename(p).startswith('~$')
    ]
    if candidates:
        candidates.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        return candidates[0]

    # 若无候选且支持生成，则尝试从默认 CSV 生成
    csv_path = os.path.join('temp-data', '新能源9月15日刊.csv')
    if csv_to_word_universal and os.path.exists(csv_path):
        return csv_to_word_universal(csv_path, 'new_energy', 'templates_config.yaml')

    raise FileNotFoundError('未找到 new_energy 文档，且无法自动生成。')


def _extract_image_info(doc_path: str) -> Tuple[List[Tuple[int, Optional[float], str, str]], List[float]]:
    """从 Word 文档中提取图片所在段落的索引、图片宽度（英寸）及上下文文本。

    参数：
        doc_path: Word 文档路径。
    返回：
        records: [(段落索引, 宽度(英寸或None), 上文文本, 下文文本)]
        widths: 所有检测到的图片宽度列表（英寸）。
    说明：
        - 通过在 run 元素中查找 w:drawing 判断该段是否含内联图片；
        - 通过 a:ext@cx 抽取宽度，换算关系：EMU -> 英寸 = cx / 914400。
    """
    doc = Document(doc_path)
    paras = doc.paragraphs
    records: List[Tuple[int, Optional[float], str, str]] = []
    widths: List[float] = []

    for i, p in enumerate(paras):
        has_img = any(r.element.xpath('.//w:drawing') for r in p.runs)
        if not has_img:
            continue

        # 找上一个与下一个非空文本段，便于人工核对占位符顺序
        prev_t = ''
        for j in range(i - 1, -1, -1):
            t = paras[j].text.strip()
            if t:
                prev_t = t[:120]
                break
        next_t = ''
        for j in range(i + 1, len(paras)):
            t = paras[j].text.strip()
            if t:
                next_t = t[:120]
                break

        # 解析图片宽度
        w_in: Optional[float] = None
        for r in p.runs:
            for d in r.element.xpath(
                './/wp:inline//a:graphic//a:graphicData//pic:pic//pic:spPr//a:xfrm//a:ext'
            ):
                try:
                    cx = int(d.get('{http://schemas.openxmlformats.org/drawingml/2006/main}cx'))
                    w_in = cx / 914400.0
                    widths.append(w_in)
                except Exception:
                    pass
        records.append((i, w_in, prev_t, next_t))
    return records, widths


def main() -> None:
    """主入口：生成或定位 new_energy 文档，提取图片信息并写出报告。"""
    os.makedirs('temp-data', exist_ok=True)
    report_path = os.path.join('temp-data', '_e2e_image_report.txt')

    lines: List[str] = []
    try:
        doc_path = _find_or_build_docx()
        lines.append(f'报告文档: {doc_path}')
        records, widths = _extract_image_info(doc_path)
        lines.append(f'图片段落数量: {len(records)}')
        for k, (pos, w, prev_t, next_t) in enumerate(records[:60], start=1):
            w_str = f'{w:.2f}' if isinstance(w, float) else 'None'
            lines.append(f'#{k} 段落索引={pos}, 宽度={w_str} in, 上文="{prev_t}", 下文="{next_t}"')
        if widths:
            avg = sum(widths) / len(widths)
            head = ', '.join(f'{x:.2f}' for x in widths[:10])
            lines.append(f'平均宽度: {avg:.2f} in; 前10: {head}')
    except Exception as e:
        lines.append(f'分析失败: {e!r}')

    with open(report_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(report_path)


if __name__ == '__main__':
    main()