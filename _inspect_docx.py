import os, time
from glob import glob
from docx import Document
from docx.shared import Inches

# 找到最新生成的 new_energy 文档
paths = sorted(glob(r'temp-data/*_new_energy.docx'), key=os.path.getmtime, reverse=True)
print('找到文档数量:', len(paths))
if not paths:
    raise SystemExit('未找到 new_energy 文档')
path = paths[0]
print('最新文档:', path)

# 打开并统计图片
doc = Document(path)
shapes = getattr(doc, 'inline_shapes', [])
try:
    count = len(shapes)
except TypeError:
    # 旧版可能不是可迭代
    count = 0
print('Inline 图片数量:', count)

# 收集图片宽度（EMU）
widths = []
try:
    for s in doc.inline_shapes:
        widths.append(s.width)
except Exception as e:
    print('遍历 inline_shapes 失败:', e)
print('图片宽度(EMU):', widths)

# 统计可能的图片段落（居中且无可见文本）
from docx.enum.text import WD_ALIGN_PARAGRAPH
img_para_guess = 0
for p in doc.paragraphs:
    txt = p.text.strip()
    if p.alignment == WD_ALIGN_PARAGRAPH.CENTER and txt == '':
        img_para_guess += 1
print('疑似图片段落数(居中且空文本):', img_para_guess)

# 列出最近的 temp-images
imgs = []
img_root = r'temp-images'
if os.path.isdir(img_root):
    for n in os.listdir(img_root):
        if n.lower().endswith(('.png','.jpg','.jpeg','.gif','.bmp','.webp')):
            p = os.path.join(img_root, n)
            imgs.append((p, os.path.getmtime(p)))
imgs.sort(key=lambda x: x[1], reverse=True)
print('最近图片:')
for p, m in imgs[:8]:
    print('  ', time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(m)), p)

# 目标宽度 5 inch = 4572000 EMU
TARGET = 4572000
if widths:
    # 允许5%偏差
    ok = all(abs(w - TARGET) <= TARGET*0.05 for w in widths if w)
    print('宽度是否约等于 5 英寸:', ok)
