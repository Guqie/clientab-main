import os, time
from glob import glob
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 选择最新的非临时文档
paths = [p for p in glob(r'temp-data/*_new_energy.docx') if not os.path.basename(p).startswith('~$')]
if not paths:
    raise SystemExit('未找到 new_energy 文档（排除临时文件后）')
paths.sort(key=os.path.getmtime, reverse=True)
path = paths[0]
print('最新文档:', path)

# 打开文档
doc = Document(path)

# 统计 inline 图片数
total_inline = len(getattr(doc, 'inline_shapes', []))
print('Inline 图片数量:', total_inline)

# 读取每个 inline 图片宽度（EMU）
widths = []
for s in doc.inline_shapes:
    widths.append(int(getattr(s, 'width', 0)))
print('图片宽度(EMU):', widths)

# 判断宽度是否约等于 5 英寸 (4572000 EMU)
TARGET = 4572000
if widths:
    ok = all(abs(w - TARGET) <= TARGET*0.05 for w in widths if w)
    print('宽度是否约等于5英寸(5%):', ok)

# 统计居中且空文本的段落（作为图片占位的近似指标）
img_para_guess = 0
for p in doc.paragraphs:
    if p.alignment == WD_ALIGN_PARAGRAPH.CENTER and p.text.strip() == '':
        img_para_guess += 1
print('疑似图片段落数(居中且空文本):', img_para_guess)

# 列出最近下载的图片
img_root = 'temp-images'
recent_imgs = []
if os.path.isdir(img_root):
    for n in os.listdir(img_root):
        if n.lower().endswith(('.png','.jpg','.jpeg','.gif','.bmp','.webp')):
            p = os.path.join(img_root, n)
            recent_imgs.append((p, os.path.getmtime(p)))
recent_imgs.sort(key=lambda x: x[1], reverse=True)
print('最近图片文件数:', len(recent_imgs))
for p, m in recent_imgs[:10]:
    print('  ', time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(m)), p)
