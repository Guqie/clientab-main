import os, sys
sys.path.insert(0, r'D:\桌面\clientab-main')
from universal_csv_to_word import csv_to_word_universal, TemplateFactory

root = r'D:\桌面\clientab-main\temp-data'
print('cwd:', os.getcwd())
print('dir exists:', os.path.isdir(root))

names = os.listdir(root)
print('files count:', len(names))
for name in names:
    print('name:', repr(name))

# 精确查找 CSV 名称并打印码点
for name in names:
    if name.endswith('.csv'):
        print('CSV file:', repr(name), 'codepoints:', [hex(ord(c)) for c in name])

# 目标路径
p = r'D:\桌面\clientab-main\temp-data\新能源9月15日刊.csv'
print('TARGET exists:', os.path.exists(p))
print('TARGET basename repr:', repr(os.path.basename(p)))
print('TARGET codepoints:', [hex(ord(c)) for c in os.path.basename(p)])

# 若存在则执行生成
if os.path.exists(p):
    print('可用模板列表:', TemplateFactory().get_available_templates())
    out = csv_to_word_universal(p, 'new_energy', r'D:\桌面\clientab-main\templates_config.yaml')
    print('生成完成:', out)
else:
    # 查找最相近的文件名（忽略全角/半角差异和空格）
    target = '新能源9月15日刊.csv'
    def normalize(s):
        return ''.join(c for c in s if not c.isspace())
    cand = [n for n in names if normalize(n) == normalize(target)]
    print('候选匹配:', cand)
