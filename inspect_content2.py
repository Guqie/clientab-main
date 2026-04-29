import pandas as pd

df = pd.read_csv(r'd:\桌面\clientab-main\temp-data\战新与未来产业月报_第四期_爬取结果.csv', dtype=str)
print('总行数:', len(df))

# 按source分组看样本
from collections import Counter
sources = df['source'].fillna('(空)').tolist()
for src, cnt in Counter(sources).most_common(20):
    print(f'  {src}: {cnt}条')

print()

# 每个source看一条content开头
seen = set()
for i in range(len(df)):
    row = df.iloc[i]
    src = row.get('source', '')
    if src in seen:
        continue
    seen.add(src)
    c = str(row['content']) if pd.notna(row['content']) else ''
    print(f'=== source={src} ===')
    print(c[:600])
    print('...(省略)...' if len(c) > 600 else '')
    print()
    if len(seen) >= 15:
        break
