import pandas as pd

df = pd.read_csv(r'd:\桌面\clientab-main\temp-data\战新与未来产业月报_第四期_爬取结果.csv', dtype=str)
print('总行数:', len(df))
print('列名:', list(df.columns))
print()

for i in [0, 5, 10, 50, 100, 500, 1000, 2000]:
    if i < len(df):
        row = df.iloc[i]
        c = str(row['content']) if pd.notna(row['content']) else ''
        src = row.get('source', '?')
        url = str(row.get('url', ''))[:70]
        print(f'=== [{i}] source={src} url={url} ===')
        print(c[:400])
        if len(c) > 400:
            print('...(省略)...')
        print()
