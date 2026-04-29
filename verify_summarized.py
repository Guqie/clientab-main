import pandas as pd

df = pd.read_excel(r'd:\桌面\clientab-main\temp-data\summaries_deepseek_final_summarized.xlsx', sheet_name=0, dtype=str)
print('总行数:', len(df))
print('summary 非空:', df['summary'].notna().sum())
print('summary 空:', df['summary'].isna().sum())

# 字数统计
lengths = df['summary'].apply(lambda x: len(str(x)) if pd.notna(x) else 0)
print('平均摘要字数:', lengths.mean())
print('最短:', lengths.min(), ' 最长:', lengths.max())

print()
print('=== 摘要示例 (前3条) ===')
for i in range(3):
    row = df.iloc[i]
    print(f'[{i}] title: {row["title"][:40]}')
    print(f'    summary: {str(row["summary"])[:120]}')
    print()
