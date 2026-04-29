import csv, re

with open('temp-data/战新与未来产业月报_第四期_爬取结果_v2.csv', encoding='utf-8-sig', errors='replace') as f:
    reader = csv.DictReader(f)
    rows = list(reader)

row = rows[2]  # row 4: 河北零碳工厂
content = row.get('content', '')
lines = [l.strip() for l in content.splitlines() if l.strip()]
print(f"First 15 lines of row 4:")
for i, line in enumerate(lines[:15]):
    print(f"  {i}: {line[:80]}")

# Check patterns
text = '\n'.join(lines[:15])
patterns = [
    '部联.*?号',
    '事项通知如下',
    '现印发',
    '附件',
    '各省',
    '第.*?条',
]
print(f"\nPattern matches in first 15 lines:")
for p in patterns:
    m = re.search(p, text)
    if m:
        print(f"  '{p}': {repr(m.group())}")
    else:
        print(f"  '{p}': NO MATCH")
