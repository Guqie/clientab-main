import csv

with open('temp-data/战新与未来产业月报_第四期_爬取结果_v2.csv', encoding='utf-8-sig', errors='replace') as f:
    reader = csv.DictReader(f)
    rows = list(reader)

with open('temp-data/check_sample.txt', 'w', encoding='utf-8') as out:
    for i, row in enumerate(rows[:5]):
        title = row.get('title', '') or ''
        c = row.get('content', '') or ''
        out.write(f'=== Row {i+2}: {title}\n')
        out.write(f'Status: {row.get("fulltext_status", "")}\n')
        out.write(f'Content ({len(c)} chars):\n')
        out.write(c[:800])
        out.write('\n\n')

print(f'Wrote {len(rows)} rows sample to check_sample.txt')
