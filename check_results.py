import csv

with open('temp-data/战新与未来产业月报_第四期_预处理结果.csv', encoding='utf-8-sig', errors='replace') as f:
    reader = csv.DictReader(f)
    rows = list(reader)

# Show samples by article type
by_type = {}
for row in rows:
    t = row.get('preprocess_type', 'unknown')
    by_type.setdefault(t, []).append(row)

for t, items in by_type.items():
    print(f"\n{'='*60}")
    print(f"TYPE: {t} ({len(items)} articles)")
    print('='*60)
    for item in items[:2]:
        title = item.get('title', '')[:50]
        c = item.get('content', '') or ''
        print(f"\n--- {title} ---")
        print(f"  {len(c)} chars")
        print(c[:600])
        print("...")
