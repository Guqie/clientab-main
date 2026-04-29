import csv, re

with open('temp-data/战新与未来产业月报_第四期_爬取结果_v2.csv', encoding='utf-8-sig', errors='replace') as f:
    reader = csv.DictReader(f)
    rows = list(reader)

# Check rows 3 and 4 (0-indexed: 1, 2) - the ones that should be gov_notice
for i in [1, 2, 3]:
    row = rows[i]
    content = row.get('content', '') or ''
    lines = [l.strip() for l in content.splitlines() if l.strip()]
    first_text = '\n'.join(lines[:20])  # check first 20 paragraphs

    article_count = len(re.findall(r'第[一二三四五六七八九十百\d]+条', first_text))

    gov_signals = [
        ('部联文号', r'部\s*联\s*\w+\[\d{4}\]\d+号'),
        ('现将如下', r'现将.*?如下[：:]'),
        ('事项通知', r'事项通知如下'),
        ('有...事项通知', r'有.*?事项通知如下'),
        ('附件', r'附件[:：]'),
        ('第X条密度', None),  # handled separately
    ]

    matches = {}
    for name, pat in gov_signals:
        if pat:
            matches[name] = bool(re.search(pat, first_text))
        else:
            matches[name] = article_count

    score = sum(1 for v in matches.values() if v)
    if article_count >= 3:
        score += 2
    elif article_count >= 1:
        score += 1

    print(f"Row {i+2}: {row.get('title', '')[:40]}")
    print(f"  First text length: {len(first_text)}")
    print(f"  Signals: {matches}")
    print(f"  article_count: {article_count}")
    print(f"  score: {score} -> {'gov_notice' if score >= 2 else 'news'}")
    print()
