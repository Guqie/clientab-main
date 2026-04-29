import csv, re

with open('temp-data/战新与未来产业月报_第四期_爬取结果_v2.csv', encoding='utf-8-sig', errors='replace') as f:
    reader = csv.DictReader(f)
    rows = list(reader)

content = rows[0].get('content', '')
lines = content.splitlines()

with open('temp-data/nav_debug.txt', 'w', encoding='utf-8') as out:
    out.write(f"Total lines: {len(lines)}\n\n")
    for i, line in enumerate(lines):
        is_nav = False
        line_stripped = line.strip()

        # Pattern checks
        if re.match(r'^\d{4}年\d{1,2}月\d{1,2}日\d{1,2}', line_stripped):
            if '来源' in line_stripped:
                is_nav = True
                reason = 'timestamp+source'
        if '订阅取消订阅' in line_stripped:
            is_nav = True
            reason = 'subscribe'
        if '已收藏收藏' in line_stripped:
            is_nav = True
            reason = 'fav'
        if re.match(r'^来源[:：][\u4e00-\u9fa5]{2,10}$', line_stripped):
            is_nav = True
            reason = 'source_only'
        if re.match(r'^\d{4}年\d{1,2}月\d{1,2}日$', line_stripped):
            is_nav = True
            reason = 'date_only'
        if re.match(r'^目\s*录$', line_stripped):
            is_nav = True
            reason = 'toc_header'
        if re.match(r'^[第][\u4e00-\u9fa5一二三四五六七八九十百\d]+[编章节节]', line_stripped):
            is_nav = True
            reason = 'toc_entry'

        if is_nav:
            out.write(f'LINE {i}: [NAV-{reason}] {line_stripped[:80]}\n')
        else:
            out.write(f'LINE {i}: [KEEP        ] {line_stripped[:80]}\n')

print('Written to nav_debug.txt')
