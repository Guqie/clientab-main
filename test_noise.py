import csv, re

with open('temp-data/战新与未来产业月报_第四期_爬取结果_v2.csv', encoding='utf-8-sig', errors='replace') as f:
    reader = csv.DictReader(f)
    rows = list(reader)

content = rows[0].get('content', '')
lines = content.splitlines()

# Run through the actual is_nav_noise logic from preprocess_content.py
def is_nav_noise(line):
    line_stripped = line.strip()
    if not line_stripped or len(line_stripped) <= 2:
        return True
    if re.fullmatch(r"[\W_]+", line_stripped):
        return True
    if re.match(r"^[a-zA-Z][a-zA-Z0-9\s,._#\[\]='-]*$", line_stripped):
        return True
    if re.match(r"^\[.+\]$", line_stripped):
        return True
    if re.search(r"订阅取消订阅|已收藏收藏|点击播报本文|大字号", line_stripped):
        return True
    # The key pattern - test on line 1
    pat = r"^\d{4}年\d{1,2}月\d{1,2}日\d{1,2}[:：]\d{2}\s+来源[:：]"
    m = re.match(pat, line_stripped)
    if m:
        return True
    if re.match(r"^来源[:：][^，\n]{2,30}$", line_stripped):
        return True
    if re.match(r"^来源[:：][\u4e00-\u9fa5]{2,10}$", line_stripped):
        return True
    if re.search(r"^\s*(央视网|人民日报|新华网|新浪财经|经济观察报|科技日报|中国经济时报)", line_stripped):
        if re.search(r"\|\s*\d{4}年", line_stripped):
            return True
    if re.match(r"^[\u4e00-\u9fa5]{2,8}\s*\|\s*\d{4}年\d{1,2}月\d{1,2}日", line_stripped):
        return True
    if re.match(r"^来源[:：][^，\n]{0,20}$", line_stripped):
        return True
    latin = sum(1 for c in line_stripped if "a" <= c <= "z" or "A" <= c <= "Z")
    chinese = sum(1 for c in line_stripped if "\u4e00" <= c <= "\u9fff")
    if chinese > 0 and latin / (chinese + latin) > 0.35:
        return True
    if len(line_stripped) < 12:
        keywords = ["印发", "发布", "出台", "规划", "方案", "意见", "通知", "标准",
                    "投资", "签约", "开工", "投产", "并网", "量产", "亿元", "兆瓦",
                    "突破", "首个", "首次", "增长", "提升", "签署", "落地", "启动"]
        if not any(k in line_stripped for k in keywords):
            return True
    return False

kept = 0
dropped = 0
for i, line in enumerate(lines):
    if is_nav_noise(line):
        dropped += 1
        if dropped <= 10:
            print(f"DROP {i}: {repr(line[:60])}")
    else:
        kept += 1
        if kept <= 5:
            print(f"KEEP {i}: {repr(line[:60])}")

print(f"\nKept: {kept}, Dropped: {dropped}")
