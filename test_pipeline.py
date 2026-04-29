import csv, re

with open('temp-data/战新与未来产业月报_第四期_爬取结果_v2.csv', encoding='utf-8-sig', errors='replace') as f:
    reader = csv.DictReader(f)
    rows = list(reader)

content = rows[0].get('content', '')
lines = [l.strip() for l in content.splitlines() if l.strip()]

# Simulate full pipeline
kept = []
for line in lines:
    if not line or len(line) <= 2:
        continue
    if re.fullmatch(r"[\W_]+", line):
        continue
    if re.match(r"^[a-zA-Z][a-zA-Z0-9\s,._#\[\]='-]*$", line):
        continue
    if re.match(r"^\[.+\]$", line):
        continue
    if re.search(r"订阅取消订阅|已收藏收藏|点击播报本文|大字号", line):
        continue
    if re.match(r"^\d{4}年\d{1,2}月\d{1,2}日\d{1,2}[:：]\d{2}\s+来源[:：]", line):
        continue
    if re.match(r"^来源[:：][^，\n]{2,30}$", line):
        continue
    if re.match(r"^来源[:：][\u4e00-\u9fa5]{2,10}$", line):
        continue
    if re.search(r"^\s*(央视网|人民日报|新华网|新浪财经|经济观察报|科技日报|中国经济时报)", line):
        if re.search(r"\|\s*\d{4}年", line):
            continue
    if re.match(r"^[\u4e00-\u9fa5]{2,8}\s*\|\s*\d{4}年\d{1,2}月\d{1,2}日", line):
        continue
    if re.match(r"^来源[:：][^，\n]{0,20}$", line):
        continue
    latin = sum(1 for c in line if "a" <= c <= "z" or "A" <= c <= "Z")
    chinese = sum(1 for c in line if "\u4e00" <= c <= "\u9fff")
    if chinese > 0 and latin / (chinese + latin) > 0.35:
        continue
    if len(line) < 12:
        keywords = ["印发", "发布", "出台", "规划", "方案", "意见", "通知", "标准",
                    "投资", "签约", "开工", "投产", "并网", "量产", "亿元", "兆瓦",
                    "突破", "首个", "首次", "增长", "提升", "签署", "落地", "启动"]
        if not any(k in line for k in keywords):
            continue
    # TOC
    if re.match(r"^目\s*录$", line):
        continue
    if re.match(r"^[第][\u4e00-\u9fa5一二三四五六七八九十百\d]+[编章节节分]", line):
        continue
    # Footer
    if re.search(r"附件[:：]|相关阅读|责任编辑|未经授权|京ICP备", line):
        continue
    # Distribution list
    provinces = ["北京", "上海", "天津", "重庆", "河北", "山西", "辽宁", "吉林", "黑龙江",
                 "江苏", "浙江", "安徽", "福建", "江西", "山东", "河南", "湖北", "湖南",
                 "广东", "海南", "四川", "贵州", "云南", "陕西", "甘肃", "青海", "台湾"]
    for p in provinces:
        if re.match(f"^{p}省", line):
            if len(line) < 100:
                break
    else:
        kept.append(line)

# Compress
MAX_CHARS = 2500
result = []
total = 0
for p in kept:
    if total + len(p) + 1 > MAX_CHARS:
        remaining = MAX_CHARS - total
        if remaining > 150:
            cut = p[:remaining]
            for punct in ["。", "！", "？", "；"]:
                idx = cut.rfind(punct)
                if idx > remaining * 0.5:
                    cut = cut[:idx + 1]
                    break
            else:
                cut = cut.rsplit("，", 1)[0] + "。"
            result.append(cut)
        break
    result.append(p)
    total += len(p) + 1

final = "\n".join(result)
print(f"Original: {len(content)} chars, {len(lines)} lines")
print(f"Kept: {len(kept)} lines")
print(f"Final: {len(final)} chars")
print()
print("=== FINAL CONTENT ===")
print(final)
