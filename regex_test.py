import re

# Test lines from the actual content
test_lines = [
    "2026年03月16日09:03 来源：人民网－人民日报222",
    "订阅取消订阅已收藏收藏大字号",
    "点击播报本文，约",
    "中华人民共和国主席令",
    "第七十号",
    "2026年3月12日",
    "目 录",
    "第一编 总 则",
    "第一章 基本规定",
]

for line in test_lines:
    line = line.strip()
    # Pattern from preprocess_content.py
    p1 = re.match(r"^\d{4}年\d{1,2}月\d{1,2}日\d{1,2}[:：]\d{2}\s+来源[:：]", line)
    # Simpler pattern
    p2 = re.search(r"\d{4}年\d{1,2}月\d{1,2}日\d{1,2}.*来源", line)
    print(f"'{line[:40]}' -> p1={bool(p1)}, p2={bool(p2)}")
    if p2:
        print(f"  matched: {repr(p2.group())}")
