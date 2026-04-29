#!/usr/bin/env python3
"""Preprocess crawled fulltext content for the strategic emerging industry monthly report.

Cleaning pipeline:
  1. Remove page chrome / metadata (nav bars, subscription prompts)
  2. Detect article type (gov_notice vs news)
  3. Strip table-of-contents blocks
  4. Remove footer/attachment sections
  5. Deduplicate and remove boilerplate
  6. Compress to MAX_CHARS
"""

import argparse
import csv
import re
import sys
from pathlib import Path
from typing import Optional

MAX_CHARS = 2500
MAX_CHARS_GOV = 4000


# ── Utility ──────────────────────────────────────────────────────────────────────

def normalize(text: str) -> str:
    text = text.replace("\r", "\n")
    text = re.sub(r"[\t\u3000 ]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return "\n".join(l.strip() for l in text.splitlines() if l.strip())


def split_paragraphs(text: str) -> list:
    """Split into lines (each line is a paragraph in crawled content)."""
    return [line.strip() for line in text.splitlines() if line.strip()]


# ── Noise classifiers ────────────────────────────────────────────────────────────

def is_nav_noise(line: str) -> bool:
    """High-confidence navigation / metadata noise lines."""
    line = line.strip()
    if not line or len(line) <= 2:
        return True

    # CSS / structural noise
    if re.fullmatch(r"[\W_]+", line):
        return True
    if re.match(r"^[a-zA-Z][a-zA-Z0-9\s,._#\[\]='-]*$", line):
        return True
    if re.match(r"^\[.+\]$", line):
        return True

    # Subscription / nav chrome / page chrome
    if re.search(r"订阅取消订阅|已收藏收藏|点击播报本文|大字号", line):
        return True

    # Short meta lines: timestamp source pattern
    # e.g. "2026年03月16日09:03 来源：人民网－人民日报"
    if re.match(r"^\d{4}年\d{1,2}月\d{1,2}日\d{1,2}[:：]\d{2}\s+来源[:：]", line):
        return True
    # e.g. "来源：人民网－人民日报" (short source-only line)
    if re.match(r"^来源[:：][^，\n]{2,30}$", line):
        return True
    # e.g. "来源：人民日报" alone
    if re.match(r"^来源[:：][\u4e00-\u9fa5]{2,10}$", line):
        return True
    # e.g. "央视网 | 2026年03月20日 09:30:26"
    if re.search(r"^\s*(央视网|人民日报|新华网|新浪财经|经济观察报|科技日报|中国经济时报)", line):
        if re.search(r"\|\s*\d{4}年", line):
            return True
    # Short lines with only source + date
    if re.match(r"^[\u4e00-\u9fa5]{2,8}\s*\|\s*\d{4}年\d{1,2}月\d{1,2}日", line):
        return True
    # "来源：" at start of line followed by short text
    if re.match(r"^来源[:：][^，\n]{0,20}$", line):
        return True

    # High density Latin mixed with Chinese
    latin = sum(1 for c in line if "a" <= c <= "z" or "A" <= c <= "Z")
    chinese = sum(1 for c in line if "\u4e00" <= c <= "\u9fff")
    if chinese > 0 and latin / (chinese + latin) > 0.35:
        return True

    # Short lines with no meaningful keywords
    if len(line) < 12:
        keywords = ["印发", "发布", "出台", "规划", "方案", "意见", "通知", "标准",
                    "投资", "签约", "开工", "投产", "并网", "量产", "亿元", "兆瓦",
                    "突破", "首个", "首次", "增长", "提升", "签署", "落地", "启动"]
        if not any(k in line for k in keywords):
            return True

    return False


def is_toc_line(line: str) -> bool:
    """Table of contents entry lines."""
    line = line.strip()
    # Pattern: 第X编 / 第X章 / 第X节 with short content
    if re.match(r"^[第][一二三四五六七八九十百\d]+[编章节节]", line):
        return True
    if re.match(r"^目\s*录$", line):
        return True
    if re.match(r"^[上中下]?[篇部章节]", line):
        return True
    return False


def is_footer(line: str) -> bool:
    """Footer / attachment section markers."""
    line = line.strip()
    for pat in [
        r"附件[:：]",
        r"相关阅读",
        r"责任编辑",
        r"编辑[:：]",
        r"校对[:：]",
        r"未经授权",
        r"转载须注明",
        r"京ICP备",
        r"举报电话",
        r"违法和不良信息举报",
        r"一键分享",
        r"分享到",
    ]:
        if re.search(pat, line):
            return True
    return False


def is_distribution_list(line: str) -> bool:
    """Government distribution lists (provinces, cities, departments)."""
    line = line.strip()
    if len(line) > 120:
        return False
    # Province-level
    provinces = [
        "北京", "上海", "天津", "重庆", "河北", "山西", "辽宁", "吉林", "黑龙江",
        "江苏", "浙江", "安徽", "福建", "江西", "山东", "河南", "湖北", "湖南",
        "广东", "海南", "四川", "贵州", "云南", "陕西", "甘肃", "青海", "台湾",
    ]
    for p in provinces:
        if re.match(f"^{p}省", line):
            return True
    # Autonomous regions
    for p in ["内蒙古", "广西", "西藏", "宁夏", "新疆"]:
        if re.match(f"^{p}自", line):
            return True
    # Special admin regions
    for p in ["香港", "澳门"]:
        if re.match(f"^{p}", line):
            return True
    # Direct-admin cities
    cities = [
        "深圳市", "广州市", "成都市", "武汉市", "西安市", "杭州市",
        "南京市", "长沙市", "郑州市", "济南市", "青岛市", "大连市",
        "沈阳市", "哈尔滨市", "长春市", "厦门市", "宁波市", "大连市",
        "青岛市", "乌鲁木齐市",
    ]
    for c in cities:
        if re.match(f"^{c}", line):
            return True
    if re.match(r"^雄安新区", line):
        return True
    if re.match(r"^定州|辛集", line):
        return True
    return False


# ── Inline noise ────────────────────────────────────────────────────────────────

def strip_inline_noise(text: str) -> str:
    """Remove noise tokens within a paragraph."""
    patterns = [
        (r"（?图源[:：].*?）?", ""),
        (r"（?图片来源[:：].*?）?", ""),
        (r"\(责编[:：].*?\)", ""),
        (r"（责编[:：].*?）", ""),
        (r"\(责任编辑[:：].*?\)", ""),
        (r"（责任编辑[:：].*?）", ""),
        (r"[（(]\s*\d{4,6}\s*[.．]\s*[A-Za-z]{1,5}\s*[）)]", ""),
        (r"（?点击.*?本文.*?）?", ""),
        (r"（?约\d+字.*?）?", ""),
        (r"https?://\S+", ""),
        # Expert patterns - remove entire matched portion
        (r"（?[^。！？\n]{0,40}(专家|业内人士|分析师|研究员|负责人|企业负责人|企业高管)(认为|表示|介绍|指出|称|分析|透露)[^。！？\n]{0,200}([。！？]|$)", r"\3"),
        (r"（?[^。！？\n]{0,30}(记者|编辑)(报道|获悉|了解到)[^。！？\n]{0,150}([。！？]|$)", r"\3"),
    ]
    for pat, repl in patterns:
        text = re.sub(pat, repl, text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


# ── Article type detection ──────────────────────────────────────────────────────

def detect_article_type(text: str) -> str:
    """Classify as gov_notice or news based on structural signals."""
    article_count = len(re.findall(r"第[一二三四五六七八九十百\d]+条", text))

    gov_signals = [
        r"部\s*联\s*\w+\[\d{4}\]\d+号",
        r"现将.*?如下[：:]",
        r"事项通知如下",
        r"现印发给你们",
        r"现印发给你们",
        r"附件[:：]",
        r"各省[、:：]",                    # "各省工业和信息化厅："
        r"第[一二三四五六七八九十百\d]+条",  # handled via density bonus below
        r"^\d+\.\s+[\u4e00-\u9fa5]{2,6}[厅局部]",  # "1. 各省工业和信息化厅"
    ]

    score = sum(1 for p in gov_signals if re.search(p, text))

    if article_count >= 3:
        score += 2
    elif article_count >= 1:
        score += 1

    return "gov_notice" if score >= 2 else "news"


# ── Core cleaning ───────────────────────────────────────────────────────────────

def clean_gov_notice(paragraphs: list) -> list:
    """Government notice / policy document pipeline."""
    result = []
    seen = set()
    in_toc = False
    toc_ended = False
    footer_started = False
    attachment_mode = False
    skipped_toc = False

    for p in paragraphs:
        p = p.strip()
        if not p:
            continue

        # Footer trigger
        if is_footer(p):
            footer_started = True
            continue
        if footer_started:
            continue

        # Distribution list skip
        if is_distribution_list(p):
            continue

        # TOC detection
        if re.match(r"^目\s*录$", p):
            in_toc = True
            skipped_toc = True
            continue
        if in_toc:
            if is_toc_line(p):
                continue
            else:
                # End of TOC section
                in_toc = False
                toc_ended = True

        # Skip policy document reference numbers (e.g. "卫办医政发〔2024〕89号")
        if re.match(r"^[各部部][办厅局司]?[医政科教工信发改]?[发函]?\[\d{4}\]\d+号", p):
            continue
        if re.match(r"^[上下][一ニ两三四五六七八九十百\d]+条", p):
            if len(p) < 200:
                continue

        # Skip nav noise lines
        if is_nav_noise(p):
            continue

        # Inline clean
        p = strip_inline_noise(p)
        if not p or len(p) < 5:
            continue

        # Deduplicate
        norm = re.sub(r"\s+", "", p)
        if norm in seen:
            continue
        seen.add(norm)

        # Skip attachment section header
        if re.match(r"^附件[一二三四\d]", p) and len(p) < 50:
            attachment_mode = True
            continue
        if attachment_mode and re.match(r"^\d+[.．、]", p):
            continue

        result.append(p)

    return result, skipped_toc


def clean_news(paragraphs: list) -> list:
    """News article pipeline."""
    result = []
    seen = set()
    footer_started = False

    for p in paragraphs:
        p = p.strip()
        if not p:
            continue

        if is_nav_noise(p):
            continue

        if is_footer(p):
            footer_started = True
            continue
        if footer_started:
            continue

        if is_distribution_list(p):
            continue

        p = strip_inline_noise(p)
        if not p or len(p) < 5:
            continue

        # Skip high-Latin-density lines
        latin = sum(1 for c in p if "a" <= c <= "z" or "A" <= c <= "Z")
        chinese = sum(1 for c in p if "\u4e00" <= c <= "\u9fff")
        if chinese > 0 and latin / (chinese + latin) > 0.4:
            continue

        norm = re.sub(r"\s+", "", p)
        if norm in seen:
            continue
        seen.add(norm)

        result.append(p)

    return result, False


def compress(paragraphs: list, max_chars: int) -> str:
    """Join paragraphs up to max_chars, cutting at sentence boundary."""
    text_parts = []
    total = 0
    for p in paragraphs:
        if total + len(p) + 1 > max_chars:
            remaining = max_chars - total
            if remaining > 150:
                cut = p[:remaining]
                # Find last sentence end
                for punct in ["。", "！", "？", "；"]:
                    idx = cut.rfind(punct)
                    if idx > remaining * 0.5:
                        cut = cut[:idx + 1]
                        break
                else:
                    cut = cut.rsplit("，", 1)[0] + "。"
                text_parts.append(cut)
            break
        text_parts.append(p)
        total += len(p) + 1
    return "\n".join(text_parts)


def preprocess(content: str) -> tuple:
    """Main pipeline. Returns (cleaned_text, article_type, skipped_toc)."""
    if not content or len(content) < 30:
        return content, "none", False

    paragraphs = split_paragraphs(content)
    if not paragraphs:
        return content[:MAX_CHARS], "unknown", False

    article_type = detect_article_type("\n".join(paragraphs[:15]))

    if article_type == "gov_notice":
        cleaned, skipped_toc = clean_gov_notice(paragraphs)
        max_c = MAX_CHARS_GOV
    else:
        cleaned, skipped_toc = clean_news(paragraphs)
        max_c = MAX_CHARS

    if not cleaned:
        return content[:MAX_CHARS], article_type, False

    return compress(cleaned, max_c), article_type, skipped_toc


# ── CLI ────────────────────────────────────────────────────────────────────────

def detect_encoding(path: Path) -> str:
    for enc in ["utf-8", "utf-8-sig", "gbk", "gb2312"]:
        try:
            with open(path, "r", encoding=enc) as f:
                f.read(1024)
            return enc
        except Exception:
            continue
    return "utf-8"


def main() -> int:
    parser = argparse.ArgumentParser(description="Preprocess crawled fulltext content")
    parser.add_argument("-i", "--input", required=True)
    parser.add_argument("-o", "--output", required=True)
    parser.add_argument("-m", "--max-chars", type=int, default=MAX_CHARS)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        raise SystemExit(f"Input not found: {args.input}")

    enc = detect_encoding(input_path)
    print(f"Input: {args.input} (encoding: {enc})")

    with open(input_path, "r", encoding=enc, newline="") as f:
        reader = csv.DictReader(f)
        fields = list(reader.fieldnames)
        rows = list(reader)

    for col in ["preprocess_type", "preprocess_toc_skipped", "preprocess_chars_in", "preprocess_chars_out"]:
        if col not in fields:
            fields.append(col)

    samples_printed = 0
    total_in = 0
    total_out = 0
    by_type = {}
    skipped = 0

    for row in rows:
        status = row.get("fulltext_status", "")
        content = row.get("content", "") or ""

        if status != "success" or not content or len(content) < 30:
            row["preprocess_type"] = "skipped"
            row["preprocess_toc_skipped"] = "N"
            row["preprocess_chars_in"] = len(content)
            row["preprocess_chars_out"] = len(content)
            skipped += 1
            continue

        total_in += len(content)
        cleaned, atype, toc_skipped = preprocess(content)

        row["content"] = cleaned
        row["preprocess_type"] = atype
        row["preprocess_toc_skipped"] = "Y" if toc_skipped else "N"
        row["preprocess_chars_in"] = len(content)
        row["preprocess_chars_out"] = len(cleaned)

        total_out += len(cleaned)
        by_type[atype] = by_type.get(atype, 0) + 1

        if samples_printed < 3:
            samples_printed += 1
            print(f"\n{'='*60}")
            print(f"Sample {samples_printed}: {row.get('title', '')[:50]}")
            print(f"Type: {atype} | TOC removed: {toc_skipped} | "
                  f"{len(content)} → {len(cleaned)} chars")
            print("-" * 40)
            print(cleaned[:500])
            print("...")

    if args.dry_run:
        print("\n=== DRY RUN - no file written ===")
        print(f"Total: {len(rows)} | Processed: {len(rows)-skipped} | Skipped: {skipped}")
        print(f"By type: {by_type}")
        return 0

    with open(args.output, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)

    print(f"\n=== Results ===")
    print(f"Total: {len(rows)} | Processed: {len(rows)-skipped} | Skipped: {skipped}")
    print(f"By type: {by_type}")
    avg_in = total_in // max(len(rows) - skipped, 1)
    avg_out = total_out // max(len(rows) - skipped, 1)
    print(f"Avg chars: {avg_in} → {avg_out} (compression: {avg_out*100//max(avg_in,1)}%)")
    print(f"Output: {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
