#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
爬虫知识库管理 CLI

用法：
    python scripts/crawler_intel_cli.py status
    python scripts/crawler_intel_cli.py selectors <site_key>
    python scripts/crawler_intel_cli.py analyze --site <site_key> --url <url>
    python scripts/crawler_intel_cli.py export --output intel_backup.json
    python scripts/crawler_intel_cli.py import --input intel_backup.json
    python scripts/crawler_intel_cli.py prune --min-rate <rate>
    python scripts/crawler_intel_cli.py samples <site_key> --success | --fail
"""

import argparse
import asyncio
import json
import sys
from datetime import datetime
from pathlib import Path

DEFAULT_DB = "temp-data/crawler_intel.db"


def _status_table(headers: list[str], rows: list[list]) -> None:
    """打印 ASCII 表格"""
    if not rows:
        print("  (空)")
        return
    col_widths = [len(h) for h in headers]
    for row in rows:
        for i, cell in enumerate(row):
            col_widths[i] = max(col_widths[i], len(str(cell)))
    sep = "+" + "+".join("-" * (w + 2) for w in col_widths) + "+"

    def fmt(row):
        return "|" + "|".join(
            f" {str(c).ljust(col_widths[i])} " for i, c in enumerate(row)
        ) + "|"

    print(sep)
    print(fmt(headers))
    print(sep.replace("-", "="))
    for row in rows:
        print(fmt(row))
    print(sep)


def cmd_status(intel) -> int:
    """展示所有站点的健康状态"""
    sites = intel.get_all_site_status()
    if not sites:
        print("知识库为空，尚未记录任何站点的抓取数据。")
        print("使用 --intel 参数运行爬虫以开始积累数据。")
        return 0

    headers = ["站点", "尝试", "成功", "成功率", "质量均分", "LLM分析", "健康状态"]
    rows = []
    for s in sites:
        rate = s["current_success_rate"] * 100
        health = s["health_status"]
        health_display = {
            "healthy": "OK",
            "degraded": "WARN",
            "critical": "FAIL",
            "unknown": "N/A",
        }.get(health, health)
        rows.append([
            s["site_key"],
            s["total_attempts"],
            s["success_count"],
            f"{rate:.1f}%",
            f"{s['avg_quality']:.2f}",
            s["llm_analysis_count"],
            health_display,
        ])
    _status_table(headers, rows)
    return 0


def cmd_selectors(intel, site_key: str) -> int:
    """展示某站点的选择器详情"""
    selectors = intel.get_site_selectors_detail(site_key)
    if not selectors:
        print(f"未找到站点 [{site_key}] 的选择器记录。")
        return 1

    headers = ["选择器", "来源", "成功", "失败", "质量", "活跃", "优先级", "更新时间"]
    rows = []
    for s in selectors:
        rows.append([
            s["selector"],
            s["source"],
            s["success_count"],
            s["fail_count"],
            f"{s['avg_quality']:.2f}",
            "Y" if s["is_active"] else "N",
            s["priority"],
            s["updated_at"][:16] if s["updated_at"] else "-",
        ])
    _status_table(headers, rows)
    return 0


def cmd_samples(intel, site_key: str, success_only: bool) -> int:
    """展示最近的质量日志样本"""
    rows = intel._db.execute("""
        SELECT url, char_count, quality_score, success, crawled_at
        FROM extraction_quality_log
        WHERE site_key = ?
        ORDER BY crawled_at DESC LIMIT 20
    """, (site_key,)).fetchall()
    if not rows:
        print(f"未找到 [{site_key}] 的质量日志。")
        return 1

    headers = ["URL", "字数", "质量", "成功", "时间"]
    table_rows = []
    for r in rows:
        url_short = (r["url"][:60] + "...") if len(r["url"]) > 63 else r["url"]
        table_rows.append([
            url_short,
            r["char_count"],
            f"{r['quality_score']:.2f}",
            "Y" if r["success"] else "N",
            r["crawled_at"][:19],
        ])
    _status_table(headers, table_rows)
    return 0


async def cmd_analyze(intel, site_key: str, url: str) -> int:
    """手动触发 LLM 分析某站点"""
    print(f"[intel] 开始分析站点 [{site_key}]，URL: {url}")
    result = await intel.analyze_and_update_selectors(url, site_key)
    if result.success:
        print(f"[intel] 分析成功，提出 {len(result.selectors)} 个选择器：")
        for i, s in enumerate(result.selectors, 1):
            print(f"  {i}. {s.selector}  ({s.reason})")
        if result.analysis:
            print(f"  分析: {result.analysis}")
        return 0
    else:
        print(f"[intel] 分析失败: {result.error}")
        return 1


def cmd_export(intel, output: str) -> int:
    """导出知识库"""
    data = intel.export_knowledge()
    path = Path(output)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"[intel] 已导出 {len(data['site_selectors'])} 条选择器，"
          f"{len(data['site_meta'])} 条站点元数据，"
          f"{len(data['llm_analysis_history'])} 条分析历史")
    print(f"[intel] 保存至: {output}")
    return 0


def cmd_import(intel, input_path: str) -> int:
    """导入知识库"""
    with open(input_path, encoding="utf-8") as f:
        data = json.load(f)
    s, m, h = intel.import_knowledge(data)
    print(f"[intel] 已导入: {s} 条选择器，{m} 条站点元数据，{h} 条分析历史")
    return 0


def cmd_prune(intel, min_rate: float) -> int:
    """清理低质量选择器"""
    count = intel.prune_selectors(min_rate)
    print(f"[intel] 已删除 {count} 条成功率低于 {min_rate:.0%} 的选择器")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="爬虫自学习知识库管理工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--db", default=DEFAULT_DB,
        help=f"知识库路径 (default: {DEFAULT_DB})",
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    # status
    sub.add_parser("status", help="查看所有站点健康状态")

    # selectors <site_key>
    sel = sub.add_parser("selectors", help="查看某站点的选择器")
    sel.add_argument("site_key", help="站点关键词，如 people.com.cn")

    # samples <site_key> [--success | --fail]
    sam = sub.add_parser("samples", help="查看最近的质量日志样本")
    sam.add_argument("site_key", help="站点关键词")
    sam.add_argument("--success", action="store_true", help="只看成功样本")
    sam.add_argument("--fail", action="store_true", help="只看失败样本")

    # analyze --site <site_key> --url <url>
    ana = sub.add_parser("analyze", help="手动触发 LLM 分析站点")
    ana.add_argument("--site", required=True, help="站点关键词")
    ana.add_argument("--url", required=True, help="待分析页面的 URL")

    # export --output <path>
    exp = sub.add_parser("export", help="导出知识库")
    exp.add_argument("--output", "-o", required=True, help="输出 JSON 路径")

    # import --input <path>
    imp = sub.add_parser("import", help="导入知识库")
    imp.add_argument("--input", "-i", required=True, help="输入 JSON 路径")

    # prune --min-rate <rate>
    pro = sub.add_parser("prune", help="清理低质量选择器")
    pro.add_argument("--min-rate", type=float, default=0.2,
                     help="删除成功率低于此值的选择器 (default: 0.2)")

    args = parser.parse_args()

    # 初始化知识库
    if args.cmd in ("analyze",) or not Path(args.db).exists():
        try:
            from csv_word_converter.crawler_intel import CrawlerIntel
        except ImportError:
            sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
            from csv_word_converter.crawler_intel import CrawlerIntel
    else:
        from csv_word_converter.crawler_intel import CrawlerIntel

    intel = CrawlerIntel(db_path=args.db)

    try:
        if args.cmd == "status":
            return cmd_status(intel)
        elif args.cmd == "selectors":
            return cmd_selectors(intel, args.site_key)
        elif args.cmd == "samples":
            return cmd_samples(intel, args.site_key, args.success)
        elif args.cmd == "analyze":
            return asyncio.run(cmd_analyze(intel, args.site_key, args.url))
        elif args.cmd == "export":
            return cmd_export(intel, args.output)
        elif args.cmd == "import":
            return cmd_import(intel, args.input)
        elif args.cmd == "prune":
            return cmd_prune(intel, args.min_rate)
        else:
            parser.print_help()
            return 1
    finally:
        intel.close()


if __name__ == "__main__":
    raise SystemExit(main())
