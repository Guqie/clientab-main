#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自学习爬虫知识库模块

为爬虫提供自学习迭代能力：
1. 持久化选择器知识（SQLite）
2. 评分每次提取结果
3. 阈值检测：成功率低于阈值时触发 LLM 分析
4. LLM 分析页面 → 生成新选择器 → 写入知识库
"""

import asyncio
import json
import logging
import re
import sqlite3
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import yaml

logger = logging.getLogger(__name__)

# ------------------------------------------------------------
# 内置选择器（作为默认种子，写入数据库时不重复插入已有记录）
# ------------------------------------------------------------
DEFAULT_SELECTORS = {
    "people.com.cn": [
        "#rmw .rm_txt", ".article .text", ".article-content",
        ".left_zw", "article", ".fl", "[class*='content']",
    ],
    "xinhuanet.com": [
        "#content", ".article", ".text", "article",
        "[class*='article']", "[class*='text']",
    ],
    "gov.cn": [
        "#UCAP_CONTENT", ".article-content", ".TRS_Editor",
        "#zoom", ".content", "article", "[class*='content']",
    ],
    "miit.gov.cn": [
        ".TRS_Editor", "#zoom", ".article-content",
        "[class*='content']", "article",
    ],
    "cctv.com": [
        "#article_content", ".content", "article",
        "[class*='article']", ".body_content",
    ],
    "chinanews.com.cn": [
        ".content", ".article-content", "article",
        "[class*='content']", ".left",
    ],
    "caixin.com": [
        ".article-content", ".content", "article",
        "[class*='text']", ".main-content",
    ],
    "yicai.com": [
        ".article-content", ".content", "article",
        ".text", "[class*='content']",
    ],
    "jiemian.com": [
        ".article-content", ".content", "article",
        "[class*='content']", ".article-body",
    ],
    "thepaper.cn": [
        ".index_content__gTMZh", "[class*='content']",
        ".article-content", "article", ".main-content",
    ],
    "163.com": [
        ".article-body", ".post_body", ".content",
        "article", "[class*='article']",
    ],
    "sina.com.cn": [
        ".article", ".article-content", "#articleContent",
        "article", "[class*='content']",
    ],
    "sohu.com": [
        "[class*='article-text']", ".article-content",
        "article", "[class*='content']", ".text",
    ],
    "qq.com": [
        ".article-content", "[class*='content']",
        "article", ".text", ".bd",
    ],
    "bjnews.com.cn": [
        ".article-content", ".content", "article",
        "[class*='content']",
    ],
    "cs.com.cn": [
        ".article-content", ".content", "article",
        "[class*='content']",
    ],
    "ccidreport.com": [
        ".TRS_Editor", ".article-content", "article",
        "[class*='content']",
    ],
    "stdaily.com": [
        ".TRS_Editor", ".article-content", "article",
        "[class*='content']",
    ],
    "kraft.com": [
        ".TRS_Editor", ".article-content", "article",
        "[class*='content']",
    ],
    "evciti.com": [
        ".TRS_Editor", ".article-content", "article",
        "[class*='content']",
    ],
}


# ------------------------------------------------------------
# LLM 分析提示词
# ------------------------------------------------------------
LLM_ANALYSIS_SYSTEM_PROMPT = """你是一个专业的网页结构分析师，擅长从 HTML 中找出新闻文章正文的最佳 CSS 选择器。
你的分析必须严谨、准确，选择器必须是真实有效的 CSS 选择器。"""

LLM_ANALYSIS_USER_TEMPLATE = """你是一个专业的网页结构分析师，擅长从 HTML 中找出新闻文章正文的最佳 CSS 选择器。

已知站点域名关键词：{site_key}
已有选择器（供参考，不一定最优）：{existing_selectors}

成功案例（有效提取的内容片段，字数多、内容相关度高）：
{success_examples}

失败案例（提取为空或质量差）：
{fail_examples}

请分析该网页的 HTML 结构，找出 3-5 个最可能命中文章正文的选择器，按置信度从高到低排序。

返回格式（必须是合法 JSON，不要包含其他内容）：
{{
  "selectors": [
    {{"selector": ".article-content", "reason": "这是该站的标准正文容器"}},
    {{"selector": "article", "reason": "HTML5 语义标签"}}
  ],
  "analysis": "简要说明分析过程（20字以内）"
}}"""


# ------------------------------------------------------------
# 数据类
# ------------------------------------------------------------

@dataclass
class ExtractionQuality:
    """单次提取的质量评估结果"""
    score: float           # 0-10 质量分
    signal_strength: float  # 内容密度
    noise_ratio: float      # 噪音行占比
    has_structured_data: bool  # 是否含结构化数据


@dataclass
class SelectorCandidate:
    """一条选择器记录"""
    selector: str
    reason: str = ""
    source: str = "llm"   # default / llm / learned


@dataclass
class AnalysisResult:
    """LLM 分析结果"""
    site_key: str
    selectors: List[SelectorCandidate]
    analysis: str
    success: bool
    error: Optional[str] = None


# ------------------------------------------------------------
# 知识库管理器
# ------------------------------------------------------------

class CrawlerIntel:
    """
    自学习爬虫知识库管理器

    职责：
    1. 持久化选择器知识（读写 SQLite）
    2. 评分每次提取结果
    3. 阈值检测：成功率低于阈值时触发 LLM 分析
    4. LLM 分析页面 → 生成新选择器 → 写入知识库
    """

    QUALITY_THRESHOLD = 0.65   # 成功率低于此值时触发 LLM 分析
    MIN_ANALYSIS_SAMPLES = 5   # 最少积累 5 条日志才开始分析

    _KEEP_KEYWORDS_RE = re.compile(
        r"(印发|发布|出台|规划|方案|意见|通知|标准|指南|目录|"
        r"投资|签约|开工|投产|并网|量产|中试|交付|首飞|下线|"
        r"产能|装机|规模|亿元|万元|千瓦|兆瓦|吉瓦|GWh|MW|GW|kW|"
        r"风电|储能|氢能|逆变器|机器人|工业软件|功率半导体|传感器|"
        r"低空|无人机|海工|船舶|新材料|污水处理|"
        r"部|局|司|委)"
    )
    _BOILERPLATE_RE = re.compile(
        r"(责任编辑|责编|编辑[:：]|校对[:：]|来源[:：]|免责声明|版权声明|"
        r"版权所有|未经.*许可.*不得.*转载|更多精彩|下载.*客户端|"
        r"打开.*APP|扫码.*关注|关注.*公众号|微信.*扫一扫|"
        r"分享至|我要评论|发表评论|相关新闻|相关阅读|推荐阅读|猜你喜欢|热门.*排行|"
        r"广告|返回首页|导航|登录.*注册|"
        r"沪深.*行情|股票代码)"
    )

    def __init__(
        self,
        db_path: str = "temp-data/crawler_intel.db",
        intel_threshold: float = 0.65,
        min_samples: int = 5,
    ):
        self._db_path = Path(db_path)
        self.QUALITY_THRESHOLD = intel_threshold
        self.MIN_ANALYSIS_SAMPLES = min_samples
        self._db = sqlite3.connect(str(self._db_path), check_same_thread=False)
        self._db.row_factory = sqlite3.Row
        self._ensure_dirs()
        self._init_schema()
        self._load_default_selectors()
        self._load_prompts()
        logger.info(f"CrawlerIntel initialized: {self._db_path}")

    # --------------------------------------------------------
    # 初始化
    # --------------------------------------------------------

    def _ensure_dirs(self) -> None:
        self._db_path.parent.mkdir(parents=True, exist_ok=True)

    def _init_schema(self) -> None:
        self._db.executescript("""
            CREATE TABLE IF NOT EXISTS site_selectors (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                site_key        TEXT    NOT NULL,
                selector        TEXT    NOT NULL,
                reason          TEXT    DEFAULT '',
                priority        INTEGER DEFAULT 0,
                success_count   INTEGER DEFAULT 0,
                fail_count      INTEGER DEFAULT 0,
                avg_quality     REAL    DEFAULT 0.0,
                source          TEXT    DEFAULT 'default',
                is_active       INTEGER DEFAULT 1,
                created_at      TEXT    NOT NULL,
                updated_at      TEXT    NOT NULL
            );
            CREATE UNIQUE INDEX IF NOT EXISTS idx_site_selector
                ON site_selectors(site_key, selector);

            CREATE TABLE IF NOT EXISTS extraction_quality_log (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                url             TEXT    NOT NULL,
                site_key        TEXT    NOT NULL,
                selector_tried  TEXT,
                method          TEXT    NOT NULL,
                char_count      INTEGER DEFAULT 0,
                quality_score   REAL    DEFAULT 0.0,
                success         INTEGER NOT NULL,
                error_msg       TEXT,
                crawled_at      TEXT    NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_quality_site
                ON extraction_quality_log(site_key, crawled_at DESC);

            CREATE TABLE IF NOT EXISTS site_meta (
                site_key        TEXT PRIMARY KEY,
                total_attempts  INTEGER DEFAULT 0,
                success_count   INTEGER DEFAULT 0,
                current_success_rate REAL DEFAULT 0.0,
                avg_quality    REAL    DEFAULT 0.0,
                llm_analysis_count INTEGER DEFAULT 0,
                last_llm_analysis TEXT,
                last_crawled    TEXT,
                health_status   TEXT    DEFAULT 'unknown',
                updated_at      TEXT    NOT NULL
            );

            CREATE TABLE IF NOT EXISTS llm_analysis_history (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                site_key        TEXT    NOT NULL,
                url             TEXT    NOT NULL,
                raw_html_length INTEGER NOT NULL,
                llm_prompt     TEXT    NOT NULL,
                llm_response   TEXT    NOT NULL,
                proposed_selectors TEXT  NOT NULL,
                accepted        INTEGER DEFAULT 0,
                improvement     REAL    DEFAULT 0.0,
                analyzed_at     TEXT    NOT NULL
            );

            CREATE TABLE IF NOT EXISTS html_samples (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                site_key        TEXT    NOT NULL,
                url             TEXT    NOT NULL,
                html_hash       TEXT    NOT NULL,
                raw_html        TEXT,
                extraction_result TEXT,
                success         INTEGER NOT NULL,
                quality_score   REAL    DEFAULT 0.0,
                created_at      TEXT    NOT NULL,
                UNIQUE(site_key, html_hash)
            );
        """)

    def _load_default_selectors(self) -> None:
        """将内置选择器种子写入数据库（已存在的跳过）"""
        now = datetime.now().isoformat()
        for site_key, selectors in DEFAULT_SELECTORS.items():
            for i, selector in enumerate(selectors):
                self._db.execute("""
                    INSERT OR IGNORE INTO site_selectors
                        (site_key, selector, priority, source, created_at, updated_at)
                    VALUES (?, ?, ?, 'default', ?, ?)
                """, (site_key, selector, i, now, now))
            self._db.execute("""
                INSERT OR IGNORE INTO site_meta (site_key, updated_at)
                VALUES (?, ?)
            """, (site_key, now))
        self._db.commit()

    def _load_prompts(self) -> None:
        """加载 YAML 提示词，无则用内置"""
        yaml_path = Path("references/deepseek-prompts.yaml")
        try:
            with open(yaml_path, encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
            self._llm_system = data.get(
                "llm_analysis", {}
            ).get(
                "system", LLM_ANALYSIS_SYSTEM_PROMPT
            )
            self._llm_user_template = data.get(
                "llm_analysis", {}
            ).get(
                "user_template", LLM_ANALYSIS_USER_TEMPLATE
            )
        except Exception:
            self._llm_system = LLM_ANALYSIS_SYSTEM_PROMPT
            self._llm_user_template = LLM_ANALYSIS_USER_TEMPLATE

    # --------------------------------------------------------
    # 工具方法
    # --------------------------------------------------------

    @staticmethod
    def match_site_key(domain: str) -> str:
        """从域名中匹配最长的 site_key（最长前缀匹配）"""
        domain = domain.lower()
        best = ""
        for key in DEFAULT_SELECTORS:
            if key in domain and len(key) > len(best):
                best = key
        return best if best else domain

    def _is_noise_line(self, line: str) -> bool:
        """判断一行是否为噪音（内部复用爬虫逻辑）"""
        if len(line) <= 2:
            return True
        if self._BOILERPLATE_RE.search(line):
            return True
        if line.count("|") >= 3 or line.count("/") >= 5:
            return True
        if len(line) < 12 and not self._KEEP_KEYWORDS_RE.search(line):
            return True
        return False

    # --------------------------------------------------------
    # 质量评分
    # --------------------------------------------------------

    def score_extraction(
        self,
        text: str,
        url: str,
        method: str,
        html_length: int = 0,
    ) -> ExtractionQuality:
        """
        评估一次提取的质量，返回 0-10 分。

        评分维度：
        - 字数合理性（100-5000 字为佳）
        - 噪音行占比
        - 结构化数据（含数字、专有名词等）
        """
        if not text or not text.strip():
            return ExtractionQuality(
                score=0.0, signal_strength=0.0,
                noise_ratio=1.0, has_structured_data=False
            )

        lines = text.split("\n")
        total_lines = max(len(lines), 1)
        noise_lines = sum(1 for l in lines if self._is_noise_line(l.strip()))
        noise_ratio = noise_lines / total_lines

        char_len = len(text)

        # 字数合理性打分（0-4 分）
        if char_len < 50:
            char_score = char_len / 50 * 2
        elif 100 <= char_len <= 5000:
            char_score = 4.0
        elif 50 <= char_len < 100:
            char_score = 2.0
        elif 5000 < char_len <= 10000:
            char_score = 3.5
        else:
            char_score = max(0, 3.0 - (char_len - 10000) / 5000)

        # 噪音惩罚（0-3 分）
        noise_score = max(0, 3.0 * (1 - noise_ratio * 2))

        # 结构化数据加分（0-1 分）
        structured = bool(
            re.search(r"\d+%", text) or
            re.search(r"[一二三四五六七八九十]、", text) or
            re.search(r"第一[、 ，]", text) or
            re.search(r"规模约\d+", text)
        )
        structured_bonus = 1.0 if structured else 0.0

        # 内容密度加分（0-2 分）
        if html_length > 0:
            density = min(char_len / html_length, 0.5)
            density_score = density * 4
        else:
            density_score = min(char_len / 5000, 1.0) * 2

        score = min(char_score + noise_score + structured_bonus + density_score, 10.0)

        return ExtractionQuality(
            score=round(score, 2),
            signal_strength=round(char_len / max(html_length, 1), 4),
            noise_ratio=round(noise_ratio, 3),
            has_structured_data=structured,
        )

    # --------------------------------------------------------
    # 日志记录
    # --------------------------------------------------------

    def log_extraction(
        self,
        url: str,
        site_key: str,
        quality: ExtractionQuality,
        success: bool,
        selector_tried: Optional[str] = None,
        method: str = "bs4",
        error_msg: Optional[str] = None,
    ) -> None:
        """每次提取后调用：写 quality_log + 更新 site_meta + 更新选择器统计"""
        now = datetime.now().isoformat()

        # 写 quality_log
        self._db.execute("""
            INSERT INTO extraction_quality_log
                (url, site_key, selector_tried, method, char_count,
                 quality_score, success, error_msg, crawled_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            url, site_key, selector_tried, method,
            0 if not success else 0,  # char_count 由爬虫侧传入，这里省略
            quality.score, int(success), error_msg, now
        ))

        # 更新 site_meta
        self._db.execute("""
            UPDATE site_meta SET
                total_attempts = total_attempts + 1,
                success_count  = success_count + ?,
                current_success_rate = CASE
                    WHEN total_attempts + 1 >= 20
                    THEN (success_count + ?) * 1.0 / 20.0
                    ELSE (success_count + ?) * 1.0 / (total_attempts + 1)
                END,
                avg_quality    = (avg_quality * total_attempts + ?) / (total_attempts + 1),
                last_crawled   = ?,
                health_status  = CASE
                    WHEN total_attempts + 1 >= 10 THEN
                        CASE
                            WHEN (success_count + ?) * 1.0 / (total_attempts + 1) >= 0.7
                            THEN 'healthy'
                            WHEN (success_count + ?) * 1.0 / (total_attempts + 1) >= 0.4
                            THEN 'degraded'
                            ELSE 'critical'
                        END
                    ELSE health_status
                END,
                updated_at     = ?
            WHERE site_key = ?
        """, (
            int(success), int(success), int(success),
            quality.score,
            now,
            int(success), int(success),
            now, site_key
        ))

        # 如果 site_meta 不存在则插入
        if self._db.execute(
            "SELECT 1 FROM site_meta WHERE site_key = ?", (site_key,)
        ).fetchone() is None:
            self._db.execute("""
                INSERT INTO site_meta (site_key, total_attempts, success_count,
                    current_success_rate, avg_quality, last_crawled, updated_at)
                VALUES (?, 1, ?, ?, ?, ?, ?)
            """, (
                site_key, int(success),
                float(int(success)), quality.score,
                now, now
            ))

        # 更新选择器统计
        if selector_tried:
            self._update_selector_stats(site_key, selector_tried, success, quality.score)

        self._db.commit()

    def _update_selector_stats(
        self, site_key: str, selector: str,
        success: bool, quality_score: float
    ) -> None:
        """更新某选择器的成功/失败计数"""
        if success:
            self._db.execute("""
                UPDATE site_selectors SET
                    success_count = success_count + 1,
                    avg_quality = (avg_quality * success_count + ?) / (success_count + 1),
                    updated_at = ?
                WHERE site_key = ? AND selector = ?
            """, (quality_score, datetime.now().isoformat(), site_key, selector))
        else:
            self._db.execute("""
                UPDATE site_selectors SET
                    fail_count = fail_count + 1,
                    updated_at = ?
                WHERE site_key = ? AND selector = ?
            """, (datetime.now().isoformat(), site_key, selector))

    # --------------------------------------------------------
    # 选择器查询
    # --------------------------------------------------------

    def get_selectors_for_site(self, domain: str) -> List[str]:
        """
        查询知识库中该站点的最优选择器列表（按成功率排序）。
        优先返回数据库中的活跃选择器；无记录时返回内置默认值。
        """
        site_key = self.match_site_key(domain)
        rows = self._db.execute("""
            SELECT selector,
                   (success_count + 1) * 1.0 / (success_count + fail_count + 2) AS win_rate
            FROM site_selectors
            WHERE site_key = ? AND is_active = 1
            ORDER BY win_rate DESC, priority ASC
        """, (site_key,)).fetchall()

        if rows:
            result = [r["selector"] for r in rows]
            logger.debug(f"知识库选择器 [{site_key}]: {result}")
            return result

        # Fallback：内置
        fallback = list(DEFAULT_SELECTORS.get(site_key, []))
        logger.debug(f"内置选择器 [{site_key}]: {fallback}")
        return fallback

    def get_site_selectors_detail(
        self, site_key: str
    ) -> List[Dict[str, Any]]:
        """获取某站点的选择器详情（CLI 展示用）"""
        rows = self._db.execute("""
            SELECT selector, reason, priority, success_count, fail_count,
                   avg_quality, source, is_active, updated_at
            FROM site_selectors
            WHERE site_key = ?
            ORDER BY success_count DESC, priority ASC
        """, (site_key,)).fetchall()
        return [dict(r) for r in rows]

    # --------------------------------------------------------
    # 阈值检测 + 触发
    # --------------------------------------------------------

    def check_and_trigger_analysis(
        self,
        site_key: str,
        url: str,
    ) -> bool:
        """
        检查站点成功率是否低于阈值。
        低于阈值时异步触发 LLM 分析（后台任务，不阻塞爬虫）。
        返回 True 表示触发了分析。
        """
        # 查询该站点最近 20 次提取的统计
        row = self._db.execute("""
            SELECT total_attempts, success_count, current_success_rate,
                   llm_analysis_count, last_llm_analysis
            FROM site_meta WHERE site_key = ?
        """, (site_key,)).fetchone()

        if row is None:
            return False

        total, success, rate, analysis_count, last_analysis = row

        if total < self.MIN_ANALYSIS_SAMPLES:
            return False

        # 频率限制：最近 30 分钟内分析过则跳过
        if last_analysis:
            try:
                last_dt = datetime.fromisoformat(last_analysis)
                if (datetime.now() - last_dt).total_seconds() < 1800:
                    return False
            except Exception:
                pass

        if rate < self.QUALITY_THRESHOLD:
            # 异步触发（不等待完成）
            try:
                loop = asyncio.get_running_loop()
                loop.create_task(
                    self._async_analyze_site(site_key, url)
                )
                logger.info(
                    f"触发 LLM 分析站点 [{site_key}]，成功率 {rate:.2%} < {self.QUALITY_THRESHOLD:.2%}"
                )
                return True
            except RuntimeError:
                # 没有运行中的事件循环（同步上下文），在新的事件循环中运行
                asyncio.run(self._async_analyze_site(site_key, url))
                return True

        return False

    # --------------------------------------------------------
    # LLM 分析（核心方法）
    # --------------------------------------------------------

    async def analyze_and_update_selectors(
        self,
        url: str,
        site_key: str,
        html: Optional[str] = None,
    ) -> AnalysisResult:
        """
        调用 LLM 分析页面 HTML，生成新选择器并写入知识库。
        """
        return await self._async_analyze_site(site_key, url, html)

    async def _async_analyze_site(
        self,
        site_key: str,
        url: str,
        html: Optional[str] = None,
    ) -> AnalysisResult:
        """内部异步分析实现"""
        from .deepseek_client import DeepSeekClient

        now = datetime.now().isoformat()

        # 获取样本
        success_samples = self._get_recent_samples(site_key, success=True, limit=3)
        fail_samples = self._get_recent_samples(site_key, success=False, limit=3)
        existing = self.get_selectors_for_site(site_key)

        if html is None and url:
            html = await self._fetch_html(url)

        # 构造提示词
        prompt_user = self._llm_user_template.format(
            site_key=site_key,
            existing_selectors=", ".join(existing) or "无",
            success_examples="\n---\n".join(
                [s["extraction_result"] or "(无内容)" for s in success_samples]
            ) or "(无成功样本)",
            fail_examples="\n---\n".join(
                [f"{s['url']}: {s['extraction_result'] or '无内容'}" for s in fail_samples]
            ) or "(无失败样本)",
        )

        try:
            client = DeepSeekClient()
            response = await client.chat_async(
                system=self._llm_system,
                user=prompt_user,
                model="deepseek-chat",
                temperature=0.3,
            )
            response = client.strip_response(response or "")
        except Exception as e:
            logger.error(f"LLM 分析失败 [{site_key}]: {e}")
            return AnalysisResult(
                site_key=site_key,
                selectors=[],
                analysis="",
                success=False,
                error=str(e),
            )

        # 解析 JSON
        result = self._parse_llm_selectors(response)

        # 写入数据库
        if result.selectors:
            self._upsert_llm_selectors(site_key, result.selectors)

        # 更新 site_meta
        self._db.execute("""
            UPDATE site_meta SET
                llm_analysis_count = llm_analysis_count + 1,
                last_llm_analysis = ?,
                updated_at = ?
            WHERE site_key = ?
        """, (now, now, site_key))
        self._db.commit()

        # 记录分析历史
        self._db.execute("""
            INSERT INTO llm_analysis_history
                (site_key, url, raw_html_length, llm_prompt, llm_response,
                 proposed_selectors, accepted, analyzed_at)
            VALUES (?, ?, ?, ?, ?, ?, 0, ?)
        """, (
            site_key, url,
            len(html) if html else 0,
            prompt_user, response,
            json.dumps(
                [{"selector": s.selector, "reason": s.reason} for s in result.selectors],
                ensure_ascii=False,
            ),
            now,
        ))
        self._db.commit()

        logger.info(
            f"LLM 分析完成 [{site_key}]，提出 {len(result.selectors)} 个选择器"
        )
        return result

    def _get_recent_samples(
        self, site_key: str, success: bool, limit: int = 3
    ) -> List[Dict[str, Any]]:
        """获取最近的样本"""
        rows = self._db.execute("""
            SELECT url, extraction_result, quality_score
            FROM html_samples
            WHERE site_key = ? AND success = ?
            ORDER BY created_at DESC LIMIT ?
        """, (site_key, int(success), limit)).fetchall()
        return [dict(r) for r in rows]

    async def _fetch_html(self, url: str) -> Optional[str]:
        """抓取指定 URL 的 HTML（供 LLM 分析用）"""
        try:
            import aiohttp
            async with aiohttp.ClientSession(
                headers={
                    "User-Agent": "Mozilla/5.0 (compatible; CrawlerIntel/1.0)",
                    "Accept": "text/html",
                },
                timeout=aiohttp.ClientTimeout(total=15),
            ) as session:
                async with session.get(url, ssl=False) as resp:
                    if resp.status == 200:
                        return await resp.text()
        except Exception as e:
            logger.warning(f"抓取 HTML 失败 [{url}]: {e}")
        return None

    def _parse_llm_selectors(self, response: str) -> AnalysisResult:
        """从 LLM 返回中解析选择器列表"""
        # 尝试提取 JSON 块
        import json as _json

        # 先去掉 markdown code fence
        response = re.sub(r"```(?:json)?\s*", "", response.strip())
        response = re.sub(r"\s*```", "", response.strip())

        try:
            data = _json.loads(response)
            selectors = [
                SelectorCandidate(
                    selector=c["selector"],
                    reason=c.get("reason", ""),
                    source="llm",
                )
                for c in data.get("selectors", [])
                if c.get("selector")
            ]
            return AnalysisResult(
                site_key="",
                selectors=selectors,
                analysis=data.get("analysis", ""),
                success=True,
            )
        except Exception:
            # 尝试正则提取 selector
            matches = re.findall(r'["\']?(\.[a-zA-Z][\w-]*|[#][a-zA-Z][\w-]*|article|\[[^\]]+\])["\']?\s*[:,]', response)
            if matches:
                selectors = [
                    SelectorCandidate(
                        selector=m.strip('"\', '),
                        reason="LLM 正则提取",
                        source="llm",
                    )
                    for m in matches[:5]
                ]
                return AnalysisResult(
                    site_key="", selectors=selectors,
                    analysis="正则提取", success=True,
                )
            return AnalysisResult(
                site_key="", selectors=[],
                analysis="",
                success=False,
                error="JSON 解析失败",
            )

    def _upsert_llm_selectors(
        self, site_key: str, candidates: List[SelectorCandidate]
    ) -> None:
        """将 LLM 生成的选择器写入数据库（source='llm'）"""
        now = datetime.now().isoformat()
        for i, cand in enumerate(candidates):
            self._db.execute("""
                INSERT INTO site_selectors
                    (site_key, selector, reason, priority, source, created_at, updated_at)
                VALUES (?, ?, ?, ?, 'llm', ?, ?)
                ON CONFLICT(site_key, selector) DO UPDATE SET
                    reason = excluded.reason,
                    priority = MIN(site_selectors.priority, excluded.priority),
                    updated_at = excluded.updated_at
            """, (site_key, cand.selector, cand.reason, i, now, now))
        self._db.commit()

    # --------------------------------------------------------
    # HTML 样本管理
    # --------------------------------------------------------

    def save_html_sample(
        self,
        site_key: str,
        url: str,
        html: str,
        extraction_result: Optional[str],
        success: bool,
        quality_score: float,
    ) -> None:
        """保存一条 HTML 样本（去重：相同 site_key + html_hash 不重复插入）"""
        import hashlib
        html_hash = hashlib.md5((html or "")[:5000].encode()).hexdigest()
        now = datetime.now().isoformat()
        try:
            self._db.execute("""
                INSERT INTO html_samples
                    (site_key, url, html_hash, raw_html, extraction_result,
                     success, quality_score, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (site_key, url, html_hash, html[:10000],
                  extraction_result, int(success), quality_score, now))
            self._db.commit()
        except sqlite3.IntegrityError:
            pass  # 已有记录，跳过

    # --------------------------------------------------------
    # 知识库管理工具
    # --------------------------------------------------------

    def get_all_site_status(self) -> List[Dict[str, Any]]:
        """获取所有站点的健康状态（CLI 展示用）"""
        rows = self._db.execute("""
            SELECT site_key, total_attempts, success_count,
                   current_success_rate, avg_quality, llm_analysis_count,
                   last_llm_analysis, health_status, last_crawled
            FROM site_meta ORDER BY current_success_rate ASC
        """).fetchall()
        return [dict(r) for r in rows]

    def prune_selectors(self, min_success_rate: float = 0.2) -> int:
        """删除成功率低于阈值的选择器"""
        cur = self._db.execute("""
            DELETE FROM site_selectors
            WHERE (success_count + 1) * 1.0 / (success_count + fail_count + 2) < ?
               AND source != 'default'
        """, (min_success_rate,))
        self._db.commit()
        return cur.rowcount

    def export_knowledge(self) -> Dict[str, Any]:
        """导出完整知识库为 dict（可序列化）"""
        selectors = self._db.execute(
            "SELECT * FROM site_selectors ORDER BY site_key, priority"
        ).fetchall()
        meta = self._db.execute(
            "SELECT * FROM site_meta ORDER BY site_key"
        ).fetchall()
        history = self._db.execute(
            "SELECT * FROM llm_analysis_history ORDER BY analyzed_at DESC LIMIT 100"
        ).fetchall()
        return {
            "exported_at": datetime.now().isoformat(),
            "site_selectors": [dict(r) for r in selectors],
            "site_meta": [dict(r) for r in meta],
            "llm_analysis_history": [dict(r) for r in history],
        }

    def import_knowledge(self, data: Dict[str, Any]) -> Tuple[int, int, int]:
        """导入知识库，返回 (selectors数, meta数, history数)"""
        count_s, count_m, count_h = 0, 0, 0
        for row in data.get("site_selectors", []):
            self._db.execute("""
                INSERT OR REPLACE INTO site_selectors
                    (site_key, selector, reason, priority, success_count,
                     fail_count, avg_quality, source, is_active, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                row["site_key"], row["selector"], row.get("reason", ""),
                row.get("priority", 0), row.get("success_count", 0),
                row.get("fail_count", 0), row.get("avg_quality", 0),
                row.get("source", "llm"), row.get("is_active", 1),
                row.get("created_at", datetime.now().isoformat()),
                row.get("updated_at", datetime.now().isoformat()),
            ))
            count_s += 1

        for row in data.get("site_meta", []):
            self._db.execute("""
                INSERT OR REPLACE INTO site_meta
                    (site_key, total_attempts, success_count,
                     current_success_rate, avg_quality, llm_analysis_count,
                     last_llm_analysis, last_crawled, health_status, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                row["site_key"], row.get("total_attempts", 0),
                row.get("success_count", 0), row.get("current_success_rate", 0),
                row.get("avg_quality", 0), row.get("llm_analysis_count", 0),
                row.get("last_llm_analysis"), row.get("last_crawled"),
                row.get("health_status", "unknown"),
                row.get("updated_at", datetime.now().isoformat()),
            ))
            count_m += 1

        for row in data.get("llm_analysis_history", []):
            self._db.execute("""
                INSERT INTO llm_analysis_history
                    (site_key, url, raw_html_length, llm_prompt, llm_response,
                     proposed_selectors, accepted, improvement, analyzed_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                row["site_key"], row.get("url", ""),
                row.get("raw_html_length", 0), row.get("llm_prompt", ""),
                row.get("llm_response", ""), row.get("proposed_selectors", "[]"),
                row.get("accepted", 0), row.get("improvement", 0),
                row.get("analyzed_at", datetime.now().isoformat()),
            ))
            count_h += 1

        self._db.commit()
        return count_s, count_m, count_h

    def close(self) -> None:
        """关闭数据库连接"""
        self._db.close()

    def __enter__(self) -> "CrawlerIntel":
        return self

    def __exit__(self, *args) -> None:
        self.close()
