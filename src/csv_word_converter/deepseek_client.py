#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepSeek API 客户端

提供异步 HTTP 调用封装，支持：
- 自动从环境变量读取配置（DEEPSEEK_API_KEY, DEEPSEEK_BASE_URL）
- 支持显式参数覆盖
- 429 限流自动退避重试
- 响应清洗（去除 Markdown 代码块）
"""

import asyncio
import json
import logging
import os
import time
from typing import Any, Dict, List, Optional

import aiohttp

logger = logging.getLogger(__name__)

DEFAULT_BASE_URL = "https://api.deepseek.com"
DEFAULT_MODEL = "deepseek-chat"
DEFAULT_TIMEOUT = 60


def _strip_model_output(text: str) -> str:
    """去除模型输出中的 Markdown 代码块包装"""
    if not text:
        return ""
    text = text.strip()
    text = text.strip("`")
    if text.startswith("json"):
        text = text[4:].strip()
    return text.strip()


class DeepSeekClient:
    """
    DeepSeek API 异步客户端

    用法：
        client = DeepSeekClient()
        async with aiohttp.ClientSession() as session:
            result = await client.chat_async(session, model="deepseek-chat",
                                             system="你是一个助手", user="你好")
    """

    def __init__(
        self,
        api_key: Optional[str] = None,
        base_url: Optional[str] = None,
        timeout: int = DEFAULT_TIMEOUT,
        max_retries: int = 3,
    ):
        self.api_key = api_key or os.environ.get("DEEPSEEK_API_KEY", "")
        self.base_url = (base_url or os.environ.get("DEEPSEEK_BASE_URL", DEFAULT_BASE_URL)).rstrip("/")
        self.timeout = timeout
        self.max_retries = max_retries

        if not self.api_key:
            logger.warning("DEEPSEEK_API_KEY 未设置，LLM 调用将失败")

    async def chat_async(
        self,
        session: aiohttp.ClientSession,
        model: str = DEFAULT_MODEL,
        system: str = "",
        user: str = "",
        messages: Optional[List[Dict[str, str]]] = None,
        temperature: float = 0.2,
        **kwargs,
    ) -> str:
        """
        发送对话请求。

        参数：
            session      — 复用的 aiohttp.ClientSession
            model        — 模型名称
            system       — system prompt（与 messages 互斥）
            user         — user prompt（与 messages 互斥）
            messages     — 完整消息列表（与 system/user 互斥）
            temperature  — 随机性参数
        """
        if messages is None:
            messages = []
            if system:
                messages.append({"role": "system", "content": system})
            messages.append({"role": "user", "content": user})

        payload = {
            "model": model,
            "messages": messages,
            "temperature": temperature,
            "stream": False,
            **kwargs,
        }
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        url = f"{self.base_url}/chat/completions"

        last_error = ""
        for attempt in range(self.max_retries + 1):
            try:
                async with session.post(
                    url, json=payload, headers=headers,
                    timeout=aiohttp.ClientTimeout(total=self.timeout),
                ) as resp:
                    text = await resp.text()
                    if resp.status == 429:
                        retry_after = resp.headers.get("Retry-After", "5")
                        wait = int(retry_after) if retry_after.isdigit() else min(
                            2 ** attempt * 2, 60
                        )
                        logger.warning(f"429 限流，等待 {wait}s 后重试（第 {attempt + 1} 次）")
                        await asyncio.sleep(wait)
                        continue
                    if resp.status >= 400:
                        raise RuntimeError(f"HTTP {resp.status}: {text[:500]}")
                    data = json.loads(text)
                    return _strip_model_output(
                        data["choices"][0]["message"]["content"]
                    )
            except RuntimeError:
                raise
            except Exception as exc:
                last_error = str(exc)
                if attempt >= self.max_retries:
                    break
                wait = min(2 ** attempt, 10) + time.time() % 1
                await asyncio.sleep(wait)

        raise RuntimeError(f"DeepSeek 请求失败: {last_error}")

    def chat(self, model: str = DEFAULT_MODEL, system: str = "", user: str = "",
             temperature: float = 0.2, **kwargs) -> str:
        """
        同步版本（内部创建事件循环）。
        """
        payload = {
            "model": model,
            "messages": [
                *([{"role": "system", "content": system}] if system else []),
                {"role": "user", "content": user},
            ],
            "temperature": temperature,
            "stream": False,
            **kwargs,
        }
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        url = f"{self.base_url}/chat/completions"

        import urllib.request
        import urllib.error

        last_error = ""
        for attempt in range(self.max_retries + 1):
            try:
                req = urllib.request.Request(
                    url,
                    data=json.dumps(payload).encode("utf-8"),
                    headers=headers,
                    method="POST",
                )
                with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
                    return _strip_model_output(
                        data["choices"][0]["message"]["content"]
                    )
            except urllib.error.HTTPError as e:
                body = e.read().decode("utf-8", errors="replace")
                if e.code == 429:
                    wait = min(2 ** attempt * 2, 60)
                    logger.warning(f"429 限流，等待 {wait}s")
                    time.sleep(wait)
                    continue
                raise RuntimeError(f"HTTP {e.code}: {body[:500]}")
            except Exception as exc:
                last_error = str(exc)
                if attempt >= self.max_retries:
                    break
                time.sleep(min(2 ** attempt, 10))

        raise RuntimeError(f"DeepSeek 请求失败: {last_error}")

    @staticmethod
    def strip_response(text: str) -> str:
        """静态方法：清洗模型输出"""
        return _strip_model_output(text)
