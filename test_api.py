"""Minimal test: can we call DeepSeek API via aiohttp?"""
import asyncio
import aiohttp
import json
import os

async def main():
    print("1. Creating session...", flush=True)
    async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=30)) as session:
        print("2. Session created", flush=True)
        url = "https://api.deepseek.com/chat/completions"
        headers = {
            "Authorization": "Bearer sk-5d7585b8e85b4d28a25e8f60e696ac78",
            "Content-Type": "application/json",
        }
        payload = {
            "model": "deepseek-v4-flash",
            "messages": [
                {"role": "user", "content": "Say hello in one word."}
            ],
            "stream": False,
        }
        print("3. Posting request...", flush=True)
        async with session.post(url, json=payload, headers=headers) as resp:
            print(f"4. Status: {resp.status}", flush=True)
            text = await resp.text()
            print(f"5. Response: {text[:200]}", flush=True)

asyncio.run(main())
print("Done.", flush=True)
