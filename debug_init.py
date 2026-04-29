"""Debug: reproduce deepseek_clean_summarize_excel init step by step"""
import sys, os, asyncio, aiohttp, json
sys.stdout.reconfigure(encoding='utf-8')
print("1. reconfigure done", flush=True)

os.environ['DEEPSEEK_API_KEY'] = 'sk-5d7585b8e85b4d28a25e8f60e696ac78'
os.environ['DEEPSEEK_BASE_URL'] = 'https://api.deepseek.com'
os.environ['DEEPSEEK_FLASH_MODEL'] = 'deepseek-v4-flash'
os.environ['DEEPSEEK_PRO_MODEL'] = 'deepseek-v4-pro'
print("2. env vars set", flush=True)

import pandas as pd
from pathlib import Path
print("3. imports done", flush=True)

path = Path(r"d:\桌面\clientab-main\temp-data\summaries_deepseek_final_cleaned.xlsx")
df = pd.read_excel(path, sheet_name=0, engine="openpyxl", dtype=str)
print(f"4. read done, shape={df.shape}", flush=True)
for col in df.columns:
    df[col] = df[col].astype(object)
print("5. astype done", flush=True)

class Args:
    api_key = None
    base_url = 'https://api.deepseek.com'
    flash_model = 'deepseek-v4-flash'
    pro_model = 'deepseek-v4-pro'
    concurrency_flash = 20
    concurrency_pro = 5
    timeout = 60
    max_retries = 3
    max_input_chars = 12000
    max_cleaned_chars = 10000
    max_summary_chars = 1200
    flash_temperature = 0.1
    pro_temperature = 0.2
    skip_flash = False
    rule_only = False
    overwrite_summary = False
    limit = 3
    concurrency = None
    minimal = False

args = Args()
print("6. args created", flush=True)

df = df.dropna(how='all').copy()
work_indices = list(df.index[:args.limit])
print(f"7. work_indices={work_indices}", flush=True)

sem_flash = asyncio.Semaphore(args.concurrency_flash)
sem_pro = asyncio.Semaphore(args.concurrency_pro)
print("8. semaphores created", flush=True)

class DeepSeekClient:
    def __init__(self, api_key, base_url, timeout=60, max_retries=3):
        self.api_key = api_key
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.max_retries = max_retries

    async def chat(self, session, model, system_prompt, user_prompt, temperature=0.2):
        url = f"{self.base_url}/chat/completions"
        payload = {"model": model, "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ], "temperature": temperature, "stream": False}
        headers = {"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"}
        last_error = ""
        for attempt in range(self.max_retries + 1):
            try:
                async with session.post(url, json=payload, headers=headers, timeout=self.timeout) as resp:
                    text = await resp.text()
                    if resp.status == 429:
                        await asyncio.sleep(min(2**attempt*2, 30))
                        continue
                    if resp.status >= 400:
                        raise RuntimeError(f"HTTP {resp.status}: {text[:500]}")
                    data = json.loads(text)
                    return data["choices"][0]["message"]["content"]
            except RuntimeError:
                raise
            except Exception as exc:
                last_error = str(exc)
                if attempt >= self.max_retries:
                    break
                await asyncio.sleep(min(2**attempt, 10))
        raise RuntimeError(last_error)

client = DeepSeekClient(os.environ['DEEPSEEK_API_KEY'], args.base_url, timeout=args.timeout, max_retries=args.max_retries)
print("9. client created", flush=True)

async def test_api():
    async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=30)) as session:
        result = await client.chat(session, args.flash_model, "You are a helpful assistant.", "Say hi in Chinese.")
        print(f"10. API test result: {result[:50]}", flush=True)

asyncio.run(test_api())
print("11. All OK!", flush=True)
