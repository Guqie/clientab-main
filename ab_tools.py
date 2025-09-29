from ast import literal_eval
import os
import requests
import time
import importlib
import random
import pandas as pd
import json
import ast
from charset_normalizer import detect
import codecs
from scraper import scrape_web_contents, parse_web_contents
from local_scraper import local_scrape_urls, format_for_word_export
from ab_time import now_in_filename, iso_date
from ab_utils import manage_thread, upload_to_container
from export_to_word import export_search_results_to_word, append_company_info_and_disclaimer

# 从Streamlit secrets获取OpenRouter API密钥的函数
import streamlit as st

def get_openrouter_api_key():
    """
    安全地获取OpenRouter API密钥
    Returns:
        str: OpenRouter API密钥
    """
    try:
        return st.secrets["openrouter_api_key"]
    except Exception as e:
        print(f"Failed to get OpenRouter API key: {e}")
        return None

# 延迟获取API密钥，避免在模块导入时就访问secrets
OPENROUTER_API_KEY = None


def execute(tool_calls):
    try:
        results = {
            f"{name}({arguments})": globals().get(name)(**literal_eval(arguments))
            for tool_call in tool_calls
            if (function := tool_call.get("function"))
            if (name := function.get("name")) and (arguments := function.get("arguments"))
            if name in globals()
        }
        return results
    except Exception as e:
        print(f"Failed to execute tool calls: {e}")
        return None


def request_llm(url, headers, data, delay=1):
    for attempt in range(3):
        try:
            print(f"Sending request to {url}")
            response = requests.post(url, headers=headers, json=data, timeout=180).json()
            print(response)
            if (message := response.get("choices", [{}])[0].get("message", {})):
                if (tool_calls := message.get("tool_calls")):
                    if (results := execute(tool_calls)):
                        return f"The following dictionary contains the results:\n{results}"
                elif (content := message.get("content")):
                    return content
            raise Exception("Invalid response or execution failed")
        except Exception as e:
            print(f"Attempt {attempt + 1} failed: {e}")
            if attempt < 2:
                time.sleep(delay)
                delay *= 2
    print("Failed to get a valid response after maximum retries")
    return None


class LLM:
    def __init__(self, url, api_key_func=None):
        """
        初始化LLM实例
        Args:
            url (str): API端点URL
            api_key_func (callable): 获取API密钥的函数，如果为None则使用get_openrouter_api_key
        """
        self.url = url
        self.api_key_func = api_key_func or get_openrouter_api_key

    def __call__(self, messages, model, temperature, top_p, response_format=None, tools=None):
        # 在实际调用时获取API密钥
        api_key = self.api_key_func()
        if not api_key:
            raise Exception("Failed to get API key")
        headers = {
            "Authorization": f"Bearer {api_key}"
        }
        data = {
            "messages": messages,
            "model": model,
            "temperature": temperature,
            "top_p": top_p,
            **({"response_format": response_format} if response_format else {}),
            **({"tools": tools} if tools else {})
        }
        return request_llm(self.url, headers, data)


# 只保留OpenRouter配置
# 创建OpenRouter LLM实例，使用延迟加载的API密钥
openrouter = LLM("https://openrouter.ai/api/v1/chat/completions")


def get_prompt(prompt, **arguments):
    if arguments:
        return getattr(importlib.import_module(f"ab_prompts.{prompt}"), prompt).format(**arguments)
    else:
        return getattr(importlib.import_module(f"ab_prompts.{prompt}"), prompt)


def get_response_format(response_format):
    if response_format:
        return getattr(importlib.import_module(f"ab_response_formats.{response_format}"), response_format)
    return None


def get_tools(tools):
    if tools:
        return [getattr(importlib.import_module("ab_tools"), tool) for tool in tools]
    return None


class Chat:
    def __call__(self, llms, messages, response_format=None, tools=None):
        for llm in llms:
            try:
                results = globals()[llm_dict[llm]["name"]](messages, **llm_dict[llm]["arguments"], response_format=response_format, tools=tools)
                if results:
                    return results
            except Exception:
                continue
        return None

chat = Chat()

def text_chat(ai, user_message):
    llms = ai_dict[ai]["llms"]
    system_message = get_prompt(ai_dict[ai]["system_message"])
    response_format = get_response_format(ai_dict[ai]["response_format"])
    tools = get_tools(ai_dict[ai]["tools"])
    messages = [{"role": "system", "content": system_message}, {"role": "user", "content": user_message}]
    return chat(llms, messages, response_format, tools)


ai_dict = {
    "GPT for text chat": {
        "category": "function_calling",
        "llms": ["gpt4o_openrouter"],
        "system_message": "online_articles_to_word",
        "response_format": None,
        "tools": ["online_articles_from_url_to_word_func", "online_articles_from_raw_to_word_func"],
        "backend_ais": None,
        "max_length": 128000,
        "intro": "OpenRouter: GPT-4o"
    },
    "GPT for extracting info from online article": {
        "category": "internal",
        "llms": ["gpt4o_openrouter"],
        "system_message": "extract_info_from_online_article",
        "response_format": "extract_info_from_online_article_json",
        "tools": None,
        "backend_ais": None,
        "max_length": None,
        "intro": "internal"
    }
}

llm_dict = {
    "gpt4o_openrouter": {
        "name": "openrouter",
        "arguments": {
            "model": "openai/gpt-4o-mini-2024-07-18",
            "temperature": 0.15,
            "top_p": 0.95
        }
    }
}

online_articles_from_url_to_word_func = {
    "type": "function",
    "function": {
        "name": "online_articles_from_url_to_word",
        "description": "Scrape the webpages for the online articles, write them into a document and return a URL for downloading the document or a CSV file.",
        "parameters": {
            "type": "object",
            "properties": {
                "search_results": {
                    "type": "object",
                    "description": "A dictionary in which each key is a category name and the corresponding value is the list of URLs under that category.",
                },
            },
            "required": ["search_results"],
            "additionalProperties": False,
        },
    }
}

online_articles_from_raw_to_word_func = {
    "type": "function",
    "function": {
        "name": "online_articles_from_raw_to_word",
        "description": "Read the Excel or CSV file for the online articles, write them into a document and return a URL for downloading it.",
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "The path of the Excel or CSV file containing the content of each online article.",
                },
            },
            "required": ["file_path"],
            "additionalProperties": False,
        },
    }
}


def search_results_to_csv(search_results):
    csv_path = f"temp-data/{now_in_filename()}.csv"
    pd.DataFrame(columns=["web_url", "web_raw_content", "heading_1", "heading_2", "source", "published_date", "web_content", "body_content"]).to_csv(csv_path, index=False, encoding="utf-8")
    df = pd.read_csv(csv_path, encoding="utf-8")
    df = pd.concat([df, pd.DataFrame([{"heading_1": heading_1, "web_url": web_url}
        for heading_1, web_urls in search_results.items()
        for web_url in web_urls]).reindex(columns=df.columns)])
    df.to_csv(csv_path, index=False, encoding="utf-8")
    return csv_path


def web_contents_from_url_to_csv(csv_path, urls_per_chunk=6, interval_seconds=5, use_local_scraper=False):
    """从CSV文件中的URL抓取网页内容并写回CSV
    
    Args:
        csv_path (str): CSV文件路径
        urls_per_chunk (int): 每批处理的URL数量
        interval_seconds (int): 批次间隔时间
        use_local_scraper (bool): 是否使用本地爬取器（不依赖API）
    
    Returns:
        int: 成功处理的URL数量
    """
    df = pd.read_csv(csv_path, encoding="utf-8")
    
    # 检查web_url列是否存在
    if "web_url" not in df.columns:
        print("Warning: 'web_url' column not found in CSV file")
        return 0
    
    # 过滤有效的URL
    valid_mask = df["web_url"].notna() & df["web_url"].astype(str).str.contains("http", case=False, na=False)
    if not valid_mask.any():
        print("Warning: No valid URLs found")
        return 0
        
    web_urls = df[valid_mask]["web_url"].tolist()
    
    if use_local_scraper:
        print(f"🕷️ 使用本地爬取器处理 {len(web_urls)} 个URL")
        # 使用本地爬取器
        scrape_results = local_scrape_urls(web_urls, delay=interval_seconds)
        
        # 将结果转换为适合的格式
        web_contents = {}
        for url, result in scrape_results.items():
            if result:
                # 格式化为与API爬取器兼容的格式
                content_dict = {
                    1: result['title'],
                    2: result['content']
                }
                web_contents[url] = content_dict
            else:
                web_contents[url] = None
    else:
        print(f"🌐 使用API爬取器处理 {len(web_urls)} 个URL")
        # 分批处理URL以避免过载
        web_url_chunks = [web_urls[i:i + urls_per_chunk] for i in range(0, len(web_urls), urls_per_chunk)]
        web_contents = {}
        for i, web_url_chunk in enumerate(web_url_chunks):
            web_contents.update(scrape_web_contents(web_url_chunk))
            if i < len(web_url_chunks) - 1:
                time.sleep(interval_seconds)
    
    # 安全地映射web_content，处理None值
    def safe_map_content(url):
        content = web_contents.get(url)
        if content is None:
            return None
        return str(content)  # 确保内容是字符串格式
    
    df.loc[valid_mask, "web_content"] = df.loc[valid_mask, "web_url"].map(safe_map_content)
    df.to_csv(csv_path, index=False, encoding="utf-8")
    return valid_mask.sum()


def ensure_csv_utf8(table_path):
    """
    确保表格文件转换为UTF-8编码的CSV格式
    支持Excel (.xlsx, .xls) 和 CSV 文件
    增强：自动检测并修复Excel中的表头问题
    
    Args:
        table_path (str): 输入文件路径
        
    Returns:
        str: 转换后的CSV文件路径，失败时返回None
    """
    if not os.path.exists(table_path):
        print(f"File not found: {table_path}")
        return None
    
    # 生成输出CSV路径
    base_name = os.path.splitext(os.path.basename(table_path))[0]
    output_dir = os.path.dirname(table_path)
    csv_path = os.path.join(output_dir, f"{base_name}.csv")
    
    try:
        if table_path.endswith(('.xlsx', '.xls')):
            # 读取Excel文件，先尝试默认方式
            df = pd.read_excel(table_path, engine="openpyxl")
            
            # 检测是否存在表头问题（列名为Unnamed或Table等，而真实列名在数据行中）
            if _has_header_issue(df):
                print("检测到Excel表头问题，尝试修复...")
                # 尝试跳过第一行重新读取
                df_fixed = pd.read_excel(table_path, header=1, engine="openpyxl")
                if not _has_header_issue(df_fixed):
                    df = df_fixed
                    print("✅ 表头问题已修复")
                else:
                    # 尝试手动修复：使用第一行数据作为列名
                    df = _manual_fix_headers(df)
                    print("✅ 手动修复表头完成")
        elif table_path.endswith(".csv"):
            # 尝试检测编码并读取CSV
            with open(table_path, "rb") as f:
                encoding = codecs.lookup(detect(f.read(min(32768, os.path.getsize(table_path))))["encoding"]).name
            df = pd.read_csv(table_path, encoding=encoding)
        else:
            return None
        
        # 保存为UTF-8 CSV
        df.to_csv(csv_path, index=False, encoding='utf-8')
        return csv_path
        
    except Exception as e:
        print(f"Failed to convert {table_path} to CSV: {e}")
        return None


def _has_header_issue(df):
    """
    检测DataFrame是否存在表头问题
    
    Args:
        df (pd.DataFrame): 待检测的DataFrame
        
    Returns:
        bool: True表示存在表头问题
    """
    if df.empty:
        return False
    
    # 检查列名是否大部分为Unnamed、Table等无意义名称
    unnamed_count = sum(1 for col in df.columns if 
                       str(col).startswith(('Unnamed', 'Table')) or 
                       pd.isna(col))
    
    # 如果超过一半的列名无意义，认为存在表头问题
    return unnamed_count > len(df.columns) / 2


def _manual_fix_headers(df):
    """
    手动修复DataFrame的表头问题
    使用第一行非空数据作为列名
    
    Args:
        df (pd.DataFrame): 原始DataFrame
        
    Returns:
        pd.DataFrame: 修复后的DataFrame
    """
    if df.empty:
        return df
    
    # 寻找第一行包含有意义列名的行
    for i in range(min(3, len(df))):
        row = df.iloc[i]
        # 检查这一行是否包含类似列名的内容
        if _looks_like_headers(row):
            # 使用这一行作为列名
            new_columns = [str(val).strip() if pd.notna(val) else f"Column_{j}" 
                          for j, val in enumerate(row)]
            df_fixed = df.iloc[i+1:].copy()
            df_fixed.columns = new_columns
            df_fixed.reset_index(drop=True, inplace=True)
            return df_fixed
    
    # 如果没找到合适的表头行，返回原DataFrame
    return df


def _looks_like_headers(row):
    """
    判断一行数据是否看起来像列名
    
    Args:
        row (pd.Series): 数据行
        
    Returns:
        bool: True表示像列名
    """
    # 常见的列名关键词
    header_keywords = {
        '标题', 'title', 'headline', '文章标题',
        'url', '链接', 'link', '网页链接', '地址',
        '内容', 'content', '正文', '原文',
        '来源', 'source', '媒体', '出处',
        '时间', 'time', 'date', '日期', '发布时间'
    }
    
    # 检查行中是否包含列名关键词
    text_values = [str(val).strip().lower() for val in row if pd.notna(val)]
    matches = sum(1 for text in text_values 
                 if any(keyword in text for keyword in header_keywords))
    
    # 如果有超过一半的值包含关键词，认为是表头
    return matches > 0 and matches >= len([v for v in row if pd.notna(v)]) / 2


def web_contents_from_raw_to_csv(csv_path):
    """从CSV文件中的原始网页内容解析出结构化内容"""
    df = pd.read_csv(csv_path, encoding="utf-8")
    
    # 检查是否存在web_raw_content列，如果不存在则跳过处理
    if "web_raw_content" not in df.columns:
        print("Warning: 'web_raw_content' column not found in CSV file")
        return
    
    valid_mask = df["web_raw_content"].notna()
    if not valid_mask.any():
        print("Warning: No valid web_raw_content found")
        return
        
    web_raw_contents = df[valid_mask]["web_raw_content"].tolist()
    web_contents = parse_web_contents(web_raw_contents)
    df.loc[valid_mask, "web_content"] = df.loc[valid_mask, "web_raw_content"].map(web_contents)
    df.to_csv(csv_path, index=False, encoding="utf-8")


def extend_body_content_bounds(web_content, body_content_bounds):
    start_bound, end_bound = body_content_bounds
    while start_bound - 1 in web_content and web_content[start_bound - 1].startswith("temp-images"):
        start_bound -= 1
    while end_bound + 1 in web_content and web_content[end_bound + 1].startswith("temp-images"):
        end_bound += 1
    return start_bound, end_bound


def extract_info_from_online_article(web_url, web_content, delay=1):
    ai = "GPT for extracting info from online article"
    user_message = f"<web_content>{(dict(list(web_content.items())[:80] + list(web_content.items())[-80:]) if len(web_content) > 160 else web_content)}</web_content>"
    for attempt in range(3):
        try:
            results = text_chat(ai, user_message)
            results = json.loads(results)
            title = results.get("title")
            source = results.get("source")
            published_date = iso_date(results.get("published_date"))
            body_content_bounds = results.get("body_content_bounds")
            if title and source and published_date and len(body_content_bounds) == 2:
                start_bound, end_bound = extend_body_content_bounds(web_content, body_content_bounds)
                body_content = {key: web_content[key] for key in range(start_bound, end_bound + 1) if key in web_content}
                return title, source, published_date, body_content
        except Exception as e:
            print(f"Attempt {attempt + 1} failed: {e}")
            if attempt < 2:
                time.sleep(delay)
                delay *= 2
    print("Failed to extract info after maximum retries")
    return None, None, None, None


def extract_info_from_online_articles(web_urls, web_contents):
    requests = [(extract_info_from_online_article, web_url, web_content) for web_url, web_content in zip(web_urls, web_contents)]
    return {arguments[0]: result for result, name, arguments in manage_thread(requests)}


def info_from_web_contents_to_csv(csv_path):
    """从网页内容中提取信息并写入CSV文件"""
    df = pd.read_csv(csv_path, encoding="utf-8")
    
    # 检查web_content列是否存在
    if "web_content" not in df.columns:
        print("Warning: 'web_content' column not found in CSV file")
        return 0
    
    # 更严格的有效性检查
    valid_mask = df["web_content"].notna() & (df["web_content"] != '') & (df["web_content"] != 'None')
    if not valid_mask.any():
        print("Warning: No valid web_content found")
        return 0
        
    web_urls = df[valid_mask]["web_url"].tolist()
    
    # 安全地处理web_content，避免None值导致的错误
    web_contents = []
    for web_content in df[valid_mask]["web_content"].tolist():
        try:
            if web_content and web_content != 'None':
                parsed_content = ast.literal_eval(web_content)
                web_contents.append(parsed_content)
            else:
                web_contents.append(None)
        except (ValueError, SyntaxError) as e:
            print(f"Warning: Failed to parse web_content: {e}")
            web_contents.append(None)
    
    # 过滤掉None值的内容
    valid_contents = [(url, content) for url, content in zip(web_urls, web_contents) if content is not None]
    if not valid_contents:
        print("Warning: No valid parsed web_content found")
        return 0
    
    valid_urls, valid_web_contents = zip(*valid_contents)
    info = extract_info_from_online_articles(list(valid_urls), list(valid_web_contents))
    
    df.loc[valid_mask, "heading_2"] = df.loc[valid_mask, "web_url"].map({web_url: values[0] for web_url, values in info.items()})
    df.loc[valid_mask, "source"] = df.loc[valid_mask, "web_url"].map({web_url: values[1] for web_url, values in info.items()})
    df.loc[valid_mask, "published_date"] = df.loc[valid_mask, "web_url"].map({web_url: values[2] for web_url, values in info.items()})
    df.loc[valid_mask, "body_content"] = df.loc[valid_mask, "web_url"].map({web_url: str(values[3]) if values[3] else None for web_url, values in info.items()})
    df.to_csv(csv_path, index=False, encoding="utf-8")
    return df.loc[valid_mask, ["heading_2", "source", "published_date", "body_content"]].notna().all(axis=1).sum()


def info_from_web_raw_contents_to_csv(csv_path):
    """从网页内容中提取信息并写入CSV文件"""
    df = pd.read_csv(csv_path, encoding="utf-8")
    
    # 检查必要的列是否存在
    required_columns = ["web_content", "web_url"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        print(f"Warning: Missing required columns: {missing_columns}")
        return
    
    # 使用web_content列而不是web_raw_content列来判断有效性
    # 同时检查web_content不为None且不为空字符串
    valid_mask = df["web_content"].notna() & (df["web_content"] != '') & (df["web_content"] != 'None')
    if not valid_mask.any():
        print("Warning: No valid web_content found")
        return
        
    web_urls = df[valid_mask]["web_url"].tolist()
    
    # 安全地处理web_content，避免None值导致的错误
    web_contents = []
    for web_content in df[valid_mask]["web_content"].tolist():
        try:
            if web_content and web_content != 'None':
                parsed_content = ast.literal_eval(web_content)
                web_contents.append(parsed_content)
            else:
                web_contents.append(None)
        except (ValueError, SyntaxError) as e:
            print(f"Warning: Failed to parse web_content: {e}")
            web_contents.append(None)
    
    # 过滤掉None值的内容
    valid_contents = [(url, content) for url, content in zip(web_urls, web_contents) if content is not None]
    if not valid_contents:
        print("Warning: No valid parsed web_content found")
        return
    
    valid_urls, valid_web_contents = zip(*valid_contents)
    info = extract_info_from_online_articles(list(valid_urls), list(valid_web_contents))
    
    df.loc[valid_mask, "heading_2"] = df.loc[valid_mask, "web_url"].map({web_url: values[0] for web_url, values in info.items()})
    df.loc[valid_mask, "source"] = df.loc[valid_mask, "web_url"].map({web_url: values[1] for web_url, values in info.items()})
    df.loc[valid_mask, "published_date"] = df.loc[valid_mask, "web_url"].map({web_url: values[2] for web_url, values in info.items()})
    df.loc[valid_mask, "body_content"] = df.loc[valid_mask, "web_url"].map({web_url: str(values[3]) if values[3] else None for web_url, values in info.items()})
    df.to_csv(csv_path, index=False, encoding="utf-8")


def online_articles_from_url_to_word(search_results):
    csv_path = search_results_to_csv(search_results)
    web_url_count = web_contents_from_url_to_csv(csv_path)
    article_info_count = info_from_web_contents_to_csv(csv_path)
    if web_url_count == article_info_count:
        doc_path = export_search_results_to_word(csv_path)
        append_company_info_and_disclaimer(doc_path)
        return upload_to_container(doc_path)
    else:
        return upload_to_container(csv_path)


def online_articles_from_raw_to_word(file_path, use_local_scraper=False):
    """从原始（Excel/CSV）文件生成排版好的Word文档（免手工预处理版）
    
    功能增强：
    1) 自动规范化列名：自动将常见同义列映射为内部标准列（如 URL/链接→web_url，标题→heading_2，原始内容→web_raw_content），免去手工改列名；
    2) 双路径自适应：优先使用 web_raw_content 解析；若无则按 web_url 在线抓取；
    3) 数据清洗与校验：对 URL 做 http(s) 过滤与去空值处理，确保后续模块稳健；
    4) 优雅降级：若Word生成失败，返回加工后的CSV下载链接。
    5) 本地爬取支持：可选择使用本地爬取器，不依赖外部API服务。
    
    Args:
        file_path (str): 输入文件路径（Excel或CSV）。允许出现"标题, URL"等常见中文列名，无需手工预处理。
        use_local_scraper (bool): 是否使用本地爬取器（默认False，使用API爬取）
        
    Returns:
        str | None: 生成的文档下载链接；若生成失败则返回CSV文件链接；若无法识别输入结构则返回None。
    """
    # 先确保转换为UTF-8 CSV
    csv_path = ensure_csv_utf8(file_path)
    if not csv_path:
        return None

    # 读取CSV
    try:
        df = pd.read_csv(csv_path, encoding="utf-8")
    except Exception as e:
        print(f"Failed to read CSV at {csv_path}: {e}")
        return None

    # -------------------- 列名规范化（免手工预处理的核心） --------------------
    # 若标准列已存在则不覆盖；从常见同义列中选择第一个命中进行映射
    def normalize_columns(df):
        # 预定义同义集合（全部小写、去空格匹配）
        synonyms = {
            "web_url": {"web_url", "url", "链接", "link", "网页链接", "网页", "网页地址", "地址", "文章链接", "原文链接", "来源链接", "website", "weburl"},
            "web_raw_content": {"web_raw_content", "raw", "raw_text", "rawtext", "raw_content", "原始内容", "抓取原文", "网页原文", "网页源码", "html", "正文raw", "原文"},
            "heading_2": {"heading_2", "标题", "title", "headline", "文章标题", "新闻标题", "报告标题"},
            "source": {"source", "来源", "媒体", "发布机构", "网站来源", "出处"},
            "published_date": {"published_date", "发布日期", "发布时间", "发布日期时间", "日期", "时间", "publish_date", "publishdate", "publish_time", "date", "time"},
            "body_content": {"body_content", "正文", "正文内容", "文章内容", "内容", "文本", "纯文本", "body", "content"},
        }

        current = set(df.columns)
        mapping = {}

        def norm(s):
            return str(s).strip().lower().replace(" ", "")

        normalized_to_originals = {}
        for c in df.columns:
            normalized_to_originals.setdefault(norm(c), []).append(c)

        for target, alias_set in synonyms.items():
            if target in current:
                continue  # 已存在标准列，不做映射
            # 在别名中寻找第一个命中的原始列
            chosen_original = None
            for alias in alias_set:
                alias_norm = norm(alias)
                if alias_norm in normalized_to_originals:
                    # 选择首个同名原始列
                    chosen_original = normalized_to_originals[alias_norm][0]
                    break
            if chosen_original and chosen_original not in mapping:
                mapping[chosen_original] = target

        if mapping:
            print(f"Normalize columns mapping: {mapping}")
            df.rename(columns=mapping, inplace=True)
        return df

    df = normalize_columns(df)

    # -------------------- 数据清洗：web_url 合法化 --------------------
    if "web_url" in df.columns:
        # 统一为字符串，去空白，将无效字符串置为NaN
        df["web_url"] = (
            df["web_url"].astype(str).str.strip().replace({"nan": None, "None": None, "": None})
        )
        # 仅保留 http(s) 开头的链接，其余置为None
        def _clean_url(u):
            if isinstance(u, str) and u.lower().startswith(("http://", "https://")):
                return u
            return None
        df["web_url"] = df["web_url"].map(_clean_url)

    # 将规范化与清洗后的数据写回，以便后续函数按标准列名读取
    try:
        df.to_csv(csv_path, index=False, encoding="utf-8")
    except Exception as e:
        print(f"Failed to persist normalized CSV at {csv_path}: {e}")
        return None

    # -------------------- 判定输入结构，选择处理路径 --------------------
    has_raw = ("web_raw_content" in df.columns) and df["web_raw_content"].notna().any()
    has_url = ("web_url" in df.columns) and df["web_url"].notna().any()

    if has_raw:
        print("Detected 'web_raw_content' column. Using raw-content parsing pipeline.")
        web_contents_from_raw_to_csv(csv_path)
        info_from_web_raw_contents_to_csv(csv_path)
    elif has_url:
        scraper_type = "本地爬取器" if use_local_scraper else "API爬取器"
        print(f"Detected 'web_url' column. Using online-scraping pipeline with {scraper_type}.")
        web_contents_from_url_to_csv(csv_path, use_local_scraper=use_local_scraper)
        info_from_web_contents_to_csv(csv_path)
    else:
        print("Neither 'web_raw_content' nor 'web_url' columns with valid values were found after normalization. Abort.")
        return None

    # -------------------- 统一导出Word，失败则降级为CSV --------------------
    doc_path = export_search_results_to_word(csv_path)
    if doc_path is None:
        print("Word document generation failed, returning CSV file instead")
        return upload_to_container(csv_path)

    append_company_info_and_disclaimer(doc_path)
    return upload_to_container(doc_path)


if __name__ == "__main__":
    csv_path = ""

   # online_articles_from_url_to_word(search_results)
    online_articles_from_raw_to_word(csv_path)
