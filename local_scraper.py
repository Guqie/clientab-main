#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本地网页爬取模块 - 不依赖外部API服务
使用requests + BeautifulSoup + readability等库实现本地爬取
"""

import requests
import time
import re
from urllib.parse import urljoin, urlparse
from bs4 import BeautifulSoup
from readability import Document
import html2text
from fake_useragent import UserAgent
import warnings
warnings.filterwarnings('ignore')

# 用户代理池
ua = UserAgent()

def get_random_headers():
    """
    获取随机请求头，模拟真实浏览器访问
    
    Returns:
        dict: 包含User-Agent等信息的请求头
    """
    return {
        'User-Agent': ua.random,
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
    }

def clean_text(text):
    """
    清理提取的文本内容
    
    Args:
        text (str): 原始文本
        
    Returns:
        str: 清理后的文本
    """
    if not text:
        return ""
    
    # 移除多余的空白字符
    text = re.sub(r'\s+', ' ', text)
    # 移除特殊字符
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', text)
    # 移除过多的换行
    text = re.sub(r'\n\s*\n', '\n\n', text)
    
    return text.strip()

def extract_with_readability(html_content, url):
    """
    使用readability库提取主要内容
    
    Args:
        html_content (str): HTML内容
        url (str): 网页URL
        
    Returns:
        tuple: (标题, 正文内容)
    """
    try:
        doc = Document(html_content)
        title = doc.title()
        content = doc.summary()
        
        # 转换为纯文本
        h = html2text.HTML2Text()
        h.ignore_links = True
        h.ignore_images = True
        h.body_width = 0
        
        text_content = h.handle(content)
        text_content = clean_text(text_content)
        
        return title, text_content
    except Exception as e:
        print(f"Readability extraction failed for {url}: {e}")
        return None, None

def extract_with_beautifulsoup(html_content, url):
    """
    使用BeautifulSoup提取内容
    
    Args:
        html_content (str): HTML内容
        url (str): 网页URL
        
    Returns:
        tuple: (标题, 正文内容)
    """
    try:
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # 移除脚本和样式标签
        for script in soup(["script", "style", "nav", "header", "footer", "aside"]):
            script.decompose()
        
        # 提取标题
        title = ""
        title_tags = soup.find_all(['title', 'h1'])
        if title_tags:
            title = title_tags[0].get_text().strip()
        
        # 尝试找到主要内容区域
        content_selectors = [
            'article', 'main', '.content', '.post-content', '.entry-content',
            '.article-content', '.post-body', '.content-body', '#content',
            '.main-content', '.article-body'
        ]
        
        main_content = None
        for selector in content_selectors:
            elements = soup.select(selector)
            if elements:
                main_content = elements[0]
                break
        
        # 如果没找到主要内容区域，使用body
        if not main_content:
            main_content = soup.find('body')
        
        if main_content:
            # 移除不需要的元素
            for unwanted in main_content.find_all(['script', 'style', 'nav', 'header', 'footer', 'aside', '.sidebar', '.advertisement']):
                unwanted.decompose()
            
            text_content = main_content.get_text()
            text_content = clean_text(text_content)
            
            return title, text_content
        
        return title, ""
        
    except Exception as e:
        print(f"BeautifulSoup extraction failed for {url}: {e}")
        return None, None

def local_scrape_single_url(url, timeout=30, max_retries=3):
    """
    本地爬取单个URL
    
    Args:
        url (str): 要爬取的URL
        timeout (int): 请求超时时间
        max_retries (int): 最大重试次数
        
    Returns:
        dict: 包含标题和内容的字典，失败返回None
    """
    print(f"🕷️ 开始本地爬取: {url}")
    
    for attempt in range(max_retries):
        try:
            headers = get_random_headers()
            
            # 发送请求
            response = requests.get(
                url, 
                headers=headers, 
                timeout=timeout,
                verify=False,  # 忽略SSL证书验证
                allow_redirects=True
            )
            
            # 检查响应状态
            if response.status_code != 200:
                print(f"❌ HTTP {response.status_code} for {url}")
                if attempt < max_retries - 1:
                    time.sleep(2 ** attempt)  # 指数退避
                    continue
                return None
            
            # 检测编码
            response.encoding = response.apparent_encoding or 'utf-8'
            html_content = response.text
            
            if len(html_content) < 100:
                print(f"❌ Content too short for {url}")
                if attempt < max_retries - 1:
                    time.sleep(2 ** attempt)
                    continue
                return None
            
            # 尝试使用readability提取内容
            title, content = extract_with_readability(html_content, url)
            
            # 如果readability失败，使用BeautifulSoup
            if not content or len(content) < 200:
                print(f"🔄 Readability failed, trying BeautifulSoup for {url}")
                title, content = extract_with_beautifulsoup(html_content, url)
            
            if content and len(content) >= 200:
                print(f"✅ Successfully scraped {url} ({len(content)} chars)")
                return {
                    'url': url,
                    'title': title or '',
                    'content': content,
                    'method': 'local_scraper',
                    'length': len(content)
                }
            else:
                print(f"❌ Extracted content too short for {url}")
                if attempt < max_retries - 1:
                    time.sleep(2 ** attempt)
                    continue
                    
        except requests.exceptions.Timeout:
            print(f"⏰ Timeout for {url} (attempt {attempt + 1})")
        except requests.exceptions.ConnectionError:
            print(f"🔌 Connection error for {url} (attempt {attempt + 1})")
        except Exception as e:
            print(f"❌ Error scraping {url} (attempt {attempt + 1}): {e}")
        
        if attempt < max_retries - 1:
            time.sleep(2 ** attempt)  # 指数退避
    
    print(f"❌ Failed to scrape {url} after {max_retries} attempts")
    return None

def local_scrape_urls(urls, delay=1):
    """
    批量本地爬取URL列表
    
    Args:
        urls (list): URL列表
        delay (float): 请求间隔时间（秒）
        
    Returns:
        dict: URL到爬取结果的映射
    """
    results = {}
    
    print(f"🚀 开始批量本地爬取 {len(urls)} 个URL")
    
    for i, url in enumerate(urls, 1):
        print(f"\n📄 处理第 {i}/{len(urls)} 个URL")
        
        result = local_scrape_single_url(url)
        results[url] = result
        
        # 添加延迟，避免被反爬
        if i < len(urls):
            time.sleep(delay)
    
    # 统计结果
    success_count = sum(1 for result in results.values() if result is not None)
    print(f"\n📊 爬取完成: {success_count}/{len(urls)} 成功")
    
    return results

def format_for_word_export(scrape_results):
    """
    将爬取结果格式化为适合Word导出的格式
    
    Args:
        scrape_results (dict): 爬取结果字典
        
    Returns:
        list: 格式化后的文章列表
    """
    articles = []
    
    for url, result in scrape_results.items():
        if result:
            article = {
                'web_url': url,
                'heading_2': result.get('title', ''),
                'source': urlparse(url).netloc,  # 使用域名作为来源
                'published_date': '',  # 本地爬取无法获取发布日期
                'body_content': result.get('content', ''),
                'web_content': result.get('content', ''),
                'scrape_method': 'local'
            }
            articles.append(article)
    
    return articles

if __name__ == "__main__":
    # 测试代码
    test_urls = [
        "https://www.example.com",
        "https://news.sina.com.cn"
    ]
    
    results = local_scrape_urls(test_urls)
    
    for url, result in results.items():
        if result:
            print(f"\n✅ {url}:")
            print(f"   标题: {result['title'][:50]}...")
            print(f"   内容长度: {len(result['content'])} 字符")
        else:
            print(f"\n❌ {url}: 爬取失败")