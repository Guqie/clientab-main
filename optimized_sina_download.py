
def optimized_sina_image_download(url, max_retries=3):
    """
    基于测试结果优化的新浪图片下载函数
    测试结果显示 basic 策略效果最佳
    
    Args:
        url: 图片URL
        max_retries: 最大重试次数
        
    Returns:
        下载结果 (success, content, error_message)
    """
    import requests
    import time
    import random
    from urllib.parse import urlparse
    
    session = requests.Session()
    
    # 动态User-Agent池
    user_agents = [
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/121.0',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    ]
    
    # 新浪Referer池
    sina_referers = [
        'https://finance.sina.com.cn/',
        'https://news.sina.com.cn/',
        'https://www.sina.com.cn/',
        'https://mobile.sina.com.cn/'
    ]
    
    for retry in range(max_retries + 1):
        try:
            # 构建请求头
            headers = {
                'User-Agent': random.choice(user_agents),
                'Accept': 'image/webp,image/apng,image/*,*/*;q=0.8',
                'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
                'Accept-Encoding': 'gzip, deflate, br',
                'Cache-Control': 'no-cache',
                'Pragma': 'no-cache',
                'Sec-Fetch-Dest': 'image',
                'Sec-Fetch-Mode': 'no-cors',
                'Sec-Fetch-Site': 'cross-site'
            }
            
            # 处理新浪图片的Referer
            parsed_url = urlparse(url)
            if 'sinaimg.cn' in parsed_url.netloc or 'sina.com' in parsed_url.netloc:
                headers['Referer'] = random.choice(sina_referers)
            else:
                headers['Referer'] = 'https://www.google.com/'
            
            # 添加重试延迟
            if retry > 0:
                delay = random.uniform(1.0, 3.0) * retry
                time.sleep(delay)
            
            # 发送请求
            response = session.get(
                url,
                headers=headers,
                timeout=30,
                stream=True
            )
            
            if response.status_code == 200:
                return True, response.content, ""
            else:
                last_error = f"HTTP {response.status_code}"
                
        except Exception as e:
            last_error = str(e)
    
    return False, None, last_error
