"""
content 预处理脚本 - 清理文章中的噪音信息
"""
import re
import pandas as pd


def clean_content(content: str) -> str:
    if pd.isna(content) or not content:
        return content

    text = str(content)

    # ============================================================
    # 阶段一：整行/段落级别的全局替换（不影响行边界）
    # ============================================================

    # 1A. 移除固定UI元素
    text = re.sub(r'订阅取消订阅已收藏收藏大字号\s*', '', text)
    text = re.sub(r'点击播报本文[，,]?(约|阅读)?\s*\n?', '', text)
    text = re.sub(r'转发分享[:：]?\s*\n?', '', text)
    text = re.sub(r'字号\s*\n?', '', text)
    text = re.sub(r'相关推荐', '', text)
    text = re.sub(r'打印本页', '', text)
    text = re.sub(r'发布于[：:]?\s*', '', text)

    # 1B. 移除嵌入在段落内的时间（行内替换）
    text = re.sub(r'\d{4}年\d{1,2}月\d{1,2}日\s+\d{2}:\d{2}:\d{2}(?!\d)', '', text)
    text = re.sub(r'\d{4}年\d{1,2}月\d{1,2}日\d{2}:\d{2}:\d{2}(?!\d)', '', text)
    text = re.sub(r'\d{4}年\d{1,2}月\d{1,2}日\s+\d{2}:\d{2}(?!\d)', '', text)
    text = re.sub(r'\d{4}年\d{1,2}月\d{1,2}日\d{2}:\d{2}(?!\d)', '', text)
    text = re.sub(r'\d{4}[年\-/]\d{1,2}[月\-/]\d{1,2}\s+\d{2}:\d{2}:\d{2}(?!\d)', '', text)
    text = re.sub(r'\d{4}[年\-/]\d{1,2}[月\-/]\d{1,2}\s+\d{2}:\d{2}(?!\d)', '', text)

    # 1C. 移除栏目标签
    text = re.sub(
        r'【[\u4e00-\u9fff]{2,6}(?:报)?(?:观察|视角|视点|关注|分析|评|看点|快讯|前沿|动态|观点|洞察|声音|对话)】',
        '', text)
    text = re.sub(r'\n?【[^】]*】\n?', '\n', text)

    # ============================================================
    # 阶段二：分行处理
    # ============================================================

    lines = text.split('\n')
    result = []

    for line in lines:
        s = line.strip()

        # ---- 2A. 跳过空行（先收集，最后统一去重） ----
        if not s:
            result.append('')
            continue

        # ---- 2B. 移除来源/作者/责任编辑 ----
        s = re.sub(r'来源[：:]\s*[\u4e00-\u9fff\w\.\-／/]+', '', s)
        s = re.sub(r'作者[：:]\s*[\u4e00-\u9fff\w]+', '', s)
        s = re.sub(r'责任编辑[：:]\s*[\u4e00-\u9fff\w]+', '', s)
        s = re.sub(r'编辑[：:]\s*[\u4e00-\u9fff\w]+', '', s)
        s = re.sub(r'校对[：:]\s*[\u4e00-\u9fff\w]+', '', s)
        s = re.sub(r'审核[：:]\s*[\u4e00-\u9fff\w]+', '', s)
        s = re.sub(r'发文[：:]\s*[\u4e00-\u9fff\w]+', '', s)
        s = re.sub(r'^[文稿文章]\s*[|｜]\s*[\u4e00-\u9fff\w]+', '', s)
        s = re.sub(r'^[发布]\s*$', '', s)

        # ---- 2C. 移除媒体名开头行（如 "人民网 |"） ----
        if re.match(r'^[\u4e00-\u9fff\w\-\.]+网\s*[|｜\-:：]*\s*$', s):
            result.append('')
            continue
        # 媒体名+网 开头
        s = re.sub(r'^[\u4e00-\u9fff\w\-\.]+网\s*[|｜\-:：]*\s*', '', s)

        # ---- 2D. 整行文号：仅当整行就是一个文号时删除 ----
        # 模式：开头可含中文/字母，后面跟〔年份〕编号
        if re.fullmatch(r'[\u4e00-\u9fff\w ]*〔\d{4}〕\d+号', s):
            result.append('')
            continue
        # 纯文号行（仅文号本身）
        if re.fullmatch(r'〔\d{4}〕\d+号', s):
            result.append('')
            continue
        # 仅含文号+空白
        if re.fullmatch(r'\s*〔\d{4}〕\d+号\s*', s):
            result.append('')
            continue

        # ---- 2E. 整行日期/时间 ----
        if re.fullmatch(r'\d{4}[年\-/]\d{1,2}[月\-/]\d{1,2}\s+\d{2}:\d{2}', s):
            result.append('')
            continue
        if re.fullmatch(r'\d{4}[年\-/]\d{1,2}[月\-/]\d{1,2}', s):
            result.append('')
            continue
        if re.fullmatch(r'\d{4}年\d{1,2}月\d{1,2}日', s):
            result.append('')
            continue
        # 独立日期行（如 "日期：2026-04-16"），拆出日期部分后整行删除
        m = re.match(r'^(?:发布|更新|日期)[时间日期]*[：:]\s*(\d{4}[年\-/]\d{1,2}[月\-/]\d{1,2}(?:\s+\d{2}:\d{2})?)$', s)
        if m:
            result.append('')
            continue
        if re.fullmatch(r'来源[：:]\s*', s):
            result.append('')
            continue
        # 联系电话行
        if re.match(r'^联系电话[：:]\s*', s):
            result.append('')
            continue

        # ---- 2F. 跳过目录章节行 ----
        # 纯章节标记行
        if re.match(r'^第[\s\d一二三四五六七八九十百千]*[编篇章节部节]$', s):
            result.append('')
            continue
        if re.match(r'^第[\s\d一二三四五六七八九十百千]*分[编篇章节部节]$', s):
            result.append('')
            continue
        # 章节标记 + 标题
        if re.match(r'^第[\s\d一二三四五六七八九十百千]*[编篇章节部节]\s+', s):
            result.append('')
            continue
        if re.match(r'^第[\s\d一二三四五六七八九十百千]*分[编篇章节部节]\s+', s):
            result.append('')
            continue
        # 目录标记
        if re.match(r'^目\s*录\s*$', s):
            result.append('')
            continue

        result.append(s)

    text = '\n'.join(result)

    # ============================================================
    # 阶段三：行内替换（处理段落内的噪音）
    # ============================================================

    # 3A. 行内文号引用（如 "（工信部联节〔2026〕44号）"）
    text = re.sub(r'【第?\d+号】', '', text)
    text = re.sub(r'〔\d{4}〕\d+号', '', text)
    text = re.sub(r'（\d{4} 年第?\d+号）', '', text)

    # 3B. 行内来源/媒体残余
    text = re.sub(r'来源[：:]\s*[\u4e00-\u9fff\w\.\-／/]+', '', text)
    text = re.sub(r'[－\-—]\s*[\u4e00-\u9fff\w]+', '', text)
    text = re.sub(r'\s*\d{3,}$', '', text)  # 末尾3位以上数字
    text = re.sub(r'\s*[\u4e00-\u9fff]+网\s*$', '', text)

    # 3C. 行内时间（独立行未覆盖的情况）
    text = re.sub(r'\d{4}年\d{1,2}月\d{1,2}日\d{2}:\d{2}(?!\d)', '', text)
    text = re.sub(r'\d{4}[年\-/]\d{1,2}[月\-/]\d{1,2}\s+\d{2}:\d{2}(?!\d)', '', text)
    text = re.sub(r'发布[时间日期][：:]\s*\d{4}[年\-/]\d{1,2}[月\-/]\d{1,2}', '', text)
    text = re.sub(r'更新[时间日期][：:]\s*\d{4}[年\-/]\d{1,2}[月\-/]\d{1,2}', '', text)
    text = re.sub(r'日期[：:]\s*\d{4}[年\-/]\d{1,2}[月\-/]\d{1,2}', '', text)

    # 3D. 清理常见目录标题词（整行）
    catalog_titles = [
        '总 则', '通 则', '监督管理', '监督管理制度',
        '规划和生态环境分区管控', '标准和监测', '生态环境影响评价',
        '生态保护补偿', '突发生态环境事件应对', '保障措施',
        '信息公开与公众参与', '污染防治', '大气污染防治',
        '水污染防治', '海洋污染防治',
    ]
    for title in catalog_titles:
        text = re.sub(rf'^\s*{re.escape(title)}\s*$', '', text, flags=re.MULTILINE)

    # 3E. 底部版权/举报
    text = re.sub(r'[\u4e00-\u9fff]{2,4}网版权所有[^\n]*', '', text)
    text = re.sub(r'举报[^\n]*', '', text)
    text = re.sub(r'未经授权[^\n]*', '', text)

    # ============================================================
    # 阶段四：最终清理
    # ============================================================

    text = re.sub(r'\n{4,}', '\n\n', text)
    text = re.sub(r' +', ' ', text)
    text = re.sub(r'\n ', '\n', text)
    text = re.sub(r' \n', '\n', text)
    text = text.strip()

    return text


if __name__ == '__main__':
    input_path = r'd:\桌面\clientab-main\temp-data\战新与未来产业月报_第四期_爬取结果.csv'
    output_path = r'd:\桌面\clientab-main\temp-data\战新与未来产业月报_第四期_爬取结果_cleaned.csv'
    report_path = r'd:\桌面\clientab-main\temp-data\cleaning_report.txt'

    df = pd.read_csv(input_path, dtype=str)
    print(f'读取 {len(df)} 条记录...')

    df['content'] = df['content'].apply(clean_content)

    df.to_csv(output_path, index=False, encoding='utf-8-sig')
    print(f'已保存到 {output_path}')

    df_orig = pd.read_csv(input_path, dtype=str)
    lines = []
    for i in range(min(10, len(df))):
        orig = str(df_orig.iloc[i]['content'])
        cleaned = str(df.iloc[i]['content'])
        lines.append('=' * 60)
        lines.append(f'[{i}] title: {df_orig.iloc[i].get("title", "")}')
        lines.append(f'原始长度: {len(orig)} | 清理后: {len(cleaned)} | 减少: {len(orig)-len(cleaned)}')
        lines.append('-' * 60)
        lines.append('--- 原始前500字 ---')
        lines.append(orig[:500])
        lines.append('')
        lines.append('--- 清理后前500字 ---')
        lines.append(cleaned[:500])
        lines.append('')

    with open(report_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(f'对比报告已保存到 {report_path}')
