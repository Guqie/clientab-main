# -*- coding: utf-8 -*-
"""
战新与未来产业月报第四期 - 完整格式校验
"""
from docx import Document
import re

def full_review():
    doc = Document(r'd:\桌面\战新与未来产业月报第四期0427.docx')
    
    all_issues = []
    
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        
        line_num = i + 1
        
        # 1. 检查中车
        zhongche_words = ['中国中车', '中车集团', '中车株所', '株机', '株洲所', '株洲电力机车', 'CRRC', '中车长客', '中车四方', '中车唐山', '中车大连']
        for word in zhongche_words:
            if word in text:
                all_issues.append({
                    'line': line_num,
                    'type': '中车新闻',
                    'severity': 'error',
                    'word': word,
                    'content': text[:80]
                })
        
        # 2. 检查标题格式
        if re.match(r'^\d+\.', text):
            if text[1] == '.' and len(text) > 2 and text[2] != ' ':
                all_issues.append({
                    'line': line_num,
                    'type': '标题格式',
                    'severity': 'warning',
                    'content': text[:60]
                })
        
        # 3. 检查URL
        urls = re.findall(r'https?://[^\s）\)]+', text)
        if urls:
            all_issues.append({
                'line': line_num,
                'type': '含URL',
                'severity': 'info',
                'urls': urls,
                'content': text[:60]
            })
    
    return all_issues

def main():
    print("=" * 70)
    print("战新与未来产业月报第四期 - 完整格式校验报告")
    print("=" * 70)
    
    issues = full_review()
    
    # 分类统计
    errors = [i for i in issues if i['severity'] == 'error']
    warnings = [i for i in issues if i['severity'] == 'warning']
    infos = [i for i in issues if i['severity'] == 'info']
    
    print("")
    print("[统计汇总]")
    print("  文档总段落数: 536")
    print("  " + "-" * 50)
    print("  错误(中车新闻): {} 处".format(len(errors)))
    print("  警告(标题格式): {} 处".format(len(warnings)))
    print("  信息(含URL):    {} 处".format(len(infos)))
    
    # 输出各分类详情
    print("")
    print("=" * 70)
    print("[错误项] - 必须修复")
    print("=" * 70)
    if errors:
        for issue in errors:
            print("")
            print("  行{}: {}".format(issue['line'], issue['type']))
            print("    内容: {}".format(issue['content']))
            print("    关键词: {}".format(issue['word']))
    else:
        print("  [PASS] 无中车相关新闻")
    
    print("")
    print("=" * 70)
    print("[警告项] - 标题格式问题")
    print("=" * 70)
    if warnings:
        for issue in warnings:
            print("")
            print("  行{}: 序号后未加空格".format(issue['line']))
            print("    内容: {}".format(issue['content']))
    else:
        print("  [PASS] 标题格式正确")
    
    print("")
    print("=" * 70)
    print("[信息项] - URL链接")
    print("=" * 70)
    if infos:
        for issue in infos:
            print("")
            print("  行{}: 含URL".format(issue['line']))
            print("    内容: {}".format(issue['content']))
            print("    URL: {}".format(issue['urls']))
    else:
        print("  [PASS] 无外部URL")
    
    # 最终判定
    print("")
    print("=" * 70)
    print("[交付判定]")
    print("=" * 70)
    
    if errors:
        print("  [FAIL] 不可交付: 存在中车相关新闻")
    elif warnings:
        print("  [REVIEW] 需修改: 存在标题格式问题")
        print("     共 {} 处需要修复".format(len(warnings)))
        print("")
        print("  标题序号后需要加空格:")
        for issue in warnings:
            content = issue['content']
            match = re.match(r'^(\d+\.)(.+)', content)
            if match:
                print("    '{}' -> '{}'".format(
                    match.group(1) + match.group(2).lstrip(),
                    match.group(1) + ' ' + match.group(2).lstrip()
                ))
    elif infos:
        print("  [PASS] 可以交付 (URL为参考信息)")
    else:
        print("  [PASS] 可以交付")

if __name__ == '__main__':
    main()
