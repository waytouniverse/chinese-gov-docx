#!/usr/bin/env python3
"""
解析PPT大纲，提取每页的标题和核心内容。
"""

import re
import sys
import json


def parse_outline(text):
    """解析Markdown大纲，返回页面列表"""
    pages = []
    lines = text.strip().split('\n')

    current_page = None
    content_buffer = []

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # 检测页面分隔
        if line == '---':
            if current_page:
                current_page['content'] = '\n'.join(content_buffer).strip()
                pages.append(current_page)
            current_page = None
            content_buffer = []
            continue

        # 检测页面标题 - 以"第X页"或一级/二级标题开头
        page_match = re.match(r'##第(\d+)页：(.+)', line) or \
                     re.match(r'##第(\d+)页[:：](.+)', line) or \
                     re.match(r'^##\s*(封面|尾页|Q\&A)', line)

        if page_match:
            if current_page:
                current_page['content'] = '\n'.join(content_buffer).strip()
                pages.append(current_page)
            if len(page_match.groups()) == 2:
                page_num = int(page_match.group(1))
                title = page_match.group(2).strip()
            else:
                page_num = len(pages) + 1
                title = page_match.group(1).strip()

            page_type = 'cover' if '封面' in title else ('ending' if any(w in title for w in ['Q&A', 'Q&A', '尾页', '感谢', '谢谢']) else 'interior')

            current_page = {
                'page_num': page_num,
                'title': title,
                'type': page_type,
                'content': ''
            }
            content_buffer = []
            continue

        # 收集内容
        if current_page and line:
            content_buffer.append(line)

    # 处理最后一个页面
    if current_page:
        current_page['content'] = '\n'.join(content_buffer).strip()
        pages.append(current_page)

    return pages


def infer_style(text):
    """根据大纲内容推断风格"""
    text_lower = text.lower()

    keywords = {
        '政务蓝': ['政务', '政府', '政策', '安全', '民生', '省市', '区县', '委员会', '数字化'],
        '科技蓝': ['AI', '人工智能', '科技', '架构', '算法', '开源', '智能体', '数据'],
        '商务金': ['金融', '银行', '投资', '风险', '财务', '营收'],
        '教育绿': ['教育', '学校', '课程', '培训', '教学', '学习'],
        '医疗白': ['医疗', '健康', '医院', '疫情', '疫苗'],
    }

    scores = {style: sum(1 for k in words if k in text_lower) for style, words in keywords.items()}
    best = max(scores, key=scores.get)

    if scores[best] == 0:
        return '通用商务'
    return best


def extract_key_points(content, max_points=5):
    """从内容中提取最关键的3-5个要点"""
    lines = [l.strip() for l in content.split('\n') if l.strip()]

    # 过滤标题行和空行
    points = []
    for line in lines:
        if line.startswith('#') or line.startswith('---'):
            continue
        # 去除markdown格式
        clean = re.sub(r'[\*\-`]', '', line).strip()
        if clean and len(clean) > 5:
            points.append(clean)

    return points[:max_points]


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python parse_outline.py <outline_file>")
        sys.exit(1)

    with open(sys.argv[1], 'r', encoding='utf-8') as f:
        text = f.read()

    pages = parse_outline(text)
    style = infer_style(text)

    result = {
        'style': style,
        'total_pages': len(pages),
        'pages': []
    }

    for page in pages:
        key_points = extract_key_points(page['content'])
        result['pages'].append({
            'page_num': page['page_num'],
            'title': page['title'],
            'type': page['type'],
            'key_points': key_points
        })

    print(json.dumps(result, ensure_ascii=False, indent=2))
