#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
対訳表 CSV版 生成スクリプト
"""
import csv
import re

EN_FILE = r"c:\Users\hyodo\Documents\claude\bilingual_text\sust-report2025_en_hybrid_xycut.txt"
JP_FILE = r"c:\Users\hyodo\Documents\claude\bilingual_text\sust-report2025_jp_hybrid_xycut.txt"
OUTPUT_FILE = r"c:\Users\hyodo\Documents\claude\bilingual_text\bilingual_table.csv"

PAGE_SHIFT_START_EN = 141
PAGE_SHIFT_FULL_EN = 147

def read_lines(path):
    with open(path, 'r', encoding='utf-8') as f:
        return [l.rstrip('\n') for l in f.readlines()]

def parse_pages(lines):
    pages = []
    current_page = None
    start = 0
    for i, l in enumerate(lines):
        m = re.match(r'<<<PAGE (\d+)>>>', l)
        if m:
            if current_page is not None:
                pages.append((current_page, start, i))
            current_page = int(m.group(1))
            start = i
    if current_page is not None:
        pages.append((current_page, start, len(lines)))
    return pages

def get_content(lines, start, end):
    result = []
    for i in range(start, end):
        l = lines[i]
        if not re.match(r'<<<PAGE \d+>>>', l):
            result.append(l)
    return result

def is_heading_en(line):
    if len(line) > 80 and not line.endswith('.'):
        return False
    heading_keywords = [
        'Contents', 'Editorial Policy', 'CEO', 'CSO', 'Key Figures', 'Highlight Topics',
        'Sustainability Management', 'Medium-', 'The Asahi', 'Material Issues',
        'Environment', 'Communities', 'Responsible Drinking', 'Health',
        'Human Rights', 'Enhancement of', 'Data Section', 'Management', 'Activities',
        'Approach', 'Outline', 'Overview', 'Key', 'Main', 'Corrective', 'Wage',
        'Heightened', 'Efforts', 'Discrimination', 'Initiatives', 'Practices',
        'Our Progress', 'Grievance', 'Engagement', 'Continuation', 'Effectiveness',
        'Framework', 'Policy', 'Targets', 'Definition', 'Progress', 'Performance',
        'Roadmap', 'Related Information', 'Regional', 'On-Site', 'Child Labor',
        'Freedom of', 'Occupational', 'Working Hours', 'Living Wages',
        'Percentage of', 'Supply Chain', 'Impacts', 'Due Diligence',
        'Reporting Channel', 'Strengthening', 'Fair Access', 'The System',
        'How "Speak', 'Receipt', 'Initial', 'Investigation', 'Feedback',
        'Gaps and', 'Examples of', 'Considerations', 'Addressing', 'Limiting',
        'Retention of', 'Fair Payment', 'Management of', 'Internal review',
        'Age Verification', 'Prohibition of', 'Controlling', 'Basic Principles',
        'Solutions of', 'Create Value', 'Respect for', 'Engagement with',
        'Fostering the', 'Continuously', 'Securing Essential', 'Respond to',
        'Sustainable Packaging', 'Natural Capital', 'Reduction of', 'Green Bonds',
        'Realize Sustainable', 'Reduction in', 'Visualizing', 'Stakeholder',
        'Governance', 'Risk Management', 'Metrics and', 'Editorial',
        'Voluntary Disclosure', 'Statutory Disclosure', 'Secondary', 'Primary',
        'Third-Party', 'Forward-Looking', 'About Stated', 'Referenced',
        'Sustainability Report', 'Integrated Report', 'Securities Report',
        'Corporate Governance', 'Convocation', 'Human Capital Report',
        'Details from', 'Respond to Climate', 'Scope of Reports',
        'Report Period', 'SASB Content Index', 'GRI Standard',
        'Activity Metric', 'Topic Activity',
        'Suppliers’ social', 'Environmental and', 'Social Impacts',
        'Percentage of beverage', 'Ingredient Sourcing',
        'Volume of products', 'Number of production', 'Total fleet',
        'Asahi GROUP', 'Asahi Group Holdings',
    ]
    for kw in heading_keywords:
        if line.startswith(kw):
            return True
    # Short line without period
    if len(line) <= 60 and not line.endswith('.') and not line.startswith('•'):
        return True
    return False

def split_sentences(text):
    result = []
    # Split by sentence boundaries
    parts = re.split(r'(?<=\.)\s+(?=[A-Z"\'(])', text)
    for p in parts:
        p = p.strip()
        if p:
            result.append(p)
    return result if result else [text.strip()]

def process_en_page(content):
    entries = []
    for l in content:
        l_stripped = l.strip()
        if not l_stripped:
            continue
        if re.match(r'^\d+$', l_stripped):
            continue
        if l_stripped.startswith('［'):
            continue
        if re.match(r'^\s*\|', l_stripped):
            # Table data
            entries.append(('段落', l_stripped))
            continue
        if l_stripped.startswith('•') or l_stripped.startswith('-'):
            entries.append(('段落', l_stripped))
            continue
        if '|' in l_stripped and len(l_stripped) < 100:
            entries.append(('段落', l_stripped))
            continue

        if is_heading_en(l_stripped):
            clean = re.sub(r'\s+', ' ', l_stripped).strip()
            entries.append(('見出し', clean))
        else:
            sents = split_sentences(l_stripped)
            for s in sents:
                if s:
                    entries.append(('段落', s))
    return entries

def get_jp_lines(content):
    result = []
    for l in content:
        l_stripped = l.strip()
        if not l_stripped:
            continue
        if re.match(r'^\d+$', l_stripped):
            continue
        if l_stripped.startswith('［'):
            continue
        result.append(l_stripped)
    return result

def main():
    print("読み込み中...")
    en_lines = read_lines(EN_FILE)
    jp_lines = read_lines(JP_FILE)

    en_pages = parse_pages(en_lines)
    jp_pages = parse_pages(jp_lines)

    print(f"英語: {len(en_pages)}ページ, 日本語: {len(jp_pages)}ページ")

    jp_map = {}
    for pn, start, end in jp_pages:
        jp_map[pn] = (start, end)

    rows = []
    seq = 0

    for en_idx, (en_pn, en_start, en_end) in enumerate(en_pages):
        # 対応する日本語ページ番号
        if en_pn < PAGE_SHIFT_START_EN:
            jp_pn = en_pn
        elif en_pn >= PAGE_SHIFT_FULL_EN:
            jp_pn = en_pn + 1
        else:
            jp_pn = en_pn

        # 日本語内容を取得
        jp_content = []
        if jp_pn in jp_map:
            js, je = jp_map[jp_pn]
            jp_content = get_content(jp_lines, js, je)

        # 英語内容を取得
        en_content = get_content(en_lines, en_start, en_end)

        # 英語エントリー
        en_entries = process_en_page(en_content)

        # 日本語行
        jp_texts = get_jp_lines(jp_content)

        if len(en_entries) == 0:
            continue

        # マッチング
        jp_idx = 0
        jp_total = len(jp_texts)

        for typ, en_text in en_entries:
            seq += 1
            if jp_idx < jp_total:
                jp_t = jp_texts[jp_idx]
                jp_idx += 1
            else:
                jp_t = ''
            rows.append([seq, typ, en_text, jp_t])

        if en_idx % 20 == 0:
            print(f"  ページ {en_pn} 処理完了 (累計 {seq})")

    print(f"全 {seq} エントリーをCSVに出力中...")

    with open(OUTPUT_FILE, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['通し番号', '種別', '英語', '日本語'])
        writer.writerows(rows)

    print(f"完了！ {OUTPUT_FILE} に保存しました。")

if __name__ == '__main__':
    main()
