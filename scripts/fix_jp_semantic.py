#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
bilingual_table.csv の日本語照合を意味的類似度で改善するスクリプト。

【処理の流れ】
1. ページごとに英語行と日本語行を取得
2. 英語行インデックス < 日本語行数 → 従来通り1:1照合（変更なし）
3. 英語行インデックス >= 日本語行数 → paraphrase-multilingual-MiniLM で
   コサイン類似度が最大の日本語行を選択（意味的照合）
"""

import csv
import re
from sentence_transformers import SentenceTransformer, util

EN_SRC  = r'c:\Users\hyodo\Documents\claude\bilingual_text\sust-report2025_en_hybrid_xycut.txt'
JP_SRC  = r'c:\Users\hyodo\Documents\claude\bilingual_text\sust-report2025_jp_hybrid_xycut.txt'
SRC_CSV = r'c:\Users\hyodo\Documents\claude\bilingual_text\bilingual_table.csv'
DST_CSV = r'c:\Users\hyodo\Documents\claude\bilingual_text\bilingual_table_v4.csv'

MODEL_NAME = 'paraphrase-multilingual-MiniLM-L12-v2'


# ─── ユーティリティ ────────────────────────────────────────

def parse_pages(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        text = f.read()
    parts = re.split(r'<<<PAGE (\d+)>>>', text)
    pages = {}
    for i in range(1, len(parts), 2):
        pages[int(parts[i])] = parts[i + 1].strip()
    return pages


def extract_lines(page_content):
    result = []
    for line in page_content.split('\n'):
        line = line.strip()
        if not line:
            continue
        if line.startswith('［') or line.startswith('【') or line.startswith('図'):
            continue
        if re.match(r'^\d+$', line):
            continue
        result.append(line)
    return result


def is_heading(text):
    t = text.strip()
    if not t:
        return False
    if len(t) <= 60 and not t.endswith('.') and not t.endswith('。'):
        return True
    heading_starts = [
        'Environment', 'Management', 'Governance', 'Strategies', 'Approach',
        'Activities', 'Related Information', 'Initiatives', 'Future Initiatives',
        'Risk Management', 'Climate Change', 'Sustainable', 'Packaging',
        'Water', 'Agricultural', 'Natural Capital', 'Biodiversity',
        'Communities', 'Responsible Drinking', 'Health', 'Human Rights',
        'Enhancement', 'Data Section', 'GRI', 'SASB',
        '環境', 'マネジメント', 'ガバナンス', '戦略', '考え方', '取り組み',
        '関連情報', '今後の取り組み', 'リスク管理', '気候変動', '持続可能',
        'コミュニティ', '責任ある飲酒', '健康', '人権', '人的資本',
        'データセクション',
    ]
    for kw in heading_starts:
        if t.startswith(kw):
            return True
    return False


def split_sentences(text):
    sentences = re.split(r'(?<=\.)\s+(?=[A-Z"\'(])', text)
    return [s.strip() for s in sentences if s.strip()]


def jp_alignment(en_page_num):
    if en_page_num <= 140:
        return en_page_num
    elif en_page_num == 141:
        return 141
    elif en_page_num == 142:
        return 141
    elif en_page_num == 143:
        return 142
    elif en_page_num == 144:
        return 143
    elif en_page_num == 145:
        return 144
    elif en_page_num == 146:
        return 145
    else:
        return en_page_num - 1


# ─── メイン ───────────────────────────────────────────────

def main():
    print(f"モデルをロード中: {MODEL_NAME}")
    model = SentenceTransformer(MODEL_NAME)
    print("  ロード完了")

    print("ソースファイルを解析中...")
    en_pages = parse_pages(EN_SRC)
    jp_pages = parse_pages(JP_SRC)
    print(f"  英語: {len(en_pages)}ページ / 日本語: {len(jp_pages)}ページ")

    with open(SRC_CSV, 'r', encoding='utf-8', newline='') as f:
        original_rows = list(csv.reader(f))
    header = original_rows[0]
    base_rows = original_rows[1:650]  # ページ1-36（seq 1-649）はそのまま流用
    print(f"  ページ1-36: {len(base_rows)}行を流用")

    new_rows = []
    seq = 649
    semantic_count = 0   # 意味的照合で埋めた行数
    exact_count = 0      # 1:1照合の行数
    empty_count = 0      # 日本語なし（JP ページ自体が存在しない）

    for en_page_num in range(37, 201):
        jp_page_num = jp_alignment(en_page_num)

        if en_page_num not in en_pages or jp_page_num not in jp_pages:
            continue

        en_lines = extract_lines(en_pages[en_page_num])
        jp_lines = extract_lines(jp_pages[jp_page_num])

        if len(en_lines) < 2:
            continue

        # JP 行が存在する場合のみ埋め込みを計算
        jp_embeddings = None
        if jp_lines:
            jp_embeddings = model.encode(jp_lines, convert_to_tensor=True, show_progress_bar=False)

        for i, en_line in enumerate(en_lines):
            if i < len(jp_lines):
                # 1:1 照合（変更なし）
                jp_text = jp_lines[i]
                exact_count += 1
            elif jp_embeddings is not None:
                # 意味的照合：この英語行に最も近い日本語行を探す
                query = en_line
                en_emb = model.encode(query, convert_to_tensor=True, show_progress_bar=False)
                sims = util.cos_sim(en_emb, jp_embeddings)[0]
                best_idx = int(sims.argmax())
                jp_text = jp_lines[best_idx]
                semantic_count += 1
            else:
                jp_text = ''
                empty_count += 1

            if is_heading(en_line):
                seq += 1
                new_rows.append([seq, '見出し', en_line, jp_text])
            else:
                sentences = split_sentences(en_line)
                for sentence in sentences:
                    seq += 1
                    new_rows.append([seq, '段落', sentence, jp_text])

        if en_page_num % 20 == 0:
            print(f"  ページ {en_page_num} 完了 (累計 {len(new_rows)} 行, 意味的照合: {semantic_count}行)")

    print(f"\n新規行数: {len(new_rows)}")
    print(f"  1:1照合: {exact_count} 行")
    print(f"  意味的照合: {semantic_count} 行")
    print(f"  空欄: {empty_count} 行")

    with open(DST_CSV, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(base_rows)
        writer.writerows(new_rows)

    total = len(base_rows) + len(new_rows)
    with_jp = sum(1 for r in base_rows + new_rows if len(r) > 3 and r[3].strip() != '')
    print(f"\n書き出し完了: {DST_CSV}")
    print(f"  総データ行: {total}")
    print(f"  日本語あり: {with_jp} ({with_jp * 100 // total}%)")


if __name__ == '__main__':
    main()
