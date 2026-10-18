#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
企業サステナビリティレポート 対訳表 生成スクリプト
英語版と日本語版のテキストファイルから対訳Excelを生成する
"""

import re
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ===== 設定 =====
EN_FILE = r"c:\Users\hyodo\Documents\claude\bilingual_text\sust-report2025_en_hybrid_xycut.txt"
JP_FILE = r"c:\Users\hyodo\Documents\claude\bilingual_text\sust-report2025_jp_hybrid_xycut.txt"
OUTPUT_FILE = r"c:\Users\hyodo\Documents\claude\bilingual_text\bilingual_table.xlsx"

# ページずれ設定：
# 英語版141ページから少しずつずれ始め、147ページ（日本語版148ページ）で完全に1ページ遅れ
PAGE_SHIFT_START_EN = 141  # ずれが始まる英語ページ番号
PAGE_SHIFT_FULL_EN = 147   # 完全に1ページ遅れになる英語ページ番号

def parse_file(filepath):
    """ファイルを読み込み、ページ構造を解析する。
    戻り値: [(page_num, start_line, text_lines), ...]
    """
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # 行番号を付加
    numbered = [(i + 1, l.rstrip('\n')) for i, l in enumerate(lines)]

    # ページ区切りを検出
    pages = []
    current_page = None
    current_start = None

    for lineno, text in numbered:
        m = re.match(r'<<<PAGE (\d+)>>>', text)
        if m:
            if current_page is not None:
                pages.append((current_page, current_start, lineno))
            current_page = int(m.group(1))
            current_start = lineno

    # 最終ページ
    if current_page is not None:
        pages.append((current_page, current_start, len(lines)))

    return pages, lines

def get_page_content(lines, start_line, end_line):
    """指定されたページ範囲のテキストを取得する（ページマーカー行を除く）"""
    content = []
    for i in range(start_line, end_line):
        line = lines[i - 1].rstrip('\n')  # start_lineは1-indexed
        if not re.match(r'<<<PAGE \d+>>>', line):
            content.append(line)
    return content

def classify_and_split_en(content_lines):
    """英語のページ内容を解析し、(type, text)のリストを返す
    - type: '段落' or '見出し'
    - 段落は一文ごとに分割、見出しは意味的なまとまりで一文に
    """
    results = []

    for line in content_lines:
        line = line.strip()
        if not line:
            continue

        # 空行でない場合、見出しと段落を判別
        # 見出しの特徴: 行が短い、大文字で始まる、ページ番号＋タイトル形式、文末にピリオドがない
        # 段落: 複数の文からなる

        # まず見出しらしいパターンをチェック
        is_heading = False

        # 目次・セクション番号で始まる見出し
        if re.match(r'^(?:Contents|Editorial|CEO|Key Figures|Highlight|Sustainability|Material Issues|Environment|Communities|Responsible|Health|Human Rights|Enhancement|Data Section|Respond|Sustainable|Natural|Reduction|Green Bonds|Basic Principles|Solutions|Create Value|Respect|Reporting Channel|Engagement|Fostering|Continuously|Securing|Medium-|The Asahi|Visualizing|Stakeholder|Governance|Risk Management|Metrics|Management|Activities|Approach|Outline|Overview|Key|Main|Corrective|Wage|Heightened|Efforts|Discrimination|Initiatives|Our|Grievance|Engagement|Continuation|The Effectiveness|Framework|Policy|Targets|Definition|Progress|Performance|Roadmap|Initiatives|Practices|Related|Regional|On-Site|Child|Freedom|Occupational|Working|Living|Percentage|Supply Chain|Impacts|Due Diligence|Reporting|Strengthening|Fair|The System|How|Receipt|Initial|Investigation|Feedback|Gaps|Examples|Considerations|Addressing|Limiting|Retention|Fair|Management|Internal|Age|Prohibition|Controlling)', line):
            is_heading = True

        # 短い行（60文字以下）で、末尾がピリオドでない、かつ文の構造を持たない
        if len(line) <= 80 and not line.endswith('.') and not line.endswith('。'):
            # ただし、箇条書き（・で始まる、数字で始まる）は段落扱い
            if line.startswith('•') or line.startswith('-') or line.startswith('|'):
                is_heading = False
            else:
                is_heading = True

        # テーブルデータなどは段落として扱う
        if '|' in line and ('Score' in line or 'Number' in line or 'Country' in line):
            results.append(('段落', line))
            continue

        # 箇条書きの項目
        if line.startswith('•') or line.startswith('-') or line.startswith('|'):
            results.append(('段落', line))
            continue

        # 数字だけで構成される行（ページ番号など）
        if re.match(r'^\d+$', line):
            continue

        # 図中テキスト
        if line.startswith('［'):
            continue

        # 表のデータ行
        if re.match(r'^[A-Za-z\s]+\|\d+', line):
            continue

        if is_heading:
            # 見出し：意味的なまとまり。空白を削除して一文に
            clean = re.sub(r'\s+', ' ', line).strip()
            # 複数行の見出しマージはここでは行わない（後処理で行う）
            results.append(('見出し', clean))
        else:
            # 段落：一文ごとに分割
            sentences = split_paragraph(line)
            for s in sentences:
                if s.strip():
                    results.append(('段落', s.strip()))

    return results

def split_paragraph(text):
    """英語の段落を一文ごとに分割する"""
    # 文の区切り（ピリオド、疑問符、感嘆符）で分割
    # ただし、略語（e.g., i.e., Dr., Mr., Ms., etc.）や数字の小数点は区切らない
    # 高度な分割を行う

    # まず文の区切り位置を特定
    parts = []

    # 簡易版の文分割
    # 正規表現で文末を検出（ピリオド+大文字/引用符/空白行）
    sentences = re.split(r'(?<=\.)\s+(?=[A-Z])', text)

    # さらに細かく分割（ピリオド、感嘆符、疑問符）
    result = []
    for s in sentences:
        # 複数文が含まれている可能性があるのでさらに分割
        sub = re.split(r'(?<=\.)(?:\s+|$)(?=[A-Z"])', s)
        for ss in sub:
            ss = ss.strip()
            if ss:
                result.append(ss)

    return result if result else [text]

def get_jp_text(content_lines):
    """日本語のページ内容を取得し、空行を除いたテキスト行を返す"""
    results = []
    for line in content_lines:
        line = line.strip()
        if line and not re.match(r'^\d+$', line) and not line.startswith('［'):
            results.append(line)
    return results

def find_page_alignment(pages_en, pages_jp):
    """英語と日本語のページ対応関係を計算する。
    ページ番号のずれを考慮する。
    """
    alignment = []  # [(en_page_num, jp_page_num, shift), ...]

    for en_pn, _, _ in pages_en:
        if en_pn < PAGE_SHIFT_START_EN:
            # ずれなし：日本語の同じページ番号と対応
            jp_pn = en_pn
            shift = 0
        elif en_pn >= PAGE_SHIFT_FULL_EN:
            # 完全に1ページ遅れ
            jp_pn = en_pn + 1
            shift = 1
        else:
            # ずれが徐々に発生している領域
            # 141: ずれ始め、147で完全に1ページ遅れ
            # 線形補間
            progress = (en_pn - PAGE_SHIFT_START_EN) / (PAGE_SHIFT_FULL_EN - PAGE_SHIFT_START_EN)
            shift = progress  # 0から1の間
            jp_pn = en_pn + shift
            # 実際の日本語ページ番号は整数
            jp_pn = en_pn + (1 if shift > 0.5 else 0)
            # 簡易版：141-146は同じページ番号（ずれは小さい）、147以降は+1
            if en_pn < PAGE_SHIFT_FULL_EN:
                jp_pn = en_pn
            else:
                jp_pn = en_pn + 1

        alignment.append((en_pn, jp_pn, shift))

    return alignment

def match_content_en_jp(en_entries, jp_lines, page_en, page_jp):
    """英語のエントリーと日本語のテキスト行をページ内でマッチングする。
    完全な自動マッチングは困難なので、ページ単位で英語の順序と日本語の順序を
    なるべく対応させる。

    ここでは、ページ内で英語のエントリー数と日本語の行数が大きく異なる場合、
    日本語行をマージまたは分割する。
    """
    # 簡易版：英語のエントリー数と日本語の行数が近い場合は1:1対応
    # 英語が一文ごと、日本語がより細かい単位の場合、日本語をマージする

    matched = []

    # 英語のエントリーをタイプ別に処理
    # 見出しは見出し同士、段落は段落同士でマッチング

    # まず日本語行をグループ化する
    # 英語の見出しと日本語の見出しをキーワードでマッチング

    # よりシンプルなアプローチ：
    # 英語の各エントリーに対して、対応する日本語テキストを
    # キーワード/位置ベースで割り当てる

    # ここでは、ページ単位で英語のエントリー数に合わせて日本語をグループ化する

    # 日本語の行を英語のエントリー数に合わせて再グループ化
    num_en = len(en_entries)

    if num_en == 0:
        return []

    if len(jp_lines) == 0:
        # 日本語がない場合
        for i, (typ, en_text) in enumerate(en_entries):
            matched.append((typ, en_text, ''))
        return matched

    # 英語のエントリーを順に処理し、日本語を割り当てる
    jp_idx = 0
    jp_total = len(jp_lines)

    for i, (typ, en_text) in enumerate(en_entries):
        # 各英語エントリーに割り当てる日本語行数の目安
        # 英語の段落（一文）には日本語1行程度
        # 英語の見出しには日本語の見出し行を割り当て

        if jp_idx >= jp_total:
            # 日本語が足りない
            matched.append((typ, en_text, ''))
            continue

        if typ == '見出し':
            # 見出しは日本語の対応する見出しを探す
            # 現在のjp_idxから日本語を1行取得
            jp_text = jp_lines[jp_idx]
            jp_idx += 1
            matched.append((typ, en_text, jp_text))
        else:
            # 段落の一文に対して、日本語を1行割り当て（または空）
            jp_text = jp_lines[jp_idx]
            jp_idx += 1
            matched.append((typ, en_text, jp_text))

    return matched


def main():
    print("ファイルを読み込み中...")
    pages_en, lines_en = parse_file(EN_FILE)
    pages_jp, lines_jp = parse_file(JP_FILE)

    print(f"英語: {len(pages_en)}ページ, 日本語: {len(pages_jp)}ページ")

    # ページ対応関係を計算
    alignment = find_page_alignment(pages_en, pages_jp)

    # Excelブックを作成
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "対訳表"

    # スタイル設定
    header_font = Font(bold=True, size=11)
    header_align = Alignment(horizontal='center', vertical='center')
    cell_font = Font(size=10)
    cell_align = Alignment(vertical='top', wrap_text=True)

    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # ヘッダー
    headers = ['通し番号', '種別', '英語', '日本語']
    ws.column_widths = [8, 10, 80, 80]

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.alignment = header_align
        cell.border = thin_border

    seq = 0
    row = 2

    # 英語のページインデックスと日本語のページインデックスをマップ
    # pages_en と pages_jp は (page_num, start_line, end_line) のリスト
    jp_page_map = {}
    for pn, start, end in pages_jp:
        jp_page_map[pn] = (start, end)

    # ずれが発生する領域の処理
    # 英語ページ141までは同じページ番号で対応
    # 141以降の処理

    for en_idx, (en_pn, en_start, en_end) in enumerate(pages_en):
        # 対応する日本語ページ番号を決定
        if en_pn < PAGE_SHIFT_START_EN:
            jp_pn = en_pn
        elif en_pn >= PAGE_SHIFT_FULL_EN:
            jp_pn = en_pn + 1
        else:
            # 141-146: 同じページ番号（ずれは軽微なのでページ番号で合わせる）
            jp_pn = en_pn

        # 日本語ページの内容を取得
        jp_content = []
        if jp_pn in jp_page_map:
            jp_start, jp_end = jp_page_map[jp_pn]
            jp_content = get_page_content(lines_jp, jp_start, jp_end)

        # 英語ページの内容を取得
        en_content = get_page_content(lines_en, en_start, en_end)

        # 英語のエントリーを生成
        en_entries = classify_and_split_en(en_content)

        # 英語と日本語をマッチング
        matched = match_content_en_jp(en_entries, jp_content, en_pn, jp_pn)

        # 結果をExcelに書き込み
        for typ, en_text, jp_text in matched:
            seq += 1
            ws.cell(row=row, column=1, value=seq).font = cell_font
            ws.cell(row=row, column=1).alignment = Alignment(vertical='top')
            ws.cell(row=row, column=1).border = thin_border

            ws.cell(row=row, column=2, value=typ).font = cell_font
            ws.cell(row=row, column=2).alignment = Alignment(vertical='top', horizontal='center')
            ws.cell(row=row, column=2).border = thin_border

            ws.cell(row=row, column=3, value=en_text).font = cell_font
            ws.cell(row=row, column=3).alignment = cell_align
            ws.cell(row=row, column=3).border = thin_border

            ws.cell(row=row, column=4, value=jp_text).font = cell_font
            ws.cell(row=row, column=4).alignment = cell_align
            ws.cell(row=row, column=4).border = thin_border

            row += 1

        if en_idx % 20 == 0:
            print(f"  ページ {en_pn} 処理中... (累計 {seq} エントリー)")

    # 列幅を調整
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 10
    ws.column_dimensions['C'].width = 80
    ws.column_dimensions['D'].width = 80

    wb.save(OUTPUT_FILE)
    print(f"\n完了！ {seq} エントリーを {OUTPUT_FILE} に保存しました。")


if __name__ == '__main__':
    main()
