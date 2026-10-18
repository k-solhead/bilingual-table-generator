# 対訳表生成ツール

英語・日本語テキストファイルを入力として、センテンス単位に分割した対訳表 Excel を生成する Streamlit アプリです。統合報告書などの英日 PDF から抽出したテキストの対訳作成を想定しています。

## 機能

- `<<<PAGE N>>>` マーカー付きテキストファイルを入力
- 英文をセンテンス単位に分割し、対応する日本語を対訳として割り当て
- 多言語意味的照合（`paraphrase-multilingual-MiniLM-L12-v2`）で高精度マッチング
- ページオフセット設定（英語と日本語でページ番号がずれている場合）
- 見出し・段落を自動判定
- Excel（.xlsx）形式でダウンロード（見出し行は太字・背景色付き）

## セットアップ

### 必要環境

- Python 3.10 以上

### インストール

```bash
# 仮想環境の作成と有効化
python -m venv .venv

# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

# 依存パッケージのインストール
pip install -r requirements.txt
```

### 起動

```bash
streamlit run bilingual_app.py
```

ブラウザで `http://localhost:8501` を開きます。

## 使い方

1. サイドバーから英語・日本語テキストファイルをアップロード
2. ページオフセットが必要な場合は設定（例：英語 147 ページ以降は日本語が 1 ページ前）
3. 「▶ 対訳表を生成」ボタンを押す
4. プレビューを確認し、Excel ファイルをダウンロード

## 入力ファイル形式

テキストファイルに `<<<PAGE N>>>` マーカーを含めてください。

```
<<<PAGE 1>>>
Introduction
This report describes our sustainability initiatives for fiscal year 2025.

<<<PAGE 2>>>
Environment
We reduced CO2 emissions by 15% compared to the previous year.
Our target is to achieve carbon neutrality by 2050.
```

## 出力 Excel の形式

| 列 | 内容 |
|---|---|
| A：通し番号 | 連番 |
| B：種別 | 見出し／段落 |
| C：英語 | センテンス単位の英文 |
| D：日本語 | 対応する日本語テキスト |

## 技術仕様

| 項目 | 内容 |
|---|---|
| 意味的照合モデル | `paraphrase-multilingual-MiniLM-L12-v2` |
| センテンス分割 | 正規表現（`. ` + 大文字で分割） |
| 見出し判定 | 行長 60 字以下・句点なし・キーワード一致 |
| マッチング方式 | 1:1 照合 + 範囲外は意味的照合（コサイン類似度） |

## ディレクトリ構成

```
bilingual-table-generator/
├── bilingual_app.py          # Streamlit アプリ本体
├── scripts/
│   ├── create_bilingual_csv.py   # CSV 生成スクリプト（バッチ処理用）
│   ├── create_bilingual_table.py # 対訳表構築スクリプト
│   └── fix_jp_semantic.py        # 意味的照合改善スクリプト
├── requirements.txt
├── .gitignore
└── README.md
```
