# 構成図自動生成ツール

社員情報CSVから、組織の案件配置構成図をExcelで自動生成するツール。

## 概要

営業区分（A〜E）ごとに取引先・顧客・案件・社員を階層構造で整理し、A4横の印刷用Excelファイルとして出力する。

### 出力構成図の階層

```
営業区分（A → B → C → D → E → 本社/その他）
  └ 取引先（人数降順）
      └ 顧客（人数降順）
          └ 案件（人数降順）
              └ 社員（自社グレード順 → BP）
```

### レイアウト特徴

- 1ページあたり最大5カラムを横に並べるフローレイアウト
- 顧客ブロック単位でカラムをまたぎ、途中で切れない
- カラムをまたぐ場合は取引先名を再出力
- 5カラムを超える場合は2ページ目に継続
- 印刷設定: A4横、48%縮小、改ページプレビュー

## フォルダ構成

```
project-allocation-map/
├── main.py           # エントリポイント
├── config.py         # 設定・マッピング定義
├── requirements.txt  # 依存パッケージ
├── input/            # 入力CSV配置フォルダ
│   └── 社員情報*.csv
├── output/           # 出力Excel格納フォルダ
└── src/
    ├── reader.py     # CSV読み込み
    ├── processor.py  # データ加工・階層構造化・ソート
    └── writer.py     # Excel出力（書式・罫線・印刷設定）
```

## 環境構築

### 1. Pythonインストール（Windows）

PowerShell またはコマンドプロンプトを開いて以下を実行:

```
winget install Python.Python.3.12
```

インストール後、ターミナル（Git Bash等）を再起動する。

### 2. 依存パッケージインストール

```bash
cd /c/work/aimarks/PRJ02_project-allocation/project-allocation-map
pip install -r requirements.txt
```

または直接指定:

```bash
pip install pandas openpyxl
```

## 実行方法

1. `input/` フォルダに `社員情報*.csv`（cp932エンコーディング）を配置
2. 以下を実行:

```bash
python main.py
```

3. `output/` フォルダに `構成図_YYYYMMDD_HHMMSS.xlsx` が生成される

## 設定変更

`config.py` で以下を変更可能:

| 設定項目 | 説明 |
|---|---|
| `SALES_DIVISION_MAP` | 取引先名 → 営業区分のマッピング |
| `GRADE_ORDER` | グレード表示順（GM → SM → MA → CF → EN → NC） |
| `LAYOUT["columns_per_page"]` | 1ページあたりのカラム数（既定: 5） |
| `LAYOUT["max_rows_per_column"]` | 1カラムあたりの最大行数（既定: 90） |
| `CHART_TITLE_DATE` | タイトル日付（Noneで自動生成: R○年○月） |
