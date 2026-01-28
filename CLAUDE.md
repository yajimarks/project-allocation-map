# プロジェクト概要

**構成図自動生成ツール** — 社員情報CSVから、組織の案件配置構成図をExcelで自動生成するPythonアプリケーション。

## 構成

```
project-allocation-map/
├── main.py           # エントリポイント
├── config.py         # 設定・マッピング定義
├── requirements.txt  # 依存パッケージ (openpyxl, pandas)
└── src/
    ├── reader.py     # CSV読み込み
    ├── processor.py  # データ加工・階層構造化
    └── writer.py     # Excel出力（書式設定含む）
```

## 処理フロー

1. **CSV読込** (`reader.py`) — `cp932`エンコーディングで社員情報CSVを読み込み
2. **データ加工** (`processor.py`) — 営業区分(A〜E) → 取引先 → 顧客 → 案件 → 社員 の階層構造に変換。グレード順ソートやBP判定も実施
3. **Excel出力** (`writer.py`) — A3横向きの構成図をExcelで生成。罫線・フォント・セル結合など詳細な書式設定付き

## 技術スタック

- **Python 3.x**
- **pandas** — CSV読み込み・データ処理
- **openpyxl** — Excel生成

## 主な特徴

- 40社以上の取引先を営業区分(A〜E)にマッピング
- グレード階層（GM → SM → MA → CF → EN → NC）による自社/BP社員のソート
- A3横印刷に最適化された段組レイアウト（5カラム/ページ）
- 出力ファイル名に和暦（令和）・タイムスタンプ付与
\r