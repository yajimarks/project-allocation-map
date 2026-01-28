"""構成図自動生成 - エントリーポイント"""
import sys
from pathlib import Path

# プロジェクトルートをパスに追加
sys.path.insert(0, str(Path(__file__).parent))

import config
from src.reader import read_csv
from src.processor import process
from src.writer import generate


def _find_csv(pattern: str) -> Path:
    """inputフォルダからパターンに一致するCSVを1件検索する。"""
    csv_files = sorted(config.CSV_DIR.glob(pattern))
    if not csv_files:
        print(f"エラー: {config.CSV_DIR} に {pattern} が見つかりません")
        sys.exit(1)
    if len(csv_files) > 1:
        print(f"エラー: {pattern} に該当するファイルが複数あります。1つにしてください:")
        for f in csv_files:
            print(f"  - {f.name}")
        sys.exit(1)
    return csv_files[0]


def main():
    # 1. CSV読込
    csv_path = _find_csv(config.CSV_PATTERN)
    print(f"CSV読込: {csv_path}")
    df = read_csv(csv_path, config.CSV_ENCODING)
    print(f"  {len(df)}件のレコードを読み込みました")

    # 2. データ加工
    divisions = process(df)
    for div in divisions:
        partner_names = ", ".join(p.display_name for p in div.partners)
        print(f"  営業:{div.key} {div.count}名 [{partner_names}]")

    # 3. Excel生成
    output_path = generate(divisions, config.OUTPUT_DIR)
    print(f"出力完了: {output_path}")


if __name__ == "__main__":
    main()
