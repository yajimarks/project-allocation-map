"""CSV読込モジュール"""
import pandas as pd
from pathlib import Path


def read_csv(csv_path: Path, encoding: str = "cp932") -> pd.DataFrame:
    """社員情報CSVを読み込みDataFrameとして返す。

    Returns:
        DataFrame with columns:
            社員番号, 名前, 所属部署, 業務コード,
            ユーザー名, 取引先名, 業務名, 状況, 役職, グレード
    """
    df = pd.read_csv(csv_path, encoding=encoding, dtype=str)
    df = df.fillna("")
    # 前後の空白を除去
    for col in df.columns:
        df[col] = df[col].str.strip()
    return df
