"""構成図自動生成 - 設定ファイル"""
from pathlib import Path

# --- 入出力パス ---
CSV_DIR = Path(__file__).parent / "input"
CSV_PATTERN = "社員情報*.csv"
OUTPUT_DIR = Path(__file__).parent / "output"
CSV_ENCODING = "cp932"

# --- 取引先名 → 出力用表示名マッピング ---
# ※ キーは株式会社・㈱除去後の名前で登録する（自動除去されてから検索される）
PARTNER_DISPLAY_MAP = {
    "SCSK　Minoriソリューションズ": "Minoriソリューションズ",
    "TISW": "TIS西日本",
    "NTTデータ フィナンシャルテクノロジー": "NFT",
    "ジェーエムエーシステムズ": "JMAS",
    "日本電気": "NEC",
    "アドヴァンスト・インフォーメイション・デザイン": "AID",
    "オービーシステム": "OBS",
    "さくらケーシーエス": "さくらKCS",
    "さくら情報システム": "SIS",
    "シーイーシー": "CEC",
    "ソーシャルトランジットオフィス": "STO",
    "-": "本社",
}

# --- 担当営業 → 取引先マッピング ---
# ※ 値は出力用表示名（PARTNER_DISPLAY_MAP適用後の名前）で登録する
# ※ どの営業にも属さない取引先は営業区分なし（末尾）に配置される
SALES_PARTNER_MAP = {
    "中村": [
        "NSD", "TISソリューションリンク", "TIS西日本", "SCSK",
        "JR九州システムソリューションズ", "TMJ", "Minoriソリューションズ",
        "JMAS", "SRA", 
    ],
    "坂口": [
        "AID", "CLIS", "情報システム工学", "OBS", "ニーズウェル",
        "バルテス", "NEC", "東邦システムサイエンス", "コスモウェーブ",
        "JSOL", "リーディング・ウィン", "JMAS"
    ],
    "原田": [
        "TMJ", "さくらKCS", "CEC", "クリエイション", "アスリーブレインズ", "アルティウスリンク",
        "TOKAIコミュニケーションズ", "クロスキャット", "テイクス", "コベルコシステム",
        "セコムトラストシステムズ", "トラストシステム", "キャノン電子テクノロジー",
    ],
    "早川": [
        "DTS", "九州DTS", "ジャステック", "SIS",
        "フォーカスシステムズ", "富士ソフト", "STO", "USEN", "DTSインサイト","NFT",
    ],
    "野地": [
        "さつき工業協同組合",
    ],
}

# --- 顧客名（ユーザー名）名寄せマッピング ---
# ※ キーは株式会社・㈱除去後の名前で登録する（自動除去されてから検索される）
CLIENT_DISPLAY_MAP = {
    "シーイーシー": "CEC",
    "ヴェオリア・ジェネッツ": "ヴェオリアジェネッツ",
    # 例: "ABC表記": "ABC",
}

# --- グレード表示変換（CSV全角 → 半角、「なし」は非表示） ---
GRADE_DISPLAY_MAP = {
    "ＧＭ": "GM",
    "ＳＭ": "SM",
    "ＭＡ": "MA",
    "ＣＦ": "CF",
    "ＥＮ": "EN",
    "ＮＣ": "NC",
    "なし": "",
}

# --- グレード表示順（上が上位） ---
GRADE_ORDER = ["GM", "SM", "MA", "CF", "EN", "NC", ""]

# --- A4横レイアウト定数 ---
LAYOUT = {
    # A4横の印刷設定
    "paper_size": "A4",
    "orientation": "landscape",

    # カラム数（1ページあたり）
    "columns_per_page": 5,

    # 1カラムあたりの最大行数（A4横48%縮小 ≒ 90行）
    "max_rows_per_column": 90,

    # A列（左マージン空列）
    "col_width_margin": 2.33,

    # 1ビジュアルカラムの列構成: 取引先 | 顧客 | 案件名 | 名前 | 所属 | 空列 | グレード
    "col_width_partner": 1.56,
    "col_width_client": 1.56,
    "col_width_project": 1.56,
    "col_width_name": 10.33,
    "col_width_dept": 23.67,
    "col_width_empty": 3.22,
    "col_width_grade": 7.11,
    # カラム間の余白列幅
    "col_width_gap": 1.22,

    # フォント
    "font_title": {"name": "ＭＳ Ｐゴシック", "size": 11, "bold": True},
    "font_division": {"name": "ＭＳ Ｐゴシック", "size": 10, "bold": True},
    "font_partner": {"name": "ＭＳ Ｐゴシック", "size": 20, "bold": True, "italic": True},
    "font_client": {"name": "ＭＳ Ｐゴシック", "size": 14, "italic": True},
    "font_project": {"name": "ＭＳ Ｐゴシック", "size": 9, "bold": False},
    "font_person": {"name": "ＭＳ Ｐゴシック", "size": 9, "bold": False},
    "font_count": {"name": "ＭＳ Ｐゴシック", "size": 9, "bold": False},
}

# --- 構成図タイトル ---
# None の場合、実行日から自動生成（例: "R8年1月"）
CHART_TITLE_DATE = None
