"""データ加工モジュール - グルーピング・集計・ソート"""
import pandas as pd
from dataclasses import dataclass, field

import config


@dataclass
class Member:
    """社員1名分の情報"""
    name: str
    dept: str
    grade: str
    is_bp: bool


@dataclass
class Project:
    """案件（業務）"""
    name: str
    members: list[Member] = field(default_factory=list)

    @property
    def count(self) -> int:
        return len(self.members)

    def row_height(self) -> int:
        """案件名行 + メンバー行数。"""
        return 1 + self.count


@dataclass
class Client:
    """顧客（ユーザー）"""
    name: str
    projects: list[Project] = field(default_factory=list)

    @property
    def count(self) -> int:
        return sum(p.count for p in self.projects)

    def row_height(self) -> int:
        """この顧客ブロックがExcel上で占める行数を返す。

        顧客名行 + 空行 + 案件ブロック群。
        """
        project_rows = sum(p.row_height() for p in self.projects)
        return 1 + 1 + project_rows


@dataclass
class Partner:
    """取引先（出力用表示名で集約）"""
    display_name: str
    clients: list[Client] = field(default_factory=list)

    @property
    def count(self) -> int:
        return sum(c.count for c in self.clients)

    def row_height(self) -> int:
        """この取引先ブロック全体がExcel上で占める行数を返す。

        取引先名行 + 空行 + 顧客ブロック群。
        """
        client_rows = sum(c.row_height() for c in self.clients)
        return 1 + 1 + client_rows


@dataclass
class Division:
    """営業区分（A, B, C, ...）"""
    key: str
    partners: list[Partner] = field(default_factory=list)

    @property
    def count(self) -> int:
        return sum(p.count for p in self.partners)


# 全角英数字 → 半角変換テーブル
_ZEN2HAN = str.maketrans(
    {chr(c): chr(c - 0xFEE0)
     for c in range(0xFF01, 0xFF5F)
     if 0xFF10 <= c <= 0xFF19      # ０-９
     or 0xFF21 <= c <= 0xFF3A      # Ａ-Ｚ
     or 0xFF41 <= c <= 0xFF5A}     # ａ-ｚ
)


def _strip_company(name: str) -> str:
    """会社名から株式会社・㈱を除去し、全角英数字を半角に変換する。"""
    for s in ("株式会社", "㈱"):
        name = name.replace(s, "")
    name = name.translate(_ZEN2HAN)
    return name.strip()


def _resolve_display_name(partner_name: str) -> str:
    """取引先名から出力用表示名を返す。

    株式会社・㈱を除去してからマッピングを検索する。
    """
    stripped = _strip_company(partner_name)
    return config.PARTNER_DISPLAY_MAP.get(stripped, stripped)


def _is_bp(dept: str) -> bool:
    """所属部署がBP社員かどうかを判定する。"""
    return dept.startswith("B推")


def _grade_sort_key(grade: str) -> int:
    """グレードのソート順を返す。"""
    try:
        return config.GRADE_ORDER.index(grade)
    except ValueError:
        return len(config.GRADE_ORDER)


def _sort_members(members: list[Member]) -> list[Member]:
    """役員 → 自社社員（グレード順）→ BP社員の順にソートする。"""
    exec_ = [m for m in members if not m.is_bp and "役員" in m.dept]
    own = [m for m in members if not m.is_bp and "役員" not in m.dept]
    bp = [m for m in members if m.is_bp]
    own.sort(key=lambda m: _grade_sort_key(m.grade))
    return exec_ + own + bp


def _resolve_client_name(name: str) -> str:
    """顧客名（ユーザー名）の名寄せ。

    株式会社・㈱を除去してからマッピングを検索する。
    """
    stripped = _strip_company(name)
    return config.CLIENT_DISPLAY_MAP.get(stripped, stripped)


def _build_clients(client_df: pd.DataFrame) -> list[Client]:
    """DataFrameからClient一覧を構築する。"""
    client_df = client_df.copy()
    client_df["出力用ユーザー名"] = client_df["ユーザー名"].apply(_resolve_client_name)
    clients = []
    for client_name, cdf in client_df.groupby("出力用ユーザー名", sort=False):
        projects = []
        for proj_name, pdf in cdf.groupby("業務名", sort=False):
            members = []
            for _, row in pdf.iterrows():
                bp = _is_bp(row["所属部署"])
                members.append(Member(
                    name=row["名前"],
                    dept=row["所属部署"],
                    grade=config.GRADE_DISPLAY_MAP.get(row["グレード"], row["グレード"]) if not bp else "",
                    is_bp=bp,
                ))
            members = _sort_members(members)
            projects.append(Project(name=proj_name, members=members))
        projects.sort(key=lambda p: p.count, reverse=True)
        clients.append(Client(name=client_name, projects=projects))
    clients.sort(key=lambda c: c.count, reverse=True)
    return clients


def process(df: pd.DataFrame) -> list[Division]:
    """DataFrameを構成図用の階層構造に加工する。

    階層: Division → Partner → Client → Project → Member
    営業区分でグルーピングし、キー昇順でソートして返す。
    """
    df = df.copy()

    # 出力用取引先名を付与
    df["出力用取引先名"] = df["取引先名"].apply(_resolve_display_name)

    # 営業区分マップを逆引き（取引先名 → 営業区分）に変換
    division_lookup = {}
    for div_key, partners in config.SALES_PARTNER_MAP.items():
        for name in partners:
            division_lookup[name] = div_key
    df["営業区分"] = df["出力用取引先名"].map(division_lookup).fillna("")

    # Division → Partner → Client の順にグルーピング
    divisions: dict[str, Division] = {}

    for (div_key, partner_display), group_df in df.groupby(
        ["営業区分", "出力用取引先名"], sort=False
    ):
        if div_key not in divisions:
            divisions[div_key] = Division(key=div_key)

        clients = _build_clients(group_df)
        partner = Partner(display_name=partner_display, clients=clients)
        divisions[div_key].partners.append(partner)

    # 各営業区分内の取引先を人数降順にソート
    for div in divisions.values():
        div.partners.sort(key=lambda p: p.count, reverse=True)

    # SALES_PARTNER_MAP の登録順でソート（未登録＝空文字キーは末尾）
    map_keys = list(config.SALES_PARTNER_MAP.keys())
    return [divisions[k] for k in sorted(
        divisions.keys(),
        key=lambda k: (k not in config.SALES_PARTNER_MAP, map_keys.index(k) if k in config.SALES_PARTNER_MAP else 0),
    )]
