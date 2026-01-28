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


def _resolve_mapping(partner_name: str) -> tuple[str, str]:
    """取引先名から (営業区分, 出力用取引先名) を返す。"""
    mapping = config.SALES_DIVISION_MAP.get(partner_name)
    if mapping:
        return mapping
    # 未登録の場合
    division = config.SALES_DIVISION_DEFAULT
    display = config.SALES_DISPLAY_NAME_DEFAULT or partner_name
    return (division, display)


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
    """自社社員（グレード順）→ BP社員の順にソートする。"""
    own = [m for m in members if not m.is_bp]
    bp = [m for m in members if m.is_bp]
    own.sort(key=lambda m: _grade_sort_key(m.grade))
    return own + bp


def _build_clients(client_df: pd.DataFrame) -> list[Client]:
    """DataFrameからClient一覧を構築する。"""
    clients = []
    for client_name, cdf in client_df.groupby("ユーザー名", sort=False):
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
        clients.append(Client(name=client_name, projects=projects))
    return clients


def process(df: pd.DataFrame) -> list[Division]:
    """DataFrameを構成図用の階層構造に加工する。

    階層: Division → Partner → Client → Project → Member
    営業区分のキー昇順（A → B → C → …）でソートして返す。
    """
    df = df.copy()

    # 営業区分と出力用取引先名を付与
    resolved = df["取引先名"].apply(_resolve_mapping)
    df["営業区分"] = resolved.apply(lambda x: x[0])
    df["出力用取引先名"] = resolved.apply(lambda x: x[1])

    # Division → Partner → Client の順にグルーピング
    divisions: dict[str, Division] = {}

    for (div_key, partner_display), group_df in df.groupby(
        ["営業区分", "出力用取引先名"], sort=False
    ):
        # Division
        if div_key not in divisions:
            divisions[div_key] = Division(key=div_key)

        # Partner
        clients = _build_clients(group_df)
        partner = Partner(display_name=partner_display, clients=clients)
        divisions[div_key].partners.append(partner)

    # キー昇順でソート
    return [divisions[k] for k in sorted(divisions.keys())]
