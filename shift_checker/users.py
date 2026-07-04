"""利用者情報の抽出（人員配置【入力用】シート）。

・入居者: 見出し行（「利用者名」）の次行から20行
・それ以降に名前がある行はショートステイ（「延べ利用回数」ラベルで打ち切り）
・施設により「重度」列の有無で列位置がずれるため、見出し行から列を自動判定する
"""

from __future__ import annotations

from dataclasses import dataclass

from .parser import ShiftFileError, load_workbook_safe
from .util import normalize, round2

USERS_SHEET_KEYWORD = "人員配置【入力用】"
RESIDENT_ROWS = 20          # 入居者の行数（C13〜C32相当）
SHORT_STAY_MAX_ROWS = 16    # ショートステイを探す最大行数


@dataclass
class UserRow:
    facility: str
    kind: str                  # 入居 / ショートステイ / 入居計 / ショートステイ計 / 施設合計
    name: str
    grade: object              # 障害区分（計の行は平均）
    severe: str                # 重度（Ⅰ/Ⅱ）。計の行は内訳
    days: object               # 延べ利用日数（計の行は合計）
    month_days: object
    rate: object               # 利用率・稼働率（%値）
    is_total: bool = False
    note: str = ""


def _load_grid(path, max_row=60, max_col=10):
    workbook = load_workbook_safe(path)
    sheet_name = None
    for name in workbook.sheetnames:
        if USERS_SHEET_KEYWORD in normalize(name):
            sheet_name = name
            break
    if not sheet_name:
        workbook.close()
        raise ShiftFileError("「人員配置【入力用】」シートが見つかりません")
    sheet = workbook[sheet_name]
    grid: dict[tuple[int, int], object] = {}
    for row in sheet.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for cell in row:
            if cell.value is not None:
                grid[(cell.row, cell.column)] = cell.value
    workbook.close()
    return grid


def _find_columns(grid):
    """見出し行（「利用者名」を含む行）から各列の位置を判定する。"""
    for row in range(8, 16):
        for col in range(2, 9):
            if "利用者名" in normalize(grid.get((row, col))):
                columns = {"header_row": row, "name": col}
                for c in range(2, 10):
                    header = normalize(grid.get((row, c)))
                    if "障害" in header and "区分" in header:
                        columns["grade"] = c
                    elif "延べ" in header and "日数" in header:
                        columns["days"] = c
                    elif header == "重度":
                        columns["severe"] = c
                if "grade" in columns and "days" in columns:
                    return columns
    raise ShiftFileError(
        "人員配置【入力用】の見出し行（利用者名・障害区分・延べ利用日数）を検出できません"
    )


def _find_month_days(grid):
    for row in range(3, 8):
        for col in range(2, 8):
            if "月の日数" in normalize(grid.get((row, col))):
                for c in range(col + 1, col + 4):
                    value = grid.get((row, c))
                    if isinstance(value, (int, float)) and not isinstance(value, bool):
                        return int(value)
    return None


def _numeric(value):
    return value if isinstance(value, (int, float)) and not isinstance(value, bool) else None


def _rate(days, month_days, count=1):
    if days is None or not month_days or not count:
        return None
    return round2(days / (month_days * count) * 100)


def _summary_row(facility, kind, entries, month_days) -> UserRow:
    grades = [_numeric(e[1]) for e in entries if _numeric(e[1]) is not None]
    days_values = [_numeric(e[3]) for e in entries if _numeric(e[3]) is not None]
    severe_counts: dict[str, int] = {}
    for e in entries:
        if e[2]:
            severe_counts[e[2]] = severe_counts.get(e[2], 0) + 1
    total_days = round2(sum(days_values)) if days_values else None
    return UserRow(
        facility=facility, kind=kind,
        name=f"{len(entries)}名",
        grade=round2(sum(grades) / len(grades)) if grades else "",
        severe="、".join(f"{k}×{v}" for k, v in sorted(severe_counts.items())),
        days=total_days,
        month_days=month_days,
        rate=_rate(total_days, month_days, len(entries)),
        is_total=True,
    )


def parse_users(path, facility: str) -> list[UserRow]:
    grid = _load_grid(path)
    columns = _find_columns(grid)
    month_days = _find_month_days(grid)
    header_row = columns["header_row"]

    def read_entry(row):
        name = normalize(grid.get((row, columns["name"])))
        if not name or "延べ利用回数" in name:
            return None
        return (
            name,
            grid.get((row, columns["grade"])),
            normalize(grid.get((row, columns["severe"]))) if "severe" in columns else "",
            _numeric(grid.get((row, columns["days"]))),
        )

    residents = []
    for row in range(header_row + 1, header_row + 1 + RESIDENT_ROWS):
        entry = read_entry(row)
        if entry:
            residents.append(entry)

    short_stays = []
    for row in range(header_row + 1 + RESIDENT_ROWS,
                     header_row + 1 + RESIDENT_ROWS + SHORT_STAY_MAX_ROWS):
        raw = normalize(grid.get((row, columns["name"])))
        if "延べ利用回数" in raw:
            break
        entry = read_entry(row)
        if entry:
            short_stays.append(entry)

    rows: list[UserRow] = []
    for kind, entries in (("入居", residents), ("ショートステイ", short_stays)):
        for name, grade, severe, days in entries:
            note = ""
            if kind == "入居" and days is not None and month_days and days < month_days:
                note = "月の日数より少（入退去・入院等の可能性）"
            rows.append(UserRow(
                facility=facility, kind=kind, name=name, grade=grade,
                severe=severe, days=days, month_days=month_days,
                rate=_rate(days, month_days), note=note,
            ))
        if entries:
            rows.append(_summary_row(
                facility, f"{kind}計", entries, month_days))
    if residents or short_stays:
        rows.append(_summary_row(facility, "施設合計", residents + short_stays, month_days))
    return rows
