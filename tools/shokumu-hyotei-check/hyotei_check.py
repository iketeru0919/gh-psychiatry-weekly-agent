# -*- coding: utf-8 -*-
"""職務評定表チェック・集約ロジック

Google Drive の「職務評定」フォルダ内の Excel(施設別・月別シート)から
氏名/順位/評価/雇用形態/入社日/職種 を抽出し、施設別に1つの Excel に
まとめる。あわせて入力エラーを検出してエラー一覧を作る。
"""

import re
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime, date
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

FORMULA_ERRORS = {"#REF!", "#VALUE!", "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!"}
VALID_GRADES = {"S", "A", "B", "C", "D", "E", "F"}

# 列の位置(1始まり): C=職種, E=雇用形態, F=入社日, H=氏名, I=順位, J=評価
COL_SHOKUSHU = 3
COL_KOYO = 5
COL_NYUSHA = 6
COL_NAME_LEFT = 2   # B列の氏名(手入力側)
COL_NAME = 8
COL_RANK = 9
COL_GRADE = 10


@dataclass
class Record:
    facility: str
    file_name: str
    sheet_name: str
    month: str          # "YYYY-MM"
    row: int
    name: str = ""
    rank: object = None
    grade: object = None
    koyo: str = ""
    nyusha: object = None
    shokushu: str = ""


@dataclass
class Issue:
    facility: str
    file_name: str
    sheet_name: str
    month: str
    row: object         # 行番号 or "-"
    kind: str
    detail: str


def normalize(s):
    """判定・比較用の正規化(全角→半角等)。表示用の値には使わない。"""
    if s is None:
        return ""
    return unicodedata.normalize("NFKC", str(s)).strip()


def clean(s):
    """表示用: 前後の空白だけ除去し、元の表記(全角スペース等)は保つ。"""
    if s is None:
        return ""
    return str(s).strip()


def facility_from_filename(name):
    m = re.match(r"[【\[](.+?)[】\]]", name)
    return m.group(1) if m else Path(name).stem


def parse_month_token(token):
    """ユーザー指定の月 ("2026-06", "2026.6", "202606", "2026/6") → "YYYY-MM"."""
    t = normalize(token).replace("年", "-").replace("月", "")
    m = re.fullmatch(r"(\d{4})[.\-/](\d{1,2})", t)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}"
    m = re.fullmatch(r"(\d{4})(\d{2})", t)
    if m and 1 <= int(m.group(2)) <= 12:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}"
    raise ValueError(f"月の指定 '{token}' を解釈できません。例: 2026-06 / 2026.6 / 202606")


def month_from_sheet_title(title):
    """シート名から (year, month) を推定。判らない要素は None。"""
    t = normalize(title)
    m = re.search(r"[RrRr]\s*(\d{1,2})[.\-年]?\s*(\d{1,2})", t)  # R8.6 (令和)
    if m:
        return 2018 + int(m.group(1)), int(m.group(2))
    m = re.search(r"(\d{4})\s*[.．\-/年]\s*(\d{1,2})", t)  # 2026.6 / 2026年6月
    if m and 1 <= int(m.group(2)) <= 12:
        return int(m.group(1)), int(m.group(2))
    m = re.fullmatch(r"(\d{4})(\d{2})", t)  # 202606
    if m and 1 <= int(m.group(2)) <= 12:
        return int(m.group(1)), int(m.group(2))
    m = re.fullmatch(r"(\d{2})(\d{2})", t)  # 2606 → 2026-06
    if m and 1 <= int(m.group(2)) <= 12:
        return 2000 + int(m.group(1)), int(m.group(2))
    m = re.search(r"(\d{1,2})\s*月", t)  # 宮下6月 → 月のみ
    if m and 1 <= int(m.group(1)) <= 12:
        return None, int(m.group(1))
    return None, None


def month_from_hyokazuki_cell(ws):
    """シート内の「評価月」ラベルの下のセルから (year, month) を取る。"""
    for row in ws.iter_rows(min_row=1, max_row=8, max_col=12):
        for cell in row:
            if normalize(cell.value) == "評価月":
                below = ws.cell(row=cell.row + 1, column=cell.column)
                v = below.value
                if isinstance(v, (datetime, date)):
                    return v.year, v.month
                m = re.search(r"(\d{4})\s*[./\-年]\s*(\d{1,2})", normalize(v))
                if m and 1 <= int(m.group(2)) <= 12:
                    return int(m.group(1)), int(m.group(2))
                return None, None
    return None, None


def sheet_month(ws):
    """シートの対象月 "YYYY-MM" を判定。評価月セル優先、シート名で補完。"""
    y, m = month_from_hyokazuki_cell(ws)
    ty, tm = month_from_sheet_title(ws.title)
    year = y if y else ty
    month = m if m else tm
    if year and month:
        return f"{year:04d}-{month:02d}"
    return None


def find_data_start(ws):
    """H/I/J のラベル行(氏名・順位・評価)を探し、その次の行番号を返す。"""
    for row in ws.iter_rows(min_row=1, max_row=15, min_col=COL_NAME, max_col=COL_GRADE):
        vals = [normalize(c.value) for c in row]
        if vals[0] == "氏名" and vals[COL_RANK - COL_NAME] == "順位":
            return row[0].row + 1
    return None


def is_formula_error(v):
    return normalize(v) in FORMULA_ERRORS


def parse_sheet(ws, facility, file_name):
    """1シートを読み、(records, issues, month) を返す。"""
    records, issues = [], []
    month = sheet_month(ws)
    month_disp = month or f"不明({ws.title})"
    start = find_data_start(ws)
    if start is None:
        issues.append(Issue(facility, file_name, ws.title, month_disp, "-",
                            "様式不一致", "氏名・順位・評価のヘッダー行が見つかりません"))
        return records, issues, month

    for r in range(start, ws.max_row + 1):
        name_l = clean(ws.cell(r, COL_NAME_LEFT).value)
        name = clean(ws.cell(r, COL_NAME).value)
        rank = ws.cell(r, COL_RANK).value
        grade = ws.cell(r, COL_GRADE).value
        koyo = clean(ws.cell(r, COL_KOYO).value)
        nyusha = ws.cell(r, COL_NYUSHA).value
        shokushu = clean(ws.cell(r, COL_SHOKUSHU).value)

        if "（例）" in name_l or "(例)" in name_l:
            continue
        grade_n = normalize(grade)
        # 空行: 数式残骸の FALSE/0 だけの行も空とみなす
        if not name_l and not name and grade_n in ("", "FALSE", "0"):
            continue

        rec = Record(facility, file_name, ws.title, month_disp, r,
                     name=name or name_l, rank=rank, grade=grade_n,
                     koyo=koyo, nyusha=nyusha, shokushu=shokushu)
        records.append(rec)

        # --- エラーチェック(行単位) ---
        def add(kind, detail):
            issues.append(Issue(facility, file_name, ws.title, month_disp, r, kind, detail))

        for label, v in (("氏名(H列)", name), ("順位(I列)", rank), ("評価(J列)", grade),
                         ("雇用形態(E列)", koyo), ("入社日(F列)", nyusha), ("職種(C列)", shokushu)):
            if is_formula_error(v):
                add("数式エラー", f"{label} が {normalize(v)} になっています")
            elif normalize(v) == "":
                add("空欄", f"{label} が空欄です(氏名: {rec.name or '不明'})")

        if name_l and name and name_l != name:
            add("不整合", f"B列の氏名『{name_l}』とH列の氏名『{name}』が一致しません")
        if not name and name_l:
            add("不整合", f"B列に『{name_l}』がありますがH列(評定計算式)に反映されていません")

        if rank is not None and not is_formula_error(rank):
            if not (isinstance(rank, (int, float)) and float(rank).is_integer()):
                add("値の異常", f"順位 '{rank}' が整数ではありません(氏名: {rec.name})")
        if grade_n and grade_n not in VALID_GRADES and not is_formula_error(grade):
            add("値の異常", f"評価 '{grade_n}' が想定値({'/'.join(sorted(VALID_GRADES))})ではありません(氏名: {rec.name})")
        if nyusha is not None and normalize(nyusha) != "" and not is_formula_error(nyusha):
            if not isinstance(nyusha, (datetime, date)):
                if not re.fullmatch(r"\d{4}[./\-]\d{1,2}[./\-]\d{1,2}", normalize(nyusha)):
                    add("値の異常", f"入社日 '{nyusha}' が日付として読めません(氏名: {rec.name})")

    # --- 順位の重複・欠番(シート単位) ---
    ranks = [int(r.rank) for r in records
             if isinstance(r.rank, (int, float)) and float(r.rank).is_integer()]
    if ranks:
        seen, dups = set(), set()
        for v in ranks:
            (dups if v in seen else seen).add(v)
        if dups:
            issues.append(Issue(facility, file_name, ws.title, month_disp, "-", "順位の重複",
                                f"順位 {sorted(dups)} が複数の職員に付いています"))
        missing = sorted(set(range(1, max(ranks) + 1)) - set(ranks))
        if missing:
            issues.append(Issue(facility, file_name, ws.title, month_disp, "-", "順位の欠番",
                                f"順位 {missing} が抜けています(1〜{max(ranks)} のうち)"))
    return records, issues, month


def process_workbook(path, target_months):
    """1ファイルを処理。target_months に一致するシートのみ読む。
    返り値: (records, issues, found_months)"""
    file_name = Path(path).name
    facility = facility_from_filename(file_name)
    wb = openpyxl.load_workbook(path, data_only=True)
    records, issues, found = [], [], set()
    for ws in wb.worksheets:
        month = sheet_month(ws)
        if month:
            found.add(month)
        if month is None:
            # 月が判定できないシート: 対象月が1つでシートも1つなら読んでみる
            if len(wb.worksheets) == 1 and len(target_months) == 1:
                recs, iss, _ = parse_sheet(ws, facility, file_name)
                issues.append(Issue(facility, file_name, ws.title, target_months[0], "-",
                                    "月判定不可", "シート名・評価月セルから月を判定できないため、"
                                    f"指定月 {target_months[0]} のデータとして扱いました"))
                for r_ in recs:
                    r_.month = target_months[0]
                records += recs
                issues += iss
                found.add(target_months[0])
            continue
        if month in target_months:
            recs, iss, _ = parse_sheet(ws, facility, file_name)
            records += recs
            issues += iss
    wb.close()
    return records, issues, found


def check_folder(xlsx_paths, target_months):
    """複数ファイルを処理し、施設単位で対象月シートの不在も検出する。"""
    all_records, all_issues = [], []
    facility_months = {}
    for p in sorted(xlsx_paths):
        name = Path(p).name
        if name.startswith("~$"):
            continue
        facility = facility_from_filename(name)
        try:
            recs, iss, found = process_workbook(p, target_months)
        except Exception as e:  # 壊れたファイル等
            all_issues.append(Issue(facility, name, "-", "-", "-", "読込エラー", str(e)))
            continue
        all_records += recs
        all_issues += iss
        facility_months.setdefault(facility, set()).update(found)

    for facility, months in sorted(facility_months.items()):
        for m in target_months:
            if m not in months:
                all_issues.append(Issue(facility, "-", "-", m, "-", "対象月シート不在",
                                        f"{m} のシートが施設『{facility}』のどのファイルにも見つかりません"))
    return all_records, all_issues


HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
ERROR_FILL = PatternFill("solid", fgColor="FCE4EC")


def _safe_sheet_name(name, used):
    s = re.sub(r"[\[\]:*?/\\]", "_", name)[:31] or "施設"
    base, i = s, 2
    while s in used:
        s = f"{base[:28]}_{i}"
        i += 1
    used.add(s)
    return s


def write_output(records, issues, target_months, out_path):
    wb = openpyxl.Workbook()

    # --- エラー一覧(先頭シート) ---
    ws = wb.active
    ws.title = "エラー一覧"
    headers = ["施設", "ファイル名", "シート名", "対象月", "行", "種別", "内容"]
    ws.append(headers)
    for i in sorted(issues, key=lambda x: (x.facility, str(x.month), str(x.row))):
        ws.append([i.facility, i.file_name, i.sheet_name, i.month, i.row, i.kind, i.detail])
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = HEADER_FILL
    if len(issues) == 0:
        ws.append(["(エラーはありませんでした)"])
    else:
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.fill = ERROR_FILL
    widths = [14, 40, 12, 10, 6, 14, 60]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    # --- 施設別シート ---
    used = {"エラー一覧"}
    by_fac = {}
    for r in records:
        by_fac.setdefault(r.facility, []).append(r)
    for facility in sorted(by_fac):
        wsf = wb.create_sheet(_safe_sheet_name(facility, used))
        wsf.append(["対象月", "氏名", "順位", "評価", "雇用形態", "入社日", "職種",
                    "ファイル名", "シート名"])
        for c in wsf[1]:
            c.font = Font(bold=True)
            c.fill = HEADER_FILL
        def sort_key(r):
            rank = r.rank if isinstance(r.rank, (int, float)) else 10 ** 6
            return (str(r.month), rank)
        for r in sorted(by_fac[facility], key=sort_key):
            nyusha = r.nyusha
            if isinstance(nyusha, datetime):
                nyusha = nyusha.date()
            row = [r.month, r.name, r.rank if not is_formula_error(r.rank) else normalize(r.rank),
                   r.grade, r.koyo, nyusha, r.shokushu, r.file_name, r.sheet_name]
            wsf.append(row)
            wsf.cell(wsf.max_row, 6).number_format = "yyyy/m/d"
        for i, w in enumerate([10, 18, 6, 6, 10, 12, 14, 40, 12], 1):
            wsf.column_dimensions[get_column_letter(i)].width = w
        wsf.freeze_panes = "A2"

    wb.save(out_path)
    return out_path
