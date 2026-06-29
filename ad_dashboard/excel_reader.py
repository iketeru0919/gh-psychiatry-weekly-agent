"""
Excel読み込みモジュール
- 1ファイル・施設ごとにシートが分かれている構成を想定
- セル位置はcell_map.jsonで管理（コードを触らずに変更可能）
"""

import json
import re
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils import column_index_from_string

CELL_MAP_PATH = Path(__file__).parent / "cell_map.json"


# ── ユーティリティ ────────────────────────────────────────

def col_letter_to_index(letter: str) -> int:
    """'A'→1, 'B'→2, 'J'→10"""
    return column_index_from_string(letter)


def get_cell_value(ws, row: int, col_letter: str) -> Any:
    """指定セルの値を取得。Noneや空文字は0に変換しない（呼び出し元で処理）"""
    return ws.cell(row=row, column=col_letter_to_index(col_letter)).value


def advance_col(col_letter: str, steps: int) -> str:
    """'C' + 2 → 'E'"""
    idx = col_letter_to_index(col_letter) + steps
    # openpyxl の get_column_letter を使う
    from openpyxl.utils import get_column_letter
    return get_column_letter(idx)


def safe_float(v) -> float | None:
    """パーセント表示・文字列・None を安全にfloatへ変換"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        # openpyxlはパーセント書式セルを0.915のように返す場合がある
        return float(v)
    s = str(v).replace("%", "").replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def safe_int(v) -> int | None:
    f = safe_float(v)
    return int(f) if f is not None else None


# ── メイン読み込み関数 ────────────────────────────────────

def load_cell_map() -> dict:
    with open(CELL_MAP_PATH, encoding="utf-8") as f:
        return json.load(f)


def read_monthly_table(ws, table_cfg: dict, months_list: list[str]) -> list[dict]:
    """
    月次テーブルを読み込む。
    months_list: 列順の月ラベルリスト（例: ['3月','4月','5月',...]）
    戻り値: [{"month": "2026-03", "residents": 20, ...}, ...]
    """
    rows_cfg = table_cfg["rows"]
    col_start = table_cfg["data_col_start"]
    defined_months = table_cfg.get("months", months_list)  # テーブル個別定義があればそれを使う

    records = []
    for i, month_label in enumerate(defined_months):
        col = advance_col(col_start, i)
        row_data: dict = {}
        any_value = False

        for field, row_num in rows_cfg.items():
            if field.startswith("_"):
                continue
            val = get_cell_value(ws, row_num, col)
            # パーセント書式対応：openpyxlが0.915を返す場合に%に変換
            if isinstance(val, float) and val < 2.0 and "rate" in field:
                val = round(val * 100, 1)
            row_data[field] = val
            if val is not None:
                any_value = True

        if not any_value:
            continue  # 未入力月はスキップ

        # month_label "3月" → "2026-03" （年はファイル名や別セルから取得できるが、ここでは固定）
        row_data["month"] = month_label_to_iso(month_label)
        records.append(row_data)

    return records


def month_label_to_iso(label: str, fiscal_year_start: int = 2026) -> str:
    """
    '3月' → '2026-03'
    会計年度: 4月始まりを想定。4〜12月→fiscal_year_start、1〜3月→fiscal_year_start+1
    """
    num = int(re.sub(r"[^\d]", "", label))
    year = fiscal_year_start if num >= 4 else fiscal_year_start + 1
    return f"{year}-{num:02d}"


def read_sheet(ws, cmap: dict, sheet_name: str) -> dict:
    """1シート（1施設）を読み込んでdictで返す"""
    result: dict = {
        "facility_id":   sheet_name,
        "facility_name": sheet_name,
        "group_company": "",
        "capacity":      0,
        "monthly":       [],
        "weekly_current": [],
    }

    # 定員
    cap_cell = cmap.get("facility_info", {}).get("capacity", "C1")
    row, col = cell_str_to_row_col(cap_cell)
    result["capacity"] = safe_int(ws.cell(row=row, column=col).value) or 0

    # 月次テーブル（重要指標推移）
    if "monthly_table" in cmap:
        monthly_basic = read_monthly_table(ws, cmap["monthly_table"], cmap["monthly_table"]["months"])
        result["monthly"] = monthly_basic
    else:
        result["monthly"] = []

    monthly_by_month = {m["month"]: m for m in result["monthly"]}

    months_list = cmap.get("monthly_table", {}).get("months", [])

    # スタッフィングテーブル（常勤換算）を月次にマージ
    if "staffing_table" in cmap:
        staffing = read_monthly_table(ws, cmap["staffing_table"], months_list)
        for s in staffing:
            month = s.pop("month")
            if month in monthly_by_month:
                monthly_by_month[month].update(s)

    # ブランディングテーブルをマージ
    if "branding_table" in cmap:
        branding = read_monthly_table(ws, cmap["branding_table"], months_list)
        for b in branding:
            month = b.pop("month")
            if month in monthly_by_month:
                monthly_by_month[month].update(b)

    # タイミーテーブルをマージ
    if "timee_table" in cmap:
        timee = read_monthly_table(ws, cmap["timee_table"], months_list)
        for t in timee:
            month = t.pop("month")
            if month in monthly_by_month:
                monthly_by_month[month].update(t)

    # 運営評価テーブルをマージ
    if "operations_table" in cmap:
        ops = read_monthly_table(ws, cmap["operations_table"], months_list)
        for o in ops:
            month = o.pop("month")
            if month in monthly_by_month:
                monthly_by_month[month].update(o)

    result["monthly"] = list(monthly_by_month.values())

    # 週次テーブル
    if "weekly_table" in cmap:
        result["weekly_current"] = read_weekly_table(ws, cmap["weekly_table"])

    return result


def read_weekly_table(ws, table_cfg: dict) -> list[dict]:
    rows_cfg = table_cfg["rows"]
    col_start = table_cfg["data_col_start"]
    weeks = table_cfg["weeks"]

    records = []
    for i, week_label in enumerate(weeks):
        col = advance_col(col_start, i)
        row_data: dict = {"week": i + 1}
        any_value = False
        for field, row_num in rows_cfg.items():
            if field.startswith("_"):
                continue
            val = get_cell_value(ws, row_num, col)
            if isinstance(val, float) and val < 2.0 and "rate" in field:
                val = round(val * 100, 1)
            row_data[field] = val
            if val is not None:
                any_value = True
        if any_value:
            records.append(row_data)
    return records


def cell_str_to_row_col(cell_str: str) -> tuple[int, int]:
    """'C5' → (row=5, col=3)"""
    match = re.match(r"([A-Za-z]+)(\d+)", cell_str)
    if not match:
        raise ValueError(f"不正なセル番地: {cell_str}")
    col = col_letter_to_index(match.group(1))
    row = int(match.group(2))
    return row, col


# ── 公開インターフェース ──────────────────────────────────

def load_from_excel(excel_path: str | Path) -> list[dict]:
    """
    Excelファイルから全施設データを読み込む。
    戻り値: [facility_dict, ...]  （app.pyのload_all_facilities()と同じ形式）
    """
    cmap = load_cell_map()
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    facilities = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        try:
            facility = read_sheet(ws, cmap, sheet_name)
            facilities.append(facility)
        except Exception as e:
            print(f"[警告] シート '{sheet_name}' の読み込みに失敗しました: {e}")

    wb.close()
    return facilities
