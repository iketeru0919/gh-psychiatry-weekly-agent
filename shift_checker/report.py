"""結果Excel（デスクトップ出力）の生成。"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .checker import FacilityResult
from .config import SUMMARY_LABELS, Settings
from .util import matches_exact

HEADER_FILL = PatternFill("solid", fgColor="16324F")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=10)
CRITICAL_FILL = PatternFill("solid", fgColor="FFDEDE")
WARNING_FILL = PatternFill("solid", fgColor="FFF3C4")
ERROR_FILL = PatternFill("solid", fgColor="F4CCCC")
OK_FILL = PatternFill("solid", fgColor="DFF3E8")
NOTE_FONT = Font(color="617080", size=9)

SEVERITY_FILLS = {"critical": CRITICAL_FILL, "warning": WARNING_FILL}


def _write_table(sheet, headers: list[str], rows: list[list], row_fills=None,
                 widths: list[int] | None = None) -> None:
    sheet.append(headers)
    for cell in sheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center")
    for index, row in enumerate(rows):
        sheet.append(row)
        fill = row_fills[index] if row_fills else None
        if fill:
            for cell in sheet[sheet.max_row]:
                cell.fill = fill
    sheet.freeze_panes = "A2"
    if widths:
        for column_number, width in enumerate(widths, start=1):
            sheet.column_dimensions[get_column_letter(column_number)].width = width
    sheet.auto_filter.ref = sheet.dimensions


def _fmt(value):
    return "" if value is None else value


def write_report(results: list[FacilityResult], settings: Settings, output_dir: Path,
                 year: int, month: int, compared: bool, previous_stamp: str) -> Path:
    workbook = Workbook()
    month_tag = f"{year % 100:02d}{month:02d}"
    now = datetime.now()

    # --- サマリー ---
    sheet = workbook.active
    sheet.title = "サマリー"
    headers = ["施設", "状態", "職員数", "重要", "要確認", "体制加算判定", "過不足結果",
               "警告・エラー内容", "使用ファイル", "候補数"]
    rows, fills = [], []
    for result in results:
        summary = result.current.facility_summary if result.current else {}
        note = result.error or " / ".join(result.warnings)
        rows.append([
            result.facility,
            result.status,
            len(result.current.staff) if result.current else "",
            result.critical_count if result.current else "",
            len(result.alerts) if result.current else "",
            _fmt(summary.get("EV78", "")),
            _fmt(summary.get("BU75", "")),
            note,
            result.source_path,
            result.duplicate_files if result.duplicate_files > 1 else "",
        ])
        if result.status != "OK":
            fills.append(ERROR_FILL)
        elif str(summary.get("EV78", "")).find("未達") >= 0 or result.critical_count:
            fills.append(WARNING_FILL)
        else:
            fills.append(OK_FILL)
    _write_table(sheet, headers, rows, fills,
                 widths=[22, 10, 8, 6, 8, 12, 10, 50, 60, 6])
    ok = sum(1 for r in results if r.status == "OK")
    footer = [
        f"実行日時: {now:%Y-%m-%d %H:%M}",
        f"対象年月: {year}年{month}月（{month_tag}）",
        f"処理結果: {ok}施設OK / {len(results) - ok}施設エラー",
        f"比較基準: {previous_stamp} 時点のスナップショット" if compared
        else "比較基準: なし（初回実行のため今回のみチェック）",
    ]
    for line in footer:
        sheet.append([line])
        sheet.cell(sheet.max_row, 1).font = NOTE_FONT

    # --- 要確認者 / 職員全件 ---
    staff_headers = ["施設", "判定", "職員氏名", "状態", "職種", "雇用区分", "所定",
                     "前回総労働", "今回総労働", "増減", "所定との差",
                     "有休(前)", "有休(今)", "有休差", "公休(前)", "公休(今)", "公休差",
                     "日別変更数", "内容"]
    staff_widths = [22, 8, 18, 9, 18, 9, 7, 9, 9, 7, 9, 7, 7, 7, 7, 7, 7, 9, 60]

    def staff_row_values(row):
        return [
            row.facility, row.judgment, row.name, row.status, row.job,
            row.employment_type, _fmt(row.prescribed),
            _fmt(row.before_hours), _fmt(row.after_hours), _fmt(row.hour_delta),
            _fmt(row.required_delta),
            row.before_paid, row.after_paid, row.paid_delta,
            row.before_vacation, row.after_vacation, row.vacation_delta,
            row.change_count, row.reason,
        ]

    for title, selector in (("要確認者", lambda r: r.alerts), ("職員全件", lambda r: r.staff_rows)):
        sheet = workbook.create_sheet(title)
        rows, fills = [], []
        for result in results:
            for row in selector(result):
                rows.append(staff_row_values(row))
                fills.append(SEVERITY_FILLS.get(row.severity))
        if rows:
            _write_table(sheet, staff_headers, rows, fills, staff_widths)
        else:
            _write_table(sheet, staff_headers, [], None, staff_widths)
            sheet.append(["該当なし"])

    # --- 日別変更（比較実行時のみ）---
    if compared:
        sheet = workbook.create_sheet("日別変更")
        headers = ["施設", "職員氏名", "日付", "曜日", "前回", "今回", "変更区分", "セル"]
        rows, fills = [], []
        for result in results:
            for change in result.change_rows:
                rows.append([
                    change.facility, change.name, change.date_label, change.weekday,
                    change.previous_value, change.current_value, change.change_type,
                    change.cell,
                ])
                fills.append(WARNING_FILL if change.change_type in ("休暇変更", "削除") else None)
        _write_table(sheet, headers, rows, fills, widths=[22, 18, 9, 6, 12, 12, 10, 8])
        if not rows:
            sheet.append(["変更はありませんでした"])

    # --- 配置体制 ---
    sheet = workbook.create_sheet("配置体制")
    keys = list(SUMMARY_LABELS.keys())
    headers = ["施設"] + [SUMMARY_LABELS[k] for k in keys]
    rows, fills = [], []
    for result in results:
        if not result.current:
            rows.append([result.facility] + ["読取エラー"] + [""] * (len(keys) - 1))
            fills.append(ERROR_FILL)
            continue
        summary = result.current.facility_summary
        rows.append([result.facility] + [_fmt(summary.get(k, "")) for k in keys])
        fills.append(WARNING_FILL if "未達" in str(summary.get("EV78", "")) else None)
    _write_table(sheet, headers, rows, fills, widths=[22] + [13] * len(keys))

    # --- 勤務記号 ---
    sheet = workbook.create_sheet("勤務記号")
    headers = ["記号", "区分", "使用施設"]
    symbol_map: dict[str, set[str]] = {}
    for result in results:
        if result.current:
            for symbol in result.current.shift_symbols:
                symbol_map.setdefault(symbol, set()).add(result.facility)
    rows = []
    for symbol in sorted(symbol_map):
        category = ("有休" if matches_exact(symbol, settings.paid_words)
                    else "公休" if matches_exact(symbol, settings.vacation_words)
                    else "勤務")
        rows.append([symbol, category, "、".join(sorted(symbol_map[symbol]))])
    _write_table(sheet, headers, rows, None, widths=[12, 8, 70])
    sheet.append([])
    sheet.append(["※「有休」「公休」の判定文字は shift_checker/config.json の "
                  "paid_words / vacation_words で変更できます"])
    sheet.cell(sheet.max_row, 1).font = NOTE_FONT

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"シフトチェック結果_{month_tag}_{now:%Y%m%d_%H%M}.xlsx"
    workbook.save(output_path)
    return output_path
