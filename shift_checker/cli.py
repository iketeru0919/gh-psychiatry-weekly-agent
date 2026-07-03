"""コマンドライン実行（手動実行の入口）。

使い方:
    python -m shift_checker            # 対象年月を対話入力
    python -m shift_checker --month 2606
    python -m shift_checker --month 2606 --base "D:\\path\\to\\シフト管理"
"""

from __future__ import annotations

import argparse
import re
import sys
from datetime import date
from pathlib import Path

from .checker import build_comparison, build_current_result, build_error_result
from .config import load_settings, resolve_output_dir, resolve_snapshot_dir
from .external import build_external_rows
from .history import build_trend, update_history
from .labor import analyze_labor
from .parser import ShiftFileError, parse_shift_book
from .report import write_report
from .scanner import scan_shift_files
from .snapshot import find_snapshot_file, latest_run_dir, month_key, save_snapshot


def parse_month(text: str) -> tuple[int, int] | None:
    """'2606' '202606' '2026-06' '2026/6' などを (2026, 6) に変換する。"""
    text = text.strip()
    match = re.fullmatch(r"(\d{2})(0[1-9]|1[0-2])", text)
    if match:
        return 2000 + int(match.group(1)), int(match.group(2))
    match = re.fullmatch(r"((?:19|20)\d{2})[-/年.]?\s*(0?[1-9]|1[0-2])月?", text)
    if match:
        return int(match.group(1)), int(match.group(2))
    return None


def ask_month() -> tuple[int, int]:
    today = date.today()
    default = f"{today.year % 100:02d}{today.month:02d}"
    while True:
        entered = input(f"対象年月を入力してください（例: 2606）[Enterで {default}]: ").strip()
        parsed = parse_month(entered or default)
        if parsed:
            return parsed
        print("  形式が読み取れません。「2606」または「2026-06」の形式で入力してください。")


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="シフト週次チェック（Excel出力）")
    parser.add_argument("--month", help="対象年月（例: 2606, 202606, 2026-06）")
    parser.add_argument("--base", help="シフト管理フォルダのパス（config.jsonより優先）")
    parser.add_argument("--out", help="結果Excelの出力先フォルダ（既定: デスクトップ）")
    parser.add_argument("--no-snapshot", action="store_true",
                        help="スナップショット保存と前回比較を行わない")
    args = parser.parse_args(argv)

    settings = load_settings()
    if args.base:
        settings.base_dir = args.base
    if args.out:
        settings.output_dir = args.out

    base_dir = Path(settings.base_dir)
    if not base_dir.is_dir():
        print(f"エラー: シフト管理フォルダが見つかりません: {base_dir}")
        print("shift_checker/config.json の base_dir を実際のパスに合わせてください。")
        return 1

    if args.month:
        parsed = parse_month(args.month)
        if not parsed:
            print(f"エラー: 対象年月の形式が読み取れません: {args.month}")
            return 1
        year, month = parsed
    else:
        year, month = ask_month()

    print(f"\n対象年月: {year}年{month}月 / フォルダ: {base_dir}")
    print("ファイルを検索しています...")
    facilities = scan_shift_files(str(base_dir), year, month)
    if not facilities:
        print(f"対象年月（{year}年{month}月）の勤務表ファイルが見つかりませんでした。")
        print("フォルダ名・ファイル名に「2606」「2026年6月」等の年月表記が含まれるか確認してください。")
        return 1
    print(f"{len(facilities)}施設分のファイルを検出しました。\n")

    snapshot_dir = resolve_snapshot_dir(settings)
    previous_run = None if args.no_snapshot else latest_run_dir(snapshot_dir, year, month)
    previous_stamp = previous_run.name if previous_run else ""
    if previous_run:
        print(f"前回スナップショット（{previous_stamp}）と比較します。")
    else:
        print("前回スナップショットが無いため、今回のみチェックを行います。")

    results = []
    snapshot_files = {}
    labor_rows = []
    for index, entry in enumerate(facilities, start=1):
        label = f"[{index}/{len(facilities)}] {entry.facility}"
        try:
            current = parse_shift_book(entry.chosen, settings)
        except ShiftFileError as error:
            print(f"{label} ... 読取エラー: {error}")
            results.append(build_error_result(entry.facility, entry.chosen, str(error)))
            continue

        result = None
        if previous_run:
            snapshot_file = find_snapshot_file(previous_run, entry.facility)
            if snapshot_file:
                try:
                    previous = parse_shift_book(snapshot_file, settings)
                    result = build_comparison(previous, current, entry.facility, settings)
                except ShiftFileError as error:
                    result = build_current_result(current, entry.facility, settings)
                    result.warnings.append(f"前回スナップショットを読めず比較省略: {error}")
            else:
                result = build_current_result(current, entry.facility, settings)
                result.warnings.append("前回スナップショットに本施設が無いため比較省略")
        if result is None:
            result = build_current_result(current, entry.facility, settings)

        result.duplicate_files = entry.duplicate_count
        if entry.duplicate_count > 1:
            result.warnings.append(
                f"対象月の候補が{entry.duplicate_count}件あり、最新更新のファイルを採用"
            )
        results.append(result)
        snapshot_files[entry.facility] = entry.chosen
        labor_rows.extend(analyze_labor(current, entry.facility, settings, month))
        alerts = len(result.alerts)
        changes = len(result.change_rows)
        detail = f"職員{len(current.staff)}名 要確認{alerts}名"
        if previous_run:
            detail += f" 日別変更{changes}件"
        print(f"{label} ... OK（{detail}）")

    labor_rows.sort(key=lambda r: ({"critical": 0, "warning": 1, "ok": 2}[r.severity],
                                   r.facility, r.name))
    trend = None
    if not args.no_snapshot:
        records = update_history(snapshot_dir, month_key(year, month), labor_rows)
        trend = build_trend(records)

    external_rows = build_external_rows(results, settings)
    external_count = sum(1 for r in external_rows if not r.is_total)
    print(f"外部人材: {external_count}枠（タイミー・派遣）")

    output_dir = resolve_output_dir(settings)
    report_path = write_report(results, settings, output_dir, year, month,
                               compared=bool(previous_run), previous_stamp=previous_stamp,
                               labor_rows=labor_rows, trend=trend, external_rows=external_rows)
    labor_alerts = sum(1 for r in labor_rows if r.severity != "ok")
    print(f"労務チェック: 対象{len(labor_rows)}名、指摘{labor_alerts}名")

    if not args.no_snapshot and snapshot_files:
        run_dir = save_snapshot(snapshot_dir, year, month, snapshot_files,
                                settings.snapshot_keep)
        print(f"\nスナップショットを保存しました: {run_dir}")

    ok = sum(1 for r in results if r.status == "OK")
    error_count = len(results) - ok
    print(f"\n処理結果: {ok}施設OK / {error_count}施設エラー")
    if error_count:
        print("エラー内容は結果Excelの「サマリー」シートを確認してください。")
    print(f"結果Excel: {report_path}")

    if sys.platform == "win32":
        try:
            import os
            os.startfile(report_path)  # 結果を自動で開く
        except OSError:
            pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
