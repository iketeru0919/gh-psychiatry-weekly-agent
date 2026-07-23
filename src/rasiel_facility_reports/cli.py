from __future__ import annotations

import argparse
from pathlib import Path

from .aggregate import (
    build_daily_summary,
    filter_reports,
    find_missing_facilities,
    list_unknown_labels,
    resolve_field_keys,
    write_daily_csv,
    write_raw_csv,
)
from .parser import parse_chat_log


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "RASIELグループホームのLINE営業活動報告チャットログを解析し、"
            "拠点ごとの日次実績をCSVに集計します。AI/外部APIは使用しません。"
        )
    )
    parser.add_argument("input", type=Path, help="チャットログのテキストファイル")
    parser.add_argument("--out-dir", type=Path, default=Path("reports"), help="CSV出力先ディレクトリ")
    parser.add_argument(
        "--default-year", default=None, help="ログに西暦が無い日付に使う年(例: 2026)。省略時は年不明のまま出力"
    )
    parser.add_argument(
        "--facility",
        action="append",
        default=None,
        help="抽出したい拠点名(部分一致、複数指定可)。省略時は全拠点",
    )
    parser.add_argument(
        "--fields",
        default=None,
        help="出力したい項目キーのカンマ区切り(例: occupancy,inquiries,tours)。省略時は全項目",
    )
    parser.add_argument(
        "--missing-on",
        default=None,
        help="指定日(YYYY-MM-DD)に報告が無い拠点を標準出力に一覧表示",
    )
    args = parser.parse_args()

    text = args.input.read_text(encoding="utf-8")
    reports = parse_chat_log(text, default_year=args.default_year)
    print(f"Parsed {len(reports)} messages.")

    reports = filter_reports(reports, facilities=args.facility)
    field_keys = resolve_field_keys(args.fields.split(",") if args.fields else None)

    raw_path = write_raw_csv(reports, args.out_dir / "raw_messages.csv", field_keys)
    print(f"Raw messages written to {raw_path}")

    daily = build_daily_summary(reports)
    daily_path = write_daily_csv(daily, args.out_dir / "daily_summary.csv", field_keys)
    print(f"Daily summary written to {daily_path} ({len(daily)} rows)")

    unknown = list_unknown_labels(reports)
    if unknown:
        print("Unrecognized labels (format drift candidates):")
        for label, count in unknown.items():
            print(f"  {label}: {count}")

    if args.missing_on:
        missing = find_missing_facilities(reports, args.missing_on)
        print(f"Facilities with no report on {args.missing_on}: {len(missing)}")
        for facility in missing:
            print(f"  - {facility}")


if __name__ == "__main__":
    main()
