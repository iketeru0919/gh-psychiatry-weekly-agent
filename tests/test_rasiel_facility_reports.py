from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from rasiel_facility_reports.aggregate import (  # noqa: E402
    build_daily_summary,
    filter_reports,
    find_missing_facilities,
    resolve_field_keys,
)
from rasiel_facility_reports.parser import extract_fields, normalize_facility, parse_chat_log  # noqa: E402

SAMPLE_PATH = Path(__file__).resolve().parent.parent / "sample_data" / "rasiel_chat_sample.txt"


def load_sample_reports():
    text = SAMPLE_PATH.read_text(encoding="utf-8")
    return parse_chat_log(text, default_year="2026")


def test_parses_all_messages_and_resolves_facility():
    reports = load_sample_reports()
    assert len(reports) == 207
    assert all(r.facility_key for r in reports)


def test_normalizes_facility_name_variants():
    assert normalize_facility("RASIEL南志賀") == normalize_facility("南志賀")
    assert normalize_facility("ラシエル谷田") == normalize_facility("RASIEL谷田") == normalize_facility("谷田")


def test_extract_fields_handles_both_morning_and_evening_templates():
    morning_body = (
        "①名前：生田裕久\n➁拠点名：RASIEL南志賀\n➂入居者数：19名\n④目標入居件数：1\n"
        "⑤営業予定時間：\n⑥営業累計時間：4\n⑦営業内容：\n⑧預り金管理：生田　畠中\n⑨カメラ録画確認：〇1"
    )
    values, unknown = extract_fields(morning_body)
    assert values["reporter_name"] == "生田裕久"
    assert values["facility_name"] == "RASIEL南志賀"
    assert values["hours_cumulative"] == "4"
    assert unknown == {}

    evening_body = (
        "①名前：福田美貴\n➁拠点名：宇都宮\n➂入居者数：19名\n④目標入居件数：1名\n"
        "⑤短期稼働日数：32\n⑥営業実績時間：\n⑦営業累計時間：3h\n⑧問合せ件数：\n"
        "⑨案件発掘：No.\n⑩見学：No.\n⑪体験：No.（無料・有料）\n⑫仮申込：No.\n⑬入居：No.\n"
        "【翌日の予定】\n　日勤1"
    )
    values, _ = extract_fields(evening_body)
    assert values["work_days"] == "32"
    assert values["tours"] == "No."
    assert values["next_day_plan"] == "日勤1"


def test_does_not_split_on_schedule_times_inside_next_day_plan():
    reports = load_sample_reports()
    match = next(
        r
        for r in reports
        if "川上" in r.reporter_name and r.date == "2026-07-20" and r.time == "18:00"
    )
    assert "PDCA面談" in match.values["next_day_plan"]
    assert "管理者会議" in match.values["next_day_plan"]


def test_time_only_boundary_reuses_previous_date():
    reports = load_sample_reports()
    match = next(
        r for r in reports if "堀尾" in r.reporter_name and r.time == "08:15"
    )
    assert match.date == "2026-07-23"


def test_filter_reports_by_facility():
    reports = load_sample_reports()
    filtered = filter_reports(reports, facilities=["南志賀"])
    assert filtered
    assert all("南志賀" in r.facility_key for r in filtered)


def test_resolve_field_keys_rejects_unknown_key():
    try:
        resolve_field_keys(["not_a_real_field"])
    except ValueError:
        pass
    else:
        raise AssertionError("expected ValueError for unknown field key")


def test_daily_summary_merges_same_day_messages():
    reports = load_sample_reports()
    daily = build_daily_summary(reports)
    row = next(r for r in daily if r["facility_key"] == "宇都宮" and r["date"] == "2026-07-20")
    # 朝(営業累計時間)と夕方(短期稼働日数・問合せ件数系)の項目が両方1行にマージされていること
    assert row["hours_cumulative"]
    assert row["work_days"]


def test_find_missing_facilities():
    reports = load_sample_reports()
    missing = find_missing_facilities(reports, "2026-07-23")
    assert "昇町" in missing
    assert "南志賀" not in missing
