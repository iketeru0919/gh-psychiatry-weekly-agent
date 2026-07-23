from __future__ import annotations

import csv
from pathlib import Path

from .labels import ALL_FIELD_KEYS, FIELD_DISPLAY_ORDER
from .models import FacilityReport
from .parser import normalize_facility

_ALWAYS_KEYS = {"date", "time", "facility_key", "facility_raw", "reporter_name", "am_name"}


def resolve_field_keys(selected: list[str] | None) -> list[str]:
    """出力に含める項目キーを決定する。未指定なら全項目。"""
    if not selected:
        return list(ALL_FIELD_KEYS)
    unknown = [key for key in selected if key not in ALL_FIELD_KEYS]
    if unknown:
        valid = ", ".join(ALL_FIELD_KEYS)
        raise ValueError(f"未知の項目キーです: {unknown}. 指定可能な項目: {valid}")
    # ALL_FIELD_KEYSの並び順を保ちつつ、必ず基本列(日付・拠点名等)は含める。
    chosen = set(selected) | _ALWAYS_KEYS
    return [key for key in ALL_FIELD_KEYS if key in chosen]


def filter_reports(
    reports: list[FacilityReport], facilities: list[str] | None = None
) -> list[FacilityReport]:
    """指定した拠点名(部分一致・表記ゆれ吸収後)だけに絞り込む。"""
    if not facilities:
        return list(reports)
    queries = [normalize_facility(f).lower() for f in facilities]
    filtered = []
    for r in reports:
        key = r.facility_key.lower()
        raw = r.facility_raw.lower()
        if any(q in key or q in raw for q in queries):
            filtered.append(r)
    return filtered


def _report_row(report: FacilityReport, field_keys: list[str]) -> dict[str, str]:
    row = {
        "date": report.date or "",
        "time": report.time or "",
        "facility_key": report.facility_key,
        "facility_raw": report.facility_raw,
        "reporter_name": report.reporter_name,
        "am_name": report.am_name or "",
    }
    row.update(report.values)
    return {key: row.get(key, "") for key in field_keys}


def write_raw_csv(reports: list[FacilityReport], path: str | Path, field_keys: list[str] | None = None) -> Path:
    """1メッセージ1行の生データCSV(監査・検証用)を書き出す。"""
    keys = field_keys or list(ALL_FIELD_KEYS)
    headers = {key: label for key, label in FIELD_DISPLAY_ORDER}
    out_path = Path(path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=keys)
        writer.writerow({key: headers.get(key, key) for key in keys})
        for report in sorted(reports, key=lambda r: (r.date or "", r.time or "", r.facility_key)):
            writer.writerow(_report_row(report, keys))
    return out_path


def build_daily_summary(reports: list[FacilityReport]) -> list[dict[str, str]]:
    """拠点×日付で当日の全メッセージをマージし、1行に集約する。

    同じ日に複数回報告される項目(累計時間など)は、その日の最後の報告値を採用する。
    朝の報告にしかない項目(カメラ確認等)と夕方の報告にしかない項目(問合せ件数等)は
    両方とも保持される。
    """
    groups: dict[tuple[str, str], list[FacilityReport]] = {}
    order: list[tuple[str, str]] = []
    for r in reports:
        if not r.date or not r.facility_key:
            continue
        key = (r.facility_key, r.date)
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(r)

    summaries: list[dict[str, str]] = []
    for facility_key, date in order:
        msgs = sorted(groups[(facility_key, date)], key=lambda r: r.time or "")
        merged: dict[str, str] = {}
        facility_raw = ""
        reporter_names: list[str] = []
        am_names: list[str] = []
        for r in msgs:
            if r.facility_raw:
                facility_raw = r.facility_raw
            if r.reporter_name and r.reporter_name not in reporter_names:
                reporter_names.append(r.reporter_name)
            if r.am_name and r.am_name not in am_names:
                am_names.append(r.am_name)
            for k, v in r.values.items():
                if v:
                    merged[k] = v
        summaries.append(
            {
                "date": date,
                "facility_key": facility_key,
                "facility_raw": facility_raw,
                "reporter_name": "・".join(reporter_names),
                "am_name": "・".join(am_names),
                "time": msgs[-1].time or "",
                "message_count": str(len(msgs)),
                **merged,
            }
        )
    return summaries


def write_daily_csv(
    summaries: list[dict[str, str]], path: str | Path, field_keys: list[str] | None = None
) -> Path:
    keys = field_keys or list(ALL_FIELD_KEYS)
    if "message_count" not in keys:
        keys = [*keys, "message_count"]
    headers = {key: label for key, label in FIELD_DISPLAY_ORDER}
    headers["message_count"] = "当日報告件数"
    out_path = Path(path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=keys)
        writer.writerow({key: headers.get(key, key) for key in keys})
        for row in sorted(summaries, key=lambda r: (r["date"], r["facility_key"])):
            writer.writerow({key: row.get(key, "") for key in keys})
    return out_path


def find_missing_facilities(reports: list[FacilityReport], date: str) -> list[str]:
    """全期間で一度でも報告のあった拠点のうち、指定日に報告が無い拠点を返す。

    上司報告での「報告漏れ拠点」の確認に使う想定。
    """
    all_facilities = {r.facility_key for r in reports if r.facility_key}
    reported_today = {r.facility_key for r in reports if r.date == date and r.facility_key}
    return sorted(all_facilities - reported_today)


def list_unknown_labels(reports: list[FacilityReport]) -> dict[str, int]:
    """ラベル辞書に無い項目の出現回数を集計する。フォーマット改善の材料に使う。"""
    counts: dict[str, int] = {}
    for r in reports:
        for label in r.unknown_fields:
            counts[label] = counts.get(label, 0) + 1
    return dict(sorted(counts.items(), key=lambda kv: -kv[1]))
