"""労務推移データの蓄積（月次集計CSV）と推移表の作成。

CSVは <snapshot_dir>/労務推移データ.csv に保存し、実行のたびに
同じ年月・施設の行を最新の集計で置き換える（他の月は保持）。
"""

from __future__ import annotations

import csv
from pathlib import Path

from .labor import LaborRow

HISTORY_FILENAME = "労務推移データ.csv"
FIELDS = ["年月", "施設", "職員氏名", "総労働時間", "所定超過", "夜勤回数", "公休数", "有休数"]


def history_path(snapshot_dir: Path) -> Path:
    return snapshot_dir / HISTORY_FILENAME


def load_history(snapshot_dir: Path) -> list[dict]:
    path = history_path(snapshot_dir)
    if not path.exists():
        return []
    try:
        with open(path, encoding="utf-8-sig", newline="") as handle:
            return [row for row in csv.DictReader(handle)]
    except (OSError, csv.Error):
        return []


def update_history(snapshot_dir: Path, month_key: str,
                   labor_rows: list[LaborRow]) -> list[dict]:
    """今回実行分の月次集計でCSVを更新し、更新後の全行を返す。"""
    records = load_history(snapshot_dir)
    updated_facilities = {row.facility for row in labor_rows}
    records = [
        r for r in records
        if not (r.get("年月") == month_key and r.get("施設") in updated_facilities)
    ]
    for row in labor_rows:
        records.append({
            "年月": month_key,
            "施設": row.facility,
            "職員氏名": row.name,
            "総労働時間": "" if row.total_hours is None else row.total_hours,
            "所定超過": "" if row.overtime is None else row.overtime,
            "夜勤回数": row.night_count,
            "公休数": row.holiday_count,
            "有休数": row.paid_count,
        })
    records.sort(key=lambda r: (str(r.get("年月")), str(r.get("施設")), str(r.get("職員氏名"))))
    snapshot_dir.mkdir(parents=True, exist_ok=True)
    with open(history_path(snapshot_dir), "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(records)
    return records


def build_trend(records: list[dict], metrics=("夜勤回数", "所定超過"),
                max_months: int = 6) -> tuple[list[str], list[list]]:
    """推移表（施設 | 職員 | 指標 | 月1..月N）のヘッダーと行を作る。"""
    months = sorted({r["年月"] for r in records if r.get("年月")})[-max_months:]
    if not months:
        return [], []
    by_person: dict[tuple[str, str], dict[str, dict]] = {}
    for record in records:
        if record.get("年月") not in months:
            continue
        key = (record.get("施設", ""), record.get("職員氏名", ""))
        by_person.setdefault(key, {})[record["年月"]] = record
    headers = ["施設", "職員氏名", "指標"] + [f"{m[:2]}年{int(m[2:])}月" for m in months]
    rows = []
    for (facility, name), monthly in sorted(by_person.items()):
        for metric in metrics:
            values = [monthly.get(m, {}).get(metric, "") for m in months]
            if any(v not in ("", None) for v in values):
                rows.append([facility, name, metric] + values)
    return headers, rows
