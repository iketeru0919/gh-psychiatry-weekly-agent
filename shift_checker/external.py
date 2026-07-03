"""外部人材（タイミー・派遣）の整理。

氏名に判定キーワードを含む職員行を外部人材として抽出する（HTML版の
「外部人材（職員氏名による判定）」と同じ方式）。
・記載名ごとに1行（日勤タイミー / 夜勤タイミー / タイミー日勤② はそれぞれ別項目）
・括弧書きで個人名があっても記載名のままタイミーに含める
・施設ごとの区分計と、全施設の区分計を出す
"""

from __future__ import annotations

from dataclasses import dataclass

from .config import Settings
from .labor import KIND_AKE, KIND_NIGHT, KIND_PAID, KIND_VACATION, KIND_WORK, classify_symbol
from .parser import ShiftBook
from .util import normalize, round2


@dataclass
class ExternalRow:
    facility: str
    category: str        # タイミー / 派遣
    name: str            # 記載名（枠名）。計の行は「タイミー計」等
    is_total: bool
    work_days: int | None
    total_hours: float | None
    night_count: int | None
    paid_count: int | None
    vacation_count: int | None
    previous_hours: float | None
    hour_delta: float | None
    change_count: int | None
    dependency: str      # 外部依存度（計の行のみ）
    note: str


def external_category(name: str, settings: Settings) -> str | None:
    text = normalize(name)
    for keyword in settings.external_keywords:
        if keyword and keyword in text:
            return keyword
    return None


def _dedupe_rows(rows):
    """シフト内容が完全一致する鏡写し行を除く（labor.pyと同じ規則）。"""
    unique, seen = [], set()
    for staff in sorted(rows, key=lambda s: s.row):
        pattern = frozenset((c.day, c.value) for c in staff.shifts.values() if c.value)
        if pattern and pattern in seen:
            continue
        seen.add(pattern)
        unique.append(staff)
    return unique


def _aggregate(rows, settings: Settings):
    """記載名グループの (勤務日数, 総労働, 夜勤, 有休, 公休) を集計する。"""
    unique = _dedupe_rows(rows)
    days: dict[int, set[str]] = {}
    for staff in unique:
        for cell in staff.shifts.values():
            if cell.day is None or not cell.value:
                continue
            days.setdefault(cell.day, set()).add(classify_symbol(cell.value, settings))
    work_days = sum(1 for kinds in days.values() if kinds & {KIND_WORK, KIND_NIGHT, KIND_AKE})
    night = sum(1 for kinds in days.values() if KIND_NIGHT in kinds)
    paid = sum(1 for kinds in days.values()
               if KIND_PAID in kinds and not kinds & {KIND_WORK, KIND_NIGHT, KIND_AKE})
    vacation = sum(1 for kinds in days.values()
                   if KIND_VACATION in kinds and not kinds & {KIND_WORK, KIND_NIGHT, KIND_AKE})
    totals = [s.total_hours for s in unique if s.total_hours is not None]
    total_hours = round2(sum(totals)) if totals else None
    return work_days, total_hours, night, paid, vacation


def _facility_total_hours(book: ShiftBook) -> float:
    """施設全体の総労働時間（鏡写し行を除いて合算）。依存度の分母に使う。"""
    from .util import base_name_key
    persons: dict[str, list] = {}
    for staff in book.staff.values():
        persons.setdefault(base_name_key(staff.name), []).append(staff)
    total = 0.0
    for rows in persons.values():
        for staff in _dedupe_rows(rows):
            if staff.total_hours is not None:
                total += staff.total_hours
    return round2(total)


def _group_by_name(book: ShiftBook, settings: Settings) -> dict[tuple[str, str], list]:
    """外部人材の職員行を (区分, 記載名) でグループ化する。"""
    groups: dict[tuple[str, str], list] = {}
    for staff in book.staff.values():
        category = external_category(staff.name, settings)
        if category:
            groups.setdefault((category, normalize(staff.name)), []).append(staff)
    return groups


def build_external_rows(results, settings: Settings) -> list[ExternalRow]:
    all_rows: list[ExternalRow] = []
    grand: dict[str, dict] = {}  # 区分 -> 全施設合計

    for result in results:
        if not result.current:
            continue
        current_groups = _group_by_name(result.current, settings)
        previous_groups = _group_by_name(result.previous, settings) if result.previous else {}
        if not current_groups and not previous_groups:
            continue

        facility_hours = _facility_total_hours(result.current)
        changes_by_name: dict[str, int] = {}
        for change in result.change_rows:
            changes_by_name[normalize(change.name)] = changes_by_name.get(normalize(change.name), 0) + 1

        facility_totals: dict[str, dict] = {}
        keys = sorted(set(current_groups) | set(previous_groups))
        for category, name in keys:
            cur_rows = current_groups.get((category, name))
            prev_rows = previous_groups.get((category, name))
            work_days = total_hours = night = paid = vacation = None
            if cur_rows:
                work_days, total_hours, night, paid, vacation = _aggregate(cur_rows, settings)
            previous_hours = None
            if prev_rows:
                _, previous_hours, _, _, _ = _aggregate(prev_rows, settings)
            delta = None
            if total_hours is not None and previous_hours is not None:
                delta = round2(total_hours - previous_hours)
            note = ""
            if cur_rows and result.previous and not prev_rows:
                note = "新規"
            elif not cur_rows:
                note = "今回なし（前回のみ）"
            all_rows.append(ExternalRow(
                facility=result.facility, category=category, name=name, is_total=False,
                work_days=work_days, total_hours=total_hours, night_count=night,
                paid_count=paid, vacation_count=vacation,
                previous_hours=previous_hours, hour_delta=delta,
                change_count=changes_by_name.get(name) if result.previous else None,
                dependency="", note=note,
            ))
            bucket = facility_totals.setdefault(category, dict(
                count=0, days=0, hours=0.0, night=0, prev_count=0, prev_hours=0.0, changes=0))
            if cur_rows:
                bucket["count"] += 1
                bucket["days"] += work_days or 0
                bucket["hours"] += total_hours or 0
                bucket["night"] += night or 0
                bucket["changes"] += changes_by_name.get(name, 0)
            if prev_rows:
                bucket["prev_count"] += 1
                bucket["prev_hours"] += previous_hours or 0

        for category in sorted(facility_totals):
            bucket = facility_totals[category]
            dependency = (
                f"{round2(bucket['hours'] / facility_hours * 100):g}%"
                if facility_hours else ""
            )
            all_rows.append(ExternalRow(
                facility=result.facility, category=category,
                name=f"{category}計（{bucket['count']}枠）", is_total=True,
                work_days=bucket["days"], total_hours=round2(bucket["hours"]),
                night_count=bucket["night"], paid_count=None, vacation_count=None,
                previous_hours=round2(bucket["prev_hours"]) if result.previous else None,
                hour_delta=(round2(bucket["hours"] - bucket["prev_hours"])
                            if result.previous else None),
                change_count=bucket["changes"] if result.previous else None,
                dependency=dependency, note="",
            ))
            g = grand.setdefault(category, dict(count=0, days=0, hours=0.0, night=0,
                                                prev_hours=0.0, changes=0, compared=False))
            g["count"] += bucket["count"]
            g["days"] += bucket["days"]
            g["hours"] += bucket["hours"]
            g["night"] += bucket["night"]
            g["prev_hours"] += bucket["prev_hours"]
            g["changes"] += bucket["changes"]
            g["compared"] = g["compared"] or bool(result.previous)

    for category in sorted(grand):
        g = grand[category]
        all_rows.append(ExternalRow(
            facility="全施設", category=category,
            name=f"{category}合計（{g['count']}枠）", is_total=True,
            work_days=g["days"], total_hours=round2(g["hours"]), night_count=g["night"],
            paid_count=None, vacation_count=None,
            previous_hours=round2(g["prev_hours"]) if g["compared"] else None,
            hour_delta=round2(g["hours"] - g["prev_hours"]) if g["compared"] else None,
            change_count=g["changes"] if g["compared"] else None,
            dependency="", note="",
        ))
    return all_rows
