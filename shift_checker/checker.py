"""チェックロジック本体（今回のみチェック＋前回スナップショット比較）。"""

from __future__ import annotations

from dataclasses import dataclass, field

from .config import Settings
from .parser import ShiftBook
from .util import format_signed, matches_exact, normalize, round2

SEVERITY_ORDER = {"critical": 0, "warning": 1, "ok": 2}
SEVERITY_LABEL = {"critical": "重要", "warning": "要確認", "ok": "問題なし"}


@dataclass
class StaffRow:
    facility: str
    severity: str
    name: str
    status: str
    employment_type: str
    job: str
    prescribed: float | None
    before_hours: float | None
    after_hours: float | None
    hour_delta: float | None
    required_delta: float | None
    before_paid: int
    after_paid: int
    paid_delta: int
    before_vacation: int
    after_vacation: int
    vacation_delta: int
    change_count: int
    reason: str

    @property
    def judgment(self) -> str:
        return SEVERITY_LABEL[self.severity]


@dataclass
class ChangeRow:
    facility: str
    name: str
    date_label: str
    weekday: str
    previous_value: str
    current_value: str
    change_type: str
    cell: str


@dataclass
class FacilityResult:
    facility: str
    status: str = "OK"           # OK / 読取エラー
    error: str = ""
    warnings: list[str] = field(default_factory=list)
    source_path: str = ""
    previous_path: str = ""
    duplicate_files: int = 0
    previous: ShiftBook | None = None
    current: ShiftBook | None = None
    staff_rows: list[StaffRow] = field(default_factory=list)
    change_rows: list[ChangeRow] = field(default_factory=list)

    @property
    def alerts(self) -> list[StaffRow]:
        return [row for row in self.staff_rows if row.severity != "ok"]

    @property
    def critical_count(self) -> int:
        return sum(1 for row in self.staff_rows if row.severity == "critical")


def is_fulltime(value: str) -> bool:
    return normalize(value) == "常勤"


def _numeric(value):
    return value if isinstance(value, (int, float)) and not isinstance(value, bool) else None


def _sort_rows(rows: list[StaffRow]) -> None:
    rows.sort(key=lambda r: (SEVERITY_ORDER[r.severity], r.name))


def _incomplete_note(required_delta, settings: Settings) -> str:
    if required_delta is not None and required_delta <= -settings.incomplete_threshold:
        return "（入力途中の可能性）"
    return ""


def build_current_result(book: ShiftBook, facility: str, settings: Settings) -> FacilityResult:
    """今回のみチェック：所定不足と日中配置数の警告。"""
    result = FacilityResult(
        facility=facility,
        current=book,
        source_path=str(book.path),
        warnings=list(book.warnings),
    )
    for person in book.staff.values():
        required_delta = None
        if is_fulltime(person.fulltime) and person.total_hours is not None and person.prescribed_hours is not None:
            required_delta = round2(person.total_hours - person.prescribed_hours)
        reasons = []
        severity = "ok"
        if required_delta is not None and required_delta < -0.5:
            severity = "critical"
            reasons.append(
                f"所定時間不足 {format_signed(required_delta)}時間{_incomplete_note(required_delta, settings)}"
            )
        allocation = _numeric(person.day_allocation)
        if allocation is not None and allocation > 1:
            if severity != "critical":
                severity = "warning"
            reasons.append(f"日中配置数 {allocation:g}")
        result.staff_rows.append(StaffRow(
            facility=facility,
            severity=severity,
            name=person.name,
            status="",
            employment_type=person.fulltime,
            job=person.job,
            prescribed=person.prescribed_hours,
            before_hours=None,
            after_hours=person.total_hours,
            hour_delta=None,
            required_delta=required_delta,
            before_paid=person.paid_count,
            after_paid=person.paid_count,
            paid_delta=0,
            before_vacation=person.vacation_count,
            after_vacation=person.vacation_count,
            vacation_delta=0,
            change_count=0,
            reason=" / ".join(reasons) or "確認事項なし",
        ))
    # 同一人物の複数行掲載（兼務）は正常運用のため警告にしない。
    # 照合は氏名の出現順で対応付け済み。
    _sort_rows(result.staff_rows)
    return result


def classify_change(previous_value: str, current_value: str, settings: Settings) -> str:
    before = normalize(previous_value)
    after = normalize(current_value)
    if not before and after:
        return "追加"
    if before and not after:
        return "削除"
    if matches_exact(before, settings.leave_words) or matches_exact(after, settings.leave_words):
        return "休暇変更"
    return "勤務変更"


def compare_person_shifts(before, after, name: str, facility: str, settings: Settings) -> list[ChangeRow]:
    keys = set()
    if before:
        keys |= set(before.shifts.keys())
    if after:
        keys |= set(after.shifts.keys())
    rows = []
    for key in keys:
        old_shift = before.shifts.get(key) if before else None
        new_shift = after.shifts.get(key) if after else None
        previous_value = old_shift.value if old_shift else ""
        current_value = new_shift.value if new_shift else ""
        if normalize(previous_value) == normalize(current_value):
            continue
        base = new_shift or old_shift
        rows.append(ChangeRow(
            facility=facility,
            name=name,
            date_label=base.date_label,
            weekday=base.weekday,
            previous_value=previous_value,
            current_value=current_value,
            change_type=classify_change(previous_value, current_value, settings),
            cell=new_shift.cell if new_shift else "",
        ))
    rows.sort(key=lambda r: (r.date_label.rjust(6), r.name))
    return rows


def build_comparison(previous: ShiftBook, current: ShiftBook, facility: str,
                     settings: Settings) -> FacilityResult:
    """前回スナップショットとの比較チェック。"""
    result = FacilityResult(
        facility=facility,
        previous=previous,
        current=current,
        source_path=str(current.path),
        previous_path=str(previous.path),
        warnings=list(current.warnings),
    )
    if previous.month_label and current.month_label and previous.month_label != current.month_label:
        result.warnings.append(
            f"対象年月注意：前回 {previous.month_label} / 今回 {current.month_label}"
        )

    keys = sorted(set(previous.staff.keys()) | set(current.staff.keys()))
    for key in keys:
        before = previous.staff.get(key)
        after = current.staff.get(key)
        name = after.name if after else before.name
        status = "職員追加" if not before else "職員削除" if not after else ""
        if before and after and before.name != after.name:
            status = "表記変更"
        prescribed = after.prescribed_hours if after else before.prescribed_hours
        fulltime = bool(after and is_fulltime(after.fulltime))
        before_hours = before.total_hours if before else None
        after_hours = after.total_hours if after else None
        hour_delta = (
            round2(after_hours - before_hours)
            if before_hours is not None and after_hours is not None else None
        )
        required_delta = (
            round2(after_hours - prescribed)
            if fulltime and after_hours is not None and prescribed is not None else None
        )
        before_paid = before.paid_count if before else 0
        after_paid = after.paid_count if after else 0
        before_vacation = before.vacation_count if before else 0
        after_vacation = after.vacation_count if after else 0
        paid_delta = after_paid - before_paid
        vacation_delta = after_vacation - before_vacation
        person_changes = compare_person_shifts(before, after, name, facility, settings)
        result.change_rows.extend(person_changes)

        reasons = []
        severity = "ok"
        if status in ("職員追加", "職員削除"):
            severity = "warning"
            reasons.append(status)
        elif status == "表記変更":
            severity = "warning"
            reasons.append(f"氏名の表記変更（{before.name} → {after.name}）")
        if required_delta is not None and required_delta < -0.5:
            severity = "critical"
            reasons.append(
                f"所定時間不足 {format_signed(required_delta)}時間{_incomplete_note(required_delta, settings)}"
            )
        if hour_delta is not None and settings.hour_threshold > 0 and abs(hour_delta) >= settings.hour_threshold:
            if severity != "critical":
                severity = "warning"
            reasons.append(f"総労働時間 {format_signed(hour_delta)}時間")
        if paid_delta != 0:
            if severity == "ok":
                severity = "warning"
            reasons.append(f"有休 {format_signed(paid_delta)}日")
        if vacation_delta != 0:
            if severity == "ok":
                severity = "warning"
            reasons.append(f"公休 {format_signed(vacation_delta)}日")
        if person_changes and severity == "ok":
            severity = "warning"
        if person_changes:
            reasons.append(f"日別変更 {len(person_changes)}件")

        result.staff_rows.append(StaffRow(
            facility=facility,
            severity=severity,
            name=name,
            status=status,
            employment_type=(after.fulltime if after else before.fulltime),
            job=(after.job if after else before.job),
            prescribed=prescribed,
            before_hours=before_hours,
            after_hours=after_hours,
            hour_delta=hour_delta,
            required_delta=required_delta,
            before_paid=before_paid,
            after_paid=after_paid,
            paid_delta=paid_delta,
            before_vacation=before_vacation,
            after_vacation=after_vacation,
            vacation_delta=vacation_delta,
            change_count=len(person_changes),
            reason=" / ".join(reasons) or "変更なし",
        ))

    _sort_rows(result.staff_rows)
    return result


def build_error_result(facility: str, path, message: str) -> FacilityResult:
    return FacilityResult(
        facility=facility,
        status="読取エラー",
        error=message,
        source_path=str(path) if path else "",
    )
