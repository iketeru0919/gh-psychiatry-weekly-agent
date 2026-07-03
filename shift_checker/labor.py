"""労務管理チェック。

職員を複数行掲載（兼務）も含めて人物単位に統合し、以下を判定する:
  1. 連続勤務日数（既定: 6日超で重要）
  2. 夜勤明けの即日勤（重要）
  3. 夜勤回数の偏り（既定: 8回超で要確認）・深夜業4回以上の健診対象目安
  4. 公休数の不足（施設別基準が設定されている場合のみ判定）
  5. 二重割当（同日に異なる勤務記号が複数行に入っている場合のみ要確認）
  6. 時間外労働（所定超過 既定: 30h以上=要確認 / 45h以上=重要）
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .config import Settings
from .parser import ShiftBook
from .util import base_name_key, format_signed, matches_exact, normalize, round2

SEVERITY_ORDER = {"critical": 0, "warning": 1, "ok": 2}
SEVERITY_LABEL = {"critical": "重要", "warning": "要確認", "ok": "問題なし"}

# 記号の区分
KIND_PAID = "有休"
KIND_VACATION = "公休"
KIND_AKE = "明け"
KIND_NIGHT = "夜勤"
KIND_WORK = "勤務"
KIND_EMPTY = ""


@dataclass
class LaborRow:
    facility: str
    severity: str
    name: str
    job: str
    employment_type: str
    row_count: int
    total_hours: float | None
    prescribed: float | None
    overtime: float | None
    max_consecutive: int
    consecutive_detail: str
    night_count: int
    night_check_target: str   # 深夜業4回以上（健診対象の目安）
    ake_violations: list[str] = field(default_factory=list)
    holiday_count: int = 0
    paid_count: int = 0
    required_holidays: object = None  # 基準未設定ならNone
    double_bookings: list[str] = field(default_factory=list)
    reasons: list[str] = field(default_factory=list)

    @property
    def judgment(self) -> str:
        return SEVERITY_LABEL[self.severity]


def classify_symbol(value: str, settings: Settings) -> str:
    text = normalize(value)
    if not text:
        return KIND_EMPTY
    if matches_exact(text, settings.paid_words):
        return KIND_PAID
    if matches_exact(text, settings.vacation_words):
        return KIND_VACATION
    # 「明」を含む記号は夜勤明け（「深夜明」等は夜勤より明けを優先）
    if any(k in text for k in settings.ake_keywords):
        return KIND_AKE
    if any(k in text for k in settings.night_keywords):
        return KIND_NIGHT
    return KIND_WORK


def required_holidays_for(facility: str, month: int, settings: Settings):
    """施設名にキーワードが含まれる設定を探す。int または {"月": 日数} 形式に対応。"""
    for keyword, value in settings.required_holidays.items():
        if keyword and keyword in facility:
            if isinstance(value, dict):
                return value.get(str(month))
            return value
    default = settings.required_holidays.get("default")
    if isinstance(default, dict):
        return default.get(str(month))
    return default


def analyze_labor(book: ShiftBook, facility: str, settings: Settings,
                  month: int) -> list[LaborRow]:
    # --- 人物単位に統合（兼務の複数行を束ねる）---
    persons: dict[str, list] = {}
    for staff in book.staff.values():
        persons.setdefault(base_name_key(staff.name), []).append(staff)

    required = required_holidays_for(facility, month, settings)
    rows: list[LaborRow] = []

    for person_rows in persons.values():
        person_rows.sort(key=lambda s: s.row)
        representative = person_rows[0]

        # タイミー等のスポット枠は複数人の集合体のため個人単位の労務判定をしない
        if any(k and k in representative.name for k in settings.labor_exclude_keywords):
            continue

        # 鏡写し行の除去: シフト内容が完全一致する行は同じ勤務の重複記載とみなし、
        # 総労働時間の合算から除く（分割記載の行はそのまま合算する）
        unique_rows = []
        seen_patterns = set()
        for staff in person_rows:
            pattern = frozenset(
                (cell.day, cell.value) for cell in staff.shifts.values() if cell.value
            )
            if pattern and pattern in seen_patterns:
                continue
            seen_patterns.add(pattern)
            unique_rows.append(staff)
        # 日別に全行の記号を集める（day番号 -> [(値, 区分)]）
        days: dict[int, list[tuple[str, str]]] = {}
        date_labels: dict[int, str] = {}
        for staff in unique_rows:
            for cell in staff.shifts.values():
                if cell.day is None:
                    continue
                kind = classify_symbol(cell.value, settings)
                days.setdefault(cell.day, [])
                date_labels.setdefault(cell.day, cell.date_label)
                if kind != KIND_EMPTY:
                    days[cell.day].append((cell.value, kind))
        ordered_days = sorted(days)

        # 日単位の人物ステータス
        def day_kinds(day: int) -> set[str]:
            return {kind for _, kind in days[day]}

        def is_work_day(day: int) -> bool:
            return bool(day_kinds(day) & {KIND_WORK, KIND_NIGHT, KIND_AKE})

        # 1. 連続勤務
        max_consecutive = 0
        consecutive_detail = ""
        streak_start = None
        streak = 0
        for day in ordered_days:
            if is_work_day(day):
                if streak == 0:
                    streak_start = day
                streak += 1
                if streak > max_consecutive:
                    max_consecutive = streak
                    consecutive_detail = (
                        f"{date_labels[streak_start]}〜{date_labels[day]}（{streak}日）"
                    )
            else:
                streak = 0

        # 2. 夜勤明けの即日勤（夜勤の翌日に日勤系が入っている場合のみ。
        #    夜勤→夜勤の連続は宿直運用で通常のため警告しない）
        ake_violations = []
        for day in ordered_days:
            if KIND_NIGHT not in day_kinds(day):
                continue
            next_day = day + 1
            if next_day not in days:
                continue
            next_kinds = day_kinds(next_day)
            if KIND_AKE in next_kinds or KIND_WORK not in next_kinds:
                continue
            next_values = "、".join(v for v, k in days[next_day] if k == KIND_WORK)
            ake_violations.append(
                f"{date_labels[day]}夜勤 → {date_labels[next_day]}「{next_values}」"
            )

        # 3. 夜勤回数
        night_count = sum(1 for day in ordered_days if KIND_NIGHT in day_kinds(day))

        # 4. 公休・有休（人物単位: 勤務が無く休み記号がある日）
        holiday_count = sum(
            1 for day in ordered_days
            if KIND_VACATION in day_kinds(day) and not is_work_day(day)
        )
        paid_count = sum(
            1 for day in ordered_days
            if KIND_PAID in day_kinds(day) and not is_work_day(day)
        )

        # 5. 二重割当（同日に異なる勤務記号。同一記号の重複記載や夜勤+明けは正常）
        double_bookings = []
        for day in ordered_days:
            work_values = {v for v, k in days[day] if k in (KIND_WORK, KIND_NIGHT)}
            if len(work_values) >= 2:
                double_bookings.append(f"{date_labels[day]}（{'＋'.join(sorted(work_values))}）")

        # 6. 時間外（常勤のみ、鏡写し行を除いた総労働時間の合算）
        totals = [s.total_hours for s in unique_rows if s.total_hours is not None]
        total_hours = round2(sum(totals)) if totals else None
        prescribed = representative.prescribed_hours
        fulltime = any(normalize(s.fulltime) == "常勤" for s in person_rows)
        overtime = None
        if fulltime and total_hours is not None and prescribed is not None:
            overtime = round2(total_hours - prescribed)

        # --- 判定 ---
        reasons = []
        severity = "ok"

        def raise_to(level: str):
            nonlocal severity
            if SEVERITY_ORDER[level] < SEVERITY_ORDER[severity]:
                severity = level

        if max_consecutive > settings.max_consecutive_days:
            raise_to("critical")
            reasons.append(f"連続勤務{max_consecutive}日 {consecutive_detail}")
        if ake_violations:
            raise_to("critical")
            reasons.append(f"夜勤明け即日勤 {len(ake_violations)}件（{'、'.join(ake_violations[:3])}）")
        if overtime is not None and overtime >= settings.overtime_warning:
            raise_to("critical")
            reasons.append(f"所定超過 {format_signed(overtime)}時間（45時間基準超）")
        elif overtime is not None and overtime >= settings.overtime_caution:
            raise_to("warning")
            reasons.append(f"所定超過 {format_signed(overtime)}時間（30時間注意）")
        if night_count > settings.night_shift_limit:
            raise_to("warning")
            reasons.append(f"夜勤{night_count}回（上限目安{settings.night_shift_limit}回超）")
        if required is not None and holiday_count < required:
            raise_to("warning")
            reasons.append(f"公休{holiday_count}日（基準{required}日に不足）")
        if double_bookings:
            raise_to("warning")
            reasons.append(f"同日に異なる勤務 {len(double_bookings)}件（{'、'.join(double_bookings[:3])}）")

        rows.append(LaborRow(
            facility=facility,
            severity=severity,
            name=representative.name,
            job=representative.job,
            employment_type=representative.fulltime,
            row_count=len(unique_rows),
            total_hours=total_hours,
            prescribed=prescribed,
            overtime=overtime,
            max_consecutive=max_consecutive,
            consecutive_detail=consecutive_detail,
            night_count=night_count,
            night_check_target="○" if night_count >= 4 else "",
            ake_violations=ake_violations,
            holiday_count=holiday_count,
            paid_count=paid_count,
            required_holidays=required,
            double_bookings=double_bookings,
            reasons=reasons,
        ))

    rows.sort(key=lambda r: (SEVERITY_ORDER[r.severity], r.name))
    return rows
