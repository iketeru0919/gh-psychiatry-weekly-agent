from __future__ import annotations

import re

from .labels import FACILITY_NOISE_TOKENS, LABELS_BY_LENGTH_DESC, LABEL_TO_FIELD
from .models import FacilityReport

_CIRCLED_NUM_CHARS = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮➀➁➂➃➄➅➆➇➈➉"
_MARKER_RE = re.compile(f"[{re.escape(_CIRCLED_NUM_CHARS)}]")
_NEXT_DAY_MARKER = "【翌日の予定】"

_DATE_TOKEN_RE = re.compile(r"^(?:(\d{4})年)?(\d{1,2})月(\d{1,2})日")
_TIME_TOKEN_RE = re.compile(r"^(\d{1,2}):(\d{2})")

_HEADER_TO_RE = re.compile(r"\[To:(?P<to_id>\d+)\]\s*(?P<to_name>[^\n]*?)さん\s*$")
_COMPANY_MARKERS = ("株式会社", "有限会社", "（", "(", "【")


def _strip_leading_timestamp(
    line: str,
) -> tuple[tuple[str | None, str, str] | None, tuple[str, str] | None, str]:
    """行頭から日付・時刻トークンを可能な限り連続で取り除く。

    LINEのコピペ由来のテキストでは、直前メッセージの末尾に次メッセージの
    タイムスタンプが連結され、さらに日付変更区切り(例:「2026年7月21日」)が
    そのまま次の送信者名に連結されることがある。そのため単発マッチではなく
    ループで貪欲に剥がす。
    """
    date_found: tuple[str | None, str, str] | None = None
    time_found: tuple[str, str] | None = None
    while True:
        m = _DATE_TOKEN_RE.match(line)
        if m:
            date_found = (m.group(1), m.group(2), m.group(3))
            line = line[m.end():].lstrip(" 　")
            continue
        m2 = _TIME_TOKEN_RE.match(line)
        if m2 and time_found is None:
            time_found = (m2.group(1), m2.group(2))
            line = line[m2.end():]
            continue
        break
    return date_found, time_found, line


def _looks_like_message_header(remainder: str) -> bool:
    stripped = remainder.rstrip()
    return "[To:" in remainder or stripped.endswith("さん")


def split_into_messages(text: str) -> list[dict]:
    """チャット全文を1メッセージ単位のブロックに分割する。"""
    messages: list[dict] = []
    current: dict | None = None

    for line in text.split("\n"):
        date_tok, time_tok, remainder = _strip_leading_timestamp(line)
        is_boundary = (date_tok is not None or time_tok is not None) and _looks_like_message_header(
            remainder
        )
        if is_boundary:
            if current is not None:
                messages.append(current)
            current = {"date": date_tok, "time": time_tok, "header": remainder, "body_lines": []}
        elif current is None:
            # 先頭行(最初のメッセージ)にはタイムスタンプが付いていないため、
            # ここが最初のヘッダー行になる。
            current = {"date": None, "time": None, "header": line, "body_lines": []}
        else:
            current["body_lines"].append(line)

    if current is not None:
        messages.append(current)
    return messages


def _format_date(date_tok: tuple[str | None, str, str] | None, default_year: str | None) -> str | None:
    if date_tok is None:
        return None
    year, month, day = date_tok
    year = year or default_year
    if year is None:
        return f"--{int(month):02d}-{int(day):02d}"
    return f"{year}-{int(month):02d}-{int(day):02d}"


def _format_time(time_tok: tuple[str, str] | None) -> str | None:
    if time_tok is None:
        return None
    hour, minute = time_tok
    return f"{int(hour):02d}:{minute}"


def parse_header(header_line: str) -> tuple[str, str, str | None, str | None]:
    """ヘッダー行から (送信者名候補, 会社・拠点補足, 宛先ID, 宛先名) を取り出す。"""
    match = _HEADER_TO_RE.search(header_line)
    to_id = match.group("to_id") if match else None
    to_name = match.group("to_name").strip() if match else None
    pre_part = header_line[: match.start()] if match else header_line

    split_index = -1
    for marker in _COMPANY_MARKERS:
        idx = pre_part.find(marker)
        if idx != -1 and (split_index == -1 or idx < split_index):
            split_index = idx

    if split_index <= 0:
        sender_name = pre_part.strip()
        company_raw = "" if split_index != 0 else pre_part.strip()
        if split_index == 0:
            sender_name = ""
    else:
        sender_name = pre_part[:split_index].strip()
        company_raw = pre_part[split_index:].strip()

    return sender_name, company_raw, to_id, to_name


def _extract_labeled_value(region: str, start: int, next_marker_start: int) -> tuple[str, str] | None:
    for label in LABELS_BY_LENGTH_DESC:
        if region.startswith(label, start):
            value_start = start + len(label)
            while value_start < len(region) and region[value_start] in "：: 　":
                value_start += 1
            value = region[value_start:next_marker_start].strip(" \n\t：:")
            value = re.sub(r"\s+", " ", value)
            return label, value
    return None


def extract_fields(body_text: str) -> tuple[dict[str, str], dict[str, str]]:
    """本文から丸数字ラベルで区切られた項目を抽出する。AIは使わずルールベースで処理する。"""
    values: dict[str, str] = {}
    unknown: dict[str, str] = {}

    next_day_idx = body_text.find(_NEXT_DAY_MARKER)
    if next_day_idx != -1:
        values["next_day_plan"] = body_text[next_day_idx + len(_NEXT_DAY_MARKER):].strip()
        numbered_region = body_text[:next_day_idx]
    else:
        numbered_region = body_text

    marker_positions = [m.start() for m in _MARKER_RE.finditer(numbered_region)]
    for i, marker_pos in enumerate(marker_positions):
        label_start = marker_pos + 1
        next_marker_start = (
            marker_positions[i + 1] if i + 1 < len(marker_positions) else len(numbered_region)
        )
        matched = _extract_labeled_value(numbered_region, label_start, next_marker_start)
        if matched is None:
            snippet = numbered_region[label_start:next_marker_start].strip()
            if snippet:
                unknown[f"unknown_{i + 1}"] = snippet
            continue
        label, value = matched
        canonical = LABEL_TO_FIELD.get(label)
        if canonical is None:
            unknown[label] = value
            continue
        # 同じ正規化フィールドが複数回現れた場合は空でない方を優先する
        if canonical not in values or not values[canonical]:
            values[canonical] = value

    return values, unknown


def normalize_facility(raw: str) -> str:
    """拠点名の表記ゆれ(「RASIEL」「株式会社ラシエル」等の有無)を吸収する。"""
    normalized = raw
    for token in FACILITY_NOISE_TOKENS:
        normalized = normalized.replace(token, "")
    normalized = normalized.strip(" 　・()（）")
    return normalized or raw.strip()


def parse_chat_log(text: str, default_year: str | None = None) -> list[FacilityReport]:
    """LINE営業活動報告のテキストを構造化されたFacilityReportのリストに変換する。

    外部API・AIモデルを一切呼び出さない、正規表現とラベル辞書だけのルールベース実装。
    オフライン環境でも実行できる。
    """
    raw_messages = split_into_messages(text)
    reports: list[FacilityReport] = []
    last_date_tok: tuple[str | None, str, str] | None = None

    for msg in raw_messages:
        if msg["date"] is not None:
            last_date_tok = msg["date"]
        date_str = _format_date(last_date_tok, default_year)
        time_str = _format_time(msg["time"])

        header_name, _company_raw, am_id, am_name = parse_header(msg["header"])
        body_text = "\n".join(msg["body_lines"])
        values, unknown = extract_fields(body_text)

        facility_raw = values.pop("facility_name", "").strip()
        facility_key = normalize_facility(facility_raw) if facility_raw else ""
        reporter_name = values.pop("reporter_name", "") or header_name

        reports.append(
            FacilityReport(
                date=date_str,
                time=time_str,
                reporter_name=reporter_name,
                facility_raw=facility_raw,
                facility_key=facility_key,
                am_id=am_id,
                am_name=am_name,
                values=values,
                unknown_fields=unknown,
                raw_header=msg["header"],
            )
        )

    return reports
