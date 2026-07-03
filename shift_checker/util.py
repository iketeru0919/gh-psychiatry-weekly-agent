"""共通ユーティリティ（正規化・数値変換・パス処理）."""

from __future__ import annotations

import os
import re
import sys
import unicodedata


def normalize(value) -> str:
    """セル値を文字列化して全角スペースを半角化し、前後の空白を除去する。"""
    if value is None:
        return ""
    return str(value).replace("　", " ").strip()


def normalize_key(value) -> str:
    return re.sub(r"\s+", "", normalize(value)).lower()


_ANNOTATION_RE = re.compile(r"[【\[（(][^】\]）)]*[】\]）)]")


def base_name_key(name) -> str:
    """氏名から注記（【社保】・(派遣)・/派遣 など）を除いた照合キーを作る。"""
    text = normalize(name)
    text = _ANNOTATION_RE.sub("", text)
    text = re.split(r"[/／]", text)[0]
    key = normalize_key(text)
    return key or normalize_key(name)


def round2(value: float) -> float:
    return round(value + sys.float_info.epsilon, 2)


def to_hours(value, number_format: str = ""):
    """総労働時間セルの値を時間数(float)へ変換する。変換不能ならNone。"""
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        # Excelの時刻シリアル値（0〜2未満かつ h/m/s 書式）は24倍して時間数へ
        if 0 <= number < 2 and re.search(r"[hms]", number_format or "", re.I):
            number *= 24
        return round2(number)
    digits = re.sub(r"[^\d.\-]", "", str(value))
    try:
        return round2(float(digits)) if digits else None
    except ValueError:
        return None


def number_or_text(value):
    if value is None or value == "":
        return ""
    if isinstance(value, bool):
        return normalize(value)
    if isinstance(value, (int, float)):
        return round2(float(value))
    text = normalize(value)
    try:
        return round2(float(text))
    except ValueError:
        return text


def numeric_or_zero(value) -> float:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    try:
        return float(normalize(value))
    except (ValueError, TypeError):
        return 0.0


def format_signed(value) -> str:
    if value is None:
        return ""
    text = f"{value:+,.2f}".rstrip("0").rstrip(".")
    return text if text not in ("+", "-", "") else "0"


def parse_words(value) -> list[str]:
    return [normalize(word) for word in str(value).split(",") if normalize(word)]


def matches_exact(value, words) -> bool:
    text = normalize(value)
    return bool(text) and text in words


def long_path(path: str) -> str:
    """Windowsの260文字制限を回避するための拡張パス表記を返す。"""
    if os.name == "nt":
        absolute = os.path.abspath(path)
        if len(absolute) > 240 and not absolute.startswith("\\\\?\\"):
            if absolute.startswith("\\\\"):
                return "\\\\?\\UNC\\" + absolute[2:]
            return "\\\\?\\" + absolute
    return path


def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", normalize(name)) or "unknown"


def nfkc(value) -> str:
    return unicodedata.normalize("NFKC", normalize(value))
