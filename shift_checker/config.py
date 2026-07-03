"""設定の読み込み。shift_checker/config.json があれば既定値を上書きする。"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from pathlib import Path

from .util import parse_words

CONFIG_FILENAME = "config.json"

DEFAULT_BASE_DIR = (
    r"C:\Users\k-takayama\Bihonest Corp. Dropbox\17.ラシエル"
    r"\00.ラシエル_General\1 各施設シフト管理"
)

DEFAULT_PAID_WORDS = "有,有休,有給"
DEFAULT_VACATION_WORDS = "休,休み,公,公休,希休,希,希望休"

# 配置・体制チェックで参照するセル（HTML版と同一）
SUMMARY_ADDRESSES = [
    "FE79", "FE82", "EV79", "EV78",
    "EA19", "EA22", "DY24", "EA30", "EA36",
    "DY40", "DU75", "BU73", "BU75",
]

# 項目名はHTML版「配置・体制チェック」画面の表示と同一にする（変更しないこと）
SUMMARY_LABELS = {
    "DL3": "所定労働時間 (DL3)",
    "FE79": "支援員 法令基準 (FE79)",
    "FE82": "世話人 法令基準 (FE82)",
    "EV79": "追加必要人数 (EV79)",
    "EV78": "体制加算判定 (EV78)",
    "DY24": "実配置数 (DY24)",
    "EA19": "日中配置換算数 (EA19)",
    "EA22": "日中配置超過換算 (EA22)",
    "DY40": "実配置数 (DY40)",
    "EA30": "日中配置換算数 (EA30)",
    "EA36": "日中配置超過換算 (EA36)",
    "DU75": "7.5:1 基準人員数 (DU75)",
    "BU73": "現在の基準値 (BU73)",
    "REAL_TOTAL": "実配置合計 (REAL_TOTAL)",
    "BU75": "過不足結果 (BU75)",
}

# HTML版と同じ4グループ構成（グループ名, 対象キー）
SUMMARY_GROUPS = [
    ("基本情報・体制加算", ["DL3", "FE79", "FE82", "EV79", "EV78"]),
    ("支援員 配置詳細", ["DY24", "EA19", "EA22"]),
    ("世話人 配置詳細", ["DY40", "EA30", "EA36"]),
    ("全体基準・過不足", ["DU75", "BU73", "REAL_TOTAL", "BU75"]),
]


@dataclass
class Settings:
    base_dir: str = DEFAULT_BASE_DIR
    output_dir: str = "desktop"  # "desktop" またはフォルダパス
    snapshot_dir: str = ""       # 空なら base_dir/_チェック履歴
    snapshot_keep: int = 10      # 対象年月ごとに残す実行履歴の数

    sheet_keyword: str = "管理用"
    start_row: int = 10
    end_row: int = 58
    name_col: str = "M"
    shift_start: str = "AI"
    shift_end: str = "CQ"
    total_col: str = "CU"
    fulltime_col: str = "AE"
    prescribed_cell: str = "DL3"
    job_col: str = "B"
    day_hours_col: str = "DU"
    day_allocation_col: str = "DL"

    paid_words: list[str] = field(default_factory=lambda: parse_words(DEFAULT_PAID_WORDS))
    vacation_words: list[str] = field(default_factory=lambda: parse_words(DEFAULT_VACATION_WORDS))
    hour_threshold: float = 8.0        # 総労働時間の変化がこれ以上なら要確認
    incomplete_threshold: float = 80.0  # 所定不足がこれ以上なら「入力途中の可能性」

    @property
    def leave_words(self) -> list[str]:
        return self.paid_words + self.vacation_words


def config_path() -> Path:
    return Path(__file__).resolve().parent / CONFIG_FILENAME


def load_settings() -> Settings:
    settings = Settings()
    path = config_path()
    if not path.exists():
        return settings
    with open(path, encoding="utf-8") as handle:
        raw = json.load(handle)
    for key, value in raw.items():
        if key in ("paid_words", "vacation_words") and isinstance(value, str):
            value = parse_words(value)
        if hasattr(settings, key):
            setattr(settings, key, value)
    return settings


def resolve_output_dir(settings: Settings) -> Path:
    if settings.output_dir and settings.output_dir.lower() != "desktop":
        return Path(settings.output_dir)
    home = Path(os.environ.get("USERPROFILE", Path.home()))
    for candidate in (home / "Desktop", home / "OneDrive" / "Desktop", home / "デスクトップ"):
        if candidate.is_dir():
            return candidate
    return home


def resolve_snapshot_dir(settings: Settings) -> Path:
    if settings.snapshot_dir:
        return Path(settings.snapshot_dir)
    return Path(settings.base_dir) / "_チェック履歴"
