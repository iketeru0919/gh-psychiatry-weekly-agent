"""Dropbox同期フォルダの再帰スキャンと対象月ファイルの選定。"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from pathlib import Path

from .util import nfkc

SHIFT_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
SKIP_DIR_PREFIXES = ("_チェック履歴", ".", "~")


@dataclass
class FacilityFiles:
    facility: str
    chosen: Path | None = None
    candidates: list[Path] = field(default_factory=list)

    @property
    def duplicate_count(self) -> int:
        return len(self.candidates)


def month_regexes(year: int, month: int) -> list[re.Pattern]:
    yy = f"{year % 100:02d}"
    mm = f"{month:02d}"
    patterns = [
        rf"(?<!\d){year}{mm}(?!\d)",       # 202606
        rf"(?<!\d){yy}{mm}(?!\d)",         # 2606
        rf"{year}\s*年\s*0?{month}\s*月",  # 2026年6月 / 2026年06月
        rf"(?<!\d){year}[-_./]0?{month}(?!\d)",  # 2026-06, 2026_6, 2026.06
    ]
    return [re.compile(p) for p in patterns]


def path_matches_month(rel_path: str, year: int, month: int) -> bool:
    text = nfkc(rel_path)
    if any(regex.search(text) for regex in month_regexes(year, month)):
        return True
    # 「…\2026年\6月\…」のように年と月が別フォルダに分かれているケース
    return bool(
        re.search(rf"{year}\s*年", text)
        and re.search(rf"(?<!\d)0?{month}\s*月", text)
    )


def scan_shift_files(base_dir: str, year: int, month: int) -> list[FacilityFiles]:
    """base_dir以下を再帰スキャンし、施設ごとに対象月の勤務表候補を集める。

    施設名は base_dir 直下の第1階層フォルダ名。フォルダの深さは問わない。
    同一施設に複数候補がある場合は更新日時が最新のものを採用する。
    """
    base = Path(base_dir)
    facilities: dict[str, FacilityFiles] = {}

    for root, dirs, files in os.walk(base):
        dirs[:] = [d for d in dirs if not d.startswith(SKIP_DIR_PREFIXES)]
        rel_root = Path(root).relative_to(base)
        for filename in files:
            if filename.startswith("~$"):  # Excelの一時ロックファイル
                continue
            if Path(filename).suffix.lower() not in SHIFT_EXTENSIONS:
                continue
            rel_path = str(rel_root / filename)
            if not path_matches_month(rel_path, year, month):
                continue
            parts = (rel_root / filename).parts
            facility = parts[0] if len(parts) > 1 else Path(filename).stem
            entry = facilities.setdefault(facility, FacilityFiles(facility=facility))
            entry.candidates.append(Path(root) / filename)

    for entry in facilities.values():
        entry.candidates.sort(key=lambda p: p.stat().st_mtime)
        entry.chosen = entry.candidates[-1] if entry.candidates else None

    return sorted(facilities.values(), key=lambda e: e.facility)
