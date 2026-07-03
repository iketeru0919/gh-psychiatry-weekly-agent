"""スナップショット（前回比較用の控えコピー）の保存・取得・整理。

構成:  <snapshot_dir>/<yymm>/<YYYY-MM-DD_HHMM>/<施設名>.xlsx
"""

from __future__ import annotations

import re
import shutil
from datetime import datetime
from pathlib import Path

from .util import long_path, sanitize_filename

RUN_DIR_RE = re.compile(r"^\d{4}-\d{2}-\d{2}_\d{4}")


def month_key(year: int, month: int) -> str:
    return f"{year % 100:02d}{month:02d}"


def latest_run_dir(snapshot_dir: Path, year: int, month: int) -> Path | None:
    base = snapshot_dir / month_key(year, month)
    if not base.is_dir():
        return None
    runs = sorted(d for d in base.iterdir() if d.is_dir() and RUN_DIR_RE.match(d.name))
    return runs[-1] if runs else None


def find_snapshot_file(run_dir: Path, facility: str) -> Path | None:
    stem = sanitize_filename(facility)
    for ext in (".xlsx", ".xlsm"):
        candidate = run_dir / f"{stem}{ext}"
        if candidate.exists():
            return candidate
    return None


def save_snapshot(snapshot_dir: Path, year: int, month: int,
                  files: dict[str, Path], keep: int) -> Path:
    """今回読んだ各施設のファイルを日時付きフォルダへコピーし、古い履歴を整理する。"""
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    run_dir = snapshot_dir / month_key(year, month) / stamp
    run_dir.mkdir(parents=True, exist_ok=True)
    for facility, source in files.items():
        target = run_dir / f"{sanitize_filename(facility)}{source.suffix.lower()}"
        shutil.copy2(long_path(str(source)), long_path(str(target)))
    prune_snapshots(snapshot_dir, year, month, keep)
    return run_dir


def prune_snapshots(snapshot_dir: Path, year: int, month: int, keep: int) -> None:
    if keep <= 0:
        return
    base = snapshot_dir / month_key(year, month)
    if not base.is_dir():
        return
    runs = sorted(d for d in base.iterdir() if d.is_dir() and RUN_DIR_RE.match(d.name))
    for old in runs[:-keep]:
        shutil.rmtree(old, ignore_errors=True)
