from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class FacilityReport:
    """1件のLINE営業活動報告メッセージから抽出した構造化データ。"""

    date: str | None
    time: str | None
    reporter_name: str
    facility_raw: str
    facility_key: str
    am_id: str | None
    am_name: str | None
    values: dict[str, str] = field(default_factory=dict)
    unknown_fields: dict[str, str] = field(default_factory=dict)
    raw_header: str = ""

    def get(self, canonical_field: str, default: str = "") -> str:
        return self.values.get(canonical_field, default)
