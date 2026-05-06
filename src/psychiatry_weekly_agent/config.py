from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv


@dataclass(frozen=True)
class Settings:
    openai_api_key: str
    ncbi_email: str
    openai_model: str = "gpt-4.1-mini"
    ncbi_tool: str = "psychiatry-weekly-gh-agent"
    report_output_dir: Path = Path("reports")
    pubmed_retmax: int = 100
    min_candidates_before_expand: int = 30


def load_settings() -> Settings:
    load_dotenv()

    openai_api_key = os.getenv("OPENAI_API_KEY", "").strip()
    ncbi_email = os.getenv("NCBI_EMAIL", "").strip()

    missing = []
    if not openai_api_key:
        missing.append("OPENAI_API_KEY")
    if not ncbi_email:
        missing.append("NCBI_EMAIL")
    if missing:
        raise RuntimeError(
            "Missing required environment variables: "
            + ", ".join(missing)
            + ". Copy .env.example to .env and fill in the values."
        )

    return Settings(
        openai_api_key=openai_api_key,
        ncbi_email=ncbi_email,
        openai_model=os.getenv("OPENAI_MODEL", "gpt-4.1-mini").strip() or "gpt-4.1-mini",
        ncbi_tool=os.getenv("NCBI_TOOL", "psychiatry-weekly-gh-agent").strip()
        or "psychiatry-weekly-gh-agent",
        report_output_dir=Path(os.getenv("REPORT_OUTPUT_DIR", "reports")),
        pubmed_retmax=int(os.getenv("PUBMED_RETMAX", "100")),
        min_candidates_before_expand=int(os.getenv("MIN_CANDIDATES_BEFORE_EXPAND", "30")),
    )
