from __future__ import annotations

import argparse
from datetime import date

from .config import load_settings
from .pubmed import fetch_recent_articles
from .report import build_markdown_report, write_report
from .scorer import score_and_select_articles


def main() -> None:
    parser = argparse.ArgumentParser(
        description="PubMedから精神科・精神薬学の新着論文を取得し、GH支援向けMarkdownレポートを作成します。"
    )
    parser.add_argument(
        "--today",
        type=date.fromisoformat,
        default=None,
        help="検索終了日をYYYY-MM-DDで指定します。省略時は実行日です。",
    )
    args = parser.parse_args()

    settings = load_settings()
    search_result = fetch_recent_articles(settings, today=args.today)
    print(f"Fetched {len(search_result.articles)} PubMed candidates.")

    selected = score_and_select_articles(settings, search_result.articles)
    print(f"Selected {len(selected)} articles for the weekly report.")

    markdown = build_markdown_report(search_result, selected)
    output_path = write_report(markdown, settings.report_output_dir)
    print(f"Report written to {output_path}")


if __name__ == "__main__":
    main()
