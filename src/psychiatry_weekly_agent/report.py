from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from pathlib import Path

from .models import ScoredArticle, SearchResult
from .constants import CATEGORY_TARGETS


def build_markdown_report(search_result: SearchResult, selected: list[ScoredArticle]) -> str:
    period = f"{search_result.start_date.isoformat()} 〜 {search_result.end_date.isoformat()}"
    expansion_note = "（候補数が少なかったため14日間に拡大）" if search_result.expanded_to_14_days else ""
    lines: list[str] = [
        "# 週刊 精神科・精神薬学・GH支援レポート",
        "",
        "## 対象期間",
        f"{period} {expansion_note}".strip(),
        "",
        "## 今週の総括",
        _summary(selected, len(search_result.articles)),
        "",
    ]

    grouped: dict[str, list[ScoredArticle]] = defaultdict(list)
    for item in selected:
        grouped[item.category].append(item)

    for category in CATEGORY_TARGETS:
        lines.extend([f"## {category}", ""])
        for item in grouped.get(category, []):
            lines.extend(_article_block(item))
            lines.append("")
        if not grouped.get(category):
            lines.extend(["該当論文なし。", ""])

    lines.extend(
        [
            "## 今週、職員へ共有するなら",
            _staff_share(selected),
            "",
            "## 医学的判断を代替しない注意書き",
            "本レポートはPubMed掲載情報とabstractをもとに、障がい者グループホームの支援・リスク管理・職員研修に役立つ可能性がある知見を整理したものです。診断、治療、処方変更、服薬中止、緊急対応の判断を代替するものではありません。利用者の症状悪化、副作用疑い、自傷他害リスク、急変がある場合は、主治医、薬剤師、訪問看護、救急、行政等の適切な専門職・機関へ速やかに相談してください。",
            "",
        ]
    )
    return "\n".join(lines)


def write_report(markdown: str, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / f"weekly_report_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.md"
    path.write_text(markdown, encoding="utf-8")
    return path


def _summary(selected: list[ScoredArticle], candidate_count: int) -> str:
    if not selected:
        return f"候補論文{candidate_count}本を確認しましたが、選定可能な論文はありませんでした。検索語や期間を見直してください。"
    top = selected[0]
    fields = ", ".join(sorted({item.field for item in selected})[:5])
    return (
        f"候補論文{candidate_count}本から、GH現場での支援、服薬・副作用理解、リスク管理、職員研修への転用可能性を重視して{len(selected)}本を選定しました。"
        f"主な分野は{fields}です。最も優先度が高い論文は「{top.article.title}」で、{top.reason}"
    )


def _article_block(item: ScoredArticle) -> list[str]:
    article = item.article
    return [
        f"### {article.title}",
        "",
        f"- **分野**: {item.field}",
        f"- **研究タイプ**: {item.research_type}",
        f"- **重要度**: {item.importance}/5（総合スコア: {item.total_score}）",
        f"- **PMID**: {article.pmid}",
        f"- **著者**: {', '.join(article.authors[:8]) if article.authors else '記載なし'}",
        f"- **雑誌名**: {article.journal}",
        f"- **発行日**: {article.publication_date}",
        "- **要点**:",
        *_bullet_lines(item.key_points, 2),
        "- **GH現場への示唆**:",
        *_bullet_lines(item.gh_implications, 2),
        "- **AMが見るポイント**:",
        *_bullet_lines(item.am_points, 2),
        "- **管理者が見るポイント**:",
        *_bullet_lines(item.manager_points, 2),
        "- **サビ管が見るポイント**:",
        *_bullet_lines(item.service_manager_points, 2),
        f"- **職員研修に使える一言**: {item.training_one_liner}",
        "- **注意点**:",
        *_bullet_lines(item.cautions, 2),
        f"- **原文URL**: {article.pubmed_url}",
    ]


def _bullet_lines(items: list[str], indent: int) -> list[str]:
    prefix = " " * indent + "- "
    return [prefix + item for item in items]


def _staff_share(selected: list[ScoredArticle]) -> str:
    if not selected:
        return "今週共有する論文はありません。"
    one_liners = [f"- {item.training_one_liner}" for item in selected[:3]]
    return "\n".join(one_liners)
