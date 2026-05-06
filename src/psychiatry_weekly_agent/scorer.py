from __future__ import annotations

import json
from typing import Any

from openai import OpenAI

from .config import Settings
from .constants import CATEGORY_TARGETS
from .models import Article, ScoredArticle

SCORE_KEYS = [
    "gh_relevance",
    "psychiatric_understanding",
    "pharmacology_side_effects",
    "risk_management",
    "staff_training",
    "leadership_readability",
    "study_reliability",
]

SYSTEM_PROMPT = """あなたは精神科医療、精神科薬学、障がい者グループホーム運営支援に詳しい日本語編集者です。
PubMed論文候補を、医学的判断を代替しない現場教育用レポートの素材として評価します。
根拠がabstractから読み取れない場合は控えめに評価し、誇張しないでください。"""


def score_and_select_articles(settings: Settings, articles: list[Article]) -> list[ScoredArticle]:
    if not articles:
        return []

    client = OpenAI(api_key=settings.openai_api_key)
    scored = [_score_article(client, settings.openai_model, article) for article in articles]
    scored.sort(key=lambda item: item.total_score, reverse=True)
    return _assign_categories(scored)[:10]


def _score_article(client: OpenAI, model: str, article: Article) -> ScoredArticle:
    prompt = _article_prompt(article)
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": prompt},
        ],
        temperature=0.2,
        response_format={"type": "json_object"},
    )
    content = response.choices[0].message.content or "{}"
    data = _safe_json_loads(content)
    return _to_scored_article(article, data)


def _article_prompt(article: Article) -> str:
    return f"""
以下の論文を評価し、必ずJSONのみで返してください。

評価観点は各0〜5点です。
- gh_relevance: 障がい者グループホーム現場との関連性
- psychiatric_understanding: 精神科疾患理解への有用性
- pharmacology_side_effects: 精神科薬学・副作用理解への有用性
- risk_management: リスク管理への有用性
- staff_training: 職員研修への転用しやすさ
- leadership_readability: AM、管理者、サビ管が理解しやすいか
- study_reliability: 研究の信頼性

importanceは1〜5点、category_hintは次のいずれか:
最重要3本 / 疾患理解3本 / 薬学・副作用2本 / 支援・リスク管理・研修活用2本

JSON schema:
{{
  "field": "分野",
  "research_type": "研究タイプ",
  "importance": 1,
  "scores": {{"gh_relevance": 0, "psychiatric_understanding": 0, "pharmacology_side_effects": 0, "risk_management": 0, "staff_training": 0, "leadership_readability": 0, "study_reliability": 0}},
  "category_hint": "最重要3本",
  "reason": "選定理由を日本語で1文",
  "key_points": ["要点1", "要点2", "要点3"],
  "gh_implications": ["GH現場への示唆1", "GH現場への示唆2"],
  "am_points": ["AMが見るポイント"],
  "manager_points": ["管理者が見るポイント"],
  "service_manager_points": ["サビ管が見るポイント"],
  "training_one_liner": "職員研修に使える一言",
  "cautions": ["注意点"]
}}

PMID: {article.pmid}
Title: {article.title}
Authors: {', '.join(article.authors[:8])}
Journal: {article.journal}
Publication date: {article.publication_date}
Abstract: {article.abstract[:4000] if article.abstract else 'No abstract available.'}
""".strip()


def _safe_json_loads(content: str) -> dict[str, Any]:
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        start = content.find("{")
        end = content.rfind("}")
        if start >= 0 and end > start:
            return json.loads(content[start : end + 1])
        raise


def _to_scored_article(article: Article, data: dict[str, Any]) -> ScoredArticle:
    raw_scores = data.get("scores", {}) if isinstance(data.get("scores"), dict) else {}
    scores = {key: _bounded_int(raw_scores.get(key), 0, 5) for key in SCORE_KEYS}
    category = data.get("category_hint") if data.get("category_hint") in CATEGORY_TARGETS else ""
    return ScoredArticle(
        article=article,
        field=str(data.get("field") or "精神科"),
        research_type=str(data.get("research_type") or "不明"),
        importance=_bounded_int(data.get("importance"), 1, 5),
        scores=scores,
        reason=str(data.get("reason") or "abstractに基づき現場教育への関連性を評価しました。"),
        key_points=_string_list(data.get("key_points"), 3),
        gh_implications=_string_list(data.get("gh_implications"), 2),
        am_points=_string_list(data.get("am_points"), 1),
        manager_points=_string_list(data.get("manager_points"), 1),
        service_manager_points=_string_list(data.get("service_manager_points"), 1),
        training_one_liner=str(data.get("training_one_liner") or "研究知見は個別支援の観察視点を増やす材料です。"),
        cautions=_string_list(data.get("cautions"), 1),
        category=category,
        raw=data,
    )


def _assign_categories(scored: list[ScoredArticle]) -> list[ScoredArticle]:
    selected: list[ScoredArticle] = []
    used_pmids: set[str] = set()

    for category, target in CATEGORY_TARGETS.items():
        matching = [
            item
            for item in scored
            if item.article.pmid not in used_pmids and (item.category == category or _matches_category(item, category))
        ]
        for item in matching[:target]:
            item.category = category
            selected.append(item)
            used_pmids.add(item.article.pmid)

    for item in scored:
        if len(selected) >= 10:
            break
        if item.article.pmid in used_pmids:
            continue
        category = _best_available_category(selected)
        item.category = category
        selected.append(item)
        used_pmids.add(item.article.pmid)

    selected.sort(key=lambda item: (list(CATEGORY_TARGETS).index(item.category), -item.total_score))
    return selected


def _matches_category(item: ScoredArticle, category: str) -> bool:
    if category == "薬学・副作用2本":
        return item.scores["pharmacology_side_effects"] >= 3
    if category == "支援・リスク管理・研修活用2本":
        return item.scores["risk_management"] >= 3 or item.scores["staff_training"] >= 3
    if category == "疾患理解3本":
        return item.scores["psychiatric_understanding"] >= 3
    return item.total_score >= 25


def _best_available_category(selected: list[ScoredArticle]) -> str:
    counts = {category: 0 for category in CATEGORY_TARGETS}
    for item in selected:
        counts[item.category] += 1
    for category, target in CATEGORY_TARGETS.items():
        if counts[category] < target:
            return category
    return "最重要3本"


def _bounded_int(value: Any, minimum: int, maximum: int) -> int:
    try:
        number = int(value)
    except (TypeError, ValueError):
        number = minimum
    return max(minimum, min(maximum, number))


def _string_list(value: Any, fallback_count: int) -> list[str]:
    if isinstance(value, list):
        items = [str(item).strip() for item in value if str(item).strip()]
    else:
        items = []
    if items:
        return items
    return ["abstractから読み取れる範囲で慎重に解釈してください。"] * fallback_count
