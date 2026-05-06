from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Any


@dataclass(frozen=True)
class Article:
    pmid: str
    title: str
    authors: list[str]
    journal: str
    publication_date: str
    abstract: str
    pubmed_url: str


@dataclass(frozen=True)
class SearchResult:
    articles: list[Article]
    start_date: date
    end_date: date
    expanded_to_14_days: bool


@dataclass
class ScoredArticle:
    article: Article
    field: str
    research_type: str
    importance: int
    scores: dict[str, int]
    reason: str
    key_points: list[str]
    gh_implications: list[str]
    am_points: list[str]
    manager_points: list[str]
    service_manager_points: list[str]
    training_one_liner: str
    cautions: list[str]
    category: str = ""
    raw: dict[str, Any] = field(default_factory=dict)

    @property
    def total_score(self) -> int:
        return sum(self.scores.values()) + self.importance
