from __future__ import annotations

from datetime import date, timedelta
from typing import Any
import time

import requests

from .config import Settings
from .models import Article, SearchResult

BASE_URL = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils"

SEARCH_TERM = " OR ".join(
    [
        '"Mental Disorders"[Mesh]',
        'psychiatry[Title/Abstract]',
        'psychiatric[Title/Abstract]',
        'schizophrenia[Title/Abstract]',
        'bipolar disorder[Title/Abstract]',
        'depression[Title/Abstract]',
        'anxiety disorder[Title/Abstract]',
        'autism spectrum disorder[Title/Abstract]',
        'intellectual disability[Title/Abstract]',
        'developmental disorder[Title/Abstract]',
        'suicide[Title/Abstract]',
        'self-harm[Title/Abstract]',
        'aggression[Title/Abstract]',
        'antipsychotic[Title/Abstract]',
        'antidepressant[Title/Abstract]',
        'mood stabilizer[Title/Abstract]',
        'benzodiazepine[Title/Abstract]',
    ]
)


def fetch_recent_articles(settings: Settings, today: date | None = None) -> SearchResult:
    end_date = today or date.today()
    first_start = end_date - timedelta(days=7)
    articles = _search_and_fetch(settings, first_start, end_date)

    expanded = False
    if len(articles) < settings.min_candidates_before_expand:
        expanded = True
        expanded_start = end_date - timedelta(days=14)
        articles = _search_and_fetch(settings, expanded_start, end_date)
        first_start = expanded_start

    return SearchResult(
        articles=articles[: settings.pubmed_retmax],
        start_date=first_start,
        end_date=end_date,
        expanded_to_14_days=expanded,
    )


def _search_and_fetch(settings: Settings, start_date: date, end_date: date) -> list[Article]:
    ids = _esearch(settings, start_date, end_date)
    if not ids:
        return []
    return _efetch(settings, ids)


def _base_params(settings: Settings) -> dict[str, str]:
    return {
        "tool": settings.ncbi_tool,
        "email": settings.ncbi_email,
    }


def _esearch(settings: Settings, start_date: date, end_date: date) -> list[str]:
    params = {
        **_base_params(settings),
        "db": "pubmed",
        "term": f"({SEARCH_TERM})",
        "datetype": "edat",
        "mindate": start_date.isoformat(),
        "maxdate": end_date.isoformat(),
        "retmode": "json",
        "retmax": str(settings.pubmed_retmax),
        "sort": "pub date",
    }
    response = requests.get(f"{BASE_URL}/esearch.fcgi", params=params, timeout=30)
    response.raise_for_status()
    payload = response.json()
    return payload.get("esearchresult", {}).get("idlist", [])


def _efetch(settings: Settings, ids: list[str]) -> list[Article]:
    params = {
        **_base_params(settings),
        "db": "pubmed",
        "id": ",".join(ids),
        "retmode": "xml",
    }
    # Stay comfortably within NCBI's no-key rate guidance for scripted runs.
    time.sleep(0.34)
    response = requests.get(f"{BASE_URL}/efetch.fcgi", params=params, timeout=60)
    response.raise_for_status()

    import xml.etree.ElementTree as ET

    root = ET.fromstring(response.text)
    articles: list[Article] = []
    for pubmed_article in root.findall(".//PubmedArticle"):
        article = _parse_pubmed_article(pubmed_article)
        if article:
            articles.append(article)
    return articles


def _parse_pubmed_article(pubmed_article: Any) -> Article | None:
    medline = pubmed_article.find("MedlineCitation")
    if medline is None:
        return None

    pmid = _text(medline.find("PMID"))
    article_node = medline.find("Article")
    if article_node is None or not pmid:
        return None

    title = " ".join("".join(article_node.findtext("ArticleTitle", default="").split()).split())
    title = _stringify_node(article_node.find("ArticleTitle")) or title or "No title"

    journal = _text(article_node.find("Journal/Title")) or _text(article_node.find("Journal/ISOAbbreviation"))
    publication_date = _publication_date(article_node)
    authors = _authors(article_node)
    abstract = _abstract(article_node)

    return Article(
        pmid=pmid,
        title=title,
        authors=authors,
        journal=journal or "Unknown journal",
        publication_date=publication_date,
        abstract=abstract,
        pubmed_url=f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/",
    )


def _authors(article_node: Any) -> list[str]:
    names: list[str] = []
    for author in article_node.findall("AuthorList/Author"):
        collective = _text(author.find("CollectiveName"))
        if collective:
            names.append(collective)
            continue
        last = _text(author.find("LastName"))
        fore = _text(author.find("ForeName")) or _text(author.find("Initials"))
        if last and fore:
            names.append(f"{fore} {last}")
        elif last:
            names.append(last)
    return names


def _abstract(article_node: Any) -> str:
    parts: list[str] = []
    for abstract_text in article_node.findall("Abstract/AbstractText"):
        label = abstract_text.attrib.get("Label")
        text = _stringify_node(abstract_text)
        if text:
            parts.append(f"{label}: {text}" if label else text)
    return "\n".join(parts)


def _publication_date(article_node: Any) -> str:
    for path in ["Journal/JournalIssue/PubDate", "ArticleDate"]:
        node = article_node.find(path)
        if node is None:
            continue
        year = _text(node.find("Year"))
        month = _text(node.find("Month"))
        day = _text(node.find("Day"))
        medline_date = _text(node.find("MedlineDate"))
        if year:
            return "-".join(part for part in [year, month, day] if part)
        if medline_date:
            return medline_date
    return "Unknown"


def _text(node: Any) -> str:
    if node is None or node.text is None:
        return ""
    return node.text.strip()


def _stringify_node(node: Any) -> str:
    if node is None:
        return ""
    return " ".join("".join(node.itertext()).split())
