"""
Запасной поиск нормативного документа в открытом интернете.

Используется, если Техэксперт недоступен (лимит подключений, сеть и т.п.).
Поиск: DuckDuckGo → HTML-страницы на доверенных доменах (pravo.gov.ru, legalacts.ru, …).
"""

from __future__ import annotations

import os
import re
from urllib.parse import parse_qs, unquote, urlparse

import requests

from sk_reporter.prescriptions.techexpert_client import (
    NormativeLookupResult,
    NormativeReference,
    _MAX_EXCERPT,
    _build_search_queries,
    _extract_doc_title,
    _extract_points_excerpt,
    _html_to_text,
    _score_document_header,
    _short_doc_title,
    parse_normative_reference,
)

_MAX_URLS = 6
_MAX_DDG_LINKS = 12
_TRUSTED_DOMAINS: dict[str, int] = {
    "publication.pravo.gov.ru": 12,
    "docs.cntd.ru": 11,
    "gost.ru": 10,
    "legalacts.ru": 9,
    "rulaws.ru": 9,
    "base.garant.ru": 7,
    "nvol.gosnadzor.ru": 8,
    "gosnadzor.ru": 8,
}


def _env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def internet_fallback_enabled() -> bool:
    val = _env("TE_EXPERT_INTERNET_FALLBACK", "1").lower()
    return val not in {"0", "false", "no"}


def _domain_score(url: str) -> int:
    try:
        host = urlparse(url).netloc.lower().removeprefix("www.")
    except Exception:
        return 0
    for domain, score in _TRUSTED_DOMAINS.items():
        if host == domain or host.endswith("." + domain):
            return score
    return 0


def _unwrap_ddg_href(href: str) -> str:
    if "uddg=" in href:
        qs = parse_qs(urlparse(href).query)
        return unquote(qs.get("uddg", [href])[0])
    return href


def _skip_url(url: str) -> bool:
    low = url.lower()
    if low.endswith(".pdf") or low.endswith(".doc") or low.endswith(".docx"):
        return True
    if any(x in low for x in ("/login", "javascript:", "facebook.com", "vk.com")):
        return True
    return _domain_score(url) <= 0


def _parse_date_parts(date: str) -> tuple[str, str, str] | None:
    if not date:
        return None
    m = re.match(r"(\d{1,2})[./](\d{1,2})[./](\d{2,4})", date.strip())
    if not m:
        return None
    dd, mm, yy = m.groups()
    if len(yy) == 2:
        yy = "20" + yy
    return dd.zfill(2), mm.zfill(2), yy


def _direct_url_candidates(reference: NormativeReference) -> list[str]:
    """Прямые URL по шаблонам (без поисковика) — надёжнее при блокировке DDG."""
    urls: list[str] = []
    number = reference.number
    if not number:
        return urls

    parts = _parse_date_parts(reference.date)
    raw_low = reference.raw.lower()
    issuer_low = reference.issuer.lower()

    if parts:
        dd, mm, yyyy = parts
        dotted = f"{int(dd)}.{int(mm)}.{yyyy}"
        compact = f"{dd}{mm}{yyyy}"

        if "ростехнадзор" in raw_low or "ростехнадзор" in issuer_low:
            urls.extend(
                [
                    f"https://rulaws.ru/acts/Prikaz-Rostehnadzora-ot-{dotted}-N-{number}/",
                    (
                        "https://legalacts.ru/doc/"
                        f"prikaz-rostekhnadzora-ot-{compact}-n-{number}-ob-utverzhdenii-federalnykh/"
                    ),
                ]
            )

        if reference.doc_kind == "Приказ":
            urls.append(
                f"https://rulaws.ru/acts/Prikaz-ot-{dotted}-N-{number}/"
            )

    return urls


def _duckduckgo_blocked(html: str, status_code: int) -> bool:
    low = html.lower()
    return status_code in {202, 403} or "captcha" in low or "anomaly" in low


def _duckduckgo_links(session: requests.Session, query: str) -> list[str]:
    try:
        resp = session.post(
            "https://html.duckduckgo.com/html/",
            data={"q": query, "b": ""},
            timeout=30,
        )
    except requests.RequestException:
        return []
    if _duckduckgo_blocked(resp.text, resp.status_code):
        return []
    if resp.status_code != 200:
        return []

    links: list[str] = []
    for href in re.findall(r'class="result__a"[^>]*href="([^"]+)"', resp.text):
        url = _unwrap_ddg_href(href.replace("&amp;", "&"))
        if url.startswith("http") and url not in links:
            links.append(url)
    return links[:_MAX_DDG_LINKS]


def _build_web_queries(reference: NormativeReference) -> list[str]:
    queries = list(_build_search_queries(reference))
    raw = re.sub(r"\s+", " ", reference.raw).strip()
    if raw and raw not in queries:
        queries.append(raw[:160])
    extra: list[str] = []
    for q in queries[:3]:
        extra.append(f"{q} site:publication.pravo.gov.ru OR site:legalacts.ru OR site:rulaws.ru")
        extra.append(f"{q} site:docs.cntd.ru OR site:gost.ru")
    out: list[str] = []
    for q in queries + extra:
        q = q.strip()
        if q and q not in out:
            out.append(q)
    return out[:6]


def _fetch_page_text(session: requests.Session, url: str) -> tuple[str, str]:
    resp = session.get(url, timeout=35)
    if resp.status_code != 200 or len(resp.text) < 300:
        return "", ""
    plain = _html_to_text(resp.text)
    if len(plain) < 200:
        return "", ""
    title = _extract_doc_title(plain, url)
    return plain, title


def lookup_normative_web(normative_text: str) -> NormativeLookupResult:
    """Поиск НД в интернете (запасной канал)."""
    reference = parse_normative_reference(normative_text)
    if not reference.raw:
        return NormativeLookupResult(
            ok=False,
            reference=reference,
            error="Пустая ссылка на нормативный документ (B19)",
            source="internet",
        )

    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            "Accept-Language": "ru-RU,ru;q=0.9",
        }
    )

    candidates: list[tuple[int, str]] = []
    seen_urls: set[str] = set()

    for url in _direct_url_candidates(reference):
        if url not in seen_urls and not _skip_url(url):
            seen_urls.add(url)
            candidates.append((_domain_score(url) + 5, url))

    for query in _build_web_queries(reference):
        for url in _duckduckgo_links(session, query):
            if url in seen_urls or _skip_url(url):
                continue
            seen_urls.add(url)
            domain_pts = _domain_score(url)
            title_pts = 2 if reference.number and reference.number in url else 0
            candidates.append((domain_pts + title_pts, url))
        if len(candidates) >= _MAX_URLS * 2:
            break

    candidates.sort(key=lambda x: x[0], reverse=True)
    urls = [u for _, u in candidates[:_MAX_URLS]]
    if not urls:
        return NormativeLookupResult(
            ok=False,
            reference=reference,
            error="В интернете не найдено подходящих страниц с нормативным документом",
            source="internet",
        )

    best_url = ""
    best_plain = ""
    best_title = ""
    best_score = -1

    for url in urls:
        plain, title = _fetch_page_text(session, url)
        if not plain:
            continue
        score = _score_document_header(plain[:15000], reference) + _domain_score(url)
        if score > best_score:
            best_score = score
            best_url = url
            best_plain = plain
            best_title = title or _extract_doc_title(plain, reference.search_query)

    if best_score < 3 or not best_plain:
        return NormativeLookupResult(
            ok=False,
            reference=reference,
            error=(
                "Найдены страницы в интернете, но ни одна не совпала с ссылкой из B19 "
                f"(проверено URL: {len(urls)})"
            ),
            source="internet",
        )

    excerpt = _extract_points_excerpt(best_plain, reference.points)
    if not excerpt:
        excerpt = best_plain[:_MAX_EXCERPT]

    short = _short_doc_title(best_title, reference) or _short_doc_title(
        _extract_doc_title(best_plain, reference.search_query), reference
    )

    return NormativeLookupResult(
        ok=True,
        reference=reference,
        excerpt=excerpt,
        doc_title=short,
        list_title=best_title,
        source_url=best_url,
        auth_ok=True,
        source="internet",
    )
