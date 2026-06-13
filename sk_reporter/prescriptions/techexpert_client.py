"""
Клиент Техэксперт / Кодекс (te-cloud).

Логин и поиск нормативного документа по ссылке из ячейки B19.
Учётные данные — только из переменных окружения (не коммитить в репозиторий):

  TE_EXPERT_BASE_URL   — например http://248960.te-cloud.ru
  TE_EXPERT_CATALOG    — виртуальный каталог, по умолчанию /docs
  TE_EXPERT_LOGIN
  TE_EXPERT_PASSWORD
  TE_EXPERT_USE_BROWSER — 1 попытаться через Playwright (запасной), 0 — HTTP API (по умолчанию)
  TE_EXPERT_INTERNET_FALLBACK — 1 (по умолчанию) искать в интернете, если Техэксперт недоступен
"""

from __future__ import annotations

import json
import os
import re
import time
from dataclasses import dataclass, field
from datetime import datetime
from html import unescape
from html.parser import HTMLParser
from typing import Any
from urllib.parse import urljoin, urlparse

import requests

from sk_reporter.prescriptions.te_env import load_te_expert_env

load_te_expert_env()

_DEFAULT_BASE = "http://248960.te-cloud.ru"
_DEFAULT_CATALOG = "/docs"
_MAX_EXCERPT = 12_000
_SEARCH_TAB_ID = "7"
_SEARCH_WAIT_S = 4.0
_MAX_SEARCH_PAGES = 6
_MAX_CANDIDATES = 40
_TERM_TITLE_MAX_LEN = 24
_MIN_DOCUMENT_PLAIN_CHARS = 400
_JUNK_PLAIN_MARKERS = (
    "найденные фразы",
    "0 из 0",
    "точное совпадение",
    "учет порядка слов",
)
_LOGIN_RETRIES = 2
_LOGIN_RETRY_DELAY_S = 4.0


@dataclass
class NormativeReference:
    raw: str
    search_query: str
    doc_kind: str = ""
    number: str = ""
    date: str = ""
    issuer: str = ""
    points: list[str] = field(default_factory=list)


@dataclass
class NormativeLookupResult:
    ok: bool
    reference: NormativeReference
    excerpt: str = ""
    doc_title: str = ""
    list_title: str = ""
    source_url: str = ""
    error: str = ""
    auth_ok: bool = False
    source: str = "techexpert"


class _TextExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self._chunks: list[str] = []
        self._skip = False

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in {"script", "style", "noscript"}:
            self._skip = True

    def handle_endtag(self, tag: str) -> None:
        if tag in {"script", "style", "noscript"}:
            self._skip = False
        if tag in {"p", "div", "br", "li", "tr", "h1", "h2", "h3", "td"}:
            self._chunks.append("\n")

    def handle_data(self, data: str) -> None:
        if not self._skip:
            text = data.strip()
            if text:
                self._chunks.append(text)

    def text(self) -> str:
        raw = unescape(" ".join(self._chunks))
        raw = re.sub(r"[ \t]+\n", "\n", raw)
        raw = re.sub(r"\n{3,}", "\n\n", raw)
        return re.sub(r" {2,}", " ", raw).strip()


def _env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def _config() -> dict[str, str]:
    base = _env("TE_EXPERT_BASE_URL", _DEFAULT_BASE).rstrip("/")
    catalog = _env("TE_EXPERT_CATALOG", _DEFAULT_CATALOG)
    if not catalog.startswith("/"):
        catalog = "/" + catalog
    return {
        "base_url": base,
        "catalog": catalog.rstrip("/") or "/docs",
        "login": _env("TE_EXPERT_LOGIN"),
        "password": _env("TE_EXPERT_PASSWORD"),
        "use_browser": _env("TE_EXPERT_USE_BROWSER", "0") not in {"0", "false", "no"},
    }


def _normalize_date_tokens(date: str) -> str:
    if not date:
        return ""
    m = re.match(r"(\d{1,2})[./](\d{1,2})[./](\d{2,4})", date.strip())
    if m:
        dd, mm, yy = m.groups()
        if len(yy) == 2:
            yy = "20" + yy
        return f"{int(dd)} {int(mm)} {yy}"
    return date.strip()


def _extract_issuer(raw: str) -> str:
    for pat in (
        r"(Ростехнадзор\w*)",
        r"(Роструд\w*)",
        r"(Мин(?:истерств\w*|природ\w*|труд\w*|здрав\w*)\w*)",
        r"(Гос(?:стро\w*|ком\w*)\w*)",
    ):
        m = re.search(pat, raw, flags=re.IGNORECASE)
        if m:
            return _normalize_issuer_name(m.group(1).strip())
    return ""


def _normalize_issuer_name(issuer: str) -> str:
    """Именительный падеж для поиска (B19 часто в родительном: «Ростехнадзора»)."""
    low = issuer.lower()
    fixes = {
        "ростехнадзора": "Ростехнадзор",
        "роструда": "Роструд",
    }
    return fixes.get(low, issuer)


def _normalize_query(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").lower().strip())


def _title_is_query_echo(title: str, queries: list[str]) -> bool:
    t = _normalize_query(title)
    if not t:
        return True
    for q in queries:
        if t == _normalize_query(q):
            return True
    return False


def _is_valid_document_plain(plain: str, reference: NormativeReference) -> bool:
    if not plain or len(plain) < _MIN_DOCUMENT_PLAIN_CHARS:
        return False
    low = plain.lower()
    if any(marker in low for marker in _JUNK_PLAIN_MARKERS):
        return False
    if reference.number and reference.number not in plain:
        return False
    if reference.doc_kind and reference.doc_kind.lower() not in low:
        return False
    return True


def _title_looks_like_document(title: str, reference: NormativeReference) -> bool:
    if not title or len(title) < 25:
        return False
    if _is_dictionary_hit(title):
        return False
    low = title.lower()
    if reference.number and reference.number not in title:
        return False
    if reference.doc_kind and reference.doc_kind.lower() not in low:
        return False
    if not any(
        word in low
        for word in (
            "федеральн",
            "россии",
            "гост",
            "снип",
            "правил",
            "постановлен",
            "приказ",
        )
    ):
        return False
    return True


def _build_search_queries(reference: NormativeReference) -> list[str]:
    """Запросы без «№» и «от …» — иначе te-cloud ищет точную фразу в кавычках."""
    queries: list[str] = []
    kind = reference.doc_kind.lower() if reference.doc_kind else ""
    number = reference.number
    date_tokens = _normalize_date_tokens(reference.date)
    issuer = reference.issuer

    core: list[str] = []
    if kind:
        core.append(kind)
    if number:
        core.append(number)
    if date_tokens:
        core.extend(date_tokens.split())

    if core:
        queries.append(" ".join(core))
    if issuer and number:
        queries.append(f"{issuer} {number}")
        if kind:
            queries.append(f"{kind.lower()} {number} {issuer}")
    elif kind == "приказ" and number:
        queries.append(f"Ростехнадзор {number}")
        queries.append(f"приказ Ростехнадзора {number}")
    if number:
        queries.append(number)
    if not queries and reference.raw:
        cleaned = re.sub(r"(?:№|N|No\.?)\s*", " ", reference.raw, flags=re.IGNORECASE)
        cleaned = re.sub(r"\bот\b", " ", cleaned, flags=re.IGNORECASE)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()
        queries.append(cleaned[:120])

    out: list[str] = []
    for q in queries:
        q = q.strip()
        if q and q not in out:
            out.append(q)
    return out[:4]


def _api_error_message(html: str) -> str:
    low = html.lower()
    if "toomanyusers" in low or "максимальное число подключений" in low:
        return (
            "Техэксперт: превышено число одновременных подключений. "
            "Закройте лишние сессии или повторите позже."
        )
    if "ошибка авторизации" in low:
        return "Неверный логин или пароль Техэксперт"
    if 'id="user"' in html and "авторизация" in low:
        return "Не удалось войти в Техэксперт (осталась форма входа)"
    return ""


def _api_timestamp() -> str:
    return datetime.now().strftime("%d-%m-%Y %H:%M:%S")


def _html_to_text(html: str) -> str:
    parser = _TextExtractor()
    parser.feed(html)
    return parser.text()


def _html_to_document_text(html: str) -> str:
    """Текст нормативного документа без панели «Найденные фразы» и прочего UI."""
    for pat in (
        r"<div[^>]+id=[\"'](?:text|document|doc|content)[\"'][^>]*>(.*?)</div>",
        r"<article[^>]*>(.*?)</article>",
    ):
        m = re.search(pat, html, flags=re.S | re.I)
        if m:
            chunk = _html_to_text(m.group(1))
            if len(chunk) >= _MIN_DOCUMENT_PLAIN_CHARS:
                return chunk

    cleaned = html
    for pat in (
        r"<div[^>]+class=[\"'][^\"']*(?:iFind|ifind|phrase|contextSearch)[^\"']*[\"'][^>]*>.*?</div>",
        r"<form[^>]*class=[\"'][^\"']*(?:iFind|search)[^\"']*[\"'][^>]*>.*?</form>",
    ):
        cleaned = re.sub(pat, "", cleaned, flags=re.S | re.I)
    return _html_to_text(cleaned)


def parse_normative_reference(text: str) -> NormativeReference:
    """Разбор ссылки из B19 для поиска в Техэксперт."""
    raw = (text or "").strip()
    if not raw:
        return NormativeReference(raw="", search_query="")

    points = re.findall(
        r"(?:"
        r"[PpПп]\.?\s*"
        r"|п\.?\s*"
        r"|пункт\s+"
        r"|пп\.?\s*"
        r"|подп\.?\s*"
        r"|§\s*"
        r")"
        r"(\d+(?:\.\d+)*)",
        raw,
        flags=re.IGNORECASE,
    )
    points = list(dict.fromkeys(points))

    number_m = re.search(
        r"(?:№|N|No\.?)\s*([0-9][0-9A-Za-z./-]*)",
        raw,
        flags=re.IGNORECASE,
    )
    number = number_m.group(1).rstrip(".,;") if number_m else ""
    if not number and re.search(r"\bприказ\b", raw, flags=re.IGNORECASE):
        plain_num = re.search(
            r"\bприказ\b[^0-9\n]{0,40}(\d{1,5})\b",
            raw,
            flags=re.IGNORECASE,
        )
        if plain_num:
            number = plain_num.group(1)

    date_m = re.search(
        r"(\d{1,2}[./]\d{1,2}[./]\d{2,4}|\d{1,2}\s+[а-яА-Я]+\s+\d{4})",
        raw,
    )
    date = date_m.group(1) if date_m else ""

    kind = ""
    for label, pattern in (
        ("ГОСТ", r"\bГОСТ\b"),
        ("СП", r"\bСП\b"),
        ("СНиП", r"\bСНиП\b"),
        ("Приказ", r"\bприказ\b"),
        ("ПБ", r"\bПБ\b"),
        ("ФЗ", r"\b\d+\s*-?\s*ФЗ\b|\bФЗ\b"),
    ):
        if re.search(pattern, raw, flags=re.IGNORECASE):
            kind = label
            break

    issuer = _extract_issuer(raw)
    ref = NormativeReference(
        raw=raw,
        search_query="",
        doc_kind=kind,
        number=number,
        date=date,
        issuer=issuer,
        points=points,
    )
    queries = _build_search_queries(ref)
    ref.search_query = queries[0] if queries else raw[:200]
    return ref


def _extract_points_excerpt(full_text: str, points: list[str]) -> str:
    if not full_text or not points:
        return full_text[:_MAX_EXCERPT]

    chunks: list[str] = []
    for pt in points[:5]:
        patterns = [
            rf"(?:п\.?\s*|пункт\s+){re.escape(pt)}\b.{{0,1200}}",
            rf"\b{re.escape(pt)}\s+[\.\)]\s*.{{0,1200}}",
            rf"(?:^|\n)\s*{re.escape(pt)}\.\s+.{{0,1200}}",
        ]
        for pat in patterns:
            m = re.search(pat, full_text, flags=re.IGNORECASE | re.DOTALL)
            if m:
                chunks.append(m.group(0).strip())
                break

    if chunks:
        return "\n\n---\n\n".join(chunks)[:_MAX_EXCERPT]
    return full_text[:_MAX_EXCERPT]


def _extract_doc_title(full_text: str, fallback: str) -> str:
    for pat in (
        r"(?:Приказ|Постановление|ГОСТ|СП|СНиП|Правила)[^\n]{10,240}",
        r"ПРИКАЗ[^\n]{10,240}",
    ):
        m = re.search(pat, full_text, flags=re.IGNORECASE)
        if m:
            return re.sub(r"\s+", " ", m.group(0)).strip()
    return fallback


def _title_from_reference(reference: NormativeReference) -> str:
    """Краткий заголовок вида «Приказ … от … N …» из разбора B19."""
    if not reference.number:
        return ""
    kind = reference.doc_kind or "Приказ"
    issuer = reference.issuer or ""
    date = reference.date or ""
    if issuer and date:
        return f"{kind} {issuer} от {date} N {reference.number}"
    if date:
        return f"{kind} от {date} N {reference.number}"
    return f"{kind} N {reference.number}"


def _short_doc_title(title: str, reference: NormativeReference | None = None) -> str:
    """
    Краткое наименование для B19 — как title в выдаче Техэксперт.
    Длинный заголовок («Об утверждении ФНП…») обрезается до «Приказ … от … N …».
    """
    t = re.sub(r"\s+", " ", (title or "").strip())
    if not t:
        return _title_from_reference(reference) if reference else ""

    if len(t) <= 100 and not re.search(
        r"Об утверждении|Зарегистрировано|Федеральных норм и правил",
        t,
        flags=re.IGNORECASE,
    ):
        return t

    m = re.match(
        r"((?:Приказ|Постановление|ГОСТ|СП|СНиП|Федеральный закон)\s+.+?\s+"
        r"от\s+\d{1,2}[./]\d{1,2}[./]\d{2,4}\s+(?:N|№|No\.?)\s*[\w./-]+)",
        t,
        flags=re.IGNORECASE,
    )
    if m:
        return m.group(1).strip()

    if reference:
        built = _title_from_reference(reference)
        if built and reference.number in t:
            return built
    return t[:100].strip()


def _parse_ifind_items(html: str) -> list[tuple[str, str, str]]:
    """Элементы выдачи поиска: (nd, заголовок, href)."""
    items: list[tuple[str, str, str]] = []
    for nd, body in re.findall(r'data-nd="(\d+)"[^>]*>(.*?)</li>', html, flags=re.S):
        href_m = re.search(r'href="([^"]+)"', body, flags=re.I)
        href = href_m.group(1) if href_m else ""
        text = _html_to_text(body)
        text = re.sub(r"\s+", " ", text).strip()
        if text or href:
            items.append((nd, text, href))
    if items:
        return items
    for nd in re.findall(r'data-nd="(\d+)"', html):
        items.append((nd, "", ""))
    return items


def _prioritize_ifind_tabs(tabs: list[dict[str, Any]]) -> list[str]:
    """Вкладки выдачи: сначала «документы», не термины/фразы."""
    if not tabs:
        return [_SEARCH_TAB_ID, "-1", "4"]
    scored: list[tuple[int, str]] = []
    for tab in tabs:
        tid = str(tab.get("id", ""))
        if not tid:
            continue
        title = str(tab.get("title") or tab.get("name") or "").lower()
        score = 0
        if any(w in title for w in ("документ", "нормат", "закон", "приказ", "постанов", "акт")):
            score += 20
        if any(w in title for w in ("термин", "словар", "фраз", "контекст")):
            score -= 20
        if tid == _SEARCH_TAB_ID:
            score += 2
        scored.append((score, tid))
    scored.sort(key=lambda x: x[0], reverse=True)
    out = [tid for _, tid in scored]
    for fallback in (_SEARCH_TAB_ID, "-1", "4"):
        if fallback not in out:
            out.append(fallback)
    return out[:6]


def _is_dictionary_hit(title: str) -> bool:
    if not title:
        return True
    compact = title.strip()
    if len(compact) <= _TERM_TITLE_MAX_LEN:
        return True
    if compact.lower() in {"приказ", "постановление", "гост", "сп", "снип"}:
        return True
    return False


def _score_title(title: str, reference: NormativeReference) -> int:
    if not _title_looks_like_document(title, reference):
        return -1
    low = title.lower()
    score = 0
    if reference.number and reference.number in title:
        score += 4
    if reference.doc_kind and reference.doc_kind.lower() in low:
        score += 2
    if reference.issuer and reference.issuer.lower()[:8] in low:
        score += 5
    if reference.date:
        ddmm = reference.date.replace("/", ".").split(".")
        if len(ddmm) >= 3 and ddmm[-1] in title:
            score += 2
    return score


def _score_document_header(header: str, reference: NormativeReference) -> int:
    if not header or len(header) < 80:
        return -1
    low = header.lower()
    score = 0
    if reference.number:
        num = reference.number.lower()
        if re.search(rf"\bn\s*{re.escape(num)}\b", low) or f"№ {num}" in low:
            score += 6
        elif num in low:
            score += 3
    if reference.issuer and reference.issuer.lower()[:10] in low:
        score += 8
    if reference.doc_kind and reference.doc_kind.lower() in low:
        score += 2
    if reference.date:
        m = re.match(r"(\d{1,2})[./](\d{1,2})[./](\d{2,4})", reference.date)
        if m:
            dd, mm, yy = m.groups()
            if len(yy) == 2:
                yy = "20" + yy
            for token in (f"{dd}.{mm}.{yy}", f"{dd} {mm} {yy}", f"{int(dd)} {mm}"):
                if token in low:
                    score += 4
                    break
    return score


def _fetch_document_plain(
    client: "TechExpertClient",
    nd: str,
    href: str = "",
) -> tuple[str, str, str]:
    """Загрузить текст документа по nd / ссылке из выдачи. Возвращает (plain, url, api_error)."""
    urls: list[str] = []
    if href and not href.lower().startswith("javascript"):
        urls.append(urljoin(client.catalog_url + "/", href.lstrip("/")))
    urls.extend(
        [
            f"{client.catalog_url}/text?nd={nd}",
            f"{client.catalog_url}/?nd={nd}",
            f"{client.catalog_url}/?frame=center&nd={nd}",
        ]
    )
    seen: set[str] = set()
    best_plain = ""
    best_url = urls[0] if urls else f"{client.catalog_url}/text?nd={nd}"

    for doc_url in urls:
        if doc_url in seen:
            continue
        seen.add(doc_url)
        try:
            resp = client.session.get(
                doc_url,
                timeout=60,
                headers={"Referer": f"{client.catalog_url}/?frame=center"},
            )
        except requests.RequestException:
            continue
        api_err = _api_error_message(resp.text)
        if api_err or resp.status_code != 200 or len(resp.text) < 200:
            continue
        plain = _html_to_document_text(resp.text)
        if len(plain) > len(best_plain):
            best_plain = plain
            best_url = doc_url

    if len(best_plain) < _MIN_DOCUMENT_PLAIN_CHARS:
        api_base = client.catalog.lstrip("/")
        ts = _api_timestamp()
        for path, payload in (
            (f"{api_base}/doc_text", {"nd": nd, "part": "0", "_t": ts}),
            (f"{api_base}/get_text", {"nd": nd, "_t": ts}),
        ):
            try:
                alt = client._api_post(path, payload)
            except requests.RequestException:
                continue
            if _api_error_message(alt.text):
                continue
            alt_plain = _html_to_document_text(alt.text)
            if len(alt_plain) > len(best_plain):
                best_plain = alt_plain
                best_url = f"{client.catalog_url}/text?nd={nd}"

    if not best_plain:
        return "", best_url, f"Пустой ответ Техэксперт для nd={nd}"
    return best_plain, best_url, ""


def _pick_document_candidate(
    client: "TechExpertClient",
    candidates: list[tuple[str, str, str]],
    reference: NormativeReference,
    queries: list[str] | None = None,
) -> tuple[str, str, str]:
    """Выбрать nd, href и краткий title из выдачи поиска."""
    if not candidates:
        return "", "", ""

    queries = queries or _build_search_queries(reference)
    ranked: list[tuple[int, str, str, str]] = []
    for nd, title, href in candidates:
        if _title_is_query_echo(title, queries):
            continue
        score = _score_title(title, reference)
        if score >= 0:
            ranked.append((score, nd, title, href))

    ranked.sort(key=lambda x: x[0], reverse=True)
    top = ranked[:12]
    if not top:
        top = [
            (0, nd, title, href)
            for nd, title, href in candidates[:12]
            if not _title_is_query_echo(title, queries)
        ]

    best_nd = ""
    best_href = ""
    best_list_title = ""
    best_score = -1
    for title_score, nd, title, href in top:
        plain, _url, api_err = _fetch_document_plain(client, nd, href)
        if api_err or not _is_valid_document_plain(plain, reference):
            continue
        header_score = max(title_score, _score_document_header(plain[:15000], reference))
        if header_score > best_score:
            best_score = header_score
            best_nd = nd
            best_href = href
            best_list_title = title

    if best_nd:
        return best_nd, best_href, best_list_title

    for nd, title, href in candidates:
        if _title_is_query_echo(title, queries):
            continue
        if reference.number and reference.number in title:
            plain, _url, api_err = _fetch_document_plain(client, nd, href)
            if not api_err and _is_valid_document_plain(plain, reference):
                return nd, href, title
    return "", "", ""


class TechExpertClient:
    def __init__(self) -> None:
        cfg = _config()
        self.base_url = cfg["base_url"]
        self.catalog = cfg["catalog"]
        self.login = cfg["login"]
        self.password = cfg["password"]
        self.use_browser = cfg["use_browser"]
        self.session = requests.Session()
        self._logged_in = False
        self.session.headers.update(
            {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/120.0.0.0 Safari/537.36"
                ),
                "Accept-Language": "ru-RU,ru;q=0.9",
            }
        )

    @property
    def catalog_url(self) -> str:
        return urljoin(self.base_url + "/", self.catalog.lstrip("/"))

    def _missing_credentials_error(self) -> str:
        if not self.login or not self.password:
            from sk_reporter.prescriptions.te_env import te_expert_env_path

            env_path = te_expert_env_path()
            hint = (
                f"Создайте {env_path} (скопируйте te_expert.env.example) "
                "и укажите TE_EXPERT_LOGIN / TE_EXPERT_PASSWORD"
            )
            if env_path.is_file():
                hint = (
                    f"Проверьте TE_EXPERT_LOGIN и TE_EXPERT_PASSWORD в {env_path} "
                    "(нужен te_expert.env, не .example)"
                )
            return f"Не заданы TE_EXPERT_LOGIN и TE_EXPERT_PASSWORD. {hint}"
        return ""

    def _login_http(self) -> tuple[bool, str]:
        err = self._missing_credentials_error()
        if err:
            return False, err

        last_err = ""
        for attempt in range(_LOGIN_RETRIES):
            if attempt > 0:
                time.sleep(_LOGIN_RETRY_DELAY_S)
            try:
                self.session.get(self.catalog_url, timeout=30)
                resp = self.session.post(
                    urljoin(self.base_url, "/users/login.asp"),
                    data={
                        "user": self.login,
                        "pass": self.password,
                        "path": self.catalog,
                    },
                    headers={
                        "Content-Type": "application/x-www-form-urlencoded",
                        "X-Requested-With": "XMLHttpRequest",
                    },
                    timeout=30,
                    allow_redirects=False,
                )
                if resp.status_code not in {200, 204, 302, 307}:
                    last_err = f"HTTP {resp.status_code} при входе в Техэксперт"
                    continue

                page = self.session.get(self.catalog_url + "/", timeout=30)
                api_err = _api_error_message(page.text)
                if api_err:
                    last_err = api_err
                    if "подключений" in api_err.lower():
                        continue
                    return False, api_err
                if "id=\"user\"" in page.text and "авторизация" in page.text.lower():
                    return False, "Не удалось войти в Техэксперт (осталась форма входа)"
                self._logged_in = True
                return True, ""
            except requests.RequestException as e:
                last_err = f"Сеть Техэксперт: {e}"
        return False, last_err or "Не удалось войти в Техэксперт"

    def _ensure_login(self) -> tuple[bool, str]:
        if self._logged_in:
            return True, ""
        return self._login_http()

    def _sync_playwright_cookies(self, page: Any) -> None:
        for cookie in page.context.cookies():
            name = cookie.get("name")
            value = cookie.get("value")
            if not name:
                continue
            self.session.cookies.set(
                name,
                value,
                domain=cookie.get("domain"),
                path=cookie.get("path") or "/",
            )

    def _playwright_cookies_from_session(self) -> list[dict[str, str]]:
        host = urlparse(self.base_url).hostname or ""
        cookies: list[dict[str, str]] = []
        for c in self.session.cookies:
            cookies.append(
                {
                    "name": c.name,
                    "value": c.value,
                    "domain": c.domain or host,
                    "path": c.path or "/",
                }
            )
        return cookies

    def _frame_by_name(self, page: Any, name: str) -> Any | None:
        for frame in page.frames:
            if f"frame={name}" in frame.url:
                return frame
        return None

    def _run_search_in_left_frame(self, page: Any, query: str) -> bool:
        left = self._frame_by_name(page, "left")
        if not left:
            return False
        return bool(
            left.evaluate(
                """(q) => {
                    const el = document.querySelector('#context');
                    if (!el) return false;
                    el.value = q;
                    el.dispatchEvent(new Event('input', {bubbles: true}));
                    const btn = document.querySelector('.iSearchForm-contextButton');
                    if (btn) btn.click();
                    else el.dispatchEvent(new KeyboardEvent('keydown', {key: 'Enter', bubbles: true}));
                    return true;
                }""",
                query,
            )
        )

    def _collect_candidates_from_page(self, page: Any) -> list[tuple[str, str, str]]:
        candidates: list[tuple[str, str, str]] = []
        seen: set[tuple[str, str]] = set()
        for frame in page.frames:
            for item in _parse_ifind_items(frame.content()):
                key = (item[0], item[1])
                if key in seen:
                    continue
                seen.add(key)
                candidates.append(item)
        return candidates

    def _click_search_result(self, page: Any, nd: str) -> bool:
        selectors = (
            f'li[data-nd="{nd}"] a',
            f'[data-nd="{nd}"] a',
            f'[data-nd="{nd}"]',
        )
        for frame in page.frames:
            for sel in selectors:
                loc = frame.locator(sel)
                if loc.count() > 0:
                    loc.first.click()
                    return True
        return False

    def _read_center_frame_text(self, page: Any) -> tuple[str, str]:
        center = self._frame_by_name(page, "center")
        if not center:
            return "", ""
        try:
            plain = center.evaluate("() => document.body ? document.body.innerText : ''")
        except Exception:
            plain = _html_to_document_text(center.content())
        return plain.strip(), center.url

    def _api_post(self, path: str, data: dict[str, str]) -> requests.Response:
        url = urljoin(self.base_url + "/", path.lstrip("/"))
        headers = {
            "X-Requested-With": "XMLHttpRequest",
            "Referer": f"{self.catalog_url}/?frame=left",
        }
        return self.session.post(url, data=data, headers=headers, timeout=45)

    def _collect_candidates(
        self, query: str, api_base: str
    ) -> tuple[list[tuple[str, str, str]], str]:
        ts = _api_timestamp()
        try:
            chk = self._api_post(
                f"{api_base}/check_iquery", {"query": query, "_t": ts}
            )
        except requests.RequestException as e:
            return [], f"Ошибка поиска в Техэксперт: {e}"

        api_err = _api_error_message(chk.text)
        if api_err:
            return [], api_err
        if chk.status_code != 200:
            return [], api_err or f"Техэксперт: HTTP {chk.status_code} на check_iquery"
        try:
            if not chk.json().get("good"):
                return [], ""
        except json.JSONDecodeError:
            return [], _api_error_message(chk.text) or "Техэксперт вернул некорректный ответ"

        try:
            ifr_resp = self._api_post(
                f"{api_base}/ifind_result",
                {
                    "query": query,
                    "searchByNames": "true",
                    "bp": "[]",
                    "archs": "",
                    "real": "",
                    "_t": ts,
                },
            )
            tabs = _prioritize_ifind_tabs(ifr_resp.json().get("tabs", []))
        except (requests.RequestException, json.JSONDecodeError):
            tabs = [_SEARCH_TAB_ID, "-1", "4"]

        candidates: list[tuple[str, str, str]] = []
        seen: set[str] = set()
        for tab_id in tabs[:4]:
            for part in range(_MAX_SEARCH_PAGES):
                try:
                    found = self._api_post(
                        f"{api_base}/ifind_list",
                        {
                            "query": query,
                            "part": str(part),
                            "id": tab_id,
                            "sp": "[]",
                            "bp": "[]",
                            "_t": ts,
                        },
                    )
                except requests.RequestException as e:
                    return candidates, f"Ошибка поиска в Техэксперт: {e}"

                api_err = _api_error_message(found.text)
                if api_err:
                    return candidates, api_err

                items = _parse_ifind_items(found.text)
                if not items:
                    break
                for nd, title, href in items:
                    if nd in seen:
                        continue
                    seen.add(nd)
                    candidates.append((nd, title, href))
                if len(candidates) >= _MAX_CANDIDATES:
                    break
            if len(candidates) >= _MAX_CANDIDATES:
                break
        return candidates, ""

    def _search_http(self, reference: NormativeReference) -> NormativeLookupResult:
        auth_ok, auth_err = self._ensure_login()
        if not auth_ok:
            return NormativeLookupResult(
                ok=False,
                reference=reference,
                error=auth_err,
                auth_ok=False,
            )

        api_base = self.catalog.lstrip("/")
        queries = _build_search_queries(reference)
        candidates: list[tuple[str, str, str]] = []
        seen_nd: set[str] = set()

        for query in queries:
            found, err = self._collect_candidates(query, api_base)
            if err and not found:
                return NormativeLookupResult(
                    ok=False,
                    reference=reference,
                    error=err,
                    auth_ok=True,
                )
            for nd, title, href in found:
                if nd not in seen_nd:
                    seen_nd.add(nd)
                    candidates.append((nd, title, href))

        if not candidates:
            tried = ", ".join(queries[:3])
            return NormativeLookupResult(
                ok=False,
                reference=reference,
                error=f"Документ не найден в Техэксперт (запросы: {tried})",
                auth_ok=True,
            )

        nd, href, list_title = _pick_document_candidate(
            self, candidates, reference, queries
        )
        if not nd:
            sample = "; ".join(title[:60] for _nd, title, _h in candidates[:3])
            return NormativeLookupResult(
                ok=False,
                reference=reference,
                error=(
                    "Не удалось выбрать документ среди результатов поиска "
                    f"(кандидатов: {len(candidates)}; примеры: {sample or '—'})"
                ),
                auth_ok=True,
            )

        plain, doc_url, fetch_err = _fetch_document_plain(self, nd, href)
        return self._result_from_plain(
            reference,
            plain,
            doc_url,
            fetch_err,
            auth_ok=True,
            list_title=list_title,
        )

    def _result_from_plain(
        self,
        reference: NormativeReference,
        plain: str,
        doc_url: str,
        fetch_err: str,
        *,
        auth_ok: bool,
        list_title: str = "",
    ) -> NormativeLookupResult:
        if fetch_err:
            return NormativeLookupResult(
                ok=False,
                reference=reference,
                error=fetch_err,
                auth_ok=auth_ok,
                source_url=doc_url,
            )
        if not _is_valid_document_plain(plain, reference):
            return NormativeLookupResult(
                ok=False,
                reference=reference,
                error=(
                    f"Техэксперт вернул не текст документа "
                    f"(символов: {len(plain)})"
                ),
                auth_ok=auth_ok,
                source_url=doc_url,
            )
        short = _short_doc_title(list_title, reference) or _short_doc_title(
            _extract_doc_title(plain, reference.search_query), reference
        )
        excerpt = _extract_points_excerpt(plain, reference.points) or plain[:_MAX_EXCERPT]
        return NormativeLookupResult(
            ok=True,
            reference=reference,
            excerpt=excerpt,
            doc_title=short,
            list_title=list_title or short,
            source_url=doc_url,
            auth_ok=auth_ok,
        )

    def _search_browser(self, reference: NormativeReference) -> NormativeLookupResult:
        """Поиск через строку на главной (frame=left) → клик по ссылке → текст в center."""
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            return NormativeLookupResult(
                ok=False,
                reference=reference,
                error=(
                    "Playwright не установлен. "
                    'pip install "sk-reporter[browser]" && playwright install chromium'
                ),
            )

        auth_ok, auth_err = self._ensure_login()
        if not auth_ok:
            return NormativeLookupResult(
                ok=False,
                reference=reference,
                error=auth_err,
                auth_ok=False,
            )

        queries = _build_search_queries(reference)
        cookies = self._playwright_cookies_from_session()

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(
                viewport={"width": 1600, "height": 1000},
                locale="ru-RU",
            )
            if cookies:
                context.add_cookies(cookies)
            page = context.new_page()
            try:
                page.goto(self.catalog_url + "/", wait_until="networkidle", timeout=90000)
                page.wait_for_timeout(int(_SEARCH_WAIT_S * 1000))

                body = page.inner_text("body")
                if "максимальное число подключений" in body.lower():
                    browser.close()
                    return NormativeLookupResult(
                        ok=False,
                        reference=reference,
                        error=(
                            "Техэксперт: превышено число одновременных подключений. "
                            "Закройте лишние вкладки Техэксперт."
                        ),
                        auth_ok=True,
                    )
                if "id=\"user\"" in page.content() and "авторизация" in body.lower():
                    browser.close()
                    return NormativeLookupResult(
                        ok=False,
                        reference=reference,
                        error="Сессия Техэксперт не открыла каталог после входа",
                        auth_ok=False,
                    )

                self._sync_playwright_cookies(page)
                candidates: list[tuple[str, str, str]] = []
                for query in queries:
                    if not self._run_search_in_left_frame(page, query):
                        continue
                    page.wait_for_timeout(int(_SEARCH_WAIT_S * 1000))
                    found = self._collect_candidates_from_page(page)
                    if found:
                        candidates = found
                        break

                if not candidates:
                    browser.close()
                    return NormativeLookupResult(
                        ok=False,
                        reference=reference,
                        error="В строке поиска не найдено ссылок на документ",
                        auth_ok=True,
                    )

                nd, href, list_title = _pick_document_candidate(
                    self, candidates, reference, queries
                )
                if not nd:
                    browser.close()
                    sample = "; ".join(t[:60] for _n, t, _h in candidates[:3])
                    return NormativeLookupResult(
                        ok=False,
                        reference=reference,
                        error=f"Среди ссылок поиска нет подходящего документа ({sample})",
                        auth_ok=True,
                    )

                if not self._click_search_result(page, nd):
                    browser.close()
                    return NormativeLookupResult(
                        ok=False,
                        reference=reference,
                        error=f"Не удалось открыть документ nd={nd} из выдачи поиска",
                        auth_ok=True,
                    )
                page.wait_for_timeout(int(_SEARCH_WAIT_S * 1000))

                plain, doc_url = self._read_center_frame_text(page)
                if not _is_valid_document_plain(plain, reference):
                    plain, doc_url, fetch_err = _fetch_document_plain(self, nd, href)
                    browser.close()
                    return self._result_from_plain(
                        reference,
                        plain,
                        doc_url,
                        fetch_err,
                        auth_ok=True,
                        list_title=list_title,
                    )

                browser.close()
                return self._result_from_plain(
                    reference,
                    plain,
                    doc_url,
                    "",
                    auth_ok=True,
                    list_title=list_title,
                )
            except Exception as e:
                browser.close()
                return NormativeLookupResult(
                    ok=False,
                    reference=reference,
                    error=f"Браузер Техэксперт: {e}",
                )

    def lookup(self, normative_text: str) -> NormativeLookupResult:
        reference = parse_normative_reference(normative_text)
        if not reference.raw:
            return NormativeLookupResult(
                ok=False,
                reference=reference,
                error="Пустая ссылка на нормативный документ (B19)",
            )

        result = self._search_http(reference)
        if result.ok:
            return result

        if self._missing_credentials_error():
            return result

        if not self.use_browser:
            return result

        # Запасной канал: строка поиска на главной → ссылка → документ в center
        ui_result = self._search_browser(reference)
        return ui_result if ui_result.ok else result


def lookup_normative(normative_text: str) -> dict[str, Any]:
    """Удобная обёртка для check_agent: Техэксперт → интернет (запасной)."""
    result = TechExpertClient().lookup(normative_text)
    te_error = ""
    if not result.ok:
        te_error = result.error
        from sk_reporter.prescriptions.normative_web_fallback import (
            internet_fallback_enabled,
            lookup_normative_web,
        )

        if internet_fallback_enabled():
            print(
                f"[TECHEXPERT] internet fallback after TE error: "
                f"{(te_error or '')[:160]}"
            )
            web = lookup_normative_web(normative_text)
            if web.ok:
                out = _normative_result_to_dict(web)
                out["te_fallback_error"] = te_error
                return out
    out = _normative_result_to_dict(result)
    if te_error:
        out["te_fallback_error"] = te_error
    return out


def _normative_result_to_dict(result: NormativeLookupResult) -> dict[str, Any]:
    return {
        "ok": result.ok,
        "reference": {
            "raw": result.reference.raw,
            "search_query": result.reference.search_query,
            "doc_kind": result.reference.doc_kind,
            "number": result.reference.number,
            "date": result.reference.date,
            "issuer": result.reference.issuer,
            "points": result.reference.points,
        },
        "excerpt": result.excerpt,
        "doc_title": result.doc_title,
        "list_title": result.list_title,
        "source_url": result.source_url,
        "error": result.error,
        "auth_ok": result.auth_ok,
        "source": result.source,
        "te_fallback_error": "",
    }
