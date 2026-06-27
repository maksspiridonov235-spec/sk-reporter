"""Разбор ссылки на нормативный документ из ячейки B19 (без Техэксперт)."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

_MAX_EXCERPT = 12_000


@dataclass
class NormativeReference:
    raw: str
    search_query: str
    doc_kind: str = ""
    number: str = ""
    date: str = ""
    issuer: str = ""
    points: list[str] = field(default_factory=list)


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
    low = issuer.lower()
    fixes = {
        "ростехнадзора": "Ростехнадзор",
        "роструда": "Роструд",
    }
    return fixes.get(low, issuer)


def _build_search_queries(reference: NormativeReference) -> list[str]:
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


def parse_normative_reference(text: str) -> NormativeReference:
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


def extract_points_excerpt(full_text: str, points: list[str]) -> str:
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


def title_from_reference(reference: NormativeReference) -> str:
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


def short_doc_title(title: str, reference: NormativeReference | None = None) -> str:
    t = re.sub(r"\s+", " ", (title or "").strip())
    if not t:
        return title_from_reference(reference) if reference else ""

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
        built = title_from_reference(reference)
        if built and reference.number in t:
            return built
    return t[:100].strip()
