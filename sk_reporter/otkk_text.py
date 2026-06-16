"""Текст ОТКК как в Word: без служебных полей ссылок Кодекса (HYPERLINK)."""

from __future__ import annotations

import re

# Поле ссылки Word/textutil: HYPERLINK "url" \o "подсказка" + видимый текст после "
_KODEX_FIELD_RE = re.compile(
    r'HYPERLINK\s+"[^"]*"\s*\\o\s*"(?:\\.|[^"\\])*"',
    re.I,
)
_STATUS_TAIL_RE = re.compile(
    r'Статус:\s*Действующий документ\.[^"]*"\s*',
    re.I,
)


def strip_kodeks_fields(text: str) -> str:
    """Убрать HYPERLINK и хвосты подсказок; оставить видимый текст карты."""
    if not text or "HYPERLINK" not in text:
        return text.replace("\x07", "")
    out = _KODEX_FIELD_RE.sub("", text)
    out = re.sub(r"\(утв\.[^)]*\)", "", out)
    out = _STATUS_TAIL_RE.sub("", out)
    return out.replace("\x07", "")


def normative_visible_text(block: str) -> str:
    """П.4: только коды СП через запятую, как в ячейке Word."""
    stripped = strip_kodeks_fields(block)
    codes = re.findall(r"(?:СП|ГОСТ|ВСН)\s*[\d][\d.\-]*", stripped)
    seen: set[str] = set()
    out: list[str] = []
    for raw in codes:
        code = re.sub(r"\s+", " ", raw.strip())
        key = code.casefold()
        if key not in seen:
            seen.add(key)
            out.append(code)
    return ", ".join(out)


def sanitize_otkk_rows(rows: list[dict]) -> list[dict]:
    """Перед записью в БД: никакого HYPERLINK в value."""
    cleaned: list[dict] = []
    for row in rows:
        label = str(row.get("label") or "")
        value = str(row.get("value") or "")
        if "Нормативные документы" in label:
            if "HYPERLINK" in value or len(value) > 200:
                value = normative_visible_text(value)
            else:
                value = strip_kodeks_fields(value)
        else:
            value = strip_kodeks_fields(value)
        cleaned.append({**row, "value": value})
    return cleaned
