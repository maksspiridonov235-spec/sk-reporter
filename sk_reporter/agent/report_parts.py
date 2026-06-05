"""Парсинг ЧАСТЕЙ 1–4 из ответа check/verify агента."""

import re

_VOLUME_LABELS = ("Проектный объем", "Объем за сутки", "Накопительный объем")

_GENERIC_DESC_RE = re.compile(
    r"проверен.*замечаний нет|работы ведутся в соответствии",
    re.IGNORECASE,
)


def split_part(search_text: str, num: int, next_num: int | None) -> list[str]:
    if next_num is not None:
        pattern = rf"ЧАСТЬ\s*{num}[^:\n]*[:\n](.*?)(?=ЧАСТЬ\s*{next_num}\b)"
    else:
        pattern = rf"ЧАСТЬ\s*{num}[^:\n]*[:\n](.*?)$"
    m = re.search(pattern, search_text, re.DOTALL | re.IGNORECASE)
    if not m:
        return []
    return [line.rstrip() for line in m.group(1).strip().splitlines()]


def parse_parts(corrected_text: str):
    """Parse parts 1–4 from LLM output. Returns (part1, part2, part3, part4) line lists."""
    cleaned = re.sub(r"\*\*([^*]+)\*\*", r"\1", corrected_text)

    section_match = re.search(
        r"##\s*ИСПРАВЛЕННЫЙ\s*ОТЧЁТ[^\n]*\n(.*?)$", cleaned, re.DOTALL | re.IGNORECASE
    )
    search_text = section_match.group(1).strip() if section_match else cleaned.strip()

    part1_lines = split_part(search_text, 1, 2)
    part2_lines = split_part(search_text, 2, 3)
    part3_lines = split_part(search_text, 3, 4)
    part4_lines = split_part(search_text, 4, None)

    if not part1_lines and not part2_lines:
        p2_start = re.search(
            r"(Наряд.допуск|Работы ведутся)", search_text, re.IGNORECASE
        )
        if p2_start:
            raw_part1 = search_text[: p2_start.start()].strip()
            raw_part2 = search_text[p2_start.start() :].strip()
            raw_part1 = re.sub(r"^ЧАСТЬ\s*1[^\n]*\n", "", raw_part1, flags=re.IGNORECASE).strip()
            part1_lines = [l.rstrip() for l in raw_part1.splitlines()] if raw_part1 else []
            part2_lines = [l.rstrip() for l in raw_part2.splitlines()] if raw_part2 else []

    return part1_lines, part2_lines, part3_lines, part4_lines


def extract_volume_from_line(line: str, label: str) -> str:
    pattern = rf"{re.escape(label)}\s*[–\-]?\s*([^\n;]+)"
    m = re.search(pattern, line, re.IGNORECASE)
    return m.group(1).strip() if m else ""


def normalize_volume_value(value: str) -> str:
    if not value:
        return ""
    v = value.lower().strip()
    v = re.sub(r"\s+", " ", v)
    v = v.replace("–", "-").replace("—", "-")
    return v


def volume_numeric(value: str) -> str | None:
    m = re.search(r"([\d]+(?:[.,]\d+)?)", normalize_volume_value(value))
    return m.group(1).replace(",", ".") if m else None


def count_numbered_items(lines: list[str]) -> int:
    return sum(1 for line in lines if re.match(r"^\s*\d+\.\s*\S", line))


def parse_numbered_items(lines: list[str]) -> dict[int, str]:
    items: dict[int, list[str]] = {}
    current_num: int | None = None
    for line in lines:
        m = re.match(r"^\s*(\d+)\.\s*(.*)$", line.strip())
        if m:
            current_num = int(m.group(1))
            rest = m.group(2).strip()
            items[current_num] = [rest] if rest else []
        elif current_num is not None and line.strip():
            items[current_num].append(line.strip())
    return {n: " ".join(parts).strip() for n, parts in items.items()}


def parse_works_from_part1_lines(lines: list[str]) -> list[dict]:
    works: list[dict] = []
    current: dict | None = None
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        m = re.match(r"^(\d+)\.\s*(.+)$", stripped)
        if m:
            if current:
                works.append(current)
            current = {
                "num": int(m.group(1)),
                "title": m.group(2).strip(),
                "volumes": {},
            }
            continue
        if not current:
            continue
        for label in _VOLUME_LABELS:
            if label.lower() in stripped.lower():
                current["volumes"][label] = extract_volume_from_line(stripped, label)
    if current:
        works.append(current)
    return works


def split_part1_part2_text(text: str) -> tuple[str, str]:
    m = re.search(r"(Наряд.допуск|Работы ведутся)", text, re.IGNORECASE)
    if not m:
        return text.strip(), ""
    return text[: m.start()].strip(), text[m.start() :].strip()


def is_substantive_engineer_description(text: str) -> bool:
    t = text.strip()
    if not t:
        return False
    if len(t) < 30:
        return False
    if _GENERIC_DESC_RE.search(t) and len(t) < 120:
        return False
    return True


def engineer_facts_preserved(original: str, corrected: str) -> bool:
    if not is_substantive_engineer_description(original):
        return True
    words = [w.lower() for w in re.findall(r"\w{5,}", original)]
    if not words:
        return True
    corrected_lower = corrected.lower()
    hits = sum(1 for w in words if w in corrected_lower)
    return hits / len(words) >= 0.25


def is_zero_daily_volume(daily: str) -> bool:
    num = volume_numeric(daily)
    if num is None:
        return False
    try:
        return float(num) == 0.0
    except ValueError:
        return False
