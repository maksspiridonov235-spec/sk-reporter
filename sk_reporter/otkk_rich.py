"""Блоки п.6 ОТКК: заголовки, списки и таблицы как в Word (из HTML .doc)."""

from __future__ import annotations

import re
from html import escape
from typing import Any

from lxml import html as lxml_html

from sk_reporter.otkk_parser import _clean_text, doc_to_html
from sk_reporter.otkk_text import strip_kodeks_fields

_TABLE_CAPTION_RE = re.compile(r"^Таблица\s*(?:№|N)?\s*\d+", re.I)
_NUMBERED_ROW_RE = re.compile(r"^\d+\.\s")
_ROMAN_SECTION_RE = re.compile(r"^[IVXLC]+\.\s")
_DEVIATION_RE = re.compile(
    r"(?:\d+\s*мм|±\s*\d|–\s*\d|0,\d+|1/\d|^\d+$)",
    re.I,
)
_CONTROL_RE = re.compile(
    r"(?:^То же$|Измерительный|Инструментальный|журнал|исполн)",
    re.I,
)


def _para_text(p) -> str:
    return _clean_text(" ".join(p.xpath(".//text()")))


def _is_header_triplet(lines: list[str], i: int) -> bool:
    if i + 2 >= len(lines):
        return False
    block = " ".join(lines[i : i + 3]).casefold()
    return "параметр" in block and "предельные" in block and "контроль" in block


def _is_deviation(line: str) -> bool:
    s = line.strip()
    if not s or _NUMBERED_ROW_RE.match(s):
        return False
    if s.casefold() == "то же":
        return True
    return bool(_DEVIATION_RE.search(s))


def _is_control(line: str) -> bool:
    s = line.strip()
    if not s:
        return False
    if s.casefold() == "то же":
        return True
    return bool(_CONTROL_RE.search(s))


def _split_table_rows(lines: list[str]) -> list[tuple[str, str, str]]:
    """Строки таблицы (параметр | отклонение | контроль) из плоского списка абзацев."""
    rows: list[tuple[str, str, str]] = []
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i].strip()
        if not line or _is_header_triplet(lines, i):
            if _is_header_triplet(lines, i):
                i += 3
            else:
                i += 1
            continue
        if not _NUMBERED_ROW_RE.match(line):
            i += 1
            continue

        param_parts = [line]
        i += 1
        while i < n:
            nxt = lines[i].strip()
            if not nxt:
                i += 1
                continue
            if _NUMBERED_ROW_RE.match(nxt):
                break
            if _is_deviation(nxt) or _is_control(nxt):
                break
            param_parts.append(nxt)
            i += 1

        dev_parts: list[str] = []
        while i < n:
            nxt = lines[i].strip()
            if not nxt:
                i += 1
                continue
            if _NUMBERED_ROW_RE.match(nxt):
                break
            if _is_control(nxt) and not _is_deviation(nxt):
                break
            if _is_deviation(nxt):
                dev_parts.append(nxt)
                i += 1
                continue
            break

        ctrl_parts: list[str] = []
        while i < n:
            nxt = lines[i].strip()
            if not nxt:
                i += 1
                continue
            if _NUMBERED_ROW_RE.match(nxt):
                break
            if _is_control(nxt):
                ctrl_parts.append(nxt)
                i += 1
                continue
            if _is_deviation(nxt):
                break
            break

        rows.append(
            (
                "\n".join(param_parts).strip(),
                "\n".join(dev_parts).strip(),
                "\n".join(ctrl_parts).strip(),
            )
        )
    return rows


def _paragraphs_from_doc_html(html: str) -> list[str]:
    root = lxml_html.fromstring(html)
    out: list[str] = []
    started = False
    for p in root.xpath("//p"):
        t = _para_text(p)
        if not t:
            continue
        if not started:
            if "Контролируемые параметры" in t:
                started = True
                rest = t.split("документация", 1)
                if len(rest) > 1 and rest[1].strip():
                    out.append(rest[1].strip())
            continue
        if t.startswith("Разработал"):
            break
        out.append(strip_kodeks_fields(t))
    return out


def _lines_to_segments(lines: list[str]) -> list[dict[str, Any]]:
    segments: list[dict[str, Any]] = []
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i].strip()
        if not line:
            i += 1
            continue

        if _TABLE_CAPTION_RE.match(line):
            caption = line
            i += 1
            body: list[str] = []
            while i < n:
                nxt = lines[i].strip()
                if not nxt:
                    i += 1
                    continue
                if _TABLE_CAPTION_RE.match(nxt) or _ROMAN_SECTION_RE.match(nxt):
                    break
                body.append(nxt)
                i += 1
            table_rows = _split_table_rows(body)
            if table_rows:
                segments.append(
                    {
                        "type": "table",
                        "caption": caption,
                        "headers": ["Параметр", "Предельные отклонения", "Контроль"],
                        "rows": [
                            {"cells": [a, b, c]} for a, b, c in table_rows
                        ],
                    }
                )
            else:
                segments.append({"type": "paragraph", "text": caption})
            continue

        if _ROMAN_SECTION_RE.match(line):
            segments.append({"type": "heading", "text": line})
            i += 1
            continue

        if line.startswith("- "):
            bullets = [line[2:].strip()]
            i += 1
            while i < n:
                nxt = lines[i].strip()
                if not nxt or not nxt.startswith("- "):
                    break
                bullets.append(nxt[2:].strip())
                i += 1
            segments.append({"type": "bullets", "items": bullets})
            continue

        para = [line]
        i += 1
        while i < n:
            nxt = lines[i].strip()
            if (
                not nxt
                or _TABLE_CAPTION_RE.match(nxt)
                or _ROMAN_SECTION_RE.match(nxt)
                or _NUMBERED_ROW_RE.match(nxt)
                or nxt.startswith("- ")
            ):
                break
            para.append(nxt)
            i += 1
        segments.append({"type": "paragraph", "text": "\n".join(para)})
    return segments


def extract_rich_segments(doc_path) -> list[dict[str, Any]]:
    html = doc_to_html(doc_path)
    lines = _paragraphs_from_doc_html(html)
    return _lines_to_segments(lines)


def segment_to_html(seg: dict[str, Any]) -> str:
    t = seg.get("type")
    if t == "heading":
        return f'<h4 class="otkk-section-heading">{escape(seg.get("text") or "")}</h4>'
    if t == "paragraph":
        return f'<p class="otkk-paragraph">{escape(seg.get("text") or "").replace(chr(10), "<br>")}</p>'
    if t == "bullets":
        items = "".join(f"<li>{escape(x)}</li>" for x in seg.get("items") or [])
        return f'<ul class="otkk-bullets">{items}</ul>'
    if t == "table":
        cap = escape(seg.get("caption") or "")
        headers = seg.get("headers") or []
        head = "".join(f"<th>{escape(h)}</th>" for h in headers)
        body_rows = []
        for row in seg.get("rows") or []:
            cells = row.get("cells") or []
            tds = "".join(
                f'<td>{escape(c or "").replace(chr(10), "<br>")}</td>' for c in cells
            )
            body_rows.append(f"<tr>{tds}</tr>")
        return (
            f'<div class="otkk-inner-table-wrap">'
            f'<div class="otkk-table-caption">{cap}</div>'
            f'<table class="otkk-inner-table"><thead><tr>{head}</tr></thead>'
            f'<tbody>{"".join(body_rows)}</tbody></table></div>'
        )
    return ""


def segments_to_html(segments: list[dict[str, Any]]) -> str:
    return "".join(segment_to_html(s) for s in segments)
