"""Блоки п.6 ОТКК: заголовки, списки и таблицы как в Word (из HTML .doc)."""

from __future__ import annotations

import re
from dataclasses import dataclass
from html import escape
from typing import Any, Callable

from lxml import html as lxml_html

from sk_reporter.otkk_parser import _clean_text, doc_to_html
from sk_reporter.otkk_text import strip_kodeks_fields

_TABLE_CAPTION_RE = re.compile(r"^Таблица\s*(?:№|N)?\s*(\d+)", re.I)
_ROMAN_SECTION_RE = re.compile(r"^[IVXLC]+\.\s")
_NUMBERED_ROW_RE = re.compile(r"^\d+\.\s")
_STATUS_JUNK_RE = re.compile(r"^\(утв\.|^Статус:|СП 70\.13330", re.I)
_SHAPE_JUNK_RE = re.compile(r"SHAPE\s*\\\*", re.I)
_CONTROL_HINT_RE = re.compile(
    r"(?:Измерительный|Инструментальный|журнал|исполн|геодезическ|То же|„)",
    re.I,
)
_HEADER_WORDS = ("параметр", "предельн", "контроль", "величина")


@dataclass(frozen=True)
class Para:
    cls: str
    text: str


def _para_text(p) -> str:
    return _clean_text(" ".join(p.xpath(".//text()")))


def _table_no(caption: str) -> str | None:
    m = _TABLE_CAPTION_RE.match(caption.strip())
    return m.group(1) if m else None


def _is_junk(p: Para) -> bool:
    t = p.text.strip()
    if not t or _SHAPE_JUNK_RE.search(t):
        return True
    return bool(_STATUS_JUNK_RE.search(t))


def _is_header_line(t: str) -> bool:
    low = t.casefold()
    return any(w in low for w in _HEADER_WORDS)


def _is_control(t: str) -> bool:
    return bool(_CONTROL_HINT_RE.search(t))


def _is_deviation_token(t: str) -> bool:
    s = t.strip()
    if not s or _NUMBERED_ROW_RE.match(s):
        return False
    if s in ('"', "“", "„"):
        return True
    if s.casefold() in ("то же", "—", "-", "–"):
        return True
    if re.search(r"(?:\d+\s*мм|±|0,\d+|1/\d|0\+\d|табл\.)", s, re.I):
        return True
    if re.fullmatch(r"[±–—]?\s*\d+([,.]\d+)?", s):
        return True
    if re.fullmatch(r"0[+\-]\d+", s):
        return True
    return False


def _paragraphs_with_class_from_doc_html(html: str) -> list[Para]:
    root = lxml_html.fromstring(html)
    out: list[Para] = []
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
                    out.append(Para("", strip_kodeks_fields(rest[1].strip())))
            continue
        if t.startswith("Разработал"):
            break
        out.append(Para(p.get("class") or "", strip_kodeks_fields(t)))
    return out


def _split_table_sections(paras: list[Para]) -> list[tuple[str, str, list[Para]]]:
    """(table_no, caption, body_paras)"""
    sections: list[tuple[str, str, list[Para]]] = []
    i = 0
    n = len(paras)
    while i < n:
        t = paras[i].text.strip()
        if _TABLE_CAPTION_RE.match(t):
            caption = t
            no = _table_no(caption) or ""
            i += 1
            body: list[Para] = []
            while i < n:
                nxt = paras[i].text.strip()
                if not nxt:
                    i += 1
                    continue
                if _TABLE_CAPTION_RE.match(nxt) or _ROMAN_SECTION_RE.match(nxt):
                    break
                body.append(paras[i])
                i += 1
            sections.append((no, caption, body))
            continue
        i += 1
    return sections


def _mk_table(
    caption: str,
    *,
    headers: list[str] | None = None,
    header_rows: list[list[str]] | None = None,
    rows: list[dict[str, Any]],
    layout: str = "standard",
) -> dict[str, Any]:
    seg: dict[str, Any] = {
        "type": "table",
        "caption": caption,
        "layout": layout,
        "rows": rows,
    }
    if header_rows:
        seg["header_rows"] = header_rows
    elif headers:
        seg["headers"] = headers
    return seg


def _row(cells: list[str], *, kind: str = "data") -> dict[str, Any]:
    return {"type": kind, "cells": cells}



def _para_role(
    p: Para,
    *,
    param_cls: set[str],
    dev_cls: set[str],
    ctrl_cls: set[str],
) -> str:
    t = p.text.strip()
    cls = p.cls
    if cls in dev_cls and (
        _is_deviation_token(t) or t in ('„', '"', "“") or t.casefold() == "то же"
    ):
        return "dev"
    if cls in ctrl_cls and _is_control(t):
        return "ctrl"
    if cls in param_cls:
        if cls in ctrl_cls and _is_control(t):
            return "ctrl"
        return "param"
    if _is_control(t):
        return "ctrl"
    if _is_deviation_token(t):
        return "dev"
    return "other"


def parse_table_class_map(body: list[Para], spec: dict[str, Any]) -> dict[str, Any]:
    """Универсальный разбор 3-колоночной таблицы по CSS-классам абзацев Word."""
    paras = [p for p in body if not _is_junk(p)]
    param_cls = set(spec["param_cls"])
    dev_cls = set(spec["dev_cls"])
    ctrl_cls = set(spec.get("ctrl_cls", ()))
    headers: list[str] = spec["headers"]
    rows: list[dict[str, Any]] = []

    i = 0
    while i < len(paras) and _is_header_line(paras[i].text):
        i += 1

    while i < len(paras):
        p = paras[i]
        t = p.text.strip()

        if _NUMBERED_ROW_RE.match(t):
            param = [t]
            i += 1
            while i < len(paras) and not _NUMBERED_ROW_RE.match(paras[i].text.strip()):
                role = _para_role(
                    paras[i],
                    param_cls=param_cls,
                    dev_cls=dev_cls,
                    ctrl_cls=ctrl_cls,
                )
                if role == "param":
                    param.append(paras[i].text.strip())
                    i += 1
                else:
                    break

            dev: list[str] = []
            while i < len(paras) and not _NUMBERED_ROW_RE.match(paras[i].text.strip()):
                role = _para_role(
                    paras[i],
                    param_cls=param_cls,
                    dev_cls=dev_cls,
                    ctrl_cls=ctrl_cls,
                )
                if role == "dev":
                    dev.append(paras[i].text.strip())
                    i += 1
                else:
                    break

            ctrl: list[str] = []
            while i < len(paras) and not _NUMBERED_ROW_RE.match(paras[i].text.strip()):
                role = _para_role(
                    paras[i],
                    param_cls=param_cls,
                    dev_cls=dev_cls,
                    ctrl_cls=ctrl_cls,
                )
                if role == "ctrl":
                    ctrl.append(paras[i].text.strip())
                    i += 1
                else:
                    break

            rows.append(
                _row(
                    [
                        "\n".join(param).strip(),
                        "\n".join(dev).strip(),
                        "\n".join(ctrl).strip(),
                    ]
                )
            )
            continue

        role = _para_role(p, param_cls=param_cls, dev_cls=dev_cls, ctrl_cls=ctrl_cls)
        if role == "param":
            rows.append(_row([t, "", ""], kind="section"))
            i += 1
            continue

        i += 1

    layout = spec.get("layout", "standard")
    if spec.get("header_rows"):
        return _mk_table("", header_rows=spec["header_rows"], rows=rows, layout=layout)
    return _mk_table("", headers=headers, rows=rows, layout=layout)


TABLE_SPECS: dict[str, dict[str, Any]] = {
    "1": {
        "headers": ["Параметр", "Предельные отклонения", "Контроль"],
        "param_cls": ("p16", "p17"),
        "dev_cls": ("p15",),
        "ctrl_cls": ("p16", "p15"),
    },
    "2": {
        "headers": [
            "Параметр",
            "Предельные отклонения, мм",
            "Контроль (метод, объем, вид регистрации)",
        ],
        "param_cls": ("p1", "p24"),
        "dev_cls": ("p23",),
        "ctrl_cls": ("p1", "p6"),
    },
    "3": {
        "headers": [
            "Параметр",
            "Предельные отклонения, мм",
            "Контроль (метод, объем, вид регистрации)",
        ],
        "param_cls": ("p28", "p35", "p1"),
        "dev_cls": ("p21",),
        "ctrl_cls": ("p30", "p14", "p31"),
    },
    "4": {
        "headers": [
            "Параметр",
            "Предельные отклонения, мм",
            "Контроль (метод, объем, вид регистрации)",
        ],
        "param_cls": ("p6",),
        "dev_cls": ("p23",),
        "ctrl_cls": ("p1",),
    },
    "6": {
        "headers": [
            "Параметр",
            "Предельн. Отклонения, мм",
            "Контроль (метод, объем, вид регистрации)",
        ],
        "param_cls": ("p1", "p24"),
        "dev_cls": ("p23",),
        "ctrl_cls": ("p1",),
    },
}


def parse_table_1(body: list[Para]) -> dict[str, Any]:
    return parse_table_class_map(body, TABLE_SPECS["1"])


def parse_table_std_sp(body: list[Para], headers: list[str]) -> dict[str, Any]:
    """Совместимость: определить spec по заголовкам или использовать table 2 defaults."""
    del headers
    return parse_table_class_map(body, TABLE_SPECS["2"])


def parse_table_5(body: list[Para]) -> dict[str, Any]:
    paras = [p for p in body if not _is_junk(p)]
    header_rows = [
        [
            "Параметр",
            "резервуаров и газгольдеров объемом, м³",
            "",
            "",
            "водонапорных башен",
            "Контроль (метод, объем, вид регистрации)",
        ],
        [
            "",
            "100–700",
            "1000–5000",
            "10000–50000 всех газгольдеров",
            "",
            "",
        ],
    ]
    rows: list[dict[str, Any]] = []

    i = 0
    while i < len(paras) and not (
        paras[i].cls == "p1" and _NUMBERED_ROW_RE.match(paras[i].text.strip())
    ):
        i += 1

    while i < len(paras):
        p = paras[i]
        t = p.text.strip()
        if p.cls == "p1" and _NUMBERED_ROW_RE.match(t):
            param = [t]
            i += 1
            while i < len(paras):
                pp = paras[i]
                pt = pp.text.strip()
                if pp.cls == "p1" and not _NUMBERED_ROW_RE.match(pt):
                    param.append(pt)
                    i += 1
                    continue
                break

            nums: list[str] = []
            while i < len(paras) and paras[i].cls == "p23":
                dt = paras[i].text.strip()
                if _NUMBERED_ROW_RE.match(dt):
                    break
                if paras[i].cls == "p1":
                    break
                if _is_deviation_token(dt) or dt in ("—", "-"):
                    nums.append(dt)
                    i += 1
                    continue
                break

            ctrl: list[str] = []
            while i < len(paras):
                pp = paras[i]
                ct = pp.text.strip()
                if pp.cls in ("p6", "p1") and _is_control(ct):
                    ctrl.append(ct)
                    i += 1
                    continue
                if pp.cls == "p23" and ct.casefold() == "то же":
                    ctrl.append(ct)
                    i += 1
                    continue
                break

            cells = ["\n".join(param).strip()]
            while len(nums) < 4:
                nums.append("")
            cells.extend(nums[:4])
            cells.append("\n".join(ctrl).strip())
            rows.append(_row(cells))
            continue
        i += 1

    return _mk_table("", header_rows=header_rows, rows=rows, layout="wide")


def parse_table_7(body: list[Para]) -> dict[str, Any]:
    paras = [p for p in body if not _is_junk(p)]
    header_rows = [
        [
            "Объем резервуара, м³",
            "Разность отметок наружного контура днища, мм",
            "",
            "",
            "",
            "Контроль (метод, объем, вид регистрации)",
        ],
        [
            "",
            "при незаполненном резервуаре",
            "",
            "при заполненном резервуаре",
            "",
            "",
        ],
        [
            "",
            "смежных точек на расстоянии 6 м по периметру",
            "любых других точек",
            "смежных точек на расстоянии 6 м по периметру",
            "любых других точек",
            "",
        ],
    ]
    volumes: list[str] = []
    values: list[str] = []
    control = ""

    for p in paras:
        if p.cls == "p6" and re.search(r"\d", p.text):
            volumes.append(p.text.strip())
        elif p.cls == "p23" and _is_deviation_token(p.text):
            values.append(p.text.strip())
        elif p.cls == "p1" and _is_control(p.text):
            control = p.text.strip()

    rows: list[dict[str, Any]] = []
    cols = 4
    for vi, vol in enumerate(volumes):
        start = vi * cols
        chunk = values[start : start + cols]
        while len(chunk) < cols:
            chunk.append("")
        rows.append(_row([vol, *chunk, control if vi == 0 else ""]))

    return _mk_table("", header_rows=header_rows, rows=rows, layout="wide")


def parse_table_8(body: list[Para]) -> list[dict[str, Any]]:
    """Матрица поясов + стандартная таблица параметров."""
    paras = [p for p in body if not _is_junk(p)]
    tables: list[dict[str, Any]] = []

    belts = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]
    header_rows = [
        [
            "Объем резервуара, м³",
            "Предельные отклонения от вертикали образующих стенки из рулонов и отдельных листов, мм",
            *([""] * 11),
            "Контроль (метод, объем, вид регистрации)",
        ],
        ["", "Номера поясов", *belts, ""],
    ]

    matrix_rows: list[dict[str, Any]] = []
    std_start = next(
        (idx for idx, p in enumerate(paras) if p.text.strip() == "Параметр" and p.cls == "p23"),
        len(paras),
    )
    matrix_paras = paras[:std_start]

    volumes: list[str] = []
    nums: list[str] = []
    matrix_ctrl = ""
    past_belts = False
    collecting_nums = False

    for p in matrix_paras:
        t = p.text.strip()
        if t in belts:
            past_belts = True
            continue
        if not past_belts:
            continue
        if not collecting_nums and p.cls in ("p1", "p6") and len(volumes) < 4 and not _is_control(t):
            if re.search(r"\d", t):
                volumes.append(t)
                continue
        if p.cls == "p23" and _is_deviation_token(t):
            collecting_nums = True
            nums.append(t)
        elif p.cls == "p6" and _is_control(t):
            matrix_ctrl = t

    belt_count = 12
    for vi, vol in enumerate(volumes[:4]):
        start = vi * belt_count
        chunk = nums[start : start + belt_count]
        while len(chunk) < belt_count:
            chunk.append("")
        matrix_rows.append(_row([vol, *chunk, matrix_ctrl if vi == 0 else ""]))

    if matrix_rows:
        tables.append(
            _mk_table(
                "Таблица №8 — отклонения от вертикали образующих (пояса)",
                header_rows=header_rows,
                rows=matrix_rows,
                layout="matrix",
            )
        )

    if std_start < len(paras):
        std = parse_table_class_map(paras[std_start:], TABLE_SPECS["2"])
        std["caption"] = "Таблица №8 — геометрические параметры"
        tables.append(std)

    return tables


def parse_table_11(body: list[Para]) -> tuple[dict[str, Any], list[Para]]:
    paras = [p for p in body if not _is_junk(p)]
    headers = ["Параметр", "Величина параметра, мм", "Контроль (метод, объем, вид регистрации)"]

    footnote_start = next(
        (
            j
            for j, p in enumerate(paras)
            if p.cls == "p1" and p.text.startswith("1. Значения допускаемых отклонений")
        ),
        len(paras),
    )

    i = 0
    while i < len(paras) and _is_header_line(paras[i].text):
        i += 1

    work = paras[i:footnote_start]
    param: list[str] = []
    vals: list[str] = []
    ctrl: list[str] = []

    j = 0
    while j < len(work) and work[j].cls == "p1":
        param.append(work[j].text.strip())
        j += 1
    while j < len(work) and work[j].cls in ("p41", "p23"):
        vals.append(work[j].text.strip())
        j += 1
    while j < len(work):
        if work[j].cls == "p39" or _is_control(work[j].text):
            ctrl.append(work[j].text.strip())
        j += 1

    rows = [
        _row(
            [
                "\n".join(param).strip(),
                "\n".join(vals).strip(),
                "\n".join(ctrl).strip(),
            ]
        )
    ]
    tail = paras[footnote_start:] if footnote_start < len(paras) else []
    table = _mk_table("Таблица 11", headers=headers, rows=rows)
    return table, tail


def parse_table_9_body(body: list[Para]) -> list[dict[str, Any]]:
    """Таблица 9 — методика и схема, не сетка значений."""
    segs: list[dict[str, Any]] = []
    lines = [p.text.strip() for p in body if not _is_junk(p)]
    if not lines:
        return segs
    buf: list[str] = []
    for line in lines:
        if line.startswith("Операционный и приемочный"):
            if buf:
                segs.append({"type": "paragraph", "text": "\n".join(buf)})
                buf = []
            segs.append({"type": "paragraph", "text": line})
        else:
            buf.append(line)
    if buf:
        segs.append({"type": "paragraph", "text": "\n".join(buf)})
    return segs


TABLE_PARSERS: dict[str, Callable[..., Any]] = {
    "1": parse_table_1,
    "5": parse_table_5,
    "7": parse_table_7,
    "9": parse_table_9_body,
    "11": parse_table_11,
}

STD_SP_TABLES = {"2", "3", "4", "6"}


def _parse_table_section(no: str, caption: str, body: list[Para]) -> list[dict[str, Any]]:
    if no == "8":
        return parse_table_8(body)

    if no == "9":
        segs = [{"type": "paragraph", "text": caption}]
        segs.extend(parse_table_9_body(body))
        return segs

    if no == "11":
        table, tail = parse_table_11(body)
        table["caption"] = caption
        segs: list[dict[str, Any]] = [table]
        for p in tail:
            if _SHAPE_JUNK_RE.search(p.text):
                continue
            segs.append({"type": "paragraph", "text": p.text})
        return segs

    if no in STD_SP_TABLES:
        table = parse_table_class_map(body, TABLE_SPECS[no])
        table["caption"] = caption
        return [table]

    parser = TABLE_PARSERS.get(no)
    if parser:
        table = parser(body)
        if isinstance(table, dict):
            table["caption"] = caption
            return [table]

    return [{"type": "paragraph", "text": caption}]


def _lines_to_segments(paras: list[Para]) -> list[dict[str, Any]]:
    segments: list[dict[str, Any]] = []
    i = 0
    n = len(paras)

    while i < n:
        t = paras[i].text.strip()
        if not t:
            i += 1
            continue

        if _TABLE_CAPTION_RE.match(t):
            no = _table_no(t) or ""
            caption = t
            i += 1
            body: list[Para] = []
            while i < n:
                nxt = paras[i].text.strip()
                if not nxt:
                    i += 1
                    continue
                if _TABLE_CAPTION_RE.match(nxt) or _ROMAN_SECTION_RE.match(nxt):
                    break
                body.append(paras[i])
                i += 1
            segments.extend(_parse_table_section(no, caption, body))
            continue

        if _ROMAN_SECTION_RE.match(t):
            segments.append({"type": "heading", "text": t})
            i += 1
            continue

        if t.startswith("- "):
            bullets = [t[2:].strip()]
            i += 1
            while i < n:
                nxt = paras[i].text.strip()
                if not nxt or not nxt.startswith("- "):
                    break
                bullets.append(nxt[2:].strip())
                i += 1
            segments.append({"type": "bullets", "items": bullets})
            continue

        para = [t]
        i += 1
        while i < n:
            nxt = paras[i].text.strip()
            if (
                not nxt
                or _TABLE_CAPTION_RE.match(nxt)
                or _ROMAN_SECTION_RE.match(nxt)
                or nxt.startswith("- ")
            ):
                break
            para.append(nxt)
            i += 1
        segments.append({"type": "paragraph", "text": "\n".join(para)})

    return segments


def extract_rich_segments(doc_path) -> list[dict[str, Any]]:
    html = doc_to_html(doc_path)
    paras = _paragraphs_with_class_from_doc_html(html)
    return _lines_to_segments(paras)


def _render_table(seg: dict[str, Any]) -> str:
    cap = escape(seg.get("caption") or "")
    header_rows = seg.get("header_rows")
    headers = seg.get("headers") or []
    thead = ""
    if header_rows:
        thead = "".join(
            "<tr>"
            + "".join(f"<th>{escape(h)}</th>" for h in row)
            + "</tr>"
            for row in header_rows
        )
    elif headers:
        thead = "<tr>" + "".join(f"<th>{escape(h)}</th>" for h in headers) + "</tr>"

    body_rows = []
    col_count = len(header_rows[0]) if header_rows else len(headers)
    for row in seg.get("rows") or []:
        kind = row.get("type") or "data"
        cells = row.get("cells") or []
        if kind == "section":
            span = col_count or len(cells) or 3
            body_rows.append(
                f'<tr class="otkk-section-row"><td colspan="{span}">{escape(cells[0] if cells else "")}</td></tr>'
            )
            continue
        tds = "".join(
            f'<td>{escape(c or "").replace(chr(10), "<br>")}</td>' for c in cells
        )
        body_rows.append(f"<tr>{tds}</tr>")

    layout = seg.get("layout") or "standard"
    cls = f"otkk-inner-table otkk-inner-table--{layout}"
    return (
        f'<div class="otkk-inner-table-wrap">'
        f'<div class="otkk-table-caption">{cap}</div>'
        f'<table class="{cls}"><thead>{thead}</thead>'
        f'<tbody>{"".join(body_rows)}</tbody></table></div>'
    )


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
        return _render_table(seg)
    return ""


def segments_to_html(segments: list[dict[str, Any]]) -> str:
    return "".join(segment_to_html(s) for s in segments)
