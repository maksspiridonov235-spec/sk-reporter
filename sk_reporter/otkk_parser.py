"""Парсинг ОТКК (.doc/.docx) в структуру как в исходной карте."""

from __future__ import annotations

import re
import shutil
import subprocess
import sys
import tempfile
from html import unescape
from pathlib import Path
from typing import Any

from lxml import html as lxml_html

from sk_reporter.engineer.doc_text import extract_doc_text
from sk_reporter.otkk_text import normative_visible_text, sanitize_otkk_rows, strip_kodeks_fields

_OTKK_CODE_RE = re.compile(r"ОТКК\s*[-–—]?\s*(\d+)", re.I)
_NORM_CODE_RE = re.compile(
    r"(?:СП|ГОСТ|ВСН|СНиП)\s*[\d][\d\.\-]*(?:\s*[\d\-]*)?",
    re.I,
)
_HYPERLINK_RE = re.compile(r'HYPERLINK\s+"[^"]*"\s*\\o\s*"[^"]*"', re.I)
_WS_RE = re.compile(r"\s+")


def otkk_id_from_path(path: Path) -> str | None:
    m = _OTKK_CODE_RE.search(path.stem)
    return f"otkk-{int(m.group(1))}" if m else None


def _clean_text(raw: str) -> str:
    text = unescape(raw or "")
    text = text.replace("\xa0", " ")
    text = strip_kodeks_fields(text)
    text = _HYPERLINK_RE.sub("", text)
    text = text.replace("\x07", " ")
    return _WS_RE.sub(" ", text).strip()


def _cell_text(td) -> str:
    parts = td.xpath(".//text()")
    return _clean_text(" ".join(parts))


def _cell_blocks(td) -> dict[str, Any]:
    paragraphs: list[str] = []
    bullets: list[str] = []
    for p in td.xpath(".//p"):
        t = _clean_text(" ".join(p.xpath(".//text()")))
        if not t:
            continue
        if t.startswith("-") or t.startswith("–"):
            bullets.append(t.lstrip("-– ").strip())
        else:
            paragraphs.append(t)
    text = _cell_text(td)
    if not paragraphs and not bullets and text:
        paragraphs = [text]
    return {"text": text, "paragraphs": paragraphs, "bullets": bullets}


def _normative_codes(text: str) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for m in _NORM_CODE_RE.finditer(text):
        code = _clean_text(m.group(0))
        key = code.casefold()
        if key not in seen:
            seen.add(key)
            out.append(code)
    return out


def doc_to_html(path: Path) -> str:
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(path)

    if path.suffix.lower() == ".docx":
        import mammoth

        with open(path, "rb") as f:
            result = mammoth.convert_to_html(f)
        html = result.value
    elif sys.platform == "darwin":
        textutil = shutil.which("textutil")
        if not textutil:
            raise RuntimeError("textutil не найден (macOS)")
        proc = subprocess.run(
            [textutil, "-convert", "html", "-stdout", str(path)],
            capture_output=True,
            text=True,
            errors="replace",
        )
        if proc.returncode != 0 or not proc.stdout.strip():
            raise RuntimeError(f"textutil не смог прочитать {path.name}")
        html = proc.stdout
    else:
        html = ""
        for binary in ("soffice", "libreoffice"):
            exe = shutil.which(binary)
            if not exe:
                continue
            with tempfile.TemporaryDirectory() as tmp:
                out_dir = Path(tmp)
                proc = subprocess.run(
                    [exe, "--headless", "--convert-to", "html", "--outdir", str(out_dir), str(path)],
                    capture_output=True,
                    text=True,
                    errors="replace",
                )
                if proc.returncode != 0:
                    continue
                html_files = list(out_dir.glob("*.html"))
                if html_files:
                    html = html_files[0].read_text(encoding="utf-8", errors="replace")
                    break
        if not html:
            raise RuntimeError(
                f"Не удалось конвертировать {path.name} в HTML: нужен LibreOffice (soffice) на сервере"
            )

    return html.replace("\x00", "")


def parse_otkk_document(path: Path) -> dict[str, Any]:
    """Структура карты: шапка, строки таблицы label/value, подпись."""
    path = Path(path)
    card_id = otkk_id_from_path(path)
    html_body = doc_to_html(path)
    root = lxml_html.fromstring(html_body)

    rows: list[dict[str, Any]] = []
    code = ""
    title = ""
    normative_text = ""
    signature: dict[str, str] | None = None

    norm_parts: list[str] = []
    capture_norm = False
    for p in root.xpath("//p"):
        t = _clean_text(" ".join(p.xpath(".//text()")))
        if not t:
            continue
        if t.startswith("Нормативные документы"):
            capture_norm = True
            norm_parts.append(t)
            continue
        if capture_norm and root.xpath("//table") and rows:
            capture_norm = False
        if capture_norm:
            norm_parts.append(t)
        if t.startswith("Разработал"):
            signature = {"label": "Разработал", "text": t}
        elif signature is not None and signature.get("text", "").rstrip(":") == "Разработал":
            signature["text"] = t

    if norm_parts:
        normative_text = _clean_text(" ".join(norm_parts))

    for table in root.xpath("//table"):
        for tr in table.xpath(".//tr"):
            tds = tr.xpath("./td")
            if len(tds) >= 2:
                label = _cell_text(tds[0])
                value = _cell_text(tds[1])
                if not label and not value:
                    continue
                if _OTKK_CODE_RE.search(label):
                    code = label
                    title = value
                row: dict[str, Any] = {"label": label, "value": value}
                if "Контролируемые параметры" in label:
                    row["body"] = _cell_blocks(tds[1])
                rows.append(row)
            elif len(tds) == 1:
                solo = _cell_text(tds[0])
                if solo and rows and not rows[-1].get("value"):
                    rows[-1]["value"] = solo

    if normative_text:
        norm_row = {
            "label": "Нормативные документы",
            "value": normative_visible_text(normative_text),
            "codes": _normative_codes(normative_text),
        }
        insert_at = 0
        for i, r in enumerate(rows):
            if r.get("label") == "Область применения":
                insert_at = i + 1
                break
        rows.insert(insert_at, norm_row)

    if not title:
        for r in rows:
            if r.get("label") and _OTKK_CODE_RE.search(r["label"]):
                code = r["label"]
                title = r.get("value", "")
                break

    if not signature:
        after_sig = False
        for p in root.xpath("//p"):
            t = _clean_text(" ".join(p.xpath(".//text()")))
            if not t or ("PAGE" in t and "NUMPAGES" in t):
                continue
            if t.startswith("Разработал"):
                signature = {"label": "Разработал", "text": t}
                after_sig = True
                continue
            if after_sig and t:
                signature["text"] = t
                break

    plain_text = extract_doc_text(path)

    return {
        "id": card_id,
        "code": code,
        "title": title,
        "file": path.name,
        "rows": sanitize_otkk_rows(rows),
        "signature": signature,
        "plain_text": strip_kodeks_fields(plain_text),
    }


def content_to_plain_text(content: dict[str, Any]) -> str:
    """Плоский текст из структуры (для сниппетов в отчёте инженера)."""
    parts: list[str] = []
    if content.get("code"):
        parts.append(str(content["code"]))
    if content.get("title"):
        parts.append(str(content["title"]))
    for row in content.get("rows") or []:
        label = row.get("label") or ""
        value = row.get("value") or ""
        if label:
            parts.append(f"{label}: {value}")
        body = row.get("body") or {}
        for p in body.get("paragraphs") or []:
            parts.append(p)
        for b in body.get("bullets") or []:
            parts.append(f"- {b}")
    sig = content.get("signature") or {}
    if sig.get("text"):
        parts.append(sig["text"])
    return "\n".join(p for p in parts if p)
