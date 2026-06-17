"""Утилиты для работы с ОТКК: конвертация .doc/.docx в HTML и плоский текст из JSON."""

from __future__ import annotations

import re
import shutil
import subprocess
import sys
import tempfile
from html import unescape
from pathlib import Path
from typing import Any

from sk_reporter.otkk_text import strip_kodeks_fields

_HYPERLINK_RE = re.compile(r'HYPERLINK\s+"[^"]*"\s*\\o\s*"[^"]*"', re.I)
_WS_RE = re.compile(r"\s+")


def _clean_text(raw: str) -> str:
    text = unescape(raw or "")
    text = text.replace("\xa0", " ")
    text = strip_kodeks_fields(text)
    text = _HYPERLINK_RE.sub("", text)
    text = text.replace("\x07", " ")
    return _WS_RE.sub(" ", text).strip()


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
