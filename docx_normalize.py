"""
Нормализация .docx перед макросами: единый «диалект» Word и чистая сетка таблицы.

Не заменяет ручную правку в Word, но убирает типичные расхождения между версиями:
- старые tcW (pct/auto), конфликтующие с фиксированной сеткой;
- лишние gridBefore на строках;
- разные метаданные приложения в docProps/app.xml.
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

# Как у типичного Word 2016+ (совпадает с большинством ваших отчётов)
APP_XML_APPLICATION = "Microsoft Office Word"
APP_XML_VERSION = "16.0000"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
COMPAT_URI = "http://schemas.microsoft.com/office/word"
COMPAT_MODE_VAL = "15"  # Word 2013+


def _patch_app_xml(app_xml: bytes) -> bytes:
    text = app_xml.decode("utf-8")

    def _set_tag(tag: str, value: str, src: str) -> str:
        pat = rf"<{tag}>[^<]*</{tag}>"
        repl = f"<{tag}>{value}</{tag}>"
        if re.search(pat, src):
            return re.sub(pat, repl, src, count=1)
        # вставить перед </Properties>
        return src.replace("</Properties>", f"  <{tag}>{value}</{tag}>\n</Properties>", 1)

    if "<Properties" in text:
        text = _set_tag("Application", APP_XML_APPLICATION, text)
        text = _set_tag("AppVersion", APP_XML_VERSION, text)
    return text.encode("utf-8")


def _patch_settings_compat(settings_xml: bytes) -> bytes:
    root = etree.fromstring(settings_xml)
    ns = {"w": W_NS}
    compat = root.find("w:compat", ns)
    if compat is None:
        compat = etree.SubElement(root, qn("w:compat"))

    setting_name = f"{{{W_NS}}}compatSetting"
    found = False
    for el in compat.findall(setting_name):
        if el.get("name") == "compatibilityMode" and el.get("uri") == COMPAT_URI:
            el.set("val", COMPAT_MODE_VAL)
            found = True
            break
    if not found:
        el = etree.SubElement(compat, qn("w:compatSetting"))
        el.set("name", "compatibilityMode")
        el.set("uri", COMPAT_URI)
        el.set("val", COMPAT_MODE_VAL)

    return etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=True,
    )


def normalize_docx_package(filepath: str | Path) -> list[str]:
    """
    Правит ZIP-пакет docx: app.xml + режим совместимости в settings.xml.
    Возвращает список коротких меток, что изменено.
    """
    path = Path(filepath)
    notes: list[str] = []
    tmp = path.with_suffix(path.suffix + ".norm.tmp")

    with zipfile.ZipFile(path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "docProps/app.xml":
                    data = _patch_app_xml(data)
                    notes.append("app")
                elif item.filename == "word/settings.xml":
                    try:
                        data = _patch_settings_compat(data)
                        notes.append("compat")
                    except etree.XMLSyntaxError:
                        pass
                zout.writestr(item, data)

    os.replace(tmp, path)
    return notes


def _row_occupied_cols(tr, ncol: int) -> int:
    """Сколько колонок сетки занимает строка (с учётом gridBefore, gridSpan, vMerge)."""
    col_idx = 0
    trPr = tr.find(qn("w:trPr"))
    if trPr is not None:
        gb = trPr.find(qn("w:gridBefore"))
        if gb is not None:
            col_idx = int(gb.get(qn("w:val"), 0))

    for tc in tr.findall(qn("w:tc")):
        if col_idx >= ncol:
            break
        tcPr = tc.find(qn("w:tcPr"))
        span = 1
        if tcPr is not None:
            gs = tcPr.find(qn("w:gridSpan"))
            if gs is not None:
                span = max(1, int(gs.get(qn("w:val"), 1)))
            vm = tcPr.find(qn("w:vMerge"))
            if vm is not None and vm.get(qn("w:val")) != "restart":
                col_idx += span
                continue
        col_idx += span
    return col_idx


def sanitize_document_tables(doc: Document, expected_cols: int | None = None) -> list[str]:
    """
    Снимает конфликтующие свойства таблиц перед apply_layout.
    Возвращает предупреждения (например, строка не на 6 колонок).
    """
    warnings: list[str] = []

    for ti, table in enumerate(doc.tables):
        tbl = table._tbl
        ncol = expected_cols
        if ncol is None:
            tbl_grid = tbl.find(qn("w:tblGrid"))
            if tbl_grid is not None:
                ncol = len(tbl_grid.findall(qn("w:gridCol")))
            else:
                ncol = 6

        for ri, row in enumerate(table.rows):
            tr = row._tr
            occupied = _row_occupied_cols(tr, ncol)
            if occupied != ncol:
                warnings.append(
                    f"табл.{ti + 1} строка {ri + 1}: {occupied} кол. (ожид. {ncol})"
                )

            trPr = tr.find(qn("w:trPr"))
            if trPr is not None:
                gb = trPr.find(qn("w:gridBefore"))
                if gb is not None:
                    # Частый артефакт после копирования между версиями Word
                    trPr.remove(gb)

            for tc in tr.findall(qn("w:tc")):
                tcPr = tc.find(qn("w:tcPr"))
                if tcPr is None:
                    continue
                tcW = tcPr.find(qn("w:tcW"))
                if tcW is not None:
                    tcPr.remove(tcW)

        tbl_layout = tbl.find(f".//{qn('w:tblLayout')}")
        if tbl_layout is None:
            tbl_pr = tbl.find(qn("w:tblPr"))
            if tbl_pr is None:
                tbl_pr = etree.SubElement(tbl, qn("w:tblPr"))
                tbl.insert(0, tbl_pr)
            tbl_layout = etree.SubElement(tbl_pr, qn("w:tblLayout"))
        tbl_layout.set(qn("w:type"), "fixed")

    return warnings


def try_libreoffice_rewrite(filepath: str | Path) -> bool:
    """
    Пересохраняет файл через LibreOffice (если установлен).
    Часто выравнивает «старые» docx под один формат.
    """
    path = Path(filepath).resolve()
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False

    with tempfile.TemporaryDirectory() as tmp:
        out_dir = Path(tmp)
        cmd = [
            soffice,
            "--headless",
            "--norestore",
            "--convert-to",
            "docx",
            "--outdir",
            str(out_dir),
            str(path),
        ]
        try:
            subprocess.run(cmd, check=True, capture_output=True, timeout=120)
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
            return False

        converted = out_dir / path.name
        if not converted.is_file():
            candidates = list(out_dir.glob("*.docx"))
            if not candidates:
                return False
            converted = candidates[0]
        shutil.copy2(converted, path)
    return True


def normalize_docx_for_layout(
    filepath: str | Path,
    expected_cols: int | None = None,
    use_libreoffice: bool = True,
) -> tuple[bool, str]:
    """
    Полная нормализация одного файла перед prepare/layout.
    Возвращает (успех, краткое сообщение для лога).
    """
    path = Path(filepath)
    parts: list[str] = []

    if use_libreoffice and try_libreoffice_rewrite(path):
        parts.append("LO")

    try:
        pkg = normalize_docx_package(path)
        if pkg:
            parts.append("+".join(pkg))
    except Exception as e:
        return False, f"пакет: {e}"

    try:
        doc = Document(os.fspath(path))
    except Exception as e:
        return False, f"открытие: {e}"

    warns = sanitize_document_tables(doc, expected_cols)
    try:
        doc.save(os.fspath(path))
    except Exception as e:
        return False, f"сохранение: {e}"

    if warns:
        parts.append(f"⚠ {warns[0]}" + (f" (+{len(warns)-1})" if len(warns) > 1 else ""))
    if not parts:
        parts.append("ок")
    return True, "норм: " + ", ".join(parts)
