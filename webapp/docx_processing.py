"""
Логика обработки .docx файлов.
Все функции — аналоги VBA-макросов из Word, переписанные на python-docx.
"""

import os
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from copy import deepcopy
from datetime import datetime, timedelta
from typing import Literal

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from lxml import etree


# ── Список компаний: единственное место для редактирования ──────────────────

COMPANIES = [
    ("Евракор",            ["евракор", "еврокор"]),
    ("Лесные технологии",  ["лесн. технологии", "лесные технологии", "лестех"]),
    ("ЮНС",                ["нткс", "юнс", "югранефтестрой", "югранефтестой"]),
    ("НГСК",               ["нгск", "ткс", "новая газовая строительная компания"]),
    ("Сибитек",            ["сибитек", "ооосибитек", "ооо сибитек"]),
    ("ЭСМ",                ["эсм", "оооэсм", "ооо эсм", "энергостроймонтаж",
                            "оооэнергостроймонтаж", "ооо энергостроймонтаж",
                            "энергостоймонтаж", "оооэнергостоймонтаж"]),
    ("НГП",                ["нгп", "нефтегазпроект"]),
    ("РОСЭКСПО",           ["росэкспо"]),
    ("ТПС",                ["тпс", "трубопроводсервис тпс"]),
    ("ТЭКПРО",             ["тэкпро", "трубопроводсервис тэкпро"]),
    ("ЮПИ",                ["юпи", "югорский проектный институт"]),
    ("УГГ",                ["угг", "уралгеогрупп"]),
    ("ЮГС",                ["югс", "юграгидрострой"]),
    ("ТВС",                ["твс", "тюменьвторсырье"]),
    ("НСС",                ["нсс", "нефтеспецстрой", "ооонефтеспецстрой",
                            "ооо нефтеспецстрой", "ооонсс", "ооо нсс"]),
    ("ОТ и ТБ",            ["от и тб", "отитб", "оттб", "отитб", "ОТИТБ", "ОТТБ"]),
    ("Стройфинансгрупп",   ["стройфинансгрупп", "ооостройфинансгрупп", "ооо стройфинансгрупп", "сфг"]),
]

# Шаблонное имя файла-болванки. Дата в имени определяется автоматически.
_DATE_RE = re.compile(r"\d{2}\.\d{2}\.\d{4}")


def _find_template_date(filename: str) -> str | None:
    """Возвращает дату из имени файла вида 'XX.XX.XXXX'."""
    m = _DATE_RE.search(filename)
    return m.group() if m else None


# ── Макрос 1: HighlightSecondRow_No5991 ────────────────────────────────────

def highlight_second_row(doc: Document) -> int:
    """
    Для каждой таблицы документа: если строк >= 2,
    красит 2-ю строку в голубой (#BDD6EE), жирный, по центру.
    Возвращает количество обработанных таблиц.
    """
    BLUE = RGBColor(0xBD, 0xD6, 0xEE)
    processed = 0

    for tbl in doc.tables:
        if len(tbl.rows) < 2:
            continue
        row = tbl.rows[1]
        for cell in row.cells:
            text = cell.text.strip()
            if not text:
                continue
            # Фон
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            shd = tcPr.find(qn("w:shd"))
            if shd is None:
                shd = etree.SubElement(tcPr, qn("w:shd"))
            color_hex = f"{BLUE.red:02X}{BLUE.green:02X}{BLUE.blue:02X}"
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), color_hex)

            # Вертикальное выравнивание
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            # Текст: жирный, по центру
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in para.runs:
                    run.bold = True

        processed += 1

    return processed


# ── Макрос 2: NewMacros (форматирование документа) ─────────────────────────

def format_document(doc: Document) -> None:
    """
    - Шрифт Times New Roman 10pt для всего текста
    - Отступы и интервалы = 0, одинарный межстрочный
    - Ширина всех таблиц = 18.33 см, выравнивание по центру
    - Все инлайн-картинки → 5.33 × 4 см
    """
    CM_TO_EMU = 914400 / 100  # 1 cm = 914400 / 100 EMU? нет, 1 cm = 360000 EMU
    # python-docx Cm() сам переводит
    TABLE_WIDTH = Cm(18.33)
    IMG_W = Cm(5.33)
    IMG_H = Cm(4.0)

    # Шрифт и интервалы для всех параграфов
    for para in doc.paragraphs:
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.line_spacing = 1.0
        for run in para.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(10)

    # То же внутри таблиц
    for tbl in doc.tables:
        # Ширина таблицы
        tbl.width = TABLE_WIDTH
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tblPr = tbl._tbl.find(qn("w:tblPr"))
        if tblPr is not None:
            tblW = tblPr.find(qn("w:tblW"))
            if tblW is None:
                tblW = etree.SubElement(tblPr, qn("w:tblW"))
            tblW.set(qn("w:w"), str(int(TABLE_WIDTH.pt * 20)))
            tblW.set(qn("w:type"), "dxa")

        for row in tbl.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.paragraph_format.space_before = Pt(0)
                    para.paragraph_format.space_after = Pt(0)
                    para.paragraph_format.line_spacing = 1.0
                    for run in para.runs:
                        run.font.name = "Times New Roman"
                        run.font.size = Pt(10)

    # Картинки
    for shape in doc.inline_shapes:
        shape.width = IMG_W
        shape.height = IMG_H


# ── Макросы 3 и 4: ReplaceDateInReportLine / ReplaceDateInReportLine2 ──────

def replace_date_in_report_line(doc: Document, mode: Literal["today", "yesterday"]) -> bool:
    """
    Находит абзац с текстом 'Отчёт строительного контроля по',
    внутри него заменяет любую дату вида DD.MM.YYYY на сегодняшнюю или вчерашнюю.
    Возвращает True если замена выполнена.
    """
    target_date = (
        datetime.now().strftime("%d.%m.%Y")
        if mode == "today"
        else (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")
    )
    MARKER = "отчёт строительного контроля по"

    for para in doc.paragraphs:
        if MARKER in para.text.lower():
            for run in para.runs:
                new_text = _DATE_RE.sub(target_date, run.text)
                if new_text != run.text:
                    run.text = new_text
            return True

    # Ищем и внутри таблиц (на случай если строка в таблице)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if MARKER in para.text.lower():
                        for run in para.runs:
                            new_text = _DATE_RE.sub(target_date, run.text)
                            if new_text != run.text:
                                run.text = new_text
                        return True
    return False


# ── Объединение отчётов ─────────────────────────────────────────────────────

def _copy_media_from_docx(src_path: str, dst_path: str, remap: dict) -> None:
    """
    Копирует медиафайлы (картинки) из src_path в dst_path (оба — .docx ZIP-архивы).
    Переименовывает их чтобы не было коллизий, заполняет remap {старое_имя: новое_имя}.
    """
    with zipfile.ZipFile(src_path, "r") as src_zip:
        src_media = [n for n in src_zip.namelist() if n.startswith("word/media/")]
        with zipfile.ZipFile(dst_path, "a") as dst_zip:
            existing = set(dst_zip.namelist())
            for src_name in src_media:
                base = os.path.basename(src_name)           # image1.png
                stem, ext = os.path.splitext(base)          # image1, .png
                # Генерируем уникальное имя если уже есть
                candidate = f"word/media/{base}"
                counter = 1
                while candidate in existing:
                    candidate = f"word/media/{stem}_{counter}{ext}"
                    counter += 1
                data = src_zip.read(src_name)
                dst_zip.writestr(candidate, data)
                existing.add(candidate)
                remap[base] = os.path.basename(candidate)


def _fix_image_refs_in_element(element, remap: dict) -> None:
    """
    Обновляет все ссылки на картинки в XML-элементе согласно remap.
    Картинки в docx ссылаются через r:embed в <a:blip> и <v:imagedata>.
    Remap применяется к именам файлов в relationships, но мы патчим
    уже resolved-имена прямо в XML через rId → имя файла.
    Этот патч применяется на уровне имён файлов в rels.
    """
    pass  # Rels патчим отдельно в merge_reports


def merge_reports(template_path: str, report_paths: list[str], output_path: str) -> int:
    """
    Объединяет отчёты из report_paths в шаблон template_path,
    сохраняет результат в output_path.
    Корректно переносит картинки через ZIP чтобы они не были битыми.
    Возвращает количество вставленных файлов.
    """
    shutil.copy2(template_path, output_path)
    inserted = 0

    for path in sorted(report_paths):
        try:
            src_doc = Document(path)
        except Exception:
            continue

        # 1. Копируем медиафайлы из src в output, получаем карту переименований
        remap: dict[str, str] = {}
        _copy_media_from_docx(path, output_path, remap)

        # 2. Читаем relationships из источника чтобы построить rId → имя файла
        src_rels: dict[str, str] = {}
        with zipfile.ZipFile(path, "r") as zf:
            rels_name = "word/_rels/document.xml.rels"
            if rels_name in zf.namelist():
                rels_xml = zf.read(rels_name).decode("utf-8")
                root = ET.fromstring(rels_xml)
                ns = "http://schemas.openxmlformats.org/package/2006/relationships"
                for rel in root.findall(f"{{{ns}}}Relationship"):
                    rid = rel.get("Id", "")
                    target = rel.get("Target", "")
                    # target вида "../media/image1.png" или "media/image1.png"
                    basename = os.path.basename(target)
                    src_rels[rid] = basename

        # 3. Добавляем relationships в output для новых медиафайлов
        #    и строим карту старых rId → новых rId
        rid_remap: dict[str, str] = {}
        if remap:
            with zipfile.ZipFile(output_path, "r") as zf:
                rels_name = "word/_rels/document.xml.rels"
                ET.register_namespace("", "http://schemas.openxmlformats.org/package/2006/relationships")
                rels_xml = zf.read(rels_name).decode("utf-8")
                rels_root = ET.fromstring(rels_xml)

            ns = "http://schemas.openxmlformats.org/package/2006/relationships"
            img_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

            # Найдём максимальный rId в output
            existing_ids = set()
            for rel in rels_root.findall(f"{{{ns}}}Relationship"):
                existing_ids.add(rel.get("Id", ""))
            rid_counter = 1
            while f"rId{rid_counter}" in existing_ids:
                rid_counter += 1

            # Для каждого rId из источника с картинкой — добавляем новый rel в output
            for old_rid, old_basename in src_rels.items():
                if old_basename not in remap:
                    continue
                new_basename = remap[old_basename]
                new_rid = f"rId{rid_counter}"
                while new_rid in existing_ids:
                    rid_counter += 1
                    new_rid = f"rId{rid_counter}"
                rid_remap[old_rid] = new_rid
                existing_ids.add(new_rid)
                rid_counter += 1
                new_rel = ET.SubElement(rels_root, f"{{{ns}}}Relationship")
                new_rel.set("Id", new_rid)
                new_rel.set("Type", img_type)
                new_rel.set("Target", f"media/{new_basename}")

            # Записываем обновлённый rels обратно в ZIP
            new_rels_xml = ET.tostring(rels_root, encoding="unicode", xml_declaration=False)
            new_rels_xml = '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n' + new_rels_xml
            _zip_replace(output_path, rels_name, new_rels_xml.encode("utf-8"))

        # 4. Копируем тело документа, патча rId в XML-элементах
        master = Document(output_path)
        master.add_page_break()

        for element in src_doc.element.body:
            tag = element.tag.split("}")[-1] if "}" in element.tag else element.tag
            if tag == "sectPr":
                continue
            el_copy = deepcopy(element)
            # Патчим все вхождения старых rId на новые
            if rid_remap:
                _patch_rids(el_copy, rid_remap)
            master.element.body.append(el_copy)

        master.save(output_path)
        inserted += 1

    return inserted


def _zip_replace(zip_path: str, inner_name: str, new_data: bytes) -> None:
    """Заменяет файл внутри ZIP-архива."""
    import tempfile
    tmp = zip_path + ".tmp"
    with zipfile.ZipFile(zip_path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == inner_name:
                    zout.writestr(item, new_data)
                else:
                    zout.writestr(item, zin.read(item.filename))
    os.replace(tmp, zip_path)


def _patch_rids(element, rid_remap: dict[str, str]) -> None:
    """Рекурсивно заменяет атрибуты r:embed, r:id, r:link на новые rId."""
    REMAP_ATTRS = (
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed",
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link",
    )
    for attr in REMAP_ATTRS:
        val = element.get(attr)
        if val and val in rid_remap:
            element.set(attr, rid_remap[val])
    for child in element:
        _patch_rids(child, rid_remap)


# ── Переименование файлов ───────────────────────────────────────────────────

def rename_files(folder: str, mode: Literal["today", "yesterday"]) -> list[str]:
    """
    Переименовывает .docx/.doc файлы в папке:
    заменяет все даты вида DD.MM.YYYY в имени файла на сегодняшнюю или вчерашнюю.
    Возвращает список строк лога.
    """
    new_date = (
        datetime.now().strftime("%d.%m.%Y")
        if mode == "today"
        else (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")
    )
    log = []
    for filename in os.listdir(folder):
        if not any(filename.lower().endswith(ext) for ext in (".docx", ".doc")):
            continue
        new_name = _DATE_RE.sub(new_date, filename)
        if new_name != filename:
            try:
                os.rename(
                    os.path.join(folder, filename),
                    os.path.join(folder, new_name),
                )
                log.append(f"Переименован: {filename} → {new_name}")
            except Exception as e:
                log.append(f"Ошибка: {filename} — {e}")
    return log


# ── Применение макроса к файлу ──────────────────────────────────────────────

def apply_macro_to_file(filepath: str, macro_name: str) -> tuple[bool, str]:
    """
    Применяет один из макросов к файлу, сохраняет его.
    Возвращает (успех, сообщение).
    """
    try:
        doc = Document(filepath)
    except Exception as e:
        return False, str(e)

    if macro_name == "HighlightSecondRow_No5991":
        n = highlight_second_row(doc)
        msg = f"Обработано таблиц: {n}"
    elif macro_name == "NewMacros":
        format_document(doc)
        msg = "Форматирование применено"
    elif macro_name == "ReplaceDateInReportLine":
        ok = replace_date_in_report_line(doc, "today")
        msg = "Дата заменена на сегодняшнюю" if ok else "Строка с датой не найдена"
    elif macro_name == "ReplaceDateInReportLine2":
        ok = replace_date_in_report_line(doc, "yesterday")
        msg = "Дата заменена на вчерашнюю" if ok else "Строка с датой не найдена"
    else:
        return False, f"Неизвестный макрос: {macro_name}"

    doc.save(filepath)
    return True, msg
