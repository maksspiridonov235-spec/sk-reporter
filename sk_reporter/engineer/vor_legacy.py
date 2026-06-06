"""Парсинг legacy .doc ВОР (таблица через textutil / LibreOffice → текст)."""

from __future__ import annotations

import re
import subprocess
import shutil
import sys
from pathlib import Path

from sk_reporter.engineer.vor_parser import (
    VorDocument,
    VorObject,
    VorStage,
    VorWorkItem,
    _STAGE_RE,
    vor_to_dict,
)

_UNIT_RE = re.compile(
    r"^(?:"
    r"шт|м|м²|м2|м³|м3|мп|т|кг|км|компл|л|"
    r"тыс\.?\s*шт|100\s*м|1000\s*м3|"
    r"шт/т|ям/м|кг/|т/|м\.?п\.?|п\.?\s*м\.?"
    r")(?:/\S+)?$",
    re.I,
)
_QTY_RE = re.compile(r"^[\d][\d\s,/.\-]*$")
_FOOTER_MARKERS = frozenset(
    {"Изм.", "Разраб.", "Пров.", "ГИП", "Н. контр", "Лист", "Подп.", "Дата", "Кол.уч.", "№док."}
)
_HEADER_TOKENS = frozenset(
    {"№ строки", "Наименование вида работ", "Ед.", "изм.", "Код", "Коли-чество", "вида работ", "ед. изм."}
)


def _extract_doc_plain(path: Path) -> str:
    if sys.platform == "darwin":
        textutil = shutil.which("textutil")
        if textutil:
            proc = subprocess.run(
                [textutil, "-convert", "txt", "-stdout", str(path)],
                capture_output=True,
                text=True,
                errors="replace",
            )
            if proc.returncode == 0 and proc.stdout.strip():
                return proc.stdout
    from sk_reporter.engineer.doc_text import extract_doc_text

    return extract_doc_text(path)


def _is_unit(line: str) -> bool:
    return bool(_UNIT_RE.match(line.strip()))


def _is_quantity(line: str) -> bool:
    s = line.strip()
    if not s or s == "-":
        return False
    return bool(_QTY_RE.match(s))


def _stage_title_from_text(text: str, fallback: str) -> str:
    for line in text.splitlines():
        line = line.strip()
        if _STAGE_RE.match(line):
            return line
    return fallback


def _content_lines(text: str) -> list[str]:
    lines = [ln.strip() for ln in text.splitlines()]
    start = 0
    for idx, line in enumerate(lines):
        if line == "Наименование вида работ":
            start = idx + 1
            break
    while start < len(lines) and lines[start] in _HEADER_TOKENS:
        start += 1
    out: list[str] = []
    for line in lines[start:]:
        if line in _FOOTER_MARKERS:
            break
        if "MERGEFORMAT" in line or line.startswith("PAGE "):
            continue
        out.append(line)
    return out


def _next_nonempty(lines: list[str], i: int) -> int:
    j = i
    while j < len(lines) and not lines[j]:
        j += 1
    return j


def _find_unit_index(lines: list[str], start: int, max_ahead: int = 8) -> int | None:
    j = start
    limit = min(len(lines), start + max_ahead)
    while j < limit:
        if lines[j] and _is_unit(lines[j]):
            return j
        j += 1
    return None


def parse_vor_legacy_doc(path: Path, stage_title: str = "") -> VorStage:
    text = _extract_doc_plain(path)
    title = stage_title or _stage_title_from_text(text, path.stem)
    lines = _content_lines(text)
    stage = VorStage(title=title)
    current_object: VorObject | None = None
    i = 0

    while i < len(lines):
        line = lines[i]
        if not line or _STAGE_RE.match(line):
            i += 1
            continue

        unit_idx = _find_unit_index(lines, i + 1, max_ahead=3)
        if unit_idx is not None:
            name_parts = [lines[k] for k in range(i, unit_idx) if lines[k]]
            name = " ".join(name_parts).strip()
            unit = lines[unit_idx]
            qty = ""
            k = unit_idx + 1
            while k < len(lines) and k <= unit_idx + 6:
                if lines[k] and _is_quantity(lines[k]):
                    qty = lines[k]
                    break
                k += 1
            item = VorWorkItem(name=name, unit=unit, quantity=qty)
            if current_object is not None:
                current_object.works.append(item)
            else:
                stage.works.append(item)
            i = k + 1 if qty else unit_idx + 1
            continue

        unit_idx = _find_unit_index(lines, i + 1, max_ahead=12)
        if unit_idx is not None and unit_idx > i + 1:
            obj = VorObject(title=line)
            stage.objects.append(obj)
            current_object = obj
            i += 1
            continue

        i += 1

    return stage


def parse_vor_file(path: Path, project_id: str = "") -> VorDocument:
    pid = project_id or path.parent.name
    if path.suffix.lower() == ".docx":
        from sk_reporter.engineer.vor_parser import parse_vor_docx

        return parse_vor_docx(path, pid)
    stage = parse_vor_legacy_doc(path)
    return VorDocument(project_id=pid, source=path.name, stages=[stage])


def build_merged_vor(project_dir: Path, meta: dict) -> VorDocument:
    pid = meta.get("id") or project_dir.name
    sources: list[str] = []
    stages: list[VorStage] = []

    vor_docx = meta.get("vor_docx")
    if vor_docx:
        path = project_dir / vor_docx
        if path.is_file():
            doc = parse_vor_file(path, pid)
            sources.append(path.name)
            stages.extend(doc.stages)

    for name in meta.get("vor_doc") or []:
        path = project_dir / name
        if not path.is_file():
            continue
        stage = parse_vor_legacy_doc(path)
        sources.append(path.name)
        stages.append(stage)

    if not stages:
        raise FileNotFoundError(f"Нет файлов ВОР в {project_dir}")

    return VorDocument(
        project_id=pid,
        source=", ".join(sources),
        stages=stages,
    )


def write_project_vor_cache(project_dir: Path, cache_name: str = "vor.json") -> Path:
    import json
    import yaml

    meta_path = project_dir / "project.yaml"
    meta: dict = {"id": project_dir.name}
    if meta_path.is_file():
        meta.update(yaml.safe_load(meta_path.read_text(encoding="utf-8")) or {})

    vor = build_merged_vor(project_dir, meta)
    out = project_dir / (meta.get("vor_cache") or cache_name)
    out.write_text(json.dumps(vor_to_dict(vor), ensure_ascii=False, indent=2), encoding="utf-8")
    return out
