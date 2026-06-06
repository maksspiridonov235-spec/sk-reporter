#!/usr/bin/env python3
"""Собрать кэши данных инженера: vor.json, tk/manifest.yaml, personnel.yaml."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

import yaml
from openpyxl import load_workbook

from sk_reporter.engineer.tk_catalog import write_manifest
from sk_reporter.engineer.vor_parser import write_vor_cache
from sk_reporter.paths import personnel_dir, project_dir, projects_dir


def export_personnel() -> Path:
    xlsx = personnel_dir() / "Справочник персонала.xlsx"
    if not xlsx.is_file():
        raise FileNotFoundError(xlsx)
    wb = load_workbook(xlsx, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Пустой лист в справочнике персонала")
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    people = []
    for row in rows[1:]:
        if not any(row):
            continue
        rec = {headers[i]: (row[i] if i < len(row) else None) for i in range(len(headers))}
        if not any(rec.values()):
            continue
        people.append(rec)
    out = personnel_dir() / "personnel.yaml"
    out.write_text(yaml.safe_dump({"people": people}, allow_unicode=True, sort_keys=False), encoding="utf-8")
    return out


def build_vor_caches(project_ids: list[str] | None = None) -> list[Path]:
    built: list[Path] = []
    for proj in projects_dir().iterdir():
        if not proj.is_dir():
            continue
        if project_ids and proj.name not in project_ids:
            continue
        meta_path = proj / "project.yaml"
        vor_name = None
        if meta_path.is_file():
            meta = yaml.safe_load(meta_path.read_text(encoding="utf-8")) or {}
            vor_name = meta.get("vor_docx")
        if vor_name:
            vor_path = proj / vor_name
            if vor_path.is_file():
                built.append(write_vor_cache(vor_path))
    return built


def main() -> None:
    parser = argparse.ArgumentParser(description="Build engineer data caches")
    parser.add_argument("--personnel", action="store_true", help="Export personnel.yaml")
    parser.add_argument("--tk", action="store_true", help="Write data/tk/manifest.yaml")
    parser.add_argument("--vor", action="store_true", help="Parse VOR docx → vor.json")
    parser.add_argument("--all", action="store_true", help="All of the above")
    parser.add_argument("--project", action="append", dest="projects", help="Project id filter")
    args = parser.parse_args()

    if args.all or not any((args.personnel, args.tk, args.vor)):
        args.personnel = args.tk = args.vor = True

    results: dict[str, str] = {}
    if args.personnel:
        results["personnel"] = str(export_personnel())
    if args.tk:
        results["tk_manifest"] = str(write_manifest())
    if args.vor:
        for p in build_vor_caches(args.projects):
            results.setdefault("vor", []).append(str(p))  # type: ignore[union-attr]

    print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
