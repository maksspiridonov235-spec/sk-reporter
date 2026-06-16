#!/usr/bin/env python3
"""Собрать кэши данных инженера: vor.json, tk/manifest.yaml; персонал → PostgreSQL."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

import yaml

from sk_reporter.engineer.tk_catalog import write_manifest
from sk_reporter.paths import personnel_dir, projects_dir


def import_personnel() -> dict:
    from sk_reporter.personnel_db import import_personnel_xlsx_to_db

    xlsx = personnel_dir() / "Справочник персонала.xlsx"
    if not xlsx.is_file():
        raise FileNotFoundError(xlsx)
    return import_personnel_xlsx_to_db(xlsx)


def build_vor_caches(project_ids: list[str] | None = None) -> list[Path]:
    from sk_reporter.engineer.vor_legacy import write_project_vor_cache

    built: list[Path] = []
    for proj in sorted(projects_dir().iterdir()):
        if not proj.is_dir() or proj.name.startswith("."):
            continue
        if project_ids and proj.name not in project_ids:
            continue
        meta_path = proj / "project.yaml"
        if not meta_path.is_file():
            continue
        meta = yaml.safe_load(meta_path.read_text(encoding="utf-8")) or {}
        has_vor = bool(meta.get("vor_docx")) or bool(meta.get("vor_doc"))
        if not has_vor:
            continue
        try:
            built.append(write_project_vor_cache(proj))
        except FileNotFoundError:
            continue
    return built


def main() -> None:
    parser = argparse.ArgumentParser(description="Build engineer data caches")
    parser.add_argument(
        "--personnel",
        action="store_true",
        help="Import data/personnel/Справочник персонала.xlsx → PostgreSQL",
    )
    parser.add_argument("--tk", action="store_true", help="Write data/tk/manifest.yaml")
    parser.add_argument("--vor", action="store_true", help="Parse VOR docx → vor.json")
    parser.add_argument("--luvr", action="store_true", help="Export luvr.yaml from xlsx")
    parser.add_argument("--all", action="store_true", help="All of the above")
    parser.add_argument("--project", action="append", dest="projects", help="Project id filter")
    args = parser.parse_args()

    if args.all or not any((args.personnel, args.tk, args.vor, args.luvr)):
        args.personnel = args.tk = args.vor = args.luvr = True

    results: dict[str, str] = {}
    if args.personnel:
        result = import_personnel()
        results["personnel"] = f"upserted={result.get('upserted', 0)}"
    if args.luvr:
        from sk_reporter.luvr_store import export_luvr

        results["luvr"] = str(export_luvr())
    if args.tk:
        results["tk_manifest"] = str(write_manifest())
    if args.vor:
        for p in build_vor_caches(args.projects):
            results[f"vor_{p.parent.name}"] = str(p)

    print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
