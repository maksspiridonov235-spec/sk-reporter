#!/usr/bin/env python3
"""Собрать кэши данных инженера: vor.json; персонал → PostgreSQL."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from sk_reporter.paths import personnel_dir, projects_dir


def import_personnel() -> dict:
    from sk_reporter.personnel_db import import_personnel_xlsx_to_db

    xlsx = personnel_dir() / "Справочник персонала.xlsx"
    if not xlsx.is_file():
        raise FileNotFoundError(xlsx)
    return import_personnel_xlsx_to_db(xlsx)


def _vor_meta_for_dir(proj: Path) -> dict | None:
    meta = {"id": proj.name}
    docx = sorted(proj.glob("*.docx"))
    if docx:
        meta["vor_docx"] = docx[0].name
        return meta
    doc = sorted(proj.glob("*.doc"))
    if doc:
        meta["vor_doc"] = [d.name for d in doc]
        return meta
    return None


def build_vor_caches(project_ids: list[str] | None = None) -> list[Path]:
    from sk_reporter.engineer.vor_legacy import write_project_vor_cache

    built: list[Path] = []
    for proj in sorted(projects_dir().iterdir()):
        if not proj.is_dir() or proj.name.startswith("."):
            continue
        if project_ids and proj.name not in project_ids:
            continue
        if not _vor_meta_for_dir(proj):
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
    parser.add_argument("--vor", action="store_true", help="Parse VOR docx → vor.json")
    parser.add_argument("--all", action="store_true", help="Personnel + VOR")
    parser.add_argument("--project", action="append", dest="projects", help="Project id filter")
    args = parser.parse_args()

    if args.all or not any((args.personnel, args.vor)):
        args.personnel = args.vor = True

    results: dict[str, str] = {}
    if args.personnel:
        result = import_personnel()
        results["personnel"] = f"upserted={result.get('upserted', 0)}"
    if args.vor:
        for p in build_vor_caches(args.projects):
            results[f"vor_{p.parent.name}"] = str(p)

    print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
