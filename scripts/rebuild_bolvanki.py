#!/usr/bin/env python3
"""
Обновление болванок data/templates/*.docx: поля страницы как у шаблона инженера,
всё остальное — как в исходной болванке (колонтитул, жирный заголовок, оформление).

Предыдущая версия скрипта зря перезаписывала болванки телом шаблона инженера
и теряла верхний колонтитул.
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from sk_reporter.paths import templates_dir

DEFAULT_ENGINEER_TEMPLATE = ROOT / "data" / "engineer" / "report_template.docx"
DEFAULT_SOURCE_GIT = "7d2f95e"


def _git_show_docx(rev: str, repo_path: Path) -> bytes:
    rel = repo_path.as_posix()
    return subprocess.check_output(["git", "show", f"{rev}:{rel}"], cwd=ROOT)


def _engineer_pg_mar(engineer_template: Path) -> etree.Element:
    doc = Document(str(engineer_template))
    sect = doc.element.body.find(qn("w:sectPr"))
    if sect is None:
        raise ValueError(f"No sectPr in {engineer_template}")
    pg_mar = sect.find(qn("w:pgMar"))
    if pg_mar is None:
        raise ValueError(f"No pgMar in {engineer_template}")
    return pg_mar


def _apply_pg_mar(docx_bytes: bytes, pg_mar_src: etree.Element) -> bytes:
    with tempfile.TemporaryDirectory() as tmpdir:
        path = Path(tmpdir) / "in.docx"
        path.write_bytes(docx_bytes)
        with zipfile.ZipFile(path, "r") as zin:
            doc_xml = zin.read("word/document.xml")
        root = etree.fromstring(doc_xml)
        sect = root.find(f".//{qn('w:sectPr')}")
        if sect is None:
            raise ValueError("sectPr not found")
        existing = sect.find(qn("w:pgMar"))
        new_pg_mar = etree.fromstring(etree.tostring(pg_mar_src))
        if existing is not None:
            sect.remove(existing)
        sect.insert(0, new_pg_mar)
        new_doc_xml = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone="yes"
        )
        out = Path(tmpdir) / "out.docx"
        with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(out, "w") as zout:
            for item in zin.infolist():
                data = new_doc_xml if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)
        return out.read_bytes()


def _bolvanka_summary(docx_bytes: bytes) -> dict:
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp.write(docx_bytes)
        tmp_path = Path(tmp.name)
    try:
        with zipfile.ZipFile(tmp_path) as z:
            has_header = "word/header1.xml" in z.namelist()
            header_text = ""
            if has_header:
                hdr = etree.fromstring(z.read("word/header1.xml"))
                header_text = "".join(t.text or "" for t in hdr.findall(".//" + qn("w:t")))
            root = etree.fromstring(z.read("word/document.xml"))
            p0 = root.findall(".//" + qn("w:p"))[0]
            title = "".join(t.text or "" for t in p0.findall(".//" + qn("w:t")))
            bold = False
            for r in p0.findall(qn("w:r")):
                r_pr = r.find(qn("w:rPr"))
                if r_pr is not None and r_pr.find(qn("w:b")) is not None:
                    bold = True
                    break
            p_pr = p0.find(qn("w:pPr"))
            jc_el = p_pr.find(qn("w:jc")) if p_pr is not None else None
            jc = jc_el.get(qn("w:val")) if jc_el is not None else None
            top_el = root.find(".//" + qn("w:pgMar"))
            top_val = top_el.get(qn("w:top")) if top_el is not None else None
        return {
            "header": header_text[:60],
            "title": title[:60],
            "bold": bold,
            "jc": jc,
            "top": top_val,
        }
    finally:
        tmp_path.unlink(missing_ok=True)


def rebuild_bolvanka_from_git(
    *,
    source_rev: str,
    engineer_template: Path,
    output_path: Path,
    repo_path: Path,
) -> dict:
    original = _git_show_docx(source_rev, repo_path)
    pg_mar = _engineer_pg_mar(engineer_template)
    patched = _apply_pg_mar(original, pg_mar)
    output_path.write_bytes(patched)
    return _bolvanka_summary(patched)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Restore bolvanki content from git and apply engineer page margins",
    )
    parser.add_argument(
        "--engineer-template",
        type=Path,
        default=DEFAULT_ENGINEER_TEMPLATE,
        help=f"Margins source (default: {DEFAULT_ENGINEER_TEMPLATE})",
    )
    parser.add_argument(
        "--source-git",
        default=DEFAULT_SOURCE_GIT,
        help=f"Git revision with original bolvanki (default: {DEFAULT_SOURCE_GIT})",
    )
    parser.add_argument(
        "--templates-dir",
        type=Path,
        default=None,
        help="Output dir (default: data/templates)",
    )
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    engineer = args.engineer_template.resolve()
    out_dir = (args.templates_dir or templates_dir()).resolve()

    if not engineer.is_file():
        print(f"[ERR] Engineer template not found: {engineer}", file=sys.stderr)
        return 1
    if not out_dir.is_dir():
        print(f"[ERR] Templates dir not found: {out_dir}", file=sys.stderr)
        return 1

    files = sorted(out_dir.glob("*.docx"))
    if not files:
        print(f"[ERR] No .docx in {out_dir}", file=sys.stderr)
        return 1

    for path in files:
        rel = path.relative_to(ROOT)
        if args.dry_run:
            original = _git_show_docx(args.source_git, rel)
            info = _bolvanka_summary(_apply_pg_mar(original, _engineer_pg_mar(engineer)))
            print(
                f"[DRY] {path.name}: header={info['header']!r} "
                f"bold={info['bold']} jc={info['jc']} top={info['top']}"
            )
            continue
        info = rebuild_bolvanka_from_git(
            source_rev=args.source_git,
            engineer_template=engineer,
            output_path=path,
            repo_path=rel,
        )
        print(
            f"[OK] {path.name}: header={info['header']!r} "
            f"bold={info['bold']} jc={info['jc']} top={info['top']}"
        )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
