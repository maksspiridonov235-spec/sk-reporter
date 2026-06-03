import io
import json
import os
import shutil
import zipfile
from datetime import datetime
from pathlib import Path

from docx import Document
from fastapi import HTTPException
from pydantic import BaseModel

from apply_template_layout import hardcoded_layout
from config import (
    AGENT_ENABLED,
    LAYOUT_TEMPLATE_FILE,
    RESULT_DIR,
    TEMPLATES_DIR,
    UPLOAD_DIR,
    detect_company,
    merge_report_into_template,
)
from docx_processing import merge_reports, prepare_uploaded_reports, rename_results, rename_templates


class PrepareBody(BaseModel):
    date: str | None = None  # YYYY-MM-DD из поля «Дата в отчёте»


def parse_report_date(body: PrepareBody | None) -> str:
    if not body or not body.date:
        raise HTTPException(
            status_code=400,
            detail="Укажите date (YYYY-MM-DD) — поле «Дата в отчёте» в блоке «Макросы (до сборки)»",
        )
    try:
        return datetime.strptime(body.date, "%Y-%m-%d").strftime("%d.%m.%Y")
    except ValueError:
        raise HTTPException(status_code=400, detail="date: формат YYYY-MM-DD")


def layout_template_path() -> Path:
    path = (TEMPLATES_DIR / LAYOUT_TEMPLATE_FILE).resolve()
    if not path.is_file():
        raise HTTPException(
            status_code=404,
            detail=f"Положите «{LAYOUT_TEMPLATE_FILE}» в папку болванок: {path}",
        )
    return path


def sse(data: dict) -> str:
    return f"data: {json.dumps(data, ensure_ascii=False)}\n\n"


def zip_files(files: list[Path]) -> io.BytesIO:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(f, f.name)
    buf.seek(0)
    return buf


def do_merge(template_path: str, report_paths: list[str], output_path: str) -> int:
    if not AGENT_ENABLED:
        return merge_reports(template_path, report_paths, output_path)

    shutil.copy2(template_path, output_path)
    inserted = 0
    for i, rp in enumerate(sorted(report_paths)):
        tmp = output_path + ".tmp.docx"
        shutil.copy2(output_path, tmp)
        master = Document(tmp)
        if i > 0:
            master.add_page_break()
        master.save(tmp)
        ok = merge_report_into_template(tmp, rp, output_path)
        os.remove(tmp)
        if ok:
            inserted += 1
    return inserted


def find_reports_for_company(company_name: str, keywords: list[str]):
    found = []
    for f in UPLOAD_DIR.iterdir():
        if f.suffix.lower() not in (".docx", ".doc"):
            continue
        if AGENT_ENABLED:
            detected = detect_company(str(f))
            if detected and detected == company_name:
                found.append(f)
        else:
            kw_lower = [k.lower() for k in keywords]
            if any(k in f.name.lower() for k in kw_lower):
                found.append(f)
    return found


__all__ = [
    "PrepareBody",
    "parse_report_date",
    "layout_template_path",
    "sse",
    "zip_files",
    "do_merge",
    "find_reports_for_company",
    "hardcoded_layout",
    "prepare_uploaded_reports",
    "rename_results",
    "rename_templates",
]
