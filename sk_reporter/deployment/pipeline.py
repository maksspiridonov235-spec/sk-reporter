"""Полный пайплайн: docx → summary → Прил.7 → расстановка → ZIP."""

from __future__ import annotations

import shutil
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Callable

from sk_reporter.deployment.lookup import known_fio_set
from sk_reporter.deployment.parser import extract_from_docx
from sk_reporter.deployment.pril7_fill import fill_pril7
from sk_reporter.deployment.rasstanovka import generate_rasstanovka
from sk_reporter.deployment.summary import validate_and_filter_rows, write_summary

def run_deployment(
    *,
    reports_dir: Path,
    pril7_path: Path,
    template_path: Path,
    output_dir: Path,
    report_date: str,
    log_func: Callable[[str], None] = print,
) -> tuple[Path | None, list[str]]:
    """Возвращает (путь к ZIP или None, строки лога)."""
    logs: list[str] = []

    def log(msg: str) -> None:
        logs.append(msg)
        log_func(msg)

    output_dir.mkdir(parents=True, exist_ok=True)
    work = output_dir / "_work"
    if work.exists():
        shutil.rmtree(work)
    work.mkdir()

    docx_files = sorted(
        f for f in reports_dir.iterdir() if f.suffix.lower() in (".docx", ".doc")
    )
    if not docx_files:
        log("ОШИБКА: нет загруженных отчётов .docx")
        return None, logs

    log(f"Парсинг отчётов: {len(docx_files)} файлов")
    parsed: list[dict[str, str]] = []
    for i, path in enumerate(docx_files, 1):
        log(f"[{i}/{len(docx_files)}] {path.name}")
        row = extract_from_docx(path)
        row["Файл"] = path.name
        parsed.append(row)

    known = known_fio_set()
    log(f"Справочник сотрудников: {len(known)} ФИО")
    accepted = validate_and_filter_rows(parsed, known_fio=known, log_func=log)
    if not accepted:
        log("ОШИБКА: после сверки не осталось строк")
        return None, logs
    log(f"Принято строк: {len(accepted)} (пропущено: {len(parsed) - len(accepted)})")

    if report_date:
        dd = datetime.strptime(report_date, "%Y-%m-%d").strftime("%d.%m.%Y")
        for row in accepted:
            row["Дата"] = dd

    summary_path = work / "summary.xlsx"
    write_summary(accepted, summary_path)
    log(f"Summary: {summary_path.name}")

    pril7_work = work / pril7_path.name
    shutil.copy2(pril7_path, pril7_work)
    if not fill_pril7(summary_path, pril7_work, log_func=log):
        log("ОШИБКА: не удалось заполнить Приложение 7")
        return None, logs

    ras_path = generate_rasstanovka(
        pril7_work,
        template_path,
        work,
        report_date=report_date,
        log_func=log,
    )
    if ras_path is None:
        log("ОШИБКА: не удалось сформировать расстановку")
        return None, logs

    date_tag = datetime.strptime(report_date, "%Y-%m-%d").strftime("%d.%m.%Y") if report_date else datetime.now().strftime("%d.%m.%Y")
    zip_name = f"расстановка_{date_tag.replace('.', '-')}.zip"
    zip_path = output_dir / zip_name
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(summary_path, summary_path.name)
        zf.write(pril7_work, pril7_work.name)
        zf.write(ras_path, ras_path.name)

    shutil.rmtree(work, ignore_errors=True)
    log(f"Архив готов: {zip_name}")
    return zip_path, logs
