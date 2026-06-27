"""Пайплайн расстановки — логика как в Noviy proekt AI (ref_* модули)."""

from __future__ import annotations

import shutil
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Callable

import pandas as pd

from sk_reporter.deployment.ref_collect import extract_from_docx
from sk_reporter.deployment.ref_pril7 import fill_pril7
from sk_reporter.deployment.ref_rasstanovka import generate_rasstanovka


def _build_summary_df(reports_dir: Path, log_func: Callable[[str], None]) -> pd.DataFrame | None:
    docx_files = sorted(
        f for f in reports_dir.iterdir() if f.suffix.lower() in (".docx", ".doc")
    )
    if not docx_files:
        log_func("ОШИБКА: нет загруженных отчётов .docx")
        return None

    log_func(f"Парсинг отчётов: {len(docx_files)} файлов")
    rows: list[dict[str, str]] = []
    for i, path in enumerate(docx_files, 1):
        log_func(f"[{i}/{len(docx_files)}] {path.name}")
        row = extract_from_docx(path)
        row["Файл"] = path.name
        rows.append(row)

    df = pd.DataFrame(rows, columns=["Файл", "Дата", "Объект", "Инженер СК", "Генподрядчик"])
    df = df.dropna(subset=["Инженер СК", "Дата", "Объект"])
    df["Инженер СК"] = df["Инженер СК"].astype(str).str.strip()
    df["Дата"] = df["Дата"].astype(str).str.strip()
    df["Объект"] = df["Объект"].astype(str).str.strip()
    df["Генподрядчик"] = df["Генподрядчик"].fillna("").astype(str).str.strip()
    df = df[(df["Инженер СК"] != "") & (df["Объект"] != "") & (~df["Дата"].str.startswith("Ошибка:"))]

    if df.empty:
        log_func("ОШИБКА: после парсинга нет валидных строк")
        return None
    log_func(f"Строк в summary: {len(df)}")
    return df


def run_deployment(
    *,
    reports_dir: Path,
    pril7_path: Path,
    template_path: Path,
    output_dir: Path,
    report_date: str,
    log_func: Callable[[str], None] = print,
) -> tuple[Path | None, list[str]]:
    logs: list[str] = []

    def log(msg: str) -> None:
        logs.append(msg)
        log_func(msg)

    output_dir.mkdir(parents=True, exist_ok=True)
    work = output_dir / "_work"
    if work.exists():
        shutil.rmtree(work)
    work.mkdir()

    df = _build_summary_df(reports_dir, log)
    if df is None:
        return None, logs

    summary_path = work / "summary.xlsx"
    df.to_excel(summary_path, index=False)
    log(f"Summary: {summary_path.name}")

    pril7_work = work / pril7_path.name
    shutil.copy2(pril7_path, pril7_work)
    if not fill_pril7(str(summary_path), str(pril7_work), log_func=log):
        log("ОШИБКА: не удалось заполнить Приложение 7")
        return None, logs

    # как в референсе: fmt=excel; дата из UI — только для чтения столбца в Прил.7
    out_path, _ = generate_rasstanovka(
        str(pril7_work),
        str(template_path),
        str(work),
        report_date=report_date,
        fmt="excel",
        log_func=log,
    )
    if not out_path:
        log("ОШИБКА: не удалось сформировать расстановку")
        return None, logs

    ras_path = Path(out_path)
    date_tag = (
        datetime.strptime(report_date, "%Y-%m-%d").strftime("%d.%m.%Y")
        if report_date
        else datetime.now().strftime("%d.%m.%Y")
    )
    zip_name = f"расстановка_{date_tag.replace('.', '-')}.zip"
    zip_path = output_dir / zip_name
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(summary_path, summary_path.name)
        zf.write(pril7_work, pril7_work.name)
        zf.write(ras_path, ras_path.name)

    shutil.rmtree(work, ignore_errors=True)
    log(f"Архив готов: {zip_name}")
    return zip_path, logs
