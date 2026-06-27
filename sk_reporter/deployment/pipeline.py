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


def _validate_row(row: dict[str, str]) -> tuple[list[str], bool]:
    """Проблемы парсинга; ok=True если строку можно писать в Прил.7."""
    problems: list[str] = []
    date = str(row.get("Дата") or "").strip()
    fio = str(row.get("Инженер СК") or "").strip()
    obj = str(row.get("Объект") or "").strip()

    if date.startswith("Ошибка:"):
        problems.append(date)
    elif not date:
        problems.append("нет даты")
    if not fio:
        problems.append("нет инженера СК")
    if not obj:
        problems.append("нет объекта")
    return problems, len(problems) == 0


def _build_summary_df(reports_dir: Path, log_func: Callable[[str], None]) -> tuple[pd.DataFrame | None, pd.DataFrame | None]:
    """Полный summary (все файлы) и подмножество для заполнения Прил.7."""
    docx_files = sorted(
        f for f in reports_dir.iterdir() if f.suffix.lower() in (".docx", ".doc")
    )
    if not docx_files:
        log_func("ОШИБКА: нет загруженных отчётов .docx")
        return None, None

    log_func(f"Парсинг отчётов: {len(docx_files)} файлов")
    rows: list[dict[str, str]] = []
    for i, path in enumerate(docx_files, 1):
        log_func(f"[{i}/{len(docx_files)}] {path.name}")
        row = extract_from_docx(path)
        row["Файл"] = path.name
        problems, ok = _validate_row(row)
        row["Проблемы"] = "; ".join(problems)
        if not ok:
            log_func(f"  ⚠ в summary с пометкой: {row['Проблемы']}")
        rows.append(row)

    cols = ["Файл", "Дата", "Объект", "Инженер СК", "Генподрядчик", "Проблемы"]
    df_all = pd.DataFrame(rows, columns=cols)
    for col in ("Дата", "Объект", "Инженер СК", "Генподрядчик"):
        df_all[col] = df_all[col].fillna("").astype(str).str.strip()

    df_fill = df_all[df_all["Проблемы"] == ""].drop(columns=["Проблемы"])
    log_func(f"Summary: {len(df_all)} строк (в Прил.7 пойдёт: {len(df_fill)})")
    if df_fill.empty:
        log_func("ОШИБКА: нет ни одной строки для заполнения Прил.7")
        return df_all, None
    return df_all, df_fill


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

    df_all, df_fill = _build_summary_df(reports_dir, log)
    if df_all is None or df_fill is None:
        return None, logs

    summary_path = work / "summary.xlsx"
    df_all.to_excel(summary_path, index=False)
    log(f"Summary: {summary_path.name} ({len(df_all)} строк)")

    summary_fill_path = work / "summary_fill.xlsx"
    df_fill.to_excel(summary_fill_path, index=False)

    pril7_work = work / pril7_path.name
    shutil.copy2(pril7_path, pril7_work)
    if not fill_pril7(str(summary_fill_path), str(pril7_work), log_func=log):
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
