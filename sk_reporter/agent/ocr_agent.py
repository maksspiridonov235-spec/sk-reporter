"""
Маршрутизация отчётов СК по болванкам: detect_company по имени файла.
Склейка docx — sk_reporter.docx_processing.merge_reports (webapp/main._do_merge).
"""

from pathlib import Path
from typing import Optional

from sk_reporter.companies import COMPANIES

PRIORITY_TYPES = ("Геодезический контроль", "ОЗОТОБОС")


def detect_company(filepath: str) -> Optional[str]:
    filename_lower = Path(filepath).name.lower()

    # 1) тип работ имеет приоритет: геодезия / озотобос по имени файла
    for company in COMPANIES:
        if company.name in PRIORITY_TYPES and any(kw in filename_lower for kw in company.keywords):
            print(f"[TYPE] {filename_lower} → {company.name}")
            return company.name

    # 2) остальное — по названию организации в имени файла
    for company in COMPANIES:
        if company.name in PRIORITY_TYPES:
            continue
        if any(kw in filename_lower for kw in company.keywords):
            print(f"[FILENAME] {filename_lower} → {company.name}")
            return company.name

    return None
