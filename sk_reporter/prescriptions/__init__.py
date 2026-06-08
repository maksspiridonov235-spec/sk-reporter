"""Проверка предписаний (Excel). Отдельно от docx / sk_reporter.agent.check_agent."""

from sk_reporter.prescriptions.check_agent import (
    check_prescription,
    extract_form_fields,
    write_checked_copy,
)

__all__ = ["check_prescription", "extract_form_fields", "write_checked_copy"]
