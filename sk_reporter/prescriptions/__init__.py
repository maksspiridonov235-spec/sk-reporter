"""Проверка предписаний (Excel). Отдельно от docx / check_agent."""

from sk_reporter.prescriptions.check import check_prescription_file, write_checked_copy

__all__ = ["check_prescription_file", "write_checked_copy"]
