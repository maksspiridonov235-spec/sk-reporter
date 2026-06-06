"""Блок «Отчёт инженера»: ВОР, ТК, сборка docx."""

from sk_reporter.engineer.vor_parser import parse_vor_docx, write_vor_cache

__all__ = ["parse_vor_docx", "write_vor_cache"]
