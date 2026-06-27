"""Сохранение xlsm через openpyxl на Linux (без Excel COM)."""

from __future__ import annotations

from pathlib import Path

from openpyxl.workbook.workbook import Workbook


def strip_legacy_vml(wb: Workbook) -> None:
    """Убирает legacy VML/комментарии — иначе openpyxl падает при save xlsm."""
    for ws in wb.worksheets:
        ws.legacy_drawing = None
        comments = getattr(ws, "_comments", None)
        if comments:
            comments.clear()


def save_xlsm_workbook(wb: Workbook, path: str | Path) -> None:
    strip_legacy_vml(wb)
    wb.save(path)
