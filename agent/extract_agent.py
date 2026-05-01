"""
Агент извлечения отчёта: docx → HTML через mammoth.
Возвращает HTML-строку для передачи в check_agent и inject_agent.
"""

import mammoth
from pathlib import Path


def extract_html(filepath: str) -> dict:
    filename = Path(filepath).name
    try:
        with open(filepath, "rb") as f:
            result = mammoth.convert_to_html(f)
        html = result.value
        if not html.strip():
            return {"ok": False, "html": "", "_source_file": filename}
        return {"ok": True, "html": html, "_source_file": filename}
    except Exception as e:
        print(f"[EXTRACT_AGENT] Error: {e}")
        return {"ok": False, "html": "", "_source_file": filename}
