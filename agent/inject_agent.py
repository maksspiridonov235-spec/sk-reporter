"""
Агент инъекции: берёт оригинальный HTML отчёта + исправленный текст от check_agent,
через LLM вставляет исправления в HTML, затем конвертирует обратно в docx через pandoc.
"""

import subprocess
import tempfile
import os
from pathlib import Path

MODEL = "gemma4:31b-cloud"

INJECT_SYSTEM_PROMPT = """Ты — редактор HTML-документа. Тебе дают:
1. Оригинальный HTML отчёта строительного контроля (таблица)
2. Исправленный текст — ЧАСТЬ 1 (работы с объёмами) и ЧАСТЬ 2 (описания)

Твоя задача: вернуть ПОЛНЫЙ оригинальный HTML, заменив только содержимое двух ячеек:
- Ячейка с работами (содержит «Инспекционный контроль» и объёмы) — замени на ЧАСТЬ 1
- Ячейка с описаниями (содержит «Наряд-допуск проверен») — замени на ЧАСТЬ 2

Правила:
- Не меняй структуру таблицы, теги, атрибуты, colspan, стили
- Не трогай шапку, подписи, статусы, даты и прочие ячейки
- Сохраняй <br> для переносов строк внутри ячеек
- Возвращай ТОЛЬКО полный HTML без markdown-обёрток и пояснений"""


def inject_corrections(original_html: str, corrected_text: str, source_filename: str) -> dict:
    filename = Path(source_filename).stem

    user_prompt = f"""Вот оригинальный HTML отчёта:

{original_html}

---ИСПРАВЛЕННЫЙ ТЕКСТ---
{corrected_text}
---КОНЕЦ ИСПРАВЛЕННОГО ТЕКСТА---

Верни полный HTML с заменёнными ячейками."""

    try:
        import ollama
        response = ollama.chat(
            model=MODEL,
            messages=[
                {"role": "system", "content": INJECT_SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            stream=False,
        )
        fixed_html = response.get("message", {}).get("content", "").strip()

        if not fixed_html:
            return {"ok": False, "error": "Пустой ответ модели", "docx_path": None}

        # Убираем markdown-обёртку если модель её добавила
        if fixed_html.startswith("```"):
            lines = fixed_html.split("\n")
            fixed_html = "\n".join(lines[1:-1] if lines[-1].strip() == "```" else lines[1:])

        return html_to_docx(fixed_html, filename)

    except Exception as e:
        print(f"[INJECT_AGENT] LLM error: {e}")
        return {"ok": False, "error": str(e), "docx_path": None}


def html_to_docx(html: str, stem: str) -> dict:
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            html_path = os.path.join(tmpdir, "report.html")
            docx_path = os.path.join(tmpdir, f"{stem}_исправлен.docx")

            with open(html_path, "w", encoding="utf-8") as f:
                f.write(f"<html><head><meta charset='utf-8'></head><body>{html}</body></html>")

            result = subprocess.run(
                ["pandoc", html_path, "-o", docx_path, "--from=html", "--to=docx"],
                capture_output=True, text=True
            )

            if result.returncode != 0:
                return {"ok": False, "error": result.stderr, "docx_path": None}

            output_dir = Path(__file__).parent.parent / "output"
            output_dir.mkdir(exist_ok=True)
            final_path = output_dir / f"{stem}_исправлен.docx"

            with open(docx_path, "rb") as src, open(final_path, "wb") as dst:
                dst.write(src.read())

            print(f"[INJECT_AGENT] Saved: {final_path}")
            return {"ok": True, "docx_path": str(final_path)}

    except Exception as e:
        print(f"[INJECT_AGENT] pandoc error: {e}")
        return {"ok": False, "error": str(e), "docx_path": None}
