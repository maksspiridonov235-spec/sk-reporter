"""
Ревизор: второй проход после check_agent.
Сверяет черновик check с оригиналом docx и выдаёт исправленный отчёт в том же формате.
"""

from pathlib import Path

from sk_reporter.agent.check_agent import MODEL, extract_full_text, extract_sk_section

SYSTEM_PROMPT = """Ты — строгий ревизор отчётов строительного контроля.
Тебе дают оригинальный отчёт и черновик исправления от первого агента.

Твоя задача — найти ошибки и упущения в черновике и выдать финальный вариант.

ОБЯЗАТЕЛЬНО ПРОВЕРЬ:
1. В ЧАСТИ 1 — все работы из оригинала (с объёмом за сутки ≠ 0); ни одну не пропусти; работы с суточным объёмом 0 не включай.
2. Объёмы (Проектный / за сутки / Накопительный) — копируй из оригинала дословно, цифры и единицы не меняй.
3. В ЧАСТИ 2 — для каждой работы есть описание; не выбрасывай факты, которые написал инженер; допиши только если описание пустое или формальное.
4. ЧАСТИ 3 и 4 — столько же пунктов, сколько работ в ЧАСТИ 1, в том же порядке.
5. Участок, ПК и Ссылка — соответствуют работе, исправь только опечатки и пробелы.

ВАЖНО:
- Не упоминай ГОСТы
- Не выдумывай работы и объёмы
- Если черновик верен — воспроизведи его с минимальными правками

ФОРМАТ ОТВЕТА (как у check_agent):

## РЕЗЮМЕ ПРОВЕРКИ
- [✓ или ⚠] Объёмы: ...
- [✓ или ⚠] Нулевые объёмы: ...
- [✓ или ⚠] Описания: ...
- [✓ или ⚠] Участок, ПК: ...
- [✓ или ⚠] Ссылка: ...

## ИСПРАВЛЕННЫЙ ОТЧЁТ
ЧАСТЬ 1 — все пункты с объёмами
ЧАСТЬ 2 — описания по каждому пункту
ЧАСТЬ 3 — Участок, ПК
ЧАСТЬ 4 — Ссылка"""


def verify_report(filepath: str, check_result: dict) -> dict:
    """Перепроверяет черновик check_agent. При сбое возвращает текст check (fallback)."""
    filename = Path(filepath).name
    draft = (check_result or {}).get("report", "").strip()

    if not draft:
        print(f"[VERIFY_AGENT] нет черновика check, пропуск: {filename}")
        return {
            "ok": False,
            "report": "",
            "fallback": True,
            "verify_ran": False,
            "_source_file": filename,
        }

    print(f"[VERIFY_AGENT] перепроверяю: {filename}")
    full_text = extract_full_text(filepath)
    sk_section = extract_sk_section(filepath)

    user_prompt = f"""Сверь оригинал и черновик первого агента. Выдай финальный ИСПРАВЛЕННЫЙ ОТЧЁТ.

---СЕКЦИЯ СК (таблица)---
{sk_section or "(не найдена)"}
---КОНЕЦ СЕКЦИИ---

---ОРИГИНАЛЬНЫЙ ТЕКСТ ОТЧЁТА---
{full_text}
---КОНЕЦ ОРИГИНАЛА---

---ЧЕРНОВИК ОТ CHECK_AGENT---
{draft}
---КОНЕЦ ЧЕРНОВИКА---

Проверь черновик по критериям из системного промпта.
Исправь пропущенные работы, искажённые объёмы, потерянный текст инженера, несовпадение числа пунктов в ЧАСТЯХ 2–4.

Выведи ответ строго в формате:

## РЕЗЮМЕ ПРОВЕРКИ
(пять строк с ✓ или ⚠)

## ИСПРАВЛЕННЫЙ ОТЧЁТ
ЧАСТЬ 1:
Инспекционный контроль по проведённым работам:
1. [название]
Проектный объем – ...
Объем за сутки – ...
Накопительный объем – ...
(все работы — ни одну не пропускай)

ЧАСТЬ 2:
Наряд-допуск проверен, замечаний нет. Работы ведутся в соответствии с технологическими картами и РД.
1. [описание для работы 1]
2. [описание для работы 2]
(все работы)

ЧАСТЬ 3:
1. [участок, ПК]
2. ...

ЧАСТЬ 4:
1. [ссылка]
2. ...

ВАЖНО: только РЕЗЮМЕ и ИСПРАВЛЕННЫЙ ОТЧЁТ с ЧАСТЯМИ 1–4."""

    try:
        import ollama

        response = ollama.chat(
            model=MODEL,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            stream=False,
        )

        report_text = response.get("message", {}).get("content", "").strip()

        if not report_text:
            print(f"[VERIFY_AGENT] пустой ответ, fallback на check: {filename}")
            return {
                "ok": False,
                "report": draft,
                "fallback": True,
                "verify_ran": True,
                "_source_file": filename,
            }

        has_issues = "в порядке" not in report_text.lower() or "нет" not in report_text.lower()
        result = {
            "ok": not has_issues,
            "report": report_text,
            "fallback": False,
            "verify_ran": True,
            "_source_file": filename,
        }
        print(f"[VERIFY_AGENT] {'OK' if result['ok'] else 'ISSUES'}: {filename}")
        return result

    except Exception as e:
        print(f"[VERIFY_AGENT] ошибка ({e}), fallback на check: {filename}")
        return {
            "ok": False,
            "report": draft,
            "fallback": True,
            "verify_ran": True,
            "_source_file": filename,
        }
