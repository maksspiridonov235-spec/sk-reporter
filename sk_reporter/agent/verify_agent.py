"""
Агент перепроверки отчётов строительного контроля.
Запускается ПОСЛЕ inject: сверяет текст инженера (до правок) с тем, что вставила модель в docx.
"""

from pathlib import Path

from sk_reporter.agent.check_agent import MODEL, extract_full_text, extract_sk_section

SYSTEM_PROMPT = """Ты — ведущий инженер строительного контроля. Перепроверяешь отчёт ПОСЛЕ вставки правок моделью.

Тебе даны два варианта секции СК:
1. ОРИГИНАЛ — как написал инженер (до правок)
2. ПОСЛЕ ПРАВОК — что вставила модель в документ

Сверь их и оцени качество правок.

КРИТЕРИИ:
1. ОПИСАНИЕ: Модель не выбросила факты инженера; описания контроля сохранены или дополнены по делу.
2. УЧАСТОК, ПК: Заполнено, соответствует работе; нет лишних опечаток.
3. ССЫЛКА: Заполнена, соответствует работе.
4. ОБЪЕМЫ: Цифры и единицы не искажены относительно оригинала инженера.
5. НУЛЕВЫЕ ОБЪЕМЫ: Работы с суточным объёмом 0 не должны остаться.

ВАЖНО:
- Не упоминай ГОСТы
- Не выдумывай работы и объёмы
- Пиши только о реально найденных расхождениях

ФОРМАТ ОТВЕТА:

## РЕЗЮМЕ ПРОВЕРКИ
- [✓ или ⚠] Объёмы: [в порядке / перечисли расхождения]
- [✓ или ⚠] Нулевые объёмы: [нет / перечисли]
- [✓ или ⚠] Описания: [в порядке / перечисли проблемы]
- [✓ или ⚠] Участок, ПК: [в порядке / перечисли]
- [✓ или ⚠] Ссылка: [в порядке / перечисли]

## ОТЧЁТ О СВЕРКЕ
Кратко по каждой работе: что было у инженера, что сделала модель, есть ли замечания.
Если всё в порядке — напиши «Расхождений нет»."""


def _build_user_prompt(
    original_full: str,
    original_sk: str,
    model_full: str,
    model_sk: str,
) -> str:
    return f"""Сверь оригинал инженера и текст после вставки правок моделью.

---СЕКЦИЯ СК: ОРИГИНАЛ ИНЖЕНЕРА---
{original_sk or "(не найдена)"}
---КОНЕЦ---

---СЕКЦИЯ СК: ПОСЛЕ ПРАВОК МОДЕЛИ---
{model_sk or "(не найдена)"}
---КОНЕЦ---

---ПОЛНЫЙ ТЕКСТ: ОРИГИНАЛ ИНЖЕНЕРА---
{original_full}
---КОНЕЦ---

---ПОЛНЫЙ ТЕКСТ: ПОСЛЕ ПРАВОК МОДЕЛИ---
{model_full}
---КОНЕЦ---

Проверь: модель не потеряла работы, не исказила объёмы, сохранила смысл описаний инженера.
Выведи только РЕЗЮМЕ ПРОВЕРКИ и ОТЧЁТ О СВЕРКЕ."""


def verify_report(filepath: str, original_snapshot: dict) -> dict:
    """Сверяет docx после inject с сохранённым снимком оригинала инженера."""
    filename = Path(filepath).name
    snap = original_snapshot or {}
    original_full = (snap.get("full_text") or "").strip()
    original_sk = (snap.get("sk_section") or "").strip()

    if not original_full:
        print(f"[VERIFY_AGENT] нет снимка оригинала, пропуск: {filename}")
        return {
            "ok": False,
            "report": "Нет снимка оригинала инженера для сверки",
            "fallback": True,
            "verify_ran": False,
            "_source_file": filename,
        }

    print(f"[VERIFY_AGENT] перепроверяю (инженер vs модель): {filename}")

    model_full = extract_full_text(filepath)
    model_sk = extract_sk_section(filepath)

    if not model_full:
        print(f"[VERIFY_AGENT] не прочитан docx после inject: {filename}")
        return {
            "ok": False,
            "report": "Не удалось прочитать документ после вставки правок",
            "fallback": True,
            "verify_ran": True,
            "_source_file": filename,
        }

    user_prompt = _build_user_prompt(original_full, original_sk, model_full, model_sk)

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
            print(f"[VERIFY_AGENT] пустой ответ: {filename}")
            return {
                "ok": False,
                "report": "Пустой ответ модели при перепроверке",
                "fallback": True,
                "verify_ran": True,
                "_source_file": filename,
            }

        has_issues = "в порядке" not in report_text.lower() or "расхождений нет" not in report_text.lower()

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
        print(f"[VERIFY_AGENT] ошибка ({e}): {filename}")
        return {
            "ok": False,
            "report": f"Ошибка перепроверки: {e}",
            "fallback": True,
            "verify_ran": True,
            "_source_file": filename,
        }
