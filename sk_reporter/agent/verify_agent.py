"""
Перепроверка ответа check_agent перед inject.
1. Детерминированные правила (объёмы, полнота пунктов, текст инженера).
2. При нарушениях — один проход LLM-repair (дописать, не переписывать с нуля).
"""

from pathlib import Path

from sk_reporter.agent.report_parts import (
    _VOLUME_LABELS,
    count_numbered_items,
    engineer_facts_preserved,
    normalize_volume_value,
    parse_numbered_items,
    parse_parts,
    parse_works_from_part1_lines,
    volume_numeric,
)
from sk_reporter.agent.sk_extract import extract_sk_original, extract_sk_section

MODEL = "gemma4:31b-cloud"

REPAIR_SYSTEM = """Ты — строгий ревизор отчётов строительного контроля.
Тебе дали оригинал, черновик исправления и список нарушений.
Исправь ТОЛЬКО перечисленные нарушения. Не меняй цифры объёмов, если они совпадают с оригиналом.
Не сокращай и не выбрасывай факты из описаний инженера без необходимости.
Сохрани формат ответа: ## РЕЗЮМЕ ПРОВЕРКИ и ## ИСПРАВЛЕННЫЙ ОТЧЁТ с ЧАСТЯМИ 1–4."""


def _volumes_match(orig_volumes: dict, corr_volumes: dict, work_num: int) -> list[str]:
    issues = []
    for label in _VOLUME_LABELS:
        orig_val = orig_volumes.get(label, "")
        corr_val = corr_volumes.get(label, "")
        if not orig_val and not corr_val:
            continue
        if not corr_val:
            issues.append(f"работа {work_num}: нет строки «{label}»")
            continue
        orig_norm = normalize_volume_value(orig_val)
        corr_norm = normalize_volume_value(corr_val)
        orig_num = volume_numeric(orig_val)
        corr_num = volume_numeric(corr_val)
        if orig_num is not None and corr_num is not None and orig_num == corr_num:
            continue
        if orig_norm != corr_norm:
            issues.append(
                f"работа {work_num}: «{label}» изменено ({orig_val!r} → {corr_val!r})"
            )
    return issues


def _run_rules(original: dict, corrected_text: str) -> tuple[list[str], dict]:
    violations: list[str] = []

    if "## ИСПРАВЛЕННЫЙ ОТЧЁТ" not in corrected_text and "## ИСПРАВЛЕННЫЙ ОТЧЕТ" not in corrected_text.upper():
        violations.append("нет раздела «## ИСПРАВЛЕННЫЙ ОТЧЁТ»")

    part1, part2, part3, part4 = parse_parts(corrected_text)
    parsed = {"part1": part1, "part2": part2, "part3": part3, "part4": part4}

    if not part1 and not part2:
        violations.append("не распарсились ЧАСТЬ 1 / ЧАСТЬ 2")
        return violations, parsed

    corr_works = parse_works_from_part1_lines(part1)
    corr_desc = parse_numbered_items(part2)
    active_orig = original.get("active_works") or []

    if len(corr_works) != len(active_orig):
        violations.append(
            f"число работ с объёмами: в оригинале {len(active_orig)} (без нулевых суточных), "
            f"в исправлении {len(corr_works)}"
        )

    n = min(len(corr_works), len(active_orig))
    for i in range(n):
        orig = active_orig[i]
        corr = corr_works[i]
        work_num = orig["num"]
        violations.extend(_volumes_match(orig.get("volumes", {}), corr.get("volumes", {}), work_num))

        orig_desc = orig.get("description", "")
        corr_desc_text = corr_desc.get(work_num, corr_desc.get(corr["num"], ""))
        if orig_desc and not engineer_facts_preserved(orig_desc, corr_desc_text):
            violations.append(
                f"работа {work_num}: описание инженера сильно сокращено или переписано"
            )

    for label, lines, name in (
        (2, part2, "описания"),
        (3, part3, "Участок, ПК"),
        (4, part4, "Ссылка"),
    ):
        if not lines:
            violations.append(f"ЧАСТЬ {label} ({name}) пуста")
            continue
        counted = count_numbered_items(lines)
        if counted and counted != len(active_orig):
            violations.append(
                f"ЧАСТЬ {label}: пунктов {counted}, ожидалось {len(active_orig)}"
            )

    zero_orig = [w for w in original.get("works", []) if w.get("zero_daily")]
    for w in zero_orig:
        for cw in corr_works:
            if cw["num"] == w["num"]:
                violations.append(
                    f"работа {w['num']}: суточный объём 0 — должна быть удалена, но осталась в ЧАСТИ 1"
                )
                break

    return violations, parsed


def _repair_report(filepath: str, original: dict, draft_text: str, violations: list[str]) -> str:
    sk_section = original.get("sk_section") or extract_sk_section(filepath)
    user_prompt = f"""Оригинал (секция СК):
{sk_section}

Черновик исправления:
{draft_text}

Нарушения перепроверки (исправь все):
{chr(10).join(f'- {v}' for v in violations)}

Верни полный ответ в формате check_agent: ## РЕЗЮМЕ ПРОВЕРКИ и ## ИСПРАВЛЕННЫЙ ОТЧЁТ с ЧАСТЯМИ 1–4.
Объёмы копируй из оригинала дословно. Работы с суточным объёмом 0 не включай в ЧАСТЬ 1."""

    import ollama

    response = ollama.chat(
        model=MODEL,
        messages=[
            {"role": "system", "content": REPAIR_SYSTEM},
            {"role": "user", "content": user_prompt},
        ],
        stream=False,
    )
    return response.get("message", {}).get("content", "").strip()


def verify_report(filepath: str, check_result: dict) -> dict:
    """
    Перепроверяет ответ check_agent.
    Возвращает ok, violations, report (текст для inject), repaired.
    """
    filename = Path(filepath).name
    draft = (check_result or {}).get("report", "").strip()

    if not draft:
        return {
            "ok": False,
            "violations": ["пустой ответ check_agent"],
            "report": "",
            "repaired": False,
            "_source_file": filename,
        }

    original = extract_sk_original(filepath)
    if not original.get("ok"):
        return {
            "ok": False,
            "violations": [original.get("error") or "не удалось прочитать оригинал"],
            "report": draft,
            "repaired": False,
            "_source_file": filename,
        }

    violations, _ = _run_rules(original, draft)
    if not violations:
        print(f"[VERIFY_AGENT] OK: {filename}")
        return {
            "ok": True,
            "violations": [],
            "report": draft,
            "repaired": False,
            "_source_file": filename,
        }

    print(f"[VERIFY_AGENT] violations ({len(violations)}): {filename}")
    for v in violations[:8]:
        print(f"  - {v}")

    repaired_text = draft
    try:
        repaired_text = _repair_report(filepath, original, draft, violations)
    except Exception as e:
        print(f"[VERIFY_AGENT] repair error: {e}")
        return {
            "ok": False,
            "violations": violations + [f"repair: {e}"],
            "report": draft,
            "repaired": False,
            "_source_file": filename,
        }

    if not repaired_text:
        return {
            "ok": False,
            "violations": violations + ["repair: пустой ответ модели"],
            "report": draft,
            "repaired": False,
            "_source_file": filename,
        }

    violations2, _ = _run_rules(original, repaired_text)
    if not violations2:
        print(f"[VERIFY_AGENT] OK after repair: {filename}")
        return {
            "ok": True,
            "violations": [],
            "report": repaired_text,
            "repaired": True,
            "_source_file": filename,
        }

    print(f"[VERIFY_AGENT] FAIL after repair ({len(violations2)}): {filename}")
    return {
        "ok": False,
        "violations": violations2,
        "report": repaired_text,
        "repaired": True,
        "_source_file": filename,
    }
