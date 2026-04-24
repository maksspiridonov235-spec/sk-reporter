"""
Pipeline: Парсер → Нормализатор → Верификатор.
Запускает всю цепочку для списка файлов и возвращает сводный результат.
"""

from pathlib import Path
from typing import Optional

from agent.report_parser import parse_report
from agent.normalizer import normalize
from agent.verifier import verify


def run_pipeline(filepaths: list[str], api_key: Optional[str] = None) -> list[dict]:
    """
    Прогоняет каждый файл через полную цепочку агентов.
    Возвращает список dict с ключами: parsed, normalized, verification, _source_file.
    """
    results = []

    for fp in filepaths:
        filename = Path(fp).name
        print(f"\n{'='*50}")
        print(f"[PIPELINE] Обрабатываю: {filename}")

        # Шаг 1: Парсинг
        parsed = parse_report(fp, api_key=api_key)
        if not parsed:
            results.append({
                "_source_file": filename,
                "_pipeline_error": "Парсер не смог извлечь данные",
            })
            continue

        # Шаг 2: Нормализация
        normalized = normalize(parsed, api_key=api_key)

        # Шаг 3: Верификация
        verification = verify(normalized, api_key=api_key)

        results.append({
            "_source_file": filename,
            "parsed": parsed,
            "normalized": normalized,
            "verification": verification,
        })

    print(f"\n[PIPELINE] Завершено. Обработано файлов: {len(results)}")
    return results


def pipeline_summary(results: list[dict]) -> dict:
    """Сводная статистика по результатам pipeline."""
    total = len(results)
    errors = sum(1 for r in results if "_pipeline_error" in r)
    ok = sum(1 for r in results if r.get("verification", {}).get("ok") is True)
    fail = total - errors - ok
    avg_score = 0
    scores = [r["verification"]["score"] for r in results if "verification" in r and "score" in r["verification"]]
    if scores:
        avg_score = round(sum(scores) / len(scores), 1)

    return {
        "total": total,
        "ok": ok,
        "fail": fail,
        "errors": errors,
        "avg_score": avg_score,
    }
