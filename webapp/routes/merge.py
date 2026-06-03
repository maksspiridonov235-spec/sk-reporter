import shutil

from fastapi import APIRouter
from fastapi.responses import StreamingResponse

from companies import COMPANIES
from config import AGENT_ENABLED, RESULT_DIR, TEMPLATES_DIR, UPLOAD_DIR, detect_company
from helpers import do_merge, find_reports_for_company, sse

router = APIRouter()


@router.get("/merge/all/stream")
async def merge_all_stream():
    async def _gen():
        results = []
        errors = []
        yield sse({
            "type": "start",
            "msg": "Формирую сводные отчёты по компаниям…",
            "total": len(COMPANIES),
        })
        done_count = 0
        for name, keywords in COMPANIES:
            kw_lower = [k.lower() for k in keywords]
            template = next(
                (f for f in TEMPLATES_DIR.iterdir()
                 if any(k in f.name.lower() for k in kw_lower) and f.suffix.lower() in (".docx", ".doc")),
                None
            )
            if not template:
                msg = f"Шаблон для «{name}» не найден — пропущено"
                yield sse({"type": "warning", "company": name, "msg": msg})
                errors.append(msg)
                done_count += 1
                yield sse({"type": "progress", "current": done_count, "total": len(COMPANIES), "msg": msg})
                continue
            reports = find_reports_for_company(name, keywords)
            yield sse({"type": "info", "company": name, "msg": f"Найдено {len(reports)} отчётов"})
            if not reports:
                yield sse({"type": "info", "company": name, "msg": "Отчёты не найдены, пропускаю"})
                done_count += 1
                yield sse({"type": "progress", "current": done_count, "total": len(COMPANIES), "msg": f"{name}: нет отчётов"})
                continue
            output_path = RESULT_DIR / f"{name}_merged.docx"
            try:
                inserted = do_merge(str(template), [str(r) for r in reports], str(output_path))
                results.append({
                    "company": name,
                    "inserted": inserted,
                    "file": f"{name}_merged.docx",
                    "reports_count": len(reports),
                })
                yield sse({"type": "success", "company": name, "msg": f"Объединено: {inserted} отчётов → {name}_merged.docx"})
            except Exception as e:
                yield sse({"type": "error", "company": name, "msg": f"'{name}': {str(e)}"})
                errors.append(str(e))
            finally:
                done_count += 1
                yield sse({
                    "type": "progress",
                    "current": done_count,
                    "total": len(COMPANIES),
                    "msg": f"Готово: {name}",
                })

        all_matched = set()
        for r in results:
            company = next((c for c in COMPANIES if c[0] == r["company"]), None)
            if not company:
                continue
            for f in UPLOAD_DIR.iterdir():
                if f.suffix.lower() not in (".docx", ".doc"):
                    continue
                if AGENT_ENABLED:
                    detected = detect_company(str(f))
                    if detected and detected == r["company"]:
                        all_matched.add(f)
                else:
                    kw_lower = [k.lower() for k in company[1]]
                    if any(k in f.name.lower() for k in kw_lower):
                        all_matched.add(f)

        unmatched = []
        unmatched_unknown = []
        for f in UPLOAD_DIR.iterdir():
            if f.suffix.lower() not in (".docx", ".doc") or f in all_matched:
                continue
            shutil.copy2(f, RESULT_DIR / f.name)
            if AGENT_ENABLED:
                detected = detect_company(str(f))
                if detected:
                    unmatched.append({"file": f.name, "company": detected, "reason": "нет болванки"})
                    yield sse({"type": "info", "msg": f"Нет шаблона для «{detected}», скопирован: {f.name}"})
                else:
                    unmatched_unknown.append(f.name)
                    yield sse({"type": "warning", "msg": f"Компания не определена, скопирован: {f.name}"})
            else:
                unmatched_unknown.append(f.name)

        yield sse({
            "type": "done",
            "results": results,
            "errors": errors,
            "total_merged": len(results),
            "unmatched": unmatched,
            "unmatched_unknown": unmatched_unknown,
            "ai_agent_active": AGENT_ENABLED,
        })

    return StreamingResponse(
        _gen(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )
