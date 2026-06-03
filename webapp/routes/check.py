from pathlib import Path

from fastapi import APIRouter
from fastapi.responses import StreamingResponse

from config import UPLOAD_DIR
from helpers import sse

router = APIRouter()


@router.post("/check/descriptions/stream")
async def check_descriptions_stream():
    from agent.check_agent import check_report
    from agent.inject_agent import inject_into_docx

    async def event_generator():
        report_files = sorted(UPLOAD_DIR.glob("*.docx"))
        yield sse({
            "type": "start",
            "msg": "Проверяю загруженные отчёты…",
            "total": len(report_files),
        })
        if not report_files:
            yield sse({"type": "error", "msg": "Отчёты не загружены"})
            return
        errors_count = 0
        for file_path in report_files:
            try:
                filename = Path(file_path).name
                yield sse({"type": "info", "filename": filename, "msg": f"{filename}: проверяю..."})
                result = check_report(str(file_path))
                has_errors = not result.get("ok", False)
                if has_errors:
                    errors_count += 1
                yield sse({
                    "type": "report",
                    "filename": filename,
                    "msg": f"{filename}: " + ("⚠️ найдены проблемы" if has_errors else "✓ ОК"),
                    "hasErrors": has_errors,
                    "result": result,
                })
                corrected_text = result.get("report", "")
                if corrected_text:
                    yield sse({"type": "info", "filename": filename, "msg": f"{filename}: вставляю правки в текст документа…"})
                    inject_result = inject_into_docx(str(file_path), corrected_text, filename)
                    if inject_result.get("ok"):
                        dl_name = Path(inject_result["docx_path"]).name
                        yield sse({
                            "type": "fixed",
                            "filename": filename,
                            "msg": f"{filename}: исправлен → {dl_name}",
                            "download": f"/download/fixed/{dl_name}",
                        })
                    else:
                        yield sse({"type": "error", "msg": f'Ошибка inject для {filename}: {inject_result.get("error")}'})
            except Exception as e:
                yield sse({"type": "error", "msg": f"Ошибка проверки {Path(file_path).name}: {str(e)}"})
        yield sse({"type": "done", "summary": {"total": len(report_files), "errors": errors_count}})

    return StreamingResponse(event_generator(), media_type="text/event-stream")
