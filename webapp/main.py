import os
import io
import json
import shutil
import tempfile
import zipfile
from pathlib import Path
import sys
import asyncio
from typing import Literal
from docx import Document

# Добавляем корень проекта в путь, чтобы видеть agent
sys.path.append(str(Path(__file__).parent.parent))

from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi import Request

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from companies import COMPANIES
from docx_processing import (
    apply_macro_to_file,
    merge_reports,
    rename_files,
    rename_results,
    rename_templates,
)

# Импорт агента
try:
    from agent.ocr_agent import detect_company, merge_report_into_template
    AGENT_ENABLED = True
    print("[INFO] AI agent connected: qwen3.5:cloud via Ollama")
except ImportError as e:
    AGENT_ENABLED = False
    print(f"[WARNING] Agent not found: {e}")

# Импорт агентов
try:
    from agent.report_parser import parse_report, parse_reports_batch
    from agent.normalizer import normalize
    from agent.verifier import verify
    from agent.pipeline import run_pipeline, pipeline_summary
    PARSER_ENABLED = True
    print("[INFO] Claude API agents ready")
except ImportError as e:
    PARSER_ENABLED = False
    print(f"[WARNING] Agents not found: {e}")

app = FastAPI(title="Объединение отчётов СК")
templates = Jinja2Templates(directory="templates")

# Папка для временных файлов сессии (uploads + results)
WORK_DIR = Path(tempfile.gettempdir()) / "sk_reports_work"
WORK_DIR.mkdir(exist_ok=True)

UPLOAD_DIR = WORK_DIR / "uploads"
RESULT_DIR = WORK_DIR / "results"
TEMPLATES_DIR = WORK_DIR / "contractor_templates"

for d in (UPLOAD_DIR, RESULT_DIR, TEMPLATES_DIR):
    d.mkdir(exist_ok=True)

# Автоматически копируем шаблоны из проекта при запуске
PROJECT_TEMPLATES = Path(__file__).parent.parent / "contractor_report" / "болванки (шаблоны не вырезать только копировать)"
if PROJECT_TEMPLATES.exists():
    for template_file in PROJECT_TEMPLATES.glob("*.docx"):
        dest = TEMPLATES_DIR / template_file.name
        shutil.copy2(template_file, dest)
    print(f"[INFO] Loaded {len(list(TEMPLATES_DIR.glob('*.docx')))} templates from project")


# ── Слияние: агент или fallback ────────────────────────────────────────────

def _do_merge(template_path: str, report_paths: list[str], output_path: str, date_mode: Literal["today", "yesterday"] = "today") -> int:
    # Меняем дату в шаблоне перед сборкой
    from docx_processing import replace_date_in_report_line
    template_doc = Document(template_path)
    replace_date_in_report_line(template_doc, date_mode)
    template_doc.save(template_path)

    if AGENT_ENABLED:
        import shutil
        shutil.copy2(template_path, output_path)
        inserted = 0
        for i, rp in enumerate(sorted(report_paths)):
            # Начиная со второго отчёта добавляем разрыв страницы
            tmp = output_path + ".tmp.docx"
            shutil.copy2(output_path, tmp)
            master = Document(tmp)
            if i > 0:
                master.add_page_break()
            master.save(tmp)
            ok = merge_report_into_template(tmp, rp, output_path)
            import os; os.remove(tmp)
            if ok:
                inserted += 1
        return inserted
    else:
        return merge_reports(template_path, report_paths, output_path)


# ── Вспомогательная функция: Умный поиск отчётов ───────────────────────────

def find_reports_for_company(company_name: str, keywords: list[str]):
    found_reports = []

    for f in UPLOAD_DIR.iterdir():
        if f.suffix.lower() not in (".docx", ".doc"):
            continue

        if AGENT_ENABLED:
            detected = detect_company(str(f))
            if detected and detected == company_name:
                found_reports.append(f)
        else:
            kw_lower = [k.lower() for k in keywords]
            if any(k in f.name.lower() for k in kw_lower):
                found_reports.append(f)

    return found_reports


# ── Главная страница ────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    companies = [name for name, _ in COMPANIES]
    return templates.TemplateResponse("index.html", {
        "request": request,
        "companies": companies,
        "agent_enabled": AGENT_ENABLED,
    })


# ── Загрузка файлов отчётов (входящие от инженеров) ────────────────────────

@app.post("/upload/reports")
async def upload_reports(files: list[UploadFile] = File(...)):
    saved = []
    for f in files:
        dest = UPLOAD_DIR / f.filename
        with open(dest, "wb") as out:
            shutil.copyfileobj(f.file, out)
        saved.append(f.filename)
    
    # Если агент включён, можно сразу проанализировать загруженные файлы
    if AGENT_ENABLED:
        print(f"[INFO] Uploaded {len(saved)} files. AI analysis will be performed when generating report.")
    
    return {"uploaded": saved, "count": len(saved)}


# ── Загрузка шаблонов-болванок подрядчиков ─────────────────────────────────

@app.post("/upload/templates")
async def upload_templates(files: list[UploadFile] = File(...)):
    saved = []
    for f in files:
        dest = TEMPLATES_DIR / f.filename
        with open(dest, "wb") as out:
            shutil.copyfileobj(f.file, out)
        saved.append(f.filename)
    return {"uploaded": saved, "count": len(saved)}


@app.post("/sync/templates")
async def sync_templates():
    """Синхронизирует шаблоны из проекта в рабочую директорию."""
    if not PROJECT_TEMPLATES.exists():
        return {"error": "Папка с шаблонами в проекте не найдена"}

    synced = []
    for template_file in PROJECT_TEMPLATES.glob("*.docx"):
        dest = TEMPLATES_DIR / template_file.name
        shutil.copy2(template_file, dest)
        synced.append(template_file.name)

    return {"synced": synced, "count": len(synced)}


# ── Список загруженных файлов ───────────────────────────────────────────────

@app.get("/files/reports")
async def list_reports():
    files = [f.name for f in UPLOAD_DIR.iterdir() if f.suffix.lower() in (".docx", ".doc")]
    return {"files": sorted(files)}


@app.get("/files/templates")
async def list_templates():
    # Синхронизируем с проектом перед возвращением списка
    if PROJECT_TEMPLATES.exists():
        for template_file in PROJECT_TEMPLATES.glob("*.docx"):
            dest = TEMPLATES_DIR / template_file.name
            shutil.copy2(template_file, dest)

    files = [f.name for f in TEMPLATES_DIR.iterdir() if f.suffix.lower() in (".docx", ".doc")]
    return {"files": sorted(files)}


# ── Применить макрос ко всем шаблонам ──────────────────────────────────────

@app.post("/macro/{macro_name}")
async def run_macro(macro_name: str):
    allowed = {
        "HighlightSecondRow_No5991",
        "NewMacros",
        "ReplaceDateInReportLine",
        "ReplaceDateInReportLine2",
    }
    if macro_name not in allowed:
        raise HTTPException(status_code=400, detail="Неизвестный макрос")

    log = []
    for f in UPLOAD_DIR.iterdir():
        if f.suffix.lower() not in (".docx", ".doc"):
            continue
        ok, msg = apply_macro_to_file(str(f), macro_name)
        status = "OK" if ok else "ERR"
        log.append(f"[{status}] {f.name}: {msg}")

    return {"log": log}


# ── Переименовать файлы шаблонов ────────────────────────────────────────────

@app.post("/rename/templates/{mode}")
async def rename_templates_only(mode: str):
    if mode not in ("today", "yesterday"):
        raise HTTPException(status_code=400, detail="mode должен быть today или yesterday")
    log = rename_templates(str(TEMPLATES_DIR), mode)
    return {"log": log}


@app.post("/rename/results/{mode}")
async def rename_results_only(mode: str):
    if mode not in ("today", "yesterday"):
        raise HTTPException(status_code=400, detail="mode должен быть today или yesterday")
    log = rename_results(str(RESULT_DIR), mode)
    return {"log": log}


@app.post("/rename/{mode}")
async def rename(mode: str):
    if mode not in ("today", "yesterday"):
        raise HTTPException(status_code=400, detail="mode должен быть today или yesterday")
    results_log = rename_results(str(RESULT_DIR), mode)
    templates_log = rename_templates(str(TEMPLATES_DIR), mode)
    return {"log": results_log + templates_log}


# ── Сформировать сводный отчёт (все компании) ──────────────────────────────

@app.post("/merge/all")
async def merge_all():
    results = []
    errors = []
    
    print(f"[DEBUG] Starting merge_all(). UPLOAD_DIR: {UPLOAD_DIR}")
    print(f"[DEBUG] Files in UPLOAD_DIR: {list(UPLOAD_DIR.iterdir()) if UPLOAD_DIR.exists() else 'DIR NOT FOUND'}")
    print(f"[DEBUG] Files in TEMPLATES_DIR: {list(TEMPLATES_DIR.iterdir()) if TEMPLATES_DIR.exists() else 'DIR NOT FOUND'}")

    for name, keywords in COMPANIES:
        kw_lower = [k.lower() for k in keywords]

        template = next(
            (f for f in TEMPLATES_DIR.iterdir()
             if any(k in f.name.lower() for k in kw_lower) and f.suffix.lower() in (".docx", ".doc")),
            None
        )
        if not template:
            msg = f"Template for '{name}' not found — skipped"
            print(f"[WARNING] {msg}")
            errors.append(msg)
            continue

        # УМНЫЙ ПОИСК ОТЧЁТОВ
        reports = find_reports_for_company(name, keywords)
        print(f"[DEBUG] Found {len(reports)} reports for '{name}'")
        
        if not reports:
            # Не считаем это ошибкой, просто пропускаем
            print(f"[INFO] No reports found for '{name}', skipping")
            continue

        output_path = RESULT_DIR / f"{name}_merged.docx"
        try:
            inserted = _do_merge(str(template), [str(r) for r in reports], str(output_path))
            results.append({
                "company": name,
                "inserted": inserted,
                "file": f"{name}_merged.docx",
                "reports_count": len(reports)
            })
            print(f"[OK] Merged '{name}': {inserted} reports")
        except Exception as e:
            msg = f"'{name}': {str(e)}"
            print(f"[ERROR] {msg}")
            errors.append(msg)

    # Определяем какие файлы вошли в сборку
    all_matched: set = set()
    for r in results:
        company_name = r["company"]
        company = next((c for c in COMPANIES if c[0] == company_name), None)
        if company:
            for f in UPLOAD_DIR.iterdir():
                if f.suffix.lower() not in (".docx", ".doc"):
                    continue
                if AGENT_ENABLED:
                    detected = detect_company(str(f))
                    if detected and detected == company_name:
                        all_matched.add(f)
                else:
                    kw_lower = [k.lower() for k in company[1]]
                    if any(k in f.name.lower() for k in kw_lower):
                        all_matched.add(f)

    # Файлы которые не попали ни в одну компанию — копируем как есть
    unmatched = []
    unmatched_unknown = []
    for f in UPLOAD_DIR.iterdir():
        if f.suffix.lower() not in (".docx", ".doc"):
            continue
        if f in all_matched:
            continue
        dest = RESULT_DIR / f.name
        shutil.copy2(f, dest)
        # Пробуем определить компанию через AI для лога
        if AGENT_ENABLED:
            detected = detect_company(str(f))
            if detected:
                unmatched.append({"file": f.name, "company": detected, "reason": "нет болванки"})
                print(f"[INFO] Нет болванки для '{detected}', скопирован: {f.name}")
            else:
                unmatched_unknown.append(f.name)
                print(f"[UNKNOWN] Компания не определена, скопирован: {f.name}")
        else:
            unmatched_unknown.append(f.name)

    print(f"[SUMMARY] Results: {len(results)}, Errors: {len(errors)}, Без болванки: {len(unmatched)}, Не распознано: {len(unmatched_unknown)}")
    return {
        "results": results,
        "errors": errors,
        "total_merged": len(results),
        "unmatched": unmatched,
        "unmatched_unknown": unmatched_unknown,
        "ai_agent_active": AGENT_ENABLED
    }


# ── Сформировать все отчёты (с SSE потоком логов) ────────────────────────────

@app.get("/merge/all/stream")
async def merge_all_stream():
    """Сформировать все отчёты с потоковым выводом логов через SSE."""

    async def _gen():
        results = []
        errors = []

        yield _sse({"type": "start", "msg": "Начинаю формирование отчётов..."})

        for name, keywords in COMPANIES:
            kw_lower = [k.lower() for k in keywords]

            template = next(
                (f for f in TEMPLATES_DIR.iterdir()
                 if any(k in f.name.lower() for k in kw_lower) and f.suffix.lower() in (".docx", ".doc")),
                None
            )
            if not template:
                msg = f"Шаблон для «{name}» не найден — пропущено"
                yield _sse({"type": "warning", "company": name, "msg": msg})
                errors.append(msg)
                continue

            # УМНЫЙ ПОИСК ОТЧЁТОВ
            reports = find_reports_for_company(name, keywords)
            yield _sse({"type": "info", "company": name, "msg": f"Найдено {len(reports)} отчётов"})

            if not reports:
                yield _sse({"type": "info", "company": name, "msg": "Отчёты не найдены, пропускаю"})
                continue

            output_path = RESULT_DIR / f"{name}_merged.docx"
            try:
                inserted = _do_merge(str(template), [str(r) for r in reports], str(output_path))
                results.append({
                    "company": name,
                    "inserted": inserted,
                    "file": f"{name}_merged.docx",
                    "reports_count": len(reports)
                })
                yield _sse({"type": "success", "company": name, "msg": f"Объединено: {inserted} отчётов → {name}_merged.docx"})
            except Exception as e:
                msg = f"'{name}': {str(e)}"
                yield _sse({"type": "error", "company": name, "msg": msg})
                errors.append(msg)

        # Определяем какие файлы вошли в сборку
        all_matched: set = set()
        for r in results:
            company_name = r["company"]
            company = next((c for c in COMPANIES if c[0] == company_name), None)
            if company:
                for f in UPLOAD_DIR.iterdir():
                    if f.suffix.lower() not in (".docx", ".doc"):
                        continue
                    if AGENT_ENABLED:
                        detected = detect_company(str(f))
                        if detected and detected == company_name:
                            all_matched.add(f)
                    else:
                        kw_lower = [k.lower() for k in company[1]]
                        if any(k in f.name.lower() for k in kw_lower):
                            all_matched.add(f)

        # Файлы которые не попали ни в одну компанию — копируем как есть
        unmatched = []
        unmatched_unknown = []
        for f in UPLOAD_DIR.iterdir():
            if f.suffix.lower() not in (".docx", ".doc"):
                continue
            if f in all_matched:
                continue
            dest = RESULT_DIR / f.name
            shutil.copy2(f, dest)
            # Пробуем определить компанию через AI для лога
            if AGENT_ENABLED:
                detected = detect_company(str(f))
                if detected:
                    unmatched.append({"file": f.name, "company": detected, "reason": "нет болванки"})
                    yield _sse({"type": "info", "msg": f"Нет шаблона для «{detected}», скопирован: {f.name}"})
                else:
                    unmatched_unknown.append(f.name)
                    yield _sse({"type": "warning", "msg": f"Компания не определена, скопирован: {f.name}"})
            else:
                unmatched_unknown.append(f.name)

        yield _sse({
            "type": "done",
            "results": results,
            "errors": errors,
            "total_merged": len(results),
            "unmatched": unmatched,
            "unmatched_unknown": unmatched_unknown,
            "ai_agent_active": AGENT_ENABLED
        })

    return StreamingResponse(
        _gen(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


# ── Объединить отчёты для одной компании ───────────────────────────────────

@app.post("/merge/{company_name}")
async def merge_one(company_name: str):
    company_name = company_name.replace("%20", " ").replace("+", " ")
    company = next((c for c in COMPANIES if c[0] == company_name), None)
    if not company:
        raise HTTPException(status_code=404, detail="Компания не найдена")

    name, keywords = company

    template = next(
        (f for f in TEMPLATES_DIR.iterdir()
         if any(k in f.name.lower() for k in [kw.lower() for kw in keywords])
         and f.suffix.lower() in (".docx", ".doc")),
        None
    )
    if not template:
        raise HTTPException(status_code=404, detail=f"Шаблон для «{name}» не найден")

    reports = find_reports_for_company(name, keywords)
    if not reports:
        raise HTTPException(status_code=404, detail=f"Отчёты для «{name}» не найдены")

    output_path = RESULT_DIR / f"{name}_merged.docx"
    inserted = _do_merge(str(template), [str(r) for r in reports], str(output_path))

    return {
        "company": name,
        "template": template.name,
        "reports_found": len(reports),
        "inserted": inserted,
        "result": f"{name}_merged.docx",
    }


# ── Скачать все результаты одним ZIP ───────────────────────────────────────

@app.get("/download/all.zip")
async def download_all():
    files = [f for f in RESULT_DIR.iterdir() if f.suffix.lower() in (".docx", ".doc")]
    if not files:
        raise HTTPException(status_code=404, detail="Нет готовых файлов")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(f, f.name)
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename*=UTF-8''%D0%BE%D1%82%D1%87%D1%91%D1%82%D1%8B.zip"},
    )


# ── Скачать результат ───────────────────────────────────────────────────────

@app.get("/download/{filename}")
async def download(filename: str):
    path = (RESULT_DIR / filename).resolve()
    if not str(path).startswith(str(RESULT_DIR.resolve())):
        raise HTTPException(status_code=400, detail="Недопустимый путь")
    if not path.exists():
        raise HTTPException(status_code=404, detail="Файл не найден")
    return FileResponse(
        path=str(path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# ── Список готовых файлов ───────────────────────────────────────────────────

@app.get("/results")
async def list_results():
    files = [f.name for f in RESULT_DIR.iterdir() if f.suffix.lower() in (".docx", ".doc")]
    return {"files": sorted(files)}


# ── Страница агентов ────────────────────────────────────────────────────────

@app.get("/agents", response_class=HTMLResponse)
async def agents_page(request: Request):
    files = sorted(f.name for f in UPLOAD_DIR.iterdir() if f.suffix.lower() in (".docx", ".doc"))
    return templates.TemplateResponse("agents.html", {
        "request": request,
        "files": files,
        "agent_ready": PARSER_ENABLED and bool(os.environ.get("ANTHROPIC_API_KEY")),
    })


# ── SSE: потоковый запуск агентов ───────────────────────────────────────────

def _sse(data: dict) -> str:
    return f"data: {json.dumps(data, ensure_ascii=False)}\n\n"


def _check_agents():
    if not PARSER_ENABLED:
        return "Агенты недоступны (import error)"
    if not os.environ.get("ANTHROPIC_API_KEY"):
        return "ANTHROPIC_API_KEY не задан"
    return None


async def _stream_agent(agent_name: str, files: list[str], agent_fn):
    """Универсальный SSE-генератор для одного агента."""
    err = _check_agents()
    if err:
        yield _sse({"type": "error", "msg": err})
        return

    if not files:
        yield _sse({"type": "error", "msg": "Нет загруженных отчётов"})
        return

    yield _sse({"type": "start", "agent": agent_name, "total": len(files)})

    for i, fp in enumerate(files):
        filename = Path(fp).name
        yield _sse({"type": "progress", "file": filename, "index": i + 1, "total": len(files)})
        loop = asyncio.get_event_loop()
        result = await loop.run_in_executor(None, agent_fn, fp)
        yield _sse({"type": "result", "file": filename, "data": result})

    yield _sse({"type": "done", "agent": agent_name})


@app.get("/agents/stream/parse")
async def stream_parse():
    files = sorted(str(f) for f in UPLOAD_DIR.iterdir() if f.suffix.lower() in (".docx", ".doc"))
    return StreamingResponse(
        _stream_agent("Парсер", files, parse_report),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.get("/agents/stream/normalize")
async def stream_normalize(file: str):
    """Нормализует один файл (принимает имя файла, читает из кэша сессии)."""
    err = _check_agents()
    if err:
        async def _err():
            yield _sse({"type": "error", "msg": err})
        return StreamingResponse(_err(), media_type="text/event-stream")

    async def _gen():
        yield _sse({"type": "start", "agent": "Нормализатор", "total": 1})
        # parsed передаётся через тело — здесь упрощённо через query
        yield _sse({"type": "info", "msg": f"Используй /agents/stream/pipeline для полного цикла"})
        yield _sse({"type": "done", "agent": "Нормализатор"})

    return StreamingResponse(_gen(), media_type="text/event-stream")


@app.get("/agents/stream/pipeline")
async def stream_pipeline():
    """Полный pipeline файл за файлом через SSE."""
    files = sorted(str(f) for f in UPLOAD_DIR.iterdir() if f.suffix.lower() in (".docx", ".doc"))

    async def _gen():
        err = _check_agents()
        if err:
            yield _sse({"type": "error", "msg": err})
            return
        if not files:
            yield _sse({"type": "error", "msg": "Нет загруженных отчётов"})
            return

        yield _sse({"type": "start", "agent": "Pipeline", "total": len(files)})
        loop = asyncio.get_event_loop()
        ok_count = 0
        fail_count = 0

        for i, fp in enumerate(files):
            filename = Path(fp).name

            yield _sse({"type": "progress", "file": filename, "index": i + 1,
                        "total": len(files), "step": "parse"})
            parsed = await loop.run_in_executor(None, parse_report, fp)
            if not parsed:
                yield _sse({"type": "result", "file": filename,
                            "error": "Парсер не смог извлечь данные"})
                fail_count += 1
                continue

            yield _sse({"type": "progress", "file": filename, "index": i + 1,
                        "total": len(files), "step": "normalize"})
            normalized = await loop.run_in_executor(None, normalize, parsed)

            yield _sse({"type": "progress", "file": filename, "index": i + 1,
                        "total": len(files), "step": "verify"})
            verification = await loop.run_in_executor(None, verify, normalized)

            if verification.get("ok"):
                ok_count += 1
            else:
                fail_count += 1

            yield _sse({"type": "result", "file": filename,
                        "parsed": parsed, "normalized": normalized,
                        "verification": verification})

        yield _sse({"type": "done", "agent": "Pipeline",
                    "ok": ok_count, "fail": fail_count, "total": len(files)})

    return StreamingResponse(
        _gen(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.post("/agents/normalize-one")
async def normalize_one(data: dict):
    """Нормализует один JSON (из кэша браузера)."""
    err = _check_agents()
    if err:
        raise HTTPException(status_code=503, detail=err)
    loop = asyncio.get_event_loop()
    result = await loop.run_in_executor(None, normalize, data)
    return result


@app.post("/agents/verify-one")
async def verify_one(data: dict):
    """Верифицирует один JSON (из кэша браузера)."""
    err = _check_agents()
    if err:
        raise HTTPException(status_code=503, detail=err)
    loop = asyncio.get_event_loop()
    result = await loop.run_in_executor(None, verify, data)
    return result


# ── Очистить загруженные файлы ──────────────────────────────────────────────

@app.delete("/clear/reports")
async def clear_reports():
    shutil.rmtree(UPLOAD_DIR)
    UPLOAD_DIR.mkdir()
    return {"ok": True}


@app.delete("/clear/results")
async def clear_results():
    shutil.rmtree(RESULT_DIR)
    RESULT_DIR.mkdir()
    return {"ok": True}


@app.delete("/clear/all")
async def clear_all():
    for d in (UPLOAD_DIR, RESULT_DIR, TEMPLATES_DIR):
        shutil.rmtree(d)
        d.mkdir()
    return {"ok": True}
