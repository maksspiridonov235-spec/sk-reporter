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

sys.path.append(str(Path(__file__).parent.parent))

from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from fastapi import Request

sys.path.insert(0, str(Path(__file__).parent.parent))
from companies import COMPANIES
from docx_processing import (
    apply_macro_to_file,
    merge_reports,
    rename_files,
    rename_results,
    rename_templates,
)

# Импорт Ollama-агента (локальный AI для определения компании)
try:
    from agent.ocr_agent import detect_company, merge_report_into_template
    AGENT_ENABLED = True
    print("[INFO] AI agent connected: qwen3.5:cloud via Ollama")
except ImportError as e:
    AGENT_ENABLED = False
    print(f"[WARNING] Agent not found: {e}")

app = FastAPI(title="Объединение отчётов СК")
templates = Jinja2Templates(directory="templates")

WORK_DIR = Path(tempfile.gettempdir()) / "sk_reports_work"
WORK_DIR.mkdir(exist_ok=True)

UPLOAD_DIR = WORK_DIR / "uploads"
RESULT_DIR = WORK_DIR / "results"
TEMPLATES_DIR = WORK_DIR / "contractor_templates"

for d in (UPLOAD_DIR, RESULT_DIR, TEMPLATES_DIR):
    d.mkdir(exist_ok=True)

PROJECT_TEMPLATES = Path(__file__).parent.parent / "contractor_report" / "болванки (шаблоны не вырезать только копировать)"
if PROJECT_TEMPLATES.exists():
    for template_file in PROJECT_TEMPLATES.glob("*.docx"):
        dest = TEMPLATES_DIR / template_file.name
        shutil.copy2(template_file, dest)
    print(f"[INFO] Loaded {len(list(TEMPLATES_DIR.glob('*.docx')))} templates from project")

def _do_merge(template_path: str, report_paths: list[str], output_path: str) -> int:
    if AGENT_ENABLED:
        import shutil
        shutil.copy2(template_path, output_path)
        inserted = 0
        for i, rp in enumerate(sorted(report_paths)):
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

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    companies = [name for name, _ in COMPANIES]
    return templates.TemplateResponse("index.html", {
        "request": request,
        "companies": companies,
        "agent_enabled": AGENT_ENABLED,
    })

@app.post("/upload/reports")
async def upload_reports(files: list[UploadFile] = File(...)):
    saved = []
    for f in files:
        dest = UPLOAD_DIR / f.filename
        with open(dest, "wb") as out:
            shutil.copyfileobj(f.file, out)
        saved.append(f.filename)
    if AGENT_ENABLED:
        print(f"[INFO] Uploaded {len(saved)} files. AI analysis will be performed when generating report.")
    return {"uploaded": saved, "count": len(saved)}

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
    if not PROJECT_TEMPLATES.exists():
        return {"error": "Папка с шаблонами в проекте не найдена"}
    synced = []
    for template_file in PROJECT_TEMPLATES.glob("*.docx"):
        dest = TEMPLATES_DIR / template_file.name
        shutil.copy2(template_file, dest)
        synced.append(template_file.name)
    return {"synced": synced, "count": len(synced)}

@app.get("/files/reports")
async def list_reports():
    files = [f.name for f in UPLOAD_DIR.iterdir() if f.suffix.lower() in (".docx", ".doc")]
    return {"files": sorted(files)}

@app.get("/files/templates")
async def list_templates():
    if PROJECT_TEMPLATES.exists():
        for template_file in PROJECT_TEMPLATES.glob("*.docx"):
            dest = TEMPLATES_DIR / template_file.name
            shutil.copy2(template_file, dest)
    files = [f.name for f in TEMPLATES_DIR.iterdir() if f.suffix.lower() in (".docx", ".doc")]
    return {"files": sorted(files)}

@app.post("/check/descriptions/stream")
async def check_descriptions_stream():
    from agent.check_agent import check_report
    async def event_generator():
        yield f"data: {json.dumps({'type': 'start', 'msg': 'Начинаю проверку отчётов...'})}\n\n"
        report_files = list(UPLOAD_DIR.glob("*.docx"))
        if not report_files:
            yield f"data: {json.dumps({'type': 'error', 'msg': 'Отчёты не загружены'})}\n\n"
            return
        total = len(report_files)
        errors_count = 0
        for file_path in report_files:
            try:
                result = check_report(str(file_path))
                has_errors = not result.get("ok", False)
                if has_errors:
                    errors_count += 1
                yield f"data: {json.dumps({'type': 'report', 'filename': Path(file_path).name, 'msg': f'{Path(file_path).name}: ' + ('⚠️ найдены проблемы' if has_errors else '✓ ОК'), 'hasErrors': has_errors, 'result': result})}\n\n"
            except Exception as e:
                yield f"data: {json.dumps({'type': 'error', 'msg': f'Ошибка проверки {Path(file_path).name}: {str(e)}'})}\n\n"
        yield f"data: {json.dumps({'type': 'done', 'summary': {'total': total, 'errors': errors_count}})}\n\n"
    return StreamingResponse(event_generator(), media_type="text/event-stream")

@app.post("/macro/{macro_name}")
async def run_macro(macro_name: str):
    allowed = {"HighlightSecondRow_No5991", "NewMacros", "ReplaceDateInReportLine", "ReplaceDateInReportLine2"}
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

@app.post("/rename/templates/{mode}")
async def rename_templates_only(mode: str):
    if mode not in ("today", "yesterday"):
        raise HTTPException(status_code=400, detail="mode должен быть today или yesterday")
    if PROJECT_TEMPLATES.exists():
        for template_file in PROJECT_TEMPLATES.glob("*.docx"):
            dest = TEMPLATES_DIR / template_file.name
            shutil.copy2(template_file, dest)
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

@app.post("/merge/all")
async def merge_all():
    results = []
    errors = []
    for name, keywords in COMPANIES:
        kw_lower = [k.lower() for k in keywords]
        template = next(
            (f for f in TEMPLATES_DIR.iterdir()
             if any(k in f.name.lower() for k in kw_lower) and f.suffix.lower() in (".docx", ".doc")),
            None
        )
        if not template:
            errors.append(f"Template for '{name}' not found — skipped")
            continue
        reports = find_reports_for_company(name, keywords)
        if not reports:
            continue
        output_path = RESULT_DIR / f"{name}_merged.docx"
        try:
            inserted = _do_merge(str(template), [str(r) for r in reports], str(output_path))
            results.append({"company": name, "inserted": inserted, "file": f"{name}_merged.docx", "reports_count": len(reports)})
        except Exception as e:
            errors.append(f"'{name}': {str(e)}")
    all_matched = set()
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
    unmatched = []
    unmatched_unknown = []
    for f in UPLOAD_DIR.iterdir():
        if f.suffix.lower() not in (".docx", ".doc"):
            continue
        if f in all_matched:
            continue
        dest = RESULT_DIR / f.name
        shutil.copy2(f, dest)
        if AGENT_ENABLED:
            detected = detect_company(str(f))
            if detected:
                unmatched.append({"file": f.name, "company": detected, "reason": "нет болванки"})
            else:
                unmatched_unknown.append(f.name)
        else:
            unmatched_unknown.append(f.name)
    return {"results": results, "errors": errors, "total_merged": len(results), "unmatched": unmatched, "unmatched_unknown": unmatched_unknown, "ai_agent_active": AGENT_ENABLED}

def _sse(data: dict) -> str:
    return f"data: {json.dumps(data, ensure_ascii=False)}\n\n"

@app.get("/merge/all/stream")
async def merge_all_stream():
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
            reports = find_reports_for_company(name, keywords)
            yield _sse({"type": "info", "company": name, "msg": f"Найдено {len(reports)} отчётов"})
            if not reports:
                yield _sse({"type": "info", "company": name, "msg": "Отчёты не найдены, пропускаю"})
                continue
            output_path = RESULT_DIR / f"{name}_merged.docx"
            try:
                inserted = _do_merge(str(template), [str(r) for r in reports], str(output_path))
                results.append({"company": name, "inserted": inserted, "file": f"{name}_merged.docx", "reports_count": len(reports)})
                yield _sse({"type": "success", "company": name, "msg": f"Объединено: {inserted} отчётов → {name}_merged.docx"})
            except Exception as e:
                yield _sse({"type": "error", "company": name, "msg": f"'{name}': {str(e)}"})
                errors.append(str(e))
        all_matched = set()
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
        unmatched = []
        unmatched_unknown = []
        for f in UPLOAD_DIR.iterdir():
            if f.suffix.lower() not in (".docx", ".doc"):
                continue
            if f in all_matched:
                continue
            dest = RESULT_DIR / f.name
            shutil.copy2(f, dest)
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
        yield _sse({"type": "done", "results": results, "errors": errors, "total_merged": len(results), "unmatched": unmatched, "unmatched_unknown": unmatched_unknown, "ai_agent_active": AGENT_ENABLED})
    return StreamingResponse(_gen(), media_type="text/event-stream", headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})

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
    return {"company": name, "template": template.name, "reports_found": len(reports), "inserted": inserted, "result": f"{name}_merged.docx"}

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
    return StreamingResponse(buf, media_type="application/zip", headers={"Content-Disposition": "attachment; filename*=UTF-8''%D0%BE%D1%82%D1%87%D1%91%D1%82%D1%8B.zip"})

@app.get("/download/{filename}")
async def download(filename: str):
    path = (RESULT_DIR / filename).resolve()
    if not str(path).startswith(str(RESULT_DIR.resolve())):
        raise HTTPException(status_code=400, detail="Недопустимый путь")
    if not path.exists():
        raise HTTPException(status_code=404, detail="Файл не найден")
    return FileResponse(path=str(path), filename=filename, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

@app.get("/results")
async def list_results():
    files = [f.name for f in RESULT_DIR.iterdir() if f.suffix.lower() in (".docx", ".doc")]
    return {"files": sorted(files)}

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

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)

# ── Переключение руководителя ───────────────────────────────────────────────

@app.post("/switch-leader/{leader}")
async def switch_leader_endpoint(leader: str):
    """Переключает руководителя в загруженном отчёте."""
    if leader not in ("aniskov", "mandzhiev"):
        raise HTTPException(status_code=400, detail="leader должен быть 'aniskov' или 'mandzhiev'")
    
    from agent.leader_switcher import switch_leader
    
    report_files = list(UPLOAD_DIR.glob("*.docx"))
    if not report_files:
        raise HTTPException(status_code=404, detail="Отчёты не загружены")
    
    # Берём первый файл (или можно передать имя файла)
    filepath = str(report_files[0])
    filename = report_files[0].name
    
    ok, msg = switch_leader(filepath, leader)
    
    if not ok:
        raise HTTPException(status_code=500, detail=msg)
    
    return {"ok": True, "message": msg, "file": filename}

@app.get("/detect-leader")
async def detect_leader():
    """Определяет текущего руководителя в загруженном отчёте."""
    from agent.leader_switcher import detect_current_leader
    
    report_files = list(UPLOAD_DIR.glob("*.docx"))
    if not report_files:
        return {"leader": "unknown", "message": "Отчёты не загружены"}
    
    filepath = str(report_files[0])
    leader = detect_current_leader(filepath)
    
    names = {
        "aniskov": "Аниськов Владимир Иванович",
        "mandzhiev": "Манджиев Игорь Александрович",
        "unknown": "Не определён"
    }
    
    return {"leader": leader, "name": names.get(leader, "Неизвестно")}
