import os
import shutil
import tempfile
from pathlib import Path
import sys

# Добавляем корень проекта в путь, чтобы видеть agent
sys.path.append(str(Path(__file__).parent.parent))

from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi import Request

from docx_processing import (
    COMPANIES,
    apply_macro_to_file,
    merge_reports,
    rename_files,
)

# Импорт агента для умного поиска компаний
try:
    from agent.ocr_agent import detect_company_hybrid
    AGENT_ENABLED = True
    print("[INFO] AI agent connected: report analysis via Ollama")
except ImportError as e:
    AGENT_ENABLED = False
    print(f"[WARNING] Agent not found, using keyword search only: {e}")

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


# ── Вспомогательная функция: Умный поиск отчётов ───────────────────────────

def find_reports_for_company(company_name: str, keywords: list[str]):
    """
    Ищет отчёты для компании. Сначала по ключевым словам в имени,
    если не найдено — использует AI-агент для анализа содержимого.
    """
    found_reports = []
    
    for f in UPLOAD_DIR.iterdir():
        if f.suffix.lower() not in (".docx", ".doc"):
            continue
        
        # 1. Быстрая проверка по имени файла
        kw_lower = [k.lower() for k in keywords]
        if any(k in f.name.lower() for k in kw_lower):
            found_reports.append(f)
            continue
        
        # 2. Если по имени не подошло — используем AI (если включён)
        if AGENT_ENABLED:
            detected = detect_company_hybrid(str(f))
            if detected and detected == company_name:
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


# ── Список загруженных файлов ───────────────────────────────────────────────

@app.get("/files/reports")
async def list_reports():
    files = [f.name for f in UPLOAD_DIR.iterdir() if f.suffix.lower() in (".docx", ".doc")]
    return {"files": sorted(files)}


@app.get("/files/templates")
async def list_templates():
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
    for f in TEMPLATES_DIR.iterdir():
        if f.suffix.lower() not in (".docx", ".doc"):
            continue
        ok, msg = apply_macro_to_file(str(f), macro_name)
        status = "OK" if ok else "ERR"
        log.append(f"[{status}] {f.name}: {msg}")

    return {"log": log}


# ── Переименовать файлы шаблонов ────────────────────────────────────────────

@app.post("/rename/{mode}")
async def rename(mode: str):
    if mode not in ("today", "yesterday"):
        raise HTTPException(status_code=400, detail="mode должен быть today или yesterday")
    log = rename_files(str(TEMPLATES_DIR), mode)  # type: ignore[arg-type]
    return {"log": log}


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
            inserted = merge_reports(str(template), [str(r) for r in reports], str(output_path))
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

    print(f"[SUMMARY] Results: {len(results)}, Errors: {len(errors)}")
    return {
        "results": results,
        "errors": errors,
        "total_merged": len(results),
        "ai_agent_active": AGENT_ENABLED
    }


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
    inserted = merge_reports(str(template), [str(r) for r in reports], str(output_path))

    return {
        "company": name,
        "template": template.name,
        "reports_found": len(reports),
        "inserted": inserted,
        "result": f"{name}_merged.docx",
    }


# ── Скачать результат ───────────────────────────────────────────────────────

@app.get("/download/{filename}")
async def download(filename: str):
    # Защита от path traversal
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
