import asyncio
import io
import json
import os
import shutil
import subprocess
import tempfile
import zipfile
import traceback
from datetime import datetime
from pathlib import Path

from docx import Document
from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel

from sk_reporter.companies import COMPANIES
from sk_reporter.docx_processing import (
    merge_reports,
    prepare_uploaded_reports,
    rename_results,
    rename_templates,
)
from sk_reporter.paths import templates_dir
from sk_reporter.template_layout import hardcoded_layout

try:
    from sk_reporter.agent.ocr_agent import detect_company, merge_report_into_template
    AGENT_ENABLED = True
    print("[INFO] AI agent connected: gemma4:31b-cloud via Ollama")
except ImportError as e:
    AGENT_ENABLED = False
    print(f"[WARNING] Agent not found: {e}")

_WEBAPP_DIR = Path(__file__).resolve().parent
_HTML_TEMPLATES_DIR = _WEBAPP_DIR / "templates"

app = FastAPI(title="Объединение отчётов СК")
templates = Jinja2Templates(directory=str(_HTML_TEMPLATES_DIR))
app.mount("/static", StaticFiles(directory=_WEBAPP_DIR / "static"), name="static")

WORK_DIR = Path(tempfile.gettempdir()) / "sk_reports_work"
UPLOAD_DIR = WORK_DIR / "uploads"
RESULT_DIR = WORK_DIR / "results"

for d in (WORK_DIR, UPLOAD_DIR, RESULT_DIR):
    d.mkdir(exist_ok=True)

TEMPLATES_DIR = templates_dir()
if not TEMPLATES_DIR.exists():
    raise RuntimeError(f"Папка с болванками не найдена: {TEMPLATES_DIR}")
print(f"[INFO] Templates dir: {TEMPLATES_DIR} ({len(list(TEMPLATES_DIR.glob('*.docx')))} шаблонов)")

for _tpl in ("home.html", "daily.html", "index.html"):
    _tpl_path = _HTML_TEMPLATES_DIR / _tpl
    if not _tpl_path.is_file():
        raise RuntimeError(f"HTML-шаблон не найден: {_tpl_path} — выполните git pull и перезапустите сервер")
print(f"[INFO] UI templates: {_HTML_TEMPLATES_DIR}")
print(f"[INFO] main.py: {Path(__file__).resolve()}")
_git_head = "unknown"
try:
    _git_head = subprocess.check_output(
        ["git", "rev-parse", "--short", "HEAD"],
        cwd=Path(__file__).resolve().parent.parent,
        stderr=subprocess.DEVNULL,
        text=True,
    ).strip()
    print(f"[INFO] git: {_git_head}")
except Exception:
    print("[INFO] git: unknown")


@app.exception_handler(Exception)
async def unhandled_exception_handler(request: Request, exc: Exception):
    traceback.print_exc()
    return JSONResponse(
        status_code=500,
        content={"detail": f"Внутренняя ошибка сервера: {str(exc)}"},
    )


def _do_merge(template_path: str, report_paths: list[str], output_path: str) -> int:
    if not AGENT_ENABLED:
        return merge_reports(template_path, report_paths, output_path)

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
        os.remove(tmp)
        if ok:
            inserted += 1
    return inserted


def find_reports_for_company(company_name: str, keywords: list[str]):
    found = []
    for f in UPLOAD_DIR.iterdir():
        if f.suffix.lower() not in (".docx", ".doc"):
            continue
        if AGENT_ENABLED:
            detected = detect_company(str(f))
            if detected and detected == company_name:
                found.append(f)
        else:
            kw_lower = [k.lower() for k in keywords]
            if any(k in f.name.lower() for k in kw_lower):
                found.append(f)
    return found


def _page_context(request: Request) -> dict:
    return {
        "request": request,
        "agent_enabled": AGENT_ENABLED,
        "git_head": _git_head,
    }


@app.get("/health")
async def health():
    """Проверка версии и наличия UI-шаблонов (для отладки после git pull)."""
    missing = [
        name for name in ("home.html", "daily.html", "index.html")
        if not (_HTML_TEMPLATES_DIR / name).is_file()
    ]
    return {
        "ok": not missing,
        "git_head": _git_head,
        "main_py": str(Path(__file__).resolve()),
        "ui_templates": str(_HTML_TEMPLATES_DIR),
        "missing_templates": missing,
    }


@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse("home.html", _page_context(request))


@app.get("/daily", response_class=HTMLResponse)
async def daily_reports(request: Request):
    return templates.TemplateResponse("daily.html", _page_context(request))


@app.post("/upload/reports")
async def upload_reports(files: list[UploadFile] = File(...)):
    saved = []
    for f in files:
        if not f.filename:
            continue
        dest = UPLOAD_DIR / f.filename
        with open(dest, "wb") as out:
            shutil.copyfileobj(f.file, out)
        saved.append(f.filename)
    return {"uploaded": saved, "count": len(saved)}


@app.get("/api/build")
async def api_build():
    return {"git": _git_head}


@app.get("/files/reports")
async def list_reports():
    files = [f.name for f in UPLOAD_DIR.iterdir() if f.suffix.lower() in (".docx", ".doc")]
    return {"files": sorted(files)}




@app.get("/diagnose/reports", tags=["dev"], include_in_schema=False)
async def diagnose_reports():
    """DEV ONLY: диагностика сетки загруженных отчётов. В UI нет кнопки; только для отладки."""
    from sk_reporter.template_layout import diagnose_document, hardcoded_layout

    layout = hardcoded_layout()
    out = []
    for f in sorted(UPLOAD_DIR.iterdir()):
        if f.suffix.lower() not in (".docx", ".doc"):
            continue
        try:
            doc = Document(os.fspath(f))
            warns = diagnose_document(doc, layout)
            out.append({
                "file": f.name,
                "tables": len(doc.tables),
                "rows": [len(t.rows) for t in doc.tables],
                "images": len(doc.inline_shapes),
                "issues": warns,
                "ok": not warns,
            })
        except Exception as e:
            out.append({"file": f.name, "ok": False, "issues": [str(e)]})
    return {"dev_only": True, "reports": out, "grid_cols": layout["grid_cols"]}

@app.post("/check/descriptions/stream")
async def check_descriptions_stream():
    from sk_reporter.agent.check_agent import check_report
    from sk_reporter.agent.inject_agent import inject_into_docx

    async def event_generator():
        report_files = sorted(UPLOAD_DIR.glob("*.docx"))
        if not report_files:
            yield _sse({"type": "error", "msg": "Отчёты не загружены"})
            return
        errors_count = 0
        promoted_count = 0
        check_results: dict[str, dict] = {}

        # Фаза 1: проверить все файлы
        print(f"[CHECK_STREAM] === фаза 1: check ({len(report_files)} файлов) ===")
        yield _sse({
            "type": "start",
            "msg": "Проверяю загруженные отчёты…",
            "total": len(report_files),
        })
        for file_path in report_files:
            try:
                filename = Path(file_path).name
                yield _sse({"type": "info", "filename": filename, "msg": f"{filename}: проверяю…"})
                await asyncio.sleep(0)
                result = await asyncio.to_thread(check_report, str(file_path))
                has_errors = not result.get("ok", False)
                if has_errors:
                    errors_count += 1
                check_results[filename] = result
                yield _sse({
                    "type": "report",
                    "filename": filename,
                    "msg": f"{filename}: " + ("⚠️ найдены проблемы" if has_errors else "✓ ОК"),
                    "hasErrors": has_errors,
                    "result": result,
                })
                await asyncio.sleep(0)
            except Exception as e:
                yield _sse({"type": "error", "msg": f"Ошибка проверки {Path(file_path).name}: {str(e)}"})

        print("[CHECK_STREAM] === фаза 1 завершена, старт фазы 2: inject ===")
        yield _sse({
            "type": "check_done",
            "msg": f"Проверка завершена ({len(report_files)} файлов). Вставляю правки…",
            "summary": {"total": len(report_files), "errors": errors_count},
        })
        await asyncio.sleep(0)

        # Фаза 2: вставить правки check во все файлы → _исправлен в загрузке
        for file_path in report_files:
            try:
                filename = Path(file_path).name
                result = check_results.get(filename)
                if not result:
                    yield _sse({"type": "error", "msg": f"{filename}: нет результата check — inject пропущен"})
                    continue
                corrected_text = (result.get("report") or "").strip()
                if not corrected_text:
                    yield _sse({"type": "error", "msg": f"{filename}: пустой текст check — inject пропущен"})
                    continue
                yield _sse({"type": "info", "filename": filename, "msg": f"{filename}: вставляю правки…"})
                await asyncio.sleep(0)
                inject_result = await asyncio.to_thread(
                    inject_into_docx, str(file_path), corrected_text, filename
                )
                if inject_result.get("ok"):
                    dl_name = inject_result.get("download_name") or _fixed_download_name(filename)
                    promoted_count += 1
                    yield _sse({
                        "type": "fixed",
                        "filename": filename,
                        "msg": f"{filename}: исправлен и записан в загрузку",
                        "download": f"/download/fixed/{dl_name}",
                        "promoted": True,
                    })
                else:
                    yield _sse({"type": "error", "msg": f'Ошибка inject для {filename}: {inject_result.get("error")}'})
                await asyncio.sleep(0)
            except Exception as e:
                yield _sse({"type": "error", "msg": f"Ошибка inject {Path(file_path).name}: {str(e)}"})

        print(f"[CHECK_STREAM] === фаза 2 завершена ({promoted_count} inject) ===")
        yield _sse({
            "type": "done",
            "summary": {
                "total": len(report_files),
                "errors": errors_count,
                "promoted": promoted_count,
            },
        })

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "Connection": "keep-alive", "X-Accel-Buffering": "no"},
    )


class PrepareBody(BaseModel):
    date: str | None = None  # YYYY-MM-DD из поля «Дата в отчёте»


def _parse_report_date(body: PrepareBody | None) -> str:
    if not body or not body.date:
        raise HTTPException(
            status_code=400,
            detail="Укажите date (YYYY-MM-DD) — поле «Дата в отчёте» в блоке «Макросы (до сборки)»",
        )
    try:
        return datetime.strptime(body.date, "%Y-%m-%d").strftime("%d.%m.%Y")
    except ValueError:
        raise HTTPException(status_code=400, detail="date: формат YYYY-MM-DD")


@app.post("/macro/prepare")
async def macro_prepare(body: PrepareBody | None = None):
    target_date = _parse_report_date(body)
    layout = hardcoded_layout()
    log = prepare_uploaded_reports(str(UPLOAD_DIR), layout, target_date)
    return {
        "log": log,
        "template": "сетка захардкожена",
        "grid_cols": layout["grid_cols"],
        "grid_cols_7": layout.get("grid_cols_7"),
        "date": target_date,
    }


@app.post("/rename/templates")
async def rename_templates_endpoint(body: PrepareBody):
    target_date = _parse_report_date(body)
    log = rename_templates(str(TEMPLATES_DIR), target_date)
    return {"log": log, "date": target_date}


@app.post("/rename/results")
async def rename_results_endpoint(body: PrepareBody):
    target_date = _parse_report_date(body)
    log = rename_results(str(RESULT_DIR), target_date)
    return {"log": log, "date": target_date}


def _sse(data: dict) -> str:
    return f"data: {json.dumps(data, ensure_ascii=False)}\n\n"


def _fixed_download_name(upload_filename: str) -> str:
    return f"{Path(upload_filename).stem}_исправлен.docx"


def _upload_path_for_fixed_download(fixed_download_name: str) -> Path | None:
    """Имя *_исправлен.docx в URL → файл в UPLOAD_DIR с тем же содержимым."""
    marker = "_исправлен.docx"
    if not fixed_download_name.endswith(marker):
        return None
    upload_name = f"{fixed_download_name[: -len(marker)]}.docx"
    path = (UPLOAD_DIR / upload_name).resolve()
    if not str(path).startswith(str(UPLOAD_DIR.resolve())):
        return None
    return path if path.is_file() else None


@app.get("/merge/all/stream")
async def merge_all_stream():
    async def _gen():
        results = []
        errors = []
        yield _sse({
            "type": "start",
            "msg": "Формирую сводные отчёты по компаниям…",
            "total": len(COMPANIES),
        })
        done_count = 0
        for company in COMPANIES:
            name = company.name
            keywords = company.keywords
            stem = company.template_stem
            template = next(
                (
                    TEMPLATES_DIR / f"{stem}{ext}"
                    for ext in (".docx", ".doc")
                    if (TEMPLATES_DIR / f"{stem}{ext}").is_file()
                ),
                None,
            )
            if template is None:
                kw_lower = [k.lower() for k in keywords]
                candidates = sorted(
                    f for f in TEMPLATES_DIR.iterdir()
                    if f.suffix.lower() in (".docx", ".doc")
                    and any(k in f.stem.lower() for k in kw_lower)
                )
                template = candidates[0] if candidates else None
            if not template:
                msg = f"Шаблон для «{name}» не найден — пропущено"
                yield _sse({"type": "warning", "company": name, "msg": msg})
                errors.append(msg)
                done_count += 1
                yield _sse({"type": "progress", "current": done_count, "total": len(COMPANIES), "msg": msg})
                continue
            reports = find_reports_for_company(name, keywords)
            yield _sse({"type": "info", "company": name, "msg": f"Найдено {len(reports)} отчётов"})
            if not reports:
                yield _sse({"type": "info", "company": name, "msg": "Отчёты не найдены, пропускаю"})
                done_count += 1
                yield _sse({"type": "progress", "current": done_count, "total": len(COMPANIES), "msg": f"{name}: нет отчётов"})
                continue
            output_path = RESULT_DIR / f"{name}_merged.docx"
            try:
                inserted = _do_merge(str(template), [str(r) for r in reports], str(output_path))
                results.append({"company": name, "inserted": inserted, "file": f"{name}_merged.docx", "reports_count": len(reports)})
                yield _sse({"type": "success", "company": name, "msg": f"Объединено: {inserted} отчётов → {name}_merged.docx"})
            except Exception as e:
                yield _sse({"type": "error", "company": name, "msg": f"'{name}': {str(e)}"})
                errors.append(str(e))
            finally:
                done_count += 1
                yield _sse({
                    "type": "progress",
                    "current": done_count,
                    "total": len(COMPANIES),
                    "msg": f"Готово: {name}",
                })

        all_matched = set()
        for r in results:
            company = next((c for c in COMPANIES if c.name == r["company"]), None)
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
                    kw_lower = [k.lower() for k in company.keywords]
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
                    yield _sse({"type": "info", "msg": f"Нет шаблона для «{detected}», скопирован: {f.name}"})
                else:
                    unmatched_unknown.append(f.name)
                    yield _sse({"type": "warning", "msg": f"Компания не определена, скопирован: {f.name}"})
            else:
                unmatched_unknown.append(f.name)

        yield _sse({"type": "done", "results": results, "errors": errors, "total_merged": len(results), "unmatched": unmatched, "unmatched_unknown": unmatched_unknown, "ai_agent_active": AGENT_ENABLED})

    return StreamingResponse(_gen(), media_type="text/event-stream", headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


def _zip_files(files: list[Path], arcnames: list[str] | None = None) -> io.BytesIO:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, f in enumerate(files):
            name = arcnames[i] if arcnames else f.name
            zf.write(f, name)
    buf.seek(0)
    return buf


_ZIP_HEADERS = {"Cache-Control": "no-store, no-cache, must-revalidate"}


def _zip_download_response(
    files: list[Path],
    arcnames: list[str] | None,
    disposition: str,
) -> Response:
    """ZIP целиком в теле ответа — без StreamingResponse(BytesIO), чтобы архив не обрезался."""
    return Response(
        content=_zip_files(files, arcnames).getvalue(),
        media_type="application/zip",
        headers={**_ZIP_HEADERS, "Content-Disposition": disposition},
    )


@app.get("/download/all.zip")
async def download_all():
    files = sorted(
        f for f in RESULT_DIR.iterdir() if f.suffix.lower() in (".docx", ".doc")
    )
    if not files:
        raise HTTPException(status_code=404, detail="Нет готовых файлов")
    return _zip_download_response(
        files,
        None,
        "attachment; filename*=UTF-8''%D0%BE%D1%82%D1%87%D1%91%D1%82%D1%8B.zip",
    )


@app.get("/download/fixed/all.zip")
async def download_fixed_all():
    files = sorted(
        f for f in UPLOAD_DIR.iterdir() if f.suffix.lower() in (".docx", ".doc")
    )
    if not files:
        raise HTTPException(status_code=404, detail="Нет загруженных отчётов")
    arcnames = [_fixed_download_name(f.name) for f in files]
    return _zip_download_response(
        files,
        arcnames,
        "attachment; filename*=UTF-8''%D0%B8%D1%81%D0%BF%D1%80%D0%B0%D0%B2%D0%BB%D0%B5%D0%BD%D0%BD%D1%8B%D0%B5.zip",
    )


@app.get("/download/fixed/{filename}")
async def download_fixed(filename: str):
    path = _upload_path_for_fixed_download(filename)
    if path is None:
        raise HTTPException(status_code=404, detail="Файл не найден")
    return FileResponse(
        path=str(path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


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
    for d in (UPLOAD_DIR, RESULT_DIR):
        shutil.rmtree(d)
        d.mkdir()
    return {"ok": True}


@app.post("/switch-leader/stream/{leader}")
async def switch_leader_stream(leader: str):
    from sk_reporter.leader_switch import switch_leader_in_docx

    if leader not in ("aniskov", "mandzhiev"):
        raise HTTPException(status_code=400, detail="leader должен быть 'aniskov' или 'mandzhiev'")

    async def event_generator():
        report_files = sorted(UPLOAD_DIR.glob("*.docx"))
        yield _sse({
            "type": "start",
            "msg": "Переключаю руководителя…",
            "total": len(report_files),
            "leader": leader,
        })
        if not report_files:
            yield _sse({"type": "error", "msg": "Отчёты не загружены"})
            return

        ok_count = 0
        fail_count = 0
        changed_count = 0
        for file_path in report_files:
            filename = file_path.name
            yield _sse({
                "type": "info",
                "filename": filename,
                "msg": f"{filename}: ищу ячейки руководителя…",
            })
            try:
                ok, msg, n = switch_leader_in_docx(str(file_path), leader)
                if ok:
                    ok_count += 1
                    if n:
                        changed_count += 1
                    yield _sse({
                        "type": "file",
                        "filename": filename,
                        "ok": True,
                        "msg": msg,
                        "changes": n,
                    })
                else:
                    fail_count += 1
                    yield _sse({
                        "type": "file",
                        "filename": filename,
                        "ok": False,
                        "msg": msg,
                        "changes": 0,
                    })
            except Exception as e:
                fail_count += 1
                yield _sse({
                    "type": "error",
                    "msg": f"Ошибка {filename}: {e}",
                })

        yield _sse({
            "type": "done",
            "summary": {
                "total": len(report_files),
                "ok": ok_count,
                "failed": fail_count,
                "changed": changed_count,
            },
        })

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
