import asyncio
import io
import json
import os
import shutil
import subprocess
import tempfile
import zipfile
import time
import traceback
from datetime import datetime
from pathlib import Path

from docx import Document
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel, Field

from sk_reporter.companies import COMPANIES, GEODESY_COMPANY
from sk_reporter.docx_processing import (
    merge_reports,
    prepare_uploaded_reports,
    remove_editing_restrictions,
    rename_results,
    rename_templates,
)
from sk_reporter.paths import templates_dir
from sk_reporter.prescriptions.normative_store import normative_store_status
from sk_reporter.template_layout import hardcoded_layout

_nd_cfg = normative_store_status()
print(
    "[INFO] Normative DB: "
    f"documents={_nd_cfg['documents_count']}, "
    f"manifest={'yes' if _nd_cfg['manifest_exists'] else 'no'}"
)
if _nd_cfg["documents_count"] == 0:
    print(
        "[WARN] Normative DB: пусто — добавьте записи в data/normative/manifest.yaml "
        "(см. data/normative/README.md)"
    )

try:
    from sk_reporter.db.config import database_enabled
    from sk_reporter.personnel_db import db_status

    if database_enabled():
        _db_st = db_status()
        print(
            "[INFO] PostgreSQL: "
            f"ok={_db_st.get('ok')}, personnel={_db_st.get('count', 0)}"
        )
        if not _db_st.get("ok"):
            print(f"[WARN] PostgreSQL: {_db_st.get('error')}")
        elif _db_st.get("count", 0) == 0:
            print(
                "[WARN] PostgreSQL: справочник сотрудников пуст — "
                "загрузите Excel на /planning/personnel"
            )
        from sk_reporter.otkk_db import (
            db_status as otkk_db_status,
            purge_empty_otkk_cards,
            seed_otkk1,
            seed_otkk2,
        )

        _purged = purge_empty_otkk_cards()
        if _purged:
            print(f"[INFO] PostgreSQL ОТКК: удалено пустых записей (без content): {_purged}")
        for _seed_fn, _overwrite in ((seed_otkk1, False), (seed_otkk2, True)):
            _seed = _seed_fn(overwrite=_overwrite)
            if _seed.get("seeded"):
                print(
                    f"[INFO] PostgreSQL ОТКК: залит эталон {_seed.get('id')} "
                    f"({_seed.get('rows')} пунктов)"
                )
        _otkk_st = otkk_db_status()
        print(
            "[INFO] PostgreSQL ОТКК: "
            f"ok={_otkk_st.get('ok')}, cards={_otkk_st.get('with_content', 0)}"
        )
        if _otkk_st.get("ok") and _otkk_st.get("with_content", 0) == 0:
            print(
                "[WARN] PostgreSQL: в базе нет ОТКК — перезапустите сервис после git pull"
            )
        from sk_reporter.contractor_db import db_status as contractor_db_status, seed_evrakor

        _ev = seed_evrakor()
        if _ev.get("seeded"):
            print(f"[INFO] PostgreSQL подрядчики: залит {_ev.get('name')} ({_ev.get('id')})")
        _ctr_st = contractor_db_status()
        print(
            "[INFO] PostgreSQL подрядчики: "
            f"ok={_ctr_st.get('ok')}, count={_ctr_st.get('count', 0)}"
        )
        from sk_reporter.project_db import db_status as project_db_status, seed_projects_pilots

        _pr = seed_projects_pilots(overwrite=True)
        for pid in _pr.get("seeded") or []:
            print(f"[INFO] PostgreSQL проекты: залит {pid}")
        _proj_st = project_db_status()
        print(
            "[INFO] PostgreSQL проекты: "
            f"ok={_proj_st.get('ok')}, cards={_proj_st.get('with_content', 0)}"
        )
        from sk_reporter.position_db import db_status as position_db_status, seed_positions_from_json

        _pos = seed_positions_from_json()
        if _pos.get("seeded"):
            print(f"[INFO] PostgreSQL должности: залито {_pos.get('count', 0)} записей")
        _pos_st = position_db_status()
        print(
            "[INFO] PostgreSQL должности: "
            f"ok={_pos_st.get('ok')}, count={_pos_st.get('count', 0)}"
        )
    else:
        print("[WARN] PostgreSQL: DATABASE_URL не задан — раздел «Сотрудники» недоступен")
except Exception as _db_err:
    print(f"[WARN] PostgreSQL: {_db_err}")

try:
    from sk_reporter.agent.ocr_agent import detect_company
    from sk_reporter.llm_client import llm_status, ping_llm

    AGENT_ENABLED = True
    _llm = llm_status()
    _llm_ok, _llm_err = ping_llm()
    print(
        "[INFO] Ollama: "
        f"mode={_llm['mode']}, host={_llm['host']}, model={_llm['model']}, "
        f"api_key={'yes' if _llm['api_key_set'] else 'no'}, "
        f"ping={'ok' if _llm_ok else 'fail'}"
    )
    if not _llm_ok:
        print(f"[WARN] Ollama недоступна: {_llm_err}")
        if _llm["mode"] == "local":
            print(
                "[WARN] Офис: запустите Ollama. Облако (Relax Dev): задайте "
                "OLLAMA_API_KEY и OLLAMA_HOST=https://ollama.com"
            )
except ImportError as e:
    AGENT_ENABLED = False
    print(f"[WARNING] Agent not found: {e}")

_WEBAPP_DIR = Path(__file__).resolve().parent
_HTML_TEMPLATES_DIR = _WEBAPP_DIR / "templates"
_APP_UI_BUILD = "home+reporting+daily+deployment+weekly+weekly-photos+prescriptions+planning+planning-sections+engineer-hub+engineer"

_PLANNING_SECTIONS = {
    "personnel": "Сотрудники",
    "contractors": "Подрядчики",
    "projects": "Проекты",
    "otkk": "ОТКК",
}


def _asset_ver(static_name: str) -> str:
    try:
        mtime = int((_WEBAPP_DIR / "static" / static_name).stat().st_mtime)
        return f"{_read_git_head()}-{mtime}"
    except OSError:
        return _read_git_head()


def _read_git_head() -> str:
    try:
        return subprocess.check_output(
            ["git", "rev-parse", "--short", "HEAD"],
            cwd=_WEBAPP_DIR.parent,
            stderr=subprocess.DEVNULL,
            text=True,
        ).strip()
    except Exception:
        return "unknown"


_PROCESS_START_TS = time.time()

app = FastAPI(title="Объединение отчётов СК")
templates = Jinja2Templates(directory=str(_HTML_TEMPLATES_DIR))
app.mount("/static", StaticFiles(directory=_WEBAPP_DIR / "static"), name="static")

WORK_DIR = Path(tempfile.gettempdir()) / "sk_reports_work"
UPLOAD_DIR = WORK_DIR / "uploads"
RESULT_DIR = WORK_DIR / "results"
PRESCRIPTIONS_UPLOAD_DIR = WORK_DIR / "prescriptions_uploads"
PRESCRIPTIONS_RESULT_DIR = WORK_DIR / "prescriptions_results"
ENGINEER_OUT_DIR = WORK_DIR / "engineer_out"
DEPLOYMENT_DIR = WORK_DIR / "deployment"
DEPLOYMENT_REPORTS_DIR = DEPLOYMENT_DIR / "reports"
DEPLOYMENT_RESULT_DIR = DEPLOYMENT_DIR / "results"
DEPLOYMENT_PRIL7_PATH = DEPLOYMENT_DIR / "pril7.xlsm"
DEPLOYMENT_TEMPLATE_PATH = DEPLOYMENT_DIR / "template.xlsm"

for d in (
    WORK_DIR,
    UPLOAD_DIR,
    RESULT_DIR,
    PRESCRIPTIONS_UPLOAD_DIR,
    PRESCRIPTIONS_RESULT_DIR,
    ENGINEER_OUT_DIR,
    DEPLOYMENT_DIR,
    DEPLOYMENT_REPORTS_DIR,
    DEPLOYMENT_RESULT_DIR,
):
    d.mkdir(exist_ok=True)

TEMPLATES_DIR = templates_dir()
if not TEMPLATES_DIR.exists():
    raise RuntimeError(f"Папка с болванками не найдена: {TEMPLATES_DIR}")
print(f"[INFO] Templates dir: {TEMPLATES_DIR} ({len(list(TEMPLATES_DIR.glob('*.docx')))} шаблонов)")

for _tpl in ("home.html", "reporting.html", "daily.html", "deployment.html", "weekly.html", "weekly_photos.html", "prescriptions.html", "planning.html", "planning_section.html", "engineer_hub.html", "engineer.html"):
    _tpl_path = _HTML_TEMPLATES_DIR / _tpl
    if not _tpl_path.is_file():
        raise RuntimeError(f"HTML-шаблон не найден: {_tpl_path} — выполните git pull и перезапустите сервер")
print(f"[INFO] UI templates: {_HTML_TEMPLATES_DIR}")
print(f"[INFO] main.py: {Path(__file__).resolve()}")
_git_head = _read_git_head()
print(f"[INFO] git: {_git_head}")
print(f"[INFO] UI build: {_APP_UI_BUILD}")
for _html in sorted(_HTML_TEMPLATES_DIR.glob("*.html")):
    print(f"[INFO]   template: {_html.name} ({_html.stat().st_size} bytes)")
try:
    templates.env.get_template("home.html")
    templates.env.get_template("daily.html")
except Exception as _tpl_err:
    raise RuntimeError(f"Jinja не видит шаблоны в {_HTML_TEMPLATES_DIR}: {_tpl_err}") from _tpl_err


@app.exception_handler(Exception)
async def unhandled_exception_handler(request: Request, exc: Exception):
    traceback.print_exc()
    return JSONResponse(
        status_code=500,
        content={
            "detail": f"Внутренняя ошибка сервера: {exc}",
            "path": str(request.url.path),
            "pid": os.getpid(),
            "git": _git_head,
            "ui_build": _APP_UI_BUILD,
        },
    )


def _do_merge(template_path: str, report_paths: list[str], output_path: str) -> int:
    """Склейка болванки с отчётами — один путь через merge_reports (ZIP + rels + media)."""
    return merge_reports(template_path, report_paths, output_path)


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


def _page_context(request: Request, breadcrumbs: list[dict] | None = None, **extra) -> dict:
    ctx = {
        "request": request,
        "agent_enabled": AGENT_ENABLED,
        "git_head": _git_head,
        "asset_ver": _asset_ver,
        "breadcrumbs": breadcrumbs or [],
    }
    ctx.update(extra)
    return ctx


@app.get("/health")
async def health():
    """На диске новый git, а процесс старый — stale_process: true → нужен перезапуск uvicorn."""
    disk_git = _read_git_head()
    main_py = Path(__file__).resolve()
    main_mtime = main_py.stat().st_mtime
    has_daily = any(getattr(r, "path", None) == "/daily" for r in app.routes)
    has_prescriptions = any(getattr(r, "path", None) == "/prescriptions" for r in app.routes)
    has_reporting = any(getattr(r, "path", None) == "/reporting" for r in app.routes)
    has_planning = any(getattr(r, "path", None) == "/planning" for r in app.routes)
    has_engineer = any(getattr(r, "path", None) == "/engineer" for r in app.routes)
    has_engineer_hub = any(getattr(r, "path", None) == "/engineer-hub" for r in app.routes)
    stale = (
        disk_git != _git_head
        or main_mtime > _PROCESS_START_TS
        or not has_daily
        or not has_prescriptions
        or not has_reporting
        or not has_planning
        or not has_engineer
        or not has_engineer_hub
    )
    db_info = None
    try:
        from sk_reporter.personnel_db import db_status

        db_info = db_status()
    except Exception:
        pass
    return {
        "ok": not stale,
        "stale_process": stale,
        "app_ui_build": _APP_UI_BUILD,
        "database": db_info,
        "has_daily_route": has_daily,
        "has_prescriptions_route": has_prescriptions,
        "has_reporting_route": has_reporting,
        "has_planning_route": has_planning,
        "has_engineer_hub_route": has_engineer_hub,
        "has_engineer_route": has_engineer,
        "pid": os.getpid(),
        "git_head_at_startup": _git_head,
        "git_head_on_disk": disk_git,
        "main_py": str(main_py),
        "main_py_mtime_after_start": main_mtime > _PROCESS_START_TS,
        "ui_templates": str(_HTML_TEMPLATES_DIR),
        "templates_on_disk": sorted(p.name for p in _HTML_TEMPLATES_DIR.glob("*.html")),
        "fix": "Shift+F5 в VS Code, затем F5 (git pull не перезапускает Python)" if stale else None,
    }


@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    print(f"[REQ] GET / pid={os.getpid()} -> home.html")
    return templates.TemplateResponse("home.html", _page_context(request))


@app.get("/reporting", response_class=HTMLResponse)
async def reporting_page(request: Request):
    print(f"[REQ] GET /reporting pid={os.getpid()} -> reporting.html")
    return templates.TemplateResponse(
        "reporting.html",
        _page_context(request, breadcrumbs=[{"label": "Отчётность"}]),
    )


@app.get("/daily", response_class=HTMLResponse)
async def daily_reports(request: Request):
    print(f"[REQ] GET /daily pid={os.getpid()} -> daily.html")
    return templates.TemplateResponse(
        "daily.html",
        _page_context(
            request,
            breadcrumbs=[
                {"label": "Отчётность", "href": "/reporting"},
                {"label": "Ежедневные отчёты"},
            ],
        ),
    )


@app.get("/deployment", response_class=HTMLResponse)
async def deployment_page(request: Request):
    print(f"[REQ] GET /deployment pid={os.getpid()} -> deployment.html")
    return templates.TemplateResponse(
        "deployment.html",
        _page_context(
            request,
            breadcrumbs=[
                {"label": "Отчётность", "href": "/reporting"},
                {"label": "Расстановка"},
            ],
        ),
    )


@app.get("/weekly", response_class=HTMLResponse)
async def weekly_reports(request: Request):
    print(f"[REQ] GET /weekly pid={os.getpid()} -> weekly.html")
    return templates.TemplateResponse(
        "weekly.html",
        _page_context(
            request,
            breadcrumbs=[
                {"label": "Отчётность", "href": "/reporting"},
                {"label": "Еженедельные отчёты"},
            ],
        ),
    )


@app.get("/weekly-photos", response_class=HTMLResponse)
async def weekly_photos_reports(request: Request):
    print(f"[REQ] GET /weekly-photos pid={os.getpid()} -> weekly_photos.html")
    return templates.TemplateResponse(
        "weekly_photos.html",
        _page_context(
            request,
            breadcrumbs=[
                {"label": "Отчётность", "href": "/reporting"},
                {"label": "Еженедельные фотоотчёты"},
            ],
        ),
    )


@app.get("/prescriptions", response_class=HTMLResponse)
async def prescriptions_page(request: Request):
    print(f"[REQ] GET /prescriptions pid={os.getpid()} -> prescriptions.html")
    return templates.TemplateResponse(
        "prescriptions.html",
        _page_context(
            request,
            breadcrumbs=[
                {"label": "Отчётность", "href": "/reporting"},
                {"label": "Предписания"},
            ],
        ),
    )


@app.get("/planning", response_class=HTMLResponse)
async def planning_page(request: Request):
    print(f"[REQ] GET /planning pid={os.getpid()} -> planning.html")
    return templates.TemplateResponse(
        "planning.html",
        _page_context(request, breadcrumbs=[{"label": "Планирование"}]),
    )


@app.get("/planning/{section}", response_class=HTMLResponse)
async def planning_section_page(request: Request, section: str):
    title = _PLANNING_SECTIONS.get(section)
    if not title:
        raise HTTPException(status_code=404, detail="Неизвестный раздел планирования")
    print(f"[REQ] GET /planning/{section} pid={os.getpid()} -> planning_section.html")
    ctx = _page_context(
        request,
        breadcrumbs=[
            {"label": "Планирование", "href": "/planning"},
            {"label": title},
        ],
        section=section,
        section_title=title,
    )
    return templates.TemplateResponse("planning_section.html", ctx)


@app.get("/api/planning/projects/{project_id}")
async def planning_project_detail(project_id: str):
    from sk_reporter.project_db import get_project_catalog

    project = await asyncio.to_thread(get_project_catalog, project_id, include_content=True)
    if not project:
        raise HTTPException(status_code=404, detail="Проект не найден")
    return project


@app.get("/api/planning/{section}")
async def planning_api(section: str):
    from sk_reporter.planning_data import planning_section

    try:
        return planning_section(section)
    except KeyError:
        raise HTTPException(status_code=404, detail="Неизвестный раздел") from None


@app.post("/api/planning/contractors/seed-from-templates")
async def planning_contractors_seed_from_templates():
    from sk_reporter.contractor_db import seed_contractors_from_templates

    try:
        return await asyncio.to_thread(seed_contractors_from_templates)
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e)) from e


@app.post("/api/planning/personnel/upload-xlsx")
async def planning_personnel_upload_xlsx(file: UploadFile = File(...)):
    from sk_reporter.db.config import database_enabled
    from sk_reporter.personnel_db import import_personnel_xlsx_to_db

    if not database_enabled():
        raise HTTPException(status_code=400, detail="DATABASE_URL не задан")
    name = (file.filename or "").lower()
    if not name.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Нужен файл Excel (.xlsx)")
    suffix = ".xlsx" if name.endswith(".xlsx") else ".xls"
    tmp = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            shutil.copyfileobj(file.file, tmp)
            tmp_path = Path(tmp.name)
        return await asyncio.to_thread(import_personnel_xlsx_to_db, tmp_path)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e)) from e
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e)) from e
    finally:
        if tmp is not None:
            try:
                Path(tmp.name).unlink(missing_ok=True)
            except OSError:
                pass


@app.post("/api/planning/core/upload-xlsx")
async def planning_core_upload_xlsx(file: UploadFile = File(...)):
    from sk_reporter.core_import import import_core_xlsx_to_db
    from sk_reporter.db.config import database_enabled

    if not database_enabled():
        raise HTTPException(status_code=400, detail="DATABASE_URL не задан")
    name = (file.filename or "").lower()
    if not name.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Нужен файл Excel (.xlsx)")
    suffix = ".xlsx" if name.endswith('.xlsx') else '.xls'
    tmp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            shutil.copyfileobj(file.file, tmp)
            tmp_path = Path(tmp.name)
        return await asyncio.to_thread(import_core_xlsx_to_db, tmp_path)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e)) from e
    except FileNotFoundError as e:
        raise HTTPException(status_code=400, detail=str(e)) from e
    except RuntimeError as e:
        raise HTTPException(status_code=503, detail=str(e)) from e
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e)) from e
    finally:
        if tmp_path is not None:
            try:
                tmp_path.unlink(missing_ok=True)
            except OSError:
                pass


@app.get("/api/planning/otkk/{card_id}")
async def planning_otkk_get(card_id: str):
    from sk_reporter.db.config import database_enabled
    from sk_reporter.otkk_store import get_card

    if not database_enabled():
        raise HTTPException(status_code=400, detail="DATABASE_URL не задан")
    card = await asyncio.to_thread(get_card, card_id, include_content=True)
    if not card:
        raise HTTPException(status_code=404, detail="Карта не найдена")
    return card


@app.get("/engineer-hub", response_class=HTMLResponse)
async def engineer_hub_page(request: Request):
    print(f"[REQ] GET /engineer-hub pid={os.getpid()} -> engineer_hub.html")
    return templates.TemplateResponse(
        "engineer_hub.html",
        _page_context(request, breadcrumbs=[{"label": "Инженер ФИО"}]),
    )


@app.get("/engineer-hub/daily-report", response_class=HTMLResponse)
async def engineer_daily_report_template_page(request: Request):
    from sk_reporter.engineer.daily_report_form import render_daily_report_page

    print(f"[REQ] GET /engineer-hub/daily-report pid={os.getpid()} -> engineer_daily_report.html")
    layout_css, table_html = render_daily_report_page()
    ctx = _page_context(
        request,
        breadcrumbs=[
            {"label": "Инженер ФИО", "href": "/engineer-hub"},
            {"label": "Ежедневный отчёт (шаблон)"},
        ],
    )
    ctx["report_layout_css"] = layout_css
    ctx["report_table_html"] = table_html
    return templates.TemplateResponse("engineer_daily_report.html", ctx)


@app.get("/api/engineer-hub")
async def api_engineer_hub():
    from sk_reporter.engineer.hub import hub_payload

    return await asyncio.to_thread(hub_payload)


@app.get("/engineer/{profile_id}", response_class=HTMLResponse)
async def engineer_profile_page(request: Request, profile_id: str):
    from sk_reporter.engineer.profile import load_profile

    print(f"[REQ] GET /engineer/{profile_id} pid={os.getpid()} -> engineer.html")
    try:
        profile = await asyncio.to_thread(load_profile, profile_id)
    except ValueError as e:
        raise HTTPException(status_code=404, detail=str(e)) from e

    fio = profile.get("name") or profile_id
    return templates.TemplateResponse(
        "engineer.html",
        _page_context(
            request,
            breadcrumbs=[
                {"label": "Инженер ФИО", "href": "/engineer-hub"},
                {"label": fio},
            ],
            header_meta_id="profileName",
            profile_id=profile_id,
        ),
    )


@app.get("/engineer", response_class=HTMLResponse)
async def engineer_page(request: Request):
    from fastapi.responses import RedirectResponse

    return RedirectResponse(url="/engineer-hub", status_code=302)


class EngineerEntry(BaseModel):
    model_config = {"populate_by_name": True}

    key: str
    name: str
    unit: str = ""
    project_qty: str = ""
    daily_qty: str = ""
    cumulative_qty: str = ""
    location: str = ""
    reference: str = ""
    stage: str = ""
    object_title: str = Field(default="", alias="object")


class EngineerBuildRequest(BaseModel):
    project_id: str
    report_date: str
    entries: list[EngineerEntry]
    profile_id: str | None = None


@app.get("/api/engineer/config")
async def engineer_config(profile_id: str | None = None):
    from sk_reporter.engineer.profile import load_profile, resolve_report_template
    from sk_reporter.engineer.vor_data import list_profile_projects

    try:
        profile = load_profile(profile_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e)) from e

    tpl = resolve_report_template(profile)
    return {
        "profile": {
            "id": profile.get("id"),
            "name": profile.get("name"),
            "position": profile.get("position"),
            "person_id": profile.get("person_id"),
        },
        "projects": list_profile_projects(profile),
        "template_ok": tpl is not None,
        "template_path": str(tpl) if tpl else None,
    }


@app.post("/api/engineer/build")
async def engineer_build(body: EngineerBuildRequest):
    from datetime import date as date_cls

    from sk_reporter.engineer.profile import load_profile, resolve_report_template
    from sk_reporter.engineer.report_builder import ReportEntry, build_report_docx
    from sk_reporter.engineer.vor_data import profile_project_ids

    if not body.entries:
        raise HTTPException(status_code=400, detail="Не выбраны работы")

    try:
        profile = load_profile(body.profile_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e)) from e

    tpl = resolve_report_template(profile)
    if tpl is None:
        raise HTTPException(
            status_code=400,
            detail="Шаблон отчёта не найден — положите data/engineer/report_template.docx",
        )

    allowed = profile_project_ids(profile)
    if body.project_id not in allowed:
        raise HTTPException(status_code=400, detail="Проект не закреплён за инженером")

    try:
        report_day = date_cls.fromisoformat(body.report_date)
    except ValueError as e:
        raise HTTPException(status_code=400, detail="Неверная дата") from e

    entries = [
        ReportEntry(
            name=e.name,
            unit=e.unit,
            project_qty=e.project_qty,
            daily_qty=e.daily_qty,
            cumulative_qty=e.cumulative_qty,
            location=e.location,
            reference=e.reference,
            stage=e.stage,
            object_title=e.object,
        )
        for e in body.entries
    ]

    safe_date = body.report_date.replace("-", "")
    out_name = f"отчёт_{body.project_id}_{safe_date}.docx"
    out_path = ENGINEER_OUT_DIR / out_name

    try:
        await asyncio.to_thread(
            build_report_docx,
            tpl,
            out_path,
            entries,
            body.project_id,
            report_day,
        )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e)) from e

    return FileResponse(
        out_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=out_name,
    )


@app.post("/upload/reports")
async def upload_reports(files: list[UploadFile] = File(...)):
    saved = []
    for f in files:
        if not f.filename:
            continue
        dest = UPLOAD_DIR / f.filename
        with open(dest, "wb") as out:
            shutil.copyfileobj(f.file, out)
        if dest.suffix.lower() == ".docx":
            remove_editing_restrictions(str(dest))
        saved.append(f.filename)
    return {"uploaded": saved, "count": len(saved)}


_DEPLOYMENT_SUFFIXES = {".xlsm", ".xlsx"}


@app.get("/api/deployment/status")
async def deployment_status():
    from sk_reporter.deployment.templates_store import pril7_status, template_status

    reports = sorted(
        f.name for f in DEPLOYMENT_REPORTS_DIR.iterdir()
        if f.suffix.lower() in (".docx", ".doc")
    )
    results = sorted(f.name for f in DEPLOYMENT_RESULT_DIR.iterdir() if f.suffix.lower() == ".zip")
    tpl = template_status(DEPLOYMENT_DIR)
    pril7 = pril7_status(DEPLOYMENT_DIR)
    return {
        "reports": reports,
        "has_pril7": pril7["available"],
        "pril7_source": pril7["source"],
        "pril7_name": pril7["name"],
        "has_template": tpl["available"],
        "template_source": tpl["source"],
        "template_name": tpl["name"],
        "results": results,
    }


@app.post("/upload/deployment/reports")
async def upload_deployment_reports(files: list[UploadFile] = File(...)):
    saved = []
    for f in files:
        if not f.filename:
            continue
        suffix = Path(f.filename).suffix.lower()
        if suffix not in (".docx", ".doc"):
            raise HTTPException(status_code=400, detail=f"Нужен .docx: {f.filename}")
        dest = DEPLOYMENT_REPORTS_DIR / f.filename
        with open(dest, "wb") as out:
            shutil.copyfileobj(f.file, out)
        saved.append(f.filename)
    return {"uploaded": saved, "count": len(saved)}


@app.post("/upload/deployment/pril7")
async def upload_deployment_pril7(file: UploadFile = File(...)):
    if not file.filename:
        raise HTTPException(status_code=400, detail="Файл не выбран")
    suffix = Path(file.filename).suffix.lower()
    if suffix not in _DEPLOYMENT_SUFFIXES:
        raise HTTPException(status_code=400, detail="Нужен .xlsm или .xlsx")
    with open(DEPLOYMENT_PRIL7_PATH, "wb") as out:
        shutil.copyfileobj(file.file, out)
    return {"uploaded": DEPLOYMENT_PRIL7_PATH.name}


@app.post("/upload/deployment/template")
async def upload_deployment_template(file: UploadFile = File(...)):
    if not file.filename:
        raise HTTPException(status_code=400, detail="Файл не выбран")
    suffix = Path(file.filename).suffix.lower()
    if suffix not in _DEPLOYMENT_SUFFIXES:
        raise HTTPException(status_code=400, detail="Нужен .xlsm или .xlsx")
    with open(DEPLOYMENT_TEMPLATE_PATH, "wb") as out:
        shutil.copyfileobj(file.file, out)
    return {"uploaded": DEPLOYMENT_TEMPLATE_PATH.name}


@app.post("/api/deployment/generate/stream")
async def deployment_generate_stream(report_date: str = Form(...)):
    from sk_reporter.deployment.pipeline import run_deployment
    from sk_reporter.deployment.templates_store import resolve_pril7_template, resolve_rasstanovka_template

    async def event_generator():
        reports = list(DEPLOYMENT_REPORTS_DIR.glob("*.docx")) + list(
            DEPLOYMENT_REPORTS_DIR.glob("*.doc")
        )
        if not reports:
            yield _sse({"type": "error", "msg": "Загрузите отчёты .docx"})
            yield _sse({"type": "done", "ok": False})
            return
        try:
            pril7_path, pril7_source = resolve_pril7_template(DEPLOYMENT_DIR)
        except FileNotFoundError:
            yield _sse({"type": "error", "msg": "Шаблон Приложения 7 не найден на сервере"})
            yield _sse({"type": "done", "ok": False})
            return
        try:
            template_path, template_source = resolve_rasstanovka_template(DEPLOYMENT_DIR)
        except FileNotFoundError:
            yield _sse({"type": "error", "msg": "Шаблон расстановки не найден на сервере"})
            yield _sse({"type": "done", "ok": False})
            return

        pril7_label = "загруженный" if pril7_source == "upload" else "с сервера"
        tpl_label = "загруженный" if template_source == "upload" else "с сервера"
        yield _sse({
            "type": "start",
            "msg": f"Формирую расстановку… (Прил.7: {pril7_label}, шаблон: {tpl_label})",
        })

        zip_path, logs = await asyncio.to_thread(
            run_deployment,
            reports_dir=DEPLOYMENT_REPORTS_DIR,
            pril7_path=pril7_path,
            template_path=template_path,
            output_dir=DEPLOYMENT_RESULT_DIR,
            report_date=report_date,
        )
        for line in logs:
            yield _sse({"type": "log", "msg": line})
            await asyncio.sleep(0)

        if zip_path and zip_path.is_file():
            yield _sse({
                "type": "done",
                "ok": True,
                "download": f"/download/deployment/{zip_path.name}",
                "filename": zip_path.name,
            })
        else:
            yield _sse({"type": "error", "msg": "Не удалось сформировать архив"})
            yield _sse({"type": "done", "ok": False})

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "Connection": "keep-alive", "X-Accel-Buffering": "no"},
    )


@app.get("/download/deployment/{filename}")
async def download_deployment(filename: str):
    if ".." in filename or "/" in filename or "\\" in filename:
        raise HTTPException(status_code=400, detail="Недопустимое имя файла")
    path = (DEPLOYMENT_RESULT_DIR / filename).resolve()
    if not str(path).startswith(str(DEPLOYMENT_RESULT_DIR.resolve())) or not path.is_file():
        raise HTTPException(status_code=404, detail="Файл не найден")
    return FileResponse(path, filename=filename)


@app.delete("/clear/deployment/reports")
async def clear_deployment_reports():
    shutil.rmtree(DEPLOYMENT_REPORTS_DIR)
    DEPLOYMENT_REPORTS_DIR.mkdir()
    return {"cleared": True}


@app.delete("/clear/deployment/all")
async def clear_deployment_all():
    for d in (DEPLOYMENT_REPORTS_DIR, DEPLOYMENT_RESULT_DIR):
        shutil.rmtree(d)
        d.mkdir()
    for p in (DEPLOYMENT_PRIL7_PATH, DEPLOYMENT_TEMPLATE_PATH):
        if p.is_file():
            p.unlink()
    return {"cleared": True}


_PRESCRIPTION_SUFFIXES = {".xlsx", ".xlsm", ".xls"}


@app.post("/upload/prescriptions")
async def upload_prescriptions(files: list[UploadFile] = File(...)):
    saved = []
    for f in files:
        if not f.filename:
            continue
        suffix = Path(f.filename).suffix.lower()
        if suffix not in _PRESCRIPTION_SUFFIXES:
            raise HTTPException(
                status_code=400,
                detail=f"Недопустимый формат: {f.filename} (нужен .xlsx или .xls)",
            )
        dest = PRESCRIPTIONS_UPLOAD_DIR / f.filename
        with open(dest, "wb") as out:
            shutil.copyfileobj(f.file, out)
        saved.append(f.filename)
    return {"uploaded": saved, "count": len(saved)}


@app.get("/prescriptions/results")
async def list_prescription_results():
    files = [
        f.name for f in PRESCRIPTIONS_RESULT_DIR.iterdir()
        if f.suffix.lower() in _PRESCRIPTION_SUFFIXES
    ]
    return {"files": sorted(files)}


def _checked_prescription_name(upload_filename: str) -> str:
    stem = Path(upload_filename).stem
    suffix = Path(upload_filename).suffix.lower() or ".xlsx"
    return f"{stem}_проверен{suffix}"


@app.post("/check/prescriptions/stream")
async def check_prescriptions_stream():
    from sk_reporter.prescriptions import check_prescription, write_checked_copy

    async def event_generator():
        files = sorted(
            f for f in PRESCRIPTIONS_UPLOAD_DIR.iterdir()
            if f.suffix.lower() in _PRESCRIPTION_SUFFIXES
        )
        if not files:
            yield _sse({"type": "error", "msg": "Excel-файлы не загружены"})
            yield _sse({"type": "done", "summary": {"total": 0, "ok": 0, "warnings": 0, "errors": 0}})
            return

        ok_count = 0
        warn_count = 0
        err_count = 0

        yield _sse({
            "type": "start",
            "msg": f"Проверяю {len(files)} файл(ов)…",
            "total": len(files),
        })

        for file_path in files:
            filename = file_path.name
            try:
                yield _sse({"type": "info", "filename": filename, "msg": f"{filename}: проверяю…"})
                await asyncio.sleep(0)
                result = await asyncio.to_thread(check_prescription, file_path)
                issues = result.get("issues") or []
                has_errors = bool(result.get("has_errors")) or any(
                    i.get("level") == "error" for i in issues
                )
                has_warnings = bool(result.get("has_warnings")) or any(
                    i.get("level") == "warn" for i in issues
                )
                if has_errors:
                    err_count += 1
                elif has_warnings:
                    warn_count += 1
                else:
                    ok_count += 1

                out_name = _checked_prescription_name(filename)
                out_path = PRESCRIPTIONS_RESULT_DIR / out_name
                await asyncio.to_thread(write_checked_copy, file_path, out_path, result)

                status = "⚠ ошибки" if has_errors else ("◐ замечания" if has_warnings else "✓ OK")
                yield _sse({
                    "type": "report",
                    "filename": out_name,
                    "msg": f"{filename}: {status}",
                    "hasErrors": has_errors,
                    "hasWarnings": has_warnings,
                    "result": result,
                    "download": f"/download/prescriptions/{out_name}",
                })
                await asyncio.sleep(0)
            except Exception as e:
                err_count += 1
                yield _sse({"type": "error", "msg": f"Ошибка {filename}: {e}"})

        yield _sse({
            "type": "done",
            "summary": {
                "total": len(files),
                "ok": ok_count,
                "warnings": warn_count,
                "errors": err_count,
            },
        })

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "Connection": "keep-alive", "X-Accel-Buffering": "no"},
    )


@app.get("/download/prescriptions/all.zip")
async def download_prescriptions_all():
    files = sorted(
        f for f in PRESCRIPTIONS_RESULT_DIR.iterdir()
        if f.suffix.lower() in _PRESCRIPTION_SUFFIXES
    )
    if not files:
        raise HTTPException(status_code=404, detail="Нет проверенных файлов")
    return _zip_download_response(
        files,
        None,
        "attachment; filename*=UTF-8''%D0%BF%D1%80%D0%B5%D0%B4%D0%BF%D0%B8%D1%81%D0%B0%D0%BD%D0%B8%D1%8F.zip",
    )


@app.get("/download/prescriptions/{filename}")
async def download_prescription(filename: str):
    if ".." in filename or "/" in filename or "\\" in filename:
        raise HTTPException(status_code=400, detail="Недопустимое имя файла")
    path = (PRESCRIPTIONS_RESULT_DIR / filename).resolve()
    if not str(path).startswith(str(PRESCRIPTIONS_RESULT_DIR.resolve())):
        raise HTTPException(status_code=400, detail="Недопустимый путь")
    if not path.is_file():
        raise HTTPException(status_code=404, detail="Файл не найден")
    media = (
        "application/vnd.ms-excel"
        if path.suffix.lower() == ".xls"
        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    return FileResponse(path=str(path), filename=filename, media_type=media)


@app.delete("/clear/prescriptions/uploads")
async def clear_prescriptions_uploads():
    shutil.rmtree(PRESCRIPTIONS_UPLOAD_DIR)
    PRESCRIPTIONS_UPLOAD_DIR.mkdir()
    return {"ok": True}


@app.delete("/clear/prescriptions/all")
async def clear_prescriptions_all():
    for d in (PRESCRIPTIONS_UPLOAD_DIR, PRESCRIPTIONS_RESULT_DIR):
        shutil.rmtree(d)
        d.mkdir()
    return {"ok": True}


@app.get("/files/reports")
async def list_reports():
    files = [f.name for f in UPLOAD_DIR.iterdir() if f.suffix.lower() in (".docx", ".doc")]
    return {"files": sorted(files)}


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
                skipped = bool(result.get("skipped"))
                has_errors = not result.get("ok", False) and not skipped
                if has_errors:
                    errors_count += 1
                check_results[filename] = result
                if skipped:
                    msg = f"{filename}: проверка пропущена ({result.get('skip_reason')})"
                else:
                    msg = f"{filename}: " + ("⚠️ найдены проблемы" if has_errors else "✓ ОК")
                yield _sse({
                    "type": "report",
                    "filename": filename,
                    "msg": msg,
                    "hasErrors": has_errors,
                    "skipped": skipped,
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
                if result.get("skipped"):
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
    weather: str | None = None  # температура, напр. «+21» → «+21℃» в отчёте


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
    weather = (body.weather if body else None) or None
    if weather is not None:
        weather = weather.strip() or None
    log = prepare_uploaded_reports(str(UPLOAD_DIR), layout, target_date, weather=weather)
    return {
        "log": log,
        "template": "сетка захардкожена",
        "grid_cols": layout["grid_cols"],
        "grid_cols_7": layout.get("grid_cols_7"),
        "date": target_date,
        "weather": weather,
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
            try:
                if name == GEODESY_COMPANY:
                    for report in sorted(reports, key=lambda p: p.name.lower()):
                        output_path = RESULT_DIR / f"{name}_merged_{report.stem}.docx"
                        inserted = _do_merge(str(template), [str(report)], str(output_path))
                        results.append({
                            "company": name,
                            "inserted": inserted,
                            "file": output_path.name,
                            "reports_count": 1,
                            "source": report.name,
                        })
                        yield _sse({
                            "type": "success",
                            "company": name,
                            "msg": f"Собран отчёт геодезии → {output_path.name}",
                        })
                else:
                    output_path = RESULT_DIR / f"{name}_merged.docx"
                    inserted = _do_merge(str(template), [str(r) for r in reports], str(output_path))
                    results.append({
                        "company": name,
                        "inserted": inserted,
                        "file": f"{name}_merged.docx",
                        "reports_count": len(reports),
                    })
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
