from fastapi import APIRouter

from config import RESULT_DIR, TEMPLATES_DIR, UPLOAD_DIR
from docx_processing import prepare_uploaded_reports, rename_results, rename_templates
from helpers import PrepareBody, hardcoded_layout, parse_report_date

router = APIRouter()


@router.post("/macro/prepare")
async def macro_prepare(body: PrepareBody | None = None):
    target_date = parse_report_date(body)
    layout = hardcoded_layout()
    log = prepare_uploaded_reports(str(UPLOAD_DIR), layout, target_date)
    return {
        "log": log,
        "template": "сетка захардкожена",
        "grid_cols": layout["grid_cols"],
        "grid_cols_7": layout.get("grid_cols_7"),
        "date": target_date,
    }


@router.post("/rename/templates")
async def rename_templates_endpoint(body: PrepareBody):
    target_date = parse_report_date(body)
    log = rename_templates(str(TEMPLATES_DIR), target_date)
    return {"log": log, "date": target_date}


@router.post("/rename/results")
async def rename_results_endpoint(body: PrepareBody):
    target_date = parse_report_date(body)
    log = rename_results(str(RESULT_DIR), target_date)
    return {"log": log, "date": target_date}
