from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse

from config import AGENT_ENABLED, templates

router = APIRouter()


@router.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {
        "request": request,
        "agent_enabled": AGENT_ENABLED,
    })
