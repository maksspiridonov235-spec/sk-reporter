import sys
import tempfile
from pathlib import Path

from fastapi.templating import Jinja2Templates

WEBAPP_DIR = Path(__file__).parent
REPO_ROOT = WEBAPP_DIR.parent
sys.path.insert(0, str(REPO_ROOT))

WORK_DIR = Path(tempfile.gettempdir()) / "sk_reports_work"
UPLOAD_DIR = WORK_DIR / "uploads"
RESULT_DIR = WORK_DIR / "results"
OUTPUT_DIR = REPO_ROOT / "output"

TEMPLATES_DIR = REPO_ROOT / "contractor_report" / "болванки (шаблоны не вырезать только копировать)"
LAYOUT_TEMPLATE_FILE = "Ежедневный отчет Шаблон.docx"

for d in (WORK_DIR, UPLOAD_DIR, RESULT_DIR):
    d.mkdir(exist_ok=True)

if not TEMPLATES_DIR.exists():
    raise RuntimeError(f"Папка с болванками не найдена: {TEMPLATES_DIR}")

print(f"[INFO] Templates dir: {TEMPLATES_DIR} ({len(list(TEMPLATES_DIR.glob('*.docx')))} шаблонов)")

try:
    from agent.ocr_agent import detect_company, merge_report_into_template  # noqa: F401

    AGENT_ENABLED = True
    print("[INFO] AI agent connected: qwen3.5:cloud via Ollama")
except ImportError as e:
    detect_company = None  # type: ignore
    merge_report_into_template = None  # type: ignore
    AGENT_ENABLED = False
    print(f"[WARNING] Agent not found: {e}")

templates = Jinja2Templates(directory=str(WEBAPP_DIR / "templates"))
