from routes.check import router as check_router
from routes.downloads import router as downloads_router
from routes.merge import router as merge_router
from routes.pages import router as pages_router
from routes.prepare import router as prepare_router
from routes.reports import router as reports_router

__all__ = [
    "pages_router",
    "reports_router",
    "check_router",
    "prepare_router",
    "merge_router",
    "downloads_router",
]
