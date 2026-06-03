from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles

from config import WEBAPP_DIR
from routes import (
    check_router,
    downloads_router,
    merge_router,
    pages_router,
    prepare_router,
    reports_router,
)

app = FastAPI(title="Объединение отчётов СК")
app.mount("/static", StaticFiles(directory=WEBAPP_DIR / "static"), name="static")

app.include_router(pages_router)
app.include_router(reports_router)
app.include_router(check_router)
app.include_router(prepare_router)
app.include_router(merge_router)
app.include_router(downloads_router)


@app.exception_handler(Exception)
async def unhandled_exception_handler(request: Request, exc: Exception):
    return JSONResponse(
        status_code=500,
        content={"detail": f"Внутренняя ошибка сервера: {str(exc)}"},
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)
