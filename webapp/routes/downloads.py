import shutil

from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse, StreamingResponse

from config import OUTPUT_DIR, RESULT_DIR, UPLOAD_DIR
from helpers import zip_files

router = APIRouter()


@router.get("/download/all.zip")
async def download_all():
    files = [f for f in RESULT_DIR.iterdir() if f.suffix.lower() in (".docx", ".doc")]
    if not files:
        raise HTTPException(status_code=404, detail="Нет готовых файлов")
    return StreamingResponse(
        zip_files(files),
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename*=UTF-8''%D0%BE%D1%82%D1%87%D1%91%D1%82%D1%8B.zip"},
    )


@router.get("/download/fixed/all.zip")
async def download_fixed_all():
    if not OUTPUT_DIR.exists():
        raise HTTPException(status_code=404, detail="Нет исправленных файлов")
    files = [f for f in OUTPUT_DIR.iterdir() if f.suffix.lower() in (".docx", ".doc")]
    if not files:
        raise HTTPException(status_code=404, detail="Нет исправленных файлов")
    return StreamingResponse(
        zip_files(files),
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename*=UTF-8''%D0%B8%D1%81%D0%BF%D1%80%D0%B0%D0%B2%D0%BB%D0%B5%D0%BD%D0%BD%D1%8B%D0%B5.zip"},
    )


@router.get("/download/fixed/{filename}")
async def download_fixed(filename: str):
    path = (OUTPUT_DIR / filename).resolve()
    if not str(path).startswith(str(OUTPUT_DIR.resolve())):
        raise HTTPException(status_code=400, detail="Недопустимый путь")
    if not path.exists():
        raise HTTPException(status_code=404, detail="Файл не найден")
    return FileResponse(
        path=str(path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@router.get("/download/{filename}")
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


@router.get("/results")
async def list_results():
    files = [f.name for f in RESULT_DIR.iterdir() if f.suffix.lower() in (".docx", ".doc")]
    return {"files": sorted(files)}


@router.delete("/clear/results")
async def clear_results():
    shutil.rmtree(RESULT_DIR)
    RESULT_DIR.mkdir()
    return {"ok": True}


@router.delete("/clear/all")
async def clear_all():
    for d in (UPLOAD_DIR, RESULT_DIR):
        shutil.rmtree(d)
        d.mkdir()
    return {"ok": True}
