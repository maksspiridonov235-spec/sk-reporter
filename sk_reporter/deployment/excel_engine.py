"""Заполнение .xlsm на Linux/macOS через LibreOffice (xlsm → xlsx → правка → xlsm)."""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Callable

_CONVERT_TIMEOUT = 180


def find_soffice() -> str | None:
    env = os.environ.get("LIBREOFFICE_PATH", "").strip()
    if env and Path(env).is_file():
        return env
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for candidate in (
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
        "/usr/lib/libreoffice/program/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if Path(candidate).is_file():
            return candidate
    return None


def libreoffice_available() -> bool:
    return find_soffice() is not None


def _convert(soffice: str, src: Path, out_dir: Path, target_ext: str) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to",
        target_ext,
        str(src.resolve()),
        "--outdir",
        str(out_dir.resolve()),
    ]
    result = subprocess.run(
        cmd,
        check=False,
        capture_output=True,
        text=True,
        timeout=_CONVERT_TIMEOUT,
    )
    if result.returncode != 0:
        err = (result.stderr or result.stdout or "").strip()
        raise RuntimeError(f"LibreOffice convert failed: {err or result.returncode}")

    ext = target_ext.split(":")[0]
    out = out_dir / f"{src.stem}.{ext}"
    if out.is_file():
        return out
    candidates = sorted(out_dir.glob(f"{src.stem}.*"))
    if len(candidates) == 1:
        return candidates[0]
    raise FileNotFoundError(f"LibreOffice не создал {out}")


def fill_xlsm_via_libreoffice(
    xlsm_path: Path,
    fill_xlsx: Callable[[Path], bool],
    log_func=print,
) -> bool:
    """Конвертирует xlsm в xlsx, вызывает fill_xlsx, конвертирует обратно в xlsm."""
    soffice = find_soffice()
    if not soffice:
        log_func(
            "ОШИБКА: LibreOffice не найден (soffice). "
            "На сервере нужен пакет libreoffice-calc (или переменная LIBREOFFICE_PATH)."
        )
        return False

    xlsm_path = Path(xlsm_path)
    with tempfile.TemporaryDirectory(prefix="sk_xlsm_") as tmp:
        work = Path(tmp)
        src = work / xlsm_path.name
        shutil.copy2(xlsm_path, src)

        log_func("LibreOffice: .xlsm → .xlsx…")
        try:
            xlsx_path = _convert(soffice, src, work, "xlsx")
        except Exception as exc:
            log_func(f"ОШИБКА конвертации: {exc}")
            return False

        if not fill_xlsx(xlsx_path):
            return False

        log_func("LibreOffice: .xlsx → .xlsm…")
        try:
            xlsm_out = _convert(soffice, xlsx_path, work, "xlsm")
        except Exception as exc:
            log_func(f"ОШИБКА конвертации: {exc}")
            return False

        shutil.copy2(xlsm_out, xlsm_path)
        log_func("Приложение 7 сохранено (.xlsm)")
        return True
