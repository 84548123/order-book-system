from __future__ import annotations

import asyncio
import json
import time
import uuid
import zipfile
from io import BytesIO
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask

from app.config import settings
from app.services.processor import ProcessingError, process_workbook

BASE_DIR = Path(__file__).resolve().parent
TEMP_DIR = BASE_DIR.parent / ".processed"
TEMP_DIR.mkdir(exist_ok=True)
LATEST_DATA_PATH = TEMP_DIR / "latest.json"
LATEST_FILE_PATH = TEMP_DIR / "latest.xlsx"

app = FastAPI(title="Order Book Data Processing", version="1.0.0")
files: dict[str, tuple[Path, float]] = {}


def _remove_file(file_id: str) -> None:
    entry = files.pop(file_id, None)
    if entry:
        entry[0].unlink(missing_ok=True)


def _cleanup_expired() -> None:
    cutoff = time.time() - settings.file_ttl_minutes * 60
    for file_id, (_, created) in list(files.items()):
        if created < cutoff:
            _remove_file(file_id)


def _validate_xlsx(content: bytes) -> None:
    if not content.startswith(b"PK"):
        raise HTTPException(400, "The uploaded file is not a valid .xlsx file.")
    try:
        with zipfile.ZipFile(BytesIO(content)) as archive:
            names = archive.namelist()
            if "[Content_Types].xml" not in names or "xl/workbook.xml" not in names:
                raise HTTPException(400, "The uploaded file is not a valid .xlsx workbook.")
            total = sum(item.file_size for item in archive.infolist())
            if total > settings.max_upload_mb * 30 * 1024 * 1024:
                raise HTTPException(413, "The expanded workbook is too large to process safely.")
    except zipfile.BadZipFile as exc:
        raise HTTPException(400, "The uploaded file is corrupted.") from exc


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "healthy"}


@app.post("/api/process")
async def process(file: UploadFile = File(...)):
    _cleanup_expired()
    if not file.filename or Path(file.filename).suffix.lower() != ".xlsx":
        raise HTTPException(400, "Only .xlsx Excel files are accepted.")
    limit = settings.max_upload_mb * 1024 * 1024
    content = await file.read(limit + 1)
    if not content:
        raise HTTPException(400, "The uploaded file is empty.")
    if len(content) > limit:
        raise HTTPException(413, f"The file exceeds the {settings.max_upload_mb} MB upload limit.")
    _validate_xlsx(content)
    try:
        result = await asyncio.to_thread(process_workbook, content, None, settings.max_rows)
    except ProcessingError as exc:
        raise HTTPException(422, str(exc)) from exc
    file_id = uuid.uuid4().hex
    path = TEMP_DIR / f"{file_id}.xlsx"
    path.write_bytes(result.content)
    files[file_id] = (path, time.time())
    payload = {
        "status": "completed",
        "file_id": "latest",
        "filename": file.filename,
        "total_records": result.total_records,
        "warnings": result.warnings[:100],
        "warning_count": len(result.warnings),
        "records": result.records,
        "detail_summary": result.detail_summary,
        "factory_summary": result.factory_summary,
    }
    # Keep the latest successful upload available to every browser using the app.
    LATEST_FILE_PATH.write_bytes(result.content)
    temporary_data = LATEST_DATA_PATH.with_suffix(".tmp")
    temporary_data.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    temporary_data.replace(LATEST_DATA_PATH)
    return payload


@app.get("/api/latest")
def latest():
    if not LATEST_DATA_PATH.exists():
        raise HTTPException(404, "No workbook has been uploaded yet.")
    try:
        return json.loads(LATEST_DATA_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise HTTPException(503, "The latest dataset is temporarily unavailable.") from exc


@app.get("/api/download/{file_id}")
def download(file_id: str):
    if file_id == "latest" and LATEST_FILE_PATH.exists():
        return FileResponse(
            LATEST_FILE_PATH,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="processed_order_book.xlsx",
        )
    _cleanup_expired()
    entry = files.get(file_id)
    if not entry or not entry[0].exists():
        raise HTTPException(404, "The processed file was not found or has expired.")
    return FileResponse(
        entry[0],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="processed_order_book.xlsx",
        background=BackgroundTask(_remove_file, file_id),
    )


app.mount("/", StaticFiles(directory=BASE_DIR / "static", html=True), name="static")
