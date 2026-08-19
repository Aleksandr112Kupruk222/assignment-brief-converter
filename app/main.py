from __future__ import annotations

import hashlib
import json
import shutil
import tempfile
import time
import uuid
from pathlib import Path

from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from .documents import generate_document, inspect_docx, match_fields


app = FastAPI(title="Assignment Brief Converter", version="1.0.0")
SESSION_ROOT = Path(tempfile.gettempdir()) / "assignment-brief-converter"
SESSION_ROOT.mkdir(parents=True, exist_ok=True)
MAX_UPLOAD_BYTES = 20 * 1024 * 1024


def _clean_old_sessions(max_age_seconds: int = 3600) -> None:
    cutoff = time.time() - max_age_seconds
    for directory in SESSION_ROOT.iterdir():
        if directory.is_dir() and directory.stat().st_mtime < cutoff:
            shutil.rmtree(directory, ignore_errors=True)


async def _save_upload(upload: UploadFile, destination: Path) -> str:
    if not upload.filename or not upload.filename.lower().endswith(".docx"):
        raise HTTPException(415, "Only .docx Word documents are supported.")
    data = await upload.read(MAX_UPLOAD_BYTES + 1)
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(413, "The document exceeds the 20 MB upload limit.")
    destination.write_bytes(data)
    return hashlib.sha256(data).hexdigest()


@app.post("/api/analyse")
async def analyse(old_document: UploadFile = File(...), new_document: UploadFile = File(...)):
    _clean_old_sessions()
    session_id = uuid.uuid4().hex
    directory = SESSION_ROOT / session_id
    directory.mkdir()
    old_path, new_path = directory / "old.docx", directory / "new.docx"
    try:
        old_hash = await _save_upload(old_document, old_path)
        new_hash = await _save_upload(new_document, new_path)
        if old_hash == new_hash:
            raise HTTPException(400, "The old assignment and new-template example must be different documents.")
        _, old_fields = inspect_docx(old_path, "old")
        _, new_fields = inspect_docx(new_path, "new", include_slots=True)
        if not new_fields:
            raise HTTPException(422, "No replaceable labelled fields or sections were detected in the new template.")
        metadata = {
            "old_filename": old_document.filename,
            "new_filename": new_document.filename,
            "old_fields": [field.public() for field in old_fields],
            "new_fields": [field.public() for field in new_fields],
        }
        (directory / "analysis.json").write_text(json.dumps(metadata), encoding="utf-8")
        mappings = match_fields(old_fields, new_fields)
        detected = sorted({field.canonical for field in old_fields if field.canonical})
        missing = [m for m in mappings if not m["source_id"]]
        return {"session_id": session_id, "detected": detected, "old_fields": metadata["old_fields"],
                "new_fields": metadata["new_fields"], "mappings": mappings, "missing": missing}
    except HTTPException:
        shutil.rmtree(directory, ignore_errors=True)
        raise
    except ValueError as exc:
        shutil.rmtree(directory, ignore_errors=True)
        raise HTTPException(422, str(exc)) from exc


@app.post("/api/generate")
async def generate(background_tasks: BackgroundTasks, session_id: str = Form(...), mappings: str = Form(...),
                   manual_values: str = Form("{}")):
    directory = SESSION_ROOT / session_id
    if not directory.is_dir() or directory.parent != SESSION_ROOT:
        raise HTTPException(404, "This conversion session has expired. Please analyse the documents again.")
    try:
        selection_data = json.loads(mappings)
        manual_data = json.loads(manual_values)
        metadata = json.loads((directory / "analysis.json").read_text(encoding="utf-8"))
        _, old_fields = inspect_docx(directory / "old.docx", "old")
        _, new_fields = inspect_docx(directory / "new.docx", "new", include_slots=True)
        output = directory / "converted-assignment-brief.docx"
        generate_document(directory / "new.docx", output, old_fields, new_fields, selection_data, manual_data)
    except (json.JSONDecodeError, KeyError, ValueError) as exc:
        raise HTTPException(400, f"Invalid conversion request: {exc}") from exc
    background_tasks.add_task(shutil.rmtree, directory, True)
    return FileResponse(output, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        filename="converted-assignment-brief.docx", background=background_tasks)


app.mount("/", StaticFiles(directory=Path(__file__).parent / "static", html=True), name="static")

