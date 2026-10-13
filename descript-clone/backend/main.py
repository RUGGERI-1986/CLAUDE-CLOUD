import json
import os
import shutil
import tempfile
import uuid
from pathlib import Path

from fastapi import BackgroundTasks, FastAPI, HTTPException, UploadFile
from fastapi.responses import FileResponse
from pydantic import BaseModel
from starlette.background import BackgroundTask

import editing
import transcribe as transcriber

DATA = Path(os.environ.get("DATA_DIR", Path(__file__).parent / "data"))
DATA.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="Descript clone")


def _dir(pid: str) -> Path:
    # uuid4 hex: evita path traversal
    if len(pid) != 32 or not all(c in "0123456789abcdef" for c in pid):
        raise HTTPException(404)
    d = DATA / pid
    if not d.is_dir():
        raise HTTPException(404)
    return d


def _write_state(d: Path, state: dict) -> None:
    tmp = d / "state.json.tmp"
    tmp.write_text(json.dumps(state))
    tmp.replace(d / "state.json")


def _read_state(d: Path) -> dict:
    return json.loads((d / "state.json").read_text())


def _run_transcription(pid: str) -> None:
    d = DATA / pid
    state = _read_state(d)
    try:
        result = transcriber.transcribe(str(d / state["file"]))
        state.update(status="done", **result)
    except Exception as e:  # el error se muestra en la UI
        state.update(status="error", error=str(e))
    _write_state(d, state)


@app.post("/api/projects")
async def create_project(file: UploadFile, tasks: BackgroundTasks):
    pid = uuid.uuid4().hex
    d = DATA / pid
    d.mkdir()
    ext = Path(file.filename or "").suffix.lower()[:10] or ".bin"
    name = f"media{ext}"
    with open(d / name, "wb") as out:
        shutil.copyfileobj(file.file, out)
    _write_state(d, {"id": pid, "title": file.filename, "file": name, "status": "transcribing", "words": []})
    tasks.add_task(_run_transcription, pid)
    return {"id": pid}


@app.get("/api/projects/{pid}")
def get_project(pid: str):
    return _read_state(_dir(pid))


@app.get("/api/projects/{pid}/media")
def get_media(pid: str):
    d = _dir(pid)
    return FileResponse(d / _read_state(d)["file"])  # soporta Range (seek en <video>)


class ExportBody(BaseModel):
    keep: list[tuple[float, float]]


@app.post("/api/projects/{pid}/export")
def export_project(pid: str, body: ExportBody):
    d = _dir(pid)
    state = _read_state(d)
    video = editing.has_video(str(d / state["file"]))
    suffix = ".mp4" if video else ".mp3"
    tmpdir = tempfile.mkdtemp()
    dst = os.path.join(tmpdir, "export" + suffix)
    try:
        editing.export(str(d / state["file"]), dst, body.keep)
    except ValueError as e:
        shutil.rmtree(tmpdir, ignore_errors=True)
        raise HTTPException(400, str(e))
    except RuntimeError as e:
        shutil.rmtree(tmpdir, ignore_errors=True)
        raise HTTPException(500, str(e))
    return FileResponse(
        dst, filename=f"{Path(state['title'] or 'export').stem}-editado{suffix}",
        background=BackgroundTask(shutil.rmtree, tmpdir, ignore_errors=True),
    )
