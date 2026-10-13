import subprocess
import time

import pytest
from fastapi.testclient import TestClient


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setenv("DATA_DIR", str(tmp_path))
    import importlib, main
    importlib.reload(main)
    import transcribe
    monkeypatch.setattr(main.transcriber, "transcribe", lambda p: {
        "language": "es", "duration": 3.0,
        "words": [{"text": "hola", "start": 0.0, "end": 1.0}, {"text": "mundo", "start": 1.0, "end": 3.0}]})
    return TestClient(main.app)


@pytest.fixture
def video(tmp_path):
    p = tmp_path / "in.mp4"
    subprocess.run(["ffmpeg", "-y", "-f", "lavfi", "-i", "testsrc=d=3:s=160x120:r=10",
                    "-f", "lavfi", "-i", "sine=d=3", "-shortest", "-pix_fmt", "yuv420p", str(p)],
                   check=True, capture_output=True)
    return p


def duration(path):
    r = subprocess.run(["ffprobe", "-v", "error", "-show_entries", "format=duration",
                        "-of", "csv=p=0", str(path)], capture_output=True, text=True)
    return float(r.stdout)


def test_merge_ranges():
    from editing import merge_ranges
    assert merge_ranges([(2, 3), (0, 1), (1, 1.5), (5, 5)]) == [[0, 1.5], [2, 3]]


def test_upload_transcribe_export(client, video, tmp_path):
    with open(video, "rb") as f:
        pid = client.post("/api/projects", files={"file": ("in.mp4", f, "video/mp4")}).json()["id"]
    p = client.get(f"/api/projects/{pid}").json()
    assert p["status"] == "done" and [w["text"] for w in p["words"]] == ["hola", "mundo"]
    assert client.get(f"/api/projects/{pid}/media").status_code == 200

    # conserva 0-1 y 2-3 (se corta 1-2): dura ~2 s
    r = client.post(f"/api/projects/{pid}/export", json={"keep": [[2, 3], [0, 1]]})
    assert r.status_code == 200
    out = tmp_path / "out.mp4"
    out.write_bytes(r.content)
    assert abs(duration(out) - 2.0) < 0.3


def test_export_empty_and_bad_id(client, video):
    with open(video, "rb") as f:
        pid = client.post("/api/projects", files={"file": ("in.mp4", f, "video/mp4")}).json()["id"]
    assert client.post(f"/api/projects/{pid}/export", json={"keep": []}).status_code == 400
    assert client.get("/api/projects/../etc").status_code == 404
    assert client.get("/api/projects/" + "0" * 32).status_code == 404
