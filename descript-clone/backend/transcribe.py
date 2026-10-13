import os

_model = None


def get_model():
    global _model
    if _model is None:
        from faster_whisper import WhisperModel

        _model = WhisperModel(
            os.environ.get("WHISPER_MODEL", "base"),
            device="auto",
            compute_type="auto",
        )
    return _model


def transcribe(path: str) -> dict:
    """Devuelve {"language": str, "duration": float, "words": [{"text","start","end"}]}."""
    segments, info = get_model().transcribe(path, word_timestamps=True, vad_filter=True)
    words = []
    for seg in segments:
        for w in seg.words or []:
            text = w.word.strip()
            if text:
                words.append({"text": text, "start": round(w.start, 3), "end": round(w.end, 3)})
    return {"language": info.language, "duration": info.duration, "words": words}
