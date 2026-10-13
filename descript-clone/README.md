# Descript clone (MVP)

Subís un video/audio, se transcribe localmente (faster-whisper) y al borrar palabras del texto se corta el medio (ffmpeg).

## Requisitos
Python 3.10+, Node 20+, ffmpeg/ffprobe en el PATH.

## Correr
```bash
# backend (puerto 8000)
cd backend && python -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --reload

# frontend (puerto 5173, proxy /api -> 8000)
cd frontend && npm install && npm run dev
```
Modelo: `WHISPER_MODEL=base` por defecto (`tiny`, `small`, `medium`...). Se descarga de HuggingFace la primera vez. Datos en `backend/data/`.

## Tests
`cd backend && pip install -r requirements-dev.txt && pytest` · `cd frontend && npm test`

## Límites conocidos
- Los cortes se hacen entre palabras según los timestamps de Whisper (error de decenas de ms); sin crossfade, puede haber clicks.
- Exportar re-codifica (libx264/aac): lento en videos largos. Sin cola ni progreso.
- Sin autenticación, sin borrado de proyectos y sin límite de tamaño de subida: solo para uso local.
- Editar solo borra/restaura palabras; no se puede reescribir ni reordenar texto.
