import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { cutRanges, keepRanges, toSrt, toTxt } from './cuts.js'

const NO_WORDS = []

function download(blob, name) {
  const a = document.createElement('a')
  a.href = URL.createObjectURL(blob)
  a.download = name
  a.click()
  URL.revokeObjectURL(a.href)
}

const newVideo = () => { location.href = location.pathname }

function Upload({ onProject }) {
  const [busy, setBusy] = useState(false)
  const [over, setOver] = useState(false)
  const [error, setError] = useState('')

  async function send(file) {
    if (!file) return
    setBusy(true)
    setError('')
    try {
      const body = new FormData()
      body.append('file', file)
      const r = await fetch('/api/projects', { method: 'POST', body })
      if (!r.ok) throw new Error(`Error ${r.status}`)
      onProject((await r.json()).id)
    } catch (e) {
      setError(e.message)
      setBusy(false)
    }
  }

  return (
    <div className="center">
      <div
        className={`drop${over ? ' over' : ''}`}
        onDragOver={(e) => { e.preventDefault(); setOver(true) }}
        onDragLeave={() => setOver(false)}
        onDrop={(e) => { e.preventDefault(); setOver(false); send(e.dataTransfer.files[0]) }}
      >
        <h2>Subí un video o audio</h2>
        <p className="hint">Se transcribe localmente. Editá el texto y se edita el medio.</p>
        <label className="btn">
          {busy ? 'Subiendo…' : 'Elegir archivo'}
          <input type="file" accept="video/*,audio/*" hidden disabled={busy} onChange={(e) => send(e.target.files[0])} />
        </label>
        {error && <p className="err">{error}</p>}
      </div>
    </div>
  )
}

function Editor({ id }) {
  const [project, setProject] = useState(null)
  const [deleted, setDeleted] = useState(() => new Set())
  const [past, setPast] = useState([])
  const [future, setFuture] = useState([])
  const [active, setActive] = useState(-1)
  const [exporting, setExporting] = useState(false)
  const [error, setError] = useState('')
  const media = useRef(null)
  const textRef = useRef(null)

  useEffect(() => {
    let stop = false
    async function poll() {
      const r = await fetch(`/api/projects/${id}`)
      const p = await r.json()
      if (stop) return
      setProject(p)
      if (p.status === 'transcribing') setTimeout(poll, 2000)
    }
    poll().catch((e) => setError(e.message))
    return () => { stop = true }
  }, [id])

  const words = project?.words ?? NO_WORDS
  const duration = project?.duration ?? 0
  const cuts = useMemo(() => cutRanges(words, deleted, duration), [words, deleted, duration])

  const commit = useCallback((next) => {
    setPast((p) => [...p, deleted])
    setFuture([])
    setDeleted(next)
  }, [deleted])

  const undo = useCallback(() => {
    if (!past.length) return
    setFuture((f) => [deleted, ...f])
    setDeleted(past[past.length - 1])
    setPast((p) => p.slice(0, -1))
  }, [past, deleted])

  const redo = useCallback(() => {
    if (!future.length) return
    setPast((p) => [...p, deleted])
    setDeleted(future[0])
    setFuture((f) => f.slice(1))
  }, [future, deleted])

  // índices de palabras dentro de la selección nativa del navegador
  function selectedIndices() {
    const sel = window.getSelection()
    if (!sel || sel.isCollapsed || !textRef.current) return []
    const out = []
    textRef.current.querySelectorAll('[data-i]').forEach((el) => {
      if (sel.containsNode(el, true)) out.push(Number(el.dataset.i))
    })
    return out
  }

  const removeSelection = useCallback(() => {
    const idx = selectedIndices()
    if (!idx.length) return
    const next = new Set(deleted)
    idx.forEach((i) => next.add(i))
    commit(next)
    window.getSelection().removeAllRanges()
  }, [deleted, commit])

  const restoreSelection = useCallback(() => {
    const idx = selectedIndices()
    if (!idx.length) return
    const next = new Set(deleted)
    idx.forEach((i) => next.delete(i))
    commit(next)
    window.getSelection().removeAllRanges()
  }, [deleted, commit])

  useEffect(() => {
    function onKey(e) {
      const mod = e.ctrlKey || e.metaKey
      if (mod && e.key.toLowerCase() === 'z') { e.preventDefault(); if (e.shiftKey) redo(); else undo() }
      else if (mod && e.key.toLowerCase() === 'y') { e.preventDefault(); redo() }
      else if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); removeSelection() }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [undo, redo, removeSelection])

  // reproducción: salta los cortes y resalta la palabra actual
  useEffect(() => {
    let raf
    function tick() {
      const m = media.current
      if (m) {
        const t = m.currentTime
        const cut = cuts.find(([s, e]) => t >= s && t < e)
        if (cut && !m.paused) m.currentTime = cut[1]
        let lo = 0, hi = words.length - 1, found = -1
        while (lo <= hi) {
          const mid = (lo + hi) >> 1
          if (words[mid].start <= t) { found = mid; lo = mid + 1 } else hi = mid - 1
        }
        setActive(found >= 0 && t <= words[found].end + 0.3 ? found : -1)
      }
      raf = requestAnimationFrame(tick)
    }
    raf = requestAnimationFrame(tick)
    return () => cancelAnimationFrame(raf)
  }, [cuts, words])

  function seek(i) {
    if (window.getSelection()?.isCollapsed === false) return // arrastrando una selección
    if (media.current) media.current.currentTime = words[i].start
  }

  async function exportMedia() {
    setExporting(true)
    setError('')
    try {
      const keep = keepRanges(cuts, duration)
      const r = await fetch(`/api/projects/${id}/export`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ keep }),
      })
      if (!r.ok) throw new Error((await r.json()).detail || `Error ${r.status}`)
      const name = /filename="?([^"]+)"?/.exec(r.headers.get('content-disposition') || '')?.[1] || 'export'
      download(await r.blob(), name)
    } catch (e) {
      setError(e.message)
    } finally {
      setExporting(false)
    }
  }

  if (!project) return <div className="center">{error ? <p className="err">{error}</p> : 'Cargando…'}</div>
  if (project.status === 'error') return (
    <div className="center">
      <div>
        <p className="err">Falló la transcripción: {project.error}</p>
        <button onClick={newVideo}>Subir otro video</button>
      </div>
    </div>
  )
  if (project.status === 'transcribing') return <div className="center"><p>Transcribiendo… (la primera vez descarga el modelo)</p></div>

  const isVideo = /\.(mp4|mov|webm|mkv|avi|m4v)$/i.test(project.file)
  const Media = isVideo ? 'video' : 'audio'
  const base = (project.title || 'transcripcion').replace(/\.[^.]+$/, '')

  return (
    <div className="app">
      <aside className="side">
        <button className="ghost" onClick={newVideo}>← Nuevo video</button>
        <Media ref={media} src={`/api/projects/${id}/media`} controls />
        <div className="tools">
          <button className="ghost" onClick={undo} disabled={!past.length}>Deshacer</button>
          <button className="ghost" onClick={redo} disabled={!future.length}>Rehacer</button>
          <button className="ghost" onClick={removeSelection}>Borrar selección</button>
          <button className="ghost" onClick={restoreSelection}>Restaurar selección</button>
        </div>
        <div className="tools">
          <button onClick={exportMedia} disabled={exporting || !words.length}>{exporting ? 'Exportando…' : 'Exportar medio editado'}</button>
          <button className="ghost" onClick={() => download(new Blob([toSrt(words, deleted, cuts)]), `${base}.srt`)}>SRT</button>
          <button className="ghost" onClick={() => download(new Blob([toTxt(words, deleted)]), `${base}.txt`)}>TXT</button>
        </div>
        {error && <p className="err">{error}</p>}
        <p className="hint">Seleccioná texto y presioná Supr para cortarlo. Clic en una palabra para ir a ese punto. Ctrl+Z deshace.</p>
      </aside>
      <main className="main">
        <div className="transcript" ref={textRef}>
          {words.length === 0 && <p className="hint">No se detectó voz en el archivo.</p>}
          {words.map((w, i) => (
            <span key={i}>
              <span
                data-i={i}
                className={`w${deleted.has(i) ? ' deleted' : ''}${i === active ? ' active' : ''}`}
                onClick={() => seek(i)}
              >{w.text}</span>{' '}
            </span>
          ))}
        </div>
      </main>
    </div>
  )
}

export default function App() {
  const [id, setId] = useState(() => new URLSearchParams(location.search).get('p'))
  const open = (pid) => { history.replaceState(null, '', `?p=${pid}`); setId(pid) }
  return id ? <Editor id={id} /> : <Upload onProject={open} />
}
