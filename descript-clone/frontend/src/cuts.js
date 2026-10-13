// words: [{text,start,end}], deleted: Set<índice>

/** Rangos a cortar: cada tramo de palabras borradas va desde su inicio hasta el inicio de la siguiente palabra conservada. */
export function cutRanges(words, deleted, duration) {
  const cuts = []
  for (let i = 0; i < words.length; i++) {
    if (!deleted.has(i)) continue
    let j = i
    while (j + 1 < words.length && deleted.has(j + 1)) j++
    const end = j + 1 < words.length ? words[j + 1].start : Math.max(duration, words[j].end)
    cuts.push([words[i].start, end])
    i = j
  }
  return cuts
}

/** Complemento de los cortes sobre [0, duration]. */
export function keepRanges(cuts, duration) {
  const keep = []
  let t = 0
  for (const [s, e] of cuts) {
    if (s > t) keep.push([t, s])
    t = Math.max(t, e)
  }
  if (t < duration) keep.push([t, duration])
  return keep
}

/** Tiempo original -> tiempo en el medio editado. */
export function remap(t, cuts) {
  let removed = 0
  for (const [s, e] of cuts) {
    if (t >= e) removed += e - s
    else if (t > s) removed += t - s
  }
  return t - removed
}

function stamp(t) {
  const ms = Math.round(t * 1000)
  const p = (n, w) => String(n).padStart(w, '0')
  return `${p(Math.floor(ms / 3600000), 2)}:${p(Math.floor(ms / 60000) % 60, 2)}:${p(Math.floor(ms / 1000) % 60, 2)},${p(ms % 1000, 3)}`
}

/** SRT sobre el medio editado; una cue cada ~8 palabras o al cerrar frase. */
export function toSrt(words, deleted, cuts) {
  const kept = words.filter((_, i) => !deleted.has(i))
  const cues = []
  let cur = []
  for (const w of kept) {
    cur.push(w)
    if (cur.length >= 8 || /[.?!]$/.test(w.text)) {
      cues.push(cur)
      cur = []
    }
  }
  if (cur.length) cues.push(cur)
  return cues
    .map((c, n) => {
      const s = remap(c[0].start, cuts)
      const e = Math.max(remap(c[c.length - 1].end, cuts), s + 0.1)
      return `${n + 1}\n${stamp(s)} --> ${stamp(e)}\n${c.map((w) => w.text).join(' ')}\n`
    })
    .join('\n')
}

export const toTxt = (words, deleted) =>
  words.filter((_, i) => !deleted.has(i)).map((w) => w.text).join(' ')
