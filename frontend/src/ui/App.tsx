import React, { useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'

const COLOR_LABELS: Record<string, string> = {
  yellow: 'Giallo',
  bright_green: 'Verde brillante',
  turquoise: 'Turchese',
  pink: 'Rosa',
  blue: 'Blu',
  red: 'Rosso',
  dark_blue: 'Blu scuro',
  teal: 'Verde acqua',
  green: 'Verde',
  violet: 'Viola',
  dark_red: 'Rosso scuro',
  dark_yellow: 'Giallo scuro',
  gray_50: 'Grigio 50%',
  gray_25: 'Grigio 25%',
  black: 'Nero',
  white: 'Bianco',
  light_gray: 'Grigio chiaro',
  dark_gray: 'Grigio scuro',
  orange: 'Arancione',
  light_orange: 'Arancione chiaro',
  light_yellow: 'Giallo chiaro',
  light_green: 'Verde chiaro',
  light_blue: 'Blu chiaro',
  light_turquoise: 'Turchese chiaro',
  magenta: 'Magenta',
  dark_teal: 'Verde acqua scuro',
  dark_magenta: 'Magenta scuro',
}

function normalizeColorKey(raw: string | null | undefined): string {
  return (raw || '')
    .toLowerCase()
    .trim()
    .replace(/\s*\(\d+\)\s*$/, '')
}

function colorLabel(raw: string | null | undefined): string {
  const key = normalizeColorKey(raw)
  if (!key) return '—'
  if (COLOR_LABELS[key]) return COLOR_LABELS[key]
  return key
    .split(/[_\s]+/)
    .map(part => (part ? part.charAt(0).toUpperCase() + part.slice(1) : part))
    .join(' ')
}

function sanitizeColorMap(map: Record<string, string> | undefined | null): Record<string, string> {
  if (!map) return {}
  let changed = false
  const result: Record<string, string> = {}
  for (const [key, value] of Object.entries(map)) {
    const trimmed = (value || '').trim()
    const keyNorm = normalizeColorKey(key)
    const valueNorm = normalizeColorKey(trimmed)
    if (!trimmed || (valueNorm && valueNorm === keyNorm) || (!valueNorm && keyNorm)) {
      const desired = colorLabel(key)
      result[key] = desired
      if (desired !== value) changed = true
    } else {
      result[key] = trimmed
      if (trimmed !== value) changed = true
    }
  }
  if (!changed) {
    let same = Object.keys(result).length === Object.keys(map).length
    if (same) {
      for (const [k, v] of Object.entries(map)) {
        if (result[k] !== v) {
          same = false
          break
        }
      }
    }
    if (same) return map
  }
  return result
}

function colorCategoryFromMap(colorKey: string | null | undefined, map: Record<string, string>): string {
  if (!colorKey) return ''
  const attempts = Array.from(
    new Set([
      colorKey,
      colorKey.toLowerCase(),
      normalizeColorKey(colorKey),
      normalizeColorKey(colorKey.toLowerCase()),
    ].filter(Boolean))
  ) as string[]
  for (const key of attempts) {
    const val = (map[key] || '').trim()
    if (val) return val
  }
  return ''
}

function sanitizeCategoryOverrides(map: Record<string, string> | undefined | null): Record<string, string> {
  if (!map) return {}
  let changed = false
  const result: Record<string, string> = {}
  for (const [key, value] of Object.entries(map)) {
    const trimmed = (value || '').trim()
    if (trimmed) {
      const normalized = normalizeColorKey(trimmed)
      if (normalized && COLOR_LABELS[normalized]) {
        const desired = colorLabel(trimmed)
        result[key] = desired
        if (desired !== value) changed = true
        continue
      }
    }
    result[key] = trimmed
    if (trimmed !== value) changed = true
  }
  if (!changed) {
    let same = Object.keys(result).length === Object.keys(map).length
    if (same) {
      for (const [k, v] of Object.entries(map)) {
        if (result[k] !== v) {
          same = false
          break
        }
      }
    }
    if (same) return map
  }
  return result
}

interface Highlight {
  filename: string
  type: 'highlight'
  highlight_color: string | null
  text: string
  context: string
  paragraph: string
  offset_start: number
  offset_end: number
  // server-provided paragraph index for stable linking
  para_index?: number | null
}

interface CommentItem {
  id: number
  author: string
  date: string
  text: string
  quoted: string
  filename?: string
  code?: string | null
}

interface ParagraphInfo {
  filename: string
  para_index: number | null
  text: string
}

interface ApiResponse {
  highlights: Highlight[]
  comments: CommentItem[]
  paragraphs?: ParagraphInfo[]
}

type FileMeta = {
  tipo: 'Intervista' | 'Focus group'
  intervistatore: string
  intervistato?: string
  ruolo?: string
  scuola: string
  gruppo?: string
  noteGruppo?: string
}

// Default to same-origin (Nginx will proxy /api to backend). Can override with VITE_API_BASE.
const API_BASE = import.meta.env.VITE_API_BASE || ''

function App() {
  const [data, setData] = useState<ApiResponse | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [colorMap, setColorMap] = useState<Record<string, string>>({})
  const [codeMap, setCodeMap] = useState<Record<string, string>>({})
  const [tab, setTab] = useState<'viewer' | 'comments' | 'dashboard' | 'kanban' | 'documenti' | 'impostazioni'>('viewer')
  const [meta, setMeta] = useState<Record<string, FileMeta>>({}) // per-filename metadata (client-side)
  const [catOverride, setCatOverride] = useState<Record<string, string>>({}) // key filename::id -> category override
  const [viewerDoc, setViewerDoc] = useState<string>('')
  // Category label → color mapping (for UI labels)
  const [categoryColors, setCategoryColors] = useState<Record<string, string>>({})
  // Documenti filters
  const [dTipo, setDTipo] = useState<string[]>([])
  const [dIntervistatore, setDIntervistatore] = useState<string[]>([])
  const [dIntervistato, setDIntervistato] = useState<string[]>([])
  const [dRuolo, setDRuolo] = useState<string[]>([])
  const [dGruppo, setDGruppo] = useState<string[]>([])
  const [dScuola, setDScuola] = useState<string[]>([])
  const [dQuery, setDQuery] = useState<string>('')

  // comment filters
  const [fDoc, setFDoc] = useState<string[]>([])
  const [fCode, setFCode] = useState<string[]>([])
  const [fColor, setFColor] = useState<string[]>([])
  const [fQuery, setFQuery] = useState<string>('')
  const [fCategory, setFCategory] = useState<string[]>([])
  const [fIntervistatore, setFIntervistatore] = useState<string[]>([])
  const [fIntervistato, setFIntervistato] = useState<string[]>([])
  const [fRuolo, setFRuolo] = useState<string[]>([])
  const [fScuola, setFScuola] = useState<string[]>([])
  const [fGruppo, setFGruppo] = useState<string[]>([])
  // Kanban filters and view
  const [kDoc, setKDoc] = useState<string[]>([])
  const [kRuolo, setKRuolo] = useState<string[]>([])
  const [kGruppo, setKGruppo] = useState<string[]>([])
  const [kCode, setKCode] = useState<string[]>([])
  const [kColor, setKColor] = useState<string[]>([])
  const [kCategory, setKCategory] = useState<string[]>([])
  const [kIntervistatore, setKIntervistatore] = useState<string[]>([])
  const [kIntervistato, setKIntervistato] = useState<string[]>([])
  const [kScuola, setKScuola] = useState<string[]>([])
  const [kQuery, setKQuery] = useState<string>("")
  const [kView, setKView] = useState<'colore' | 'xx'>('colore')

  const paraRef = useRef<HTMLDivElement>(null)
  const pendingJump = useRef<Highlight | null>(null)
  const [jumpSeq, setJumpSeq] = useState(0)

  const paragraphs = useMemo(() => {
    if (!data) return [] as { filename: string; text: string; highlights: Highlight[] }[]
    // group by filename + para_index when available, otherwise fallback to paragraph text
    const groups: Record<string, { filename: string; text: string; highs: Highlight[] }> = {}
    for (const h of data.highlights) {
      const key = `${h.filename}#${h.para_index ?? 'np'}#${h.paragraph}`
      if (!groups[key]) groups[key] = { filename: h.filename, text: h.paragraph, highs: [] }
      groups[key].highs.push(h)
    }
    return Object.values(groups).map(({ filename, text, highs }) => ({ filename, text, highlights: highs.sort((a,b)=>a.offset_start-b.offset_start) }))
  }, [data])

  const highlightsByFile = useMemo(() => {
    if (!data) return {} as Record<string, Highlight[]>
    const map: Record<string, Highlight[]> = {}
    for (const h of data.highlights) {
      ;(map[h.filename] ||= []).push(h)
    }
    return map
  }, [data])

  const paragraphsByFile = useMemo(() => {
    if (!data || !data.paragraphs) return {} as Record<string, ParagraphInfo[]>
    const map: Record<string, ParagraphInfo[]> = {}
    for (const p of data.paragraphs) {
      ;(map[p.filename] ||= []).push(p)
    }
    return map
  }, [data])

  // Keep viewerDoc aligned with uploaded data
  React.useEffect(() => {
    if (!data) return
    const files = Array.from(new Set(data.highlights.map(h => h.filename)))
    if (!viewerDoc || !files.includes(viewerDoc)) {
      setViewerDoc(files[0] || '')
    }
  }, [data])

  const categories = useMemo(() => {
    if (!data) return [] as string[]
    const colors = Array.from(new Set(data.highlights.map(h => (h.highlight_color||'').toLowerCase()).filter(Boolean)))
    // ensure map has keys
    const m: Record<string,string> = { ...colorMap }
    colors.forEach(c => {
      const current = (m[c] || '').trim()
      if (!current || normalizeColorKey(current) === normalizeColorKey(c)) {
        m[c] = colorLabel(c)
      }
    })
    if (JSON.stringify(m) !== JSON.stringify(colorMap)) setColorMap(m)
    return colors
  }, [data])

  React.useEffect(() => {
    const sanitized = sanitizeColorMap(colorMap)
    if (sanitized !== colorMap) {
      setColorMap(sanitized)
    }
  }, [colorMap])

  React.useEffect(() => {
    const sanitized = sanitizeCategoryOverrides(catOverride)
    if (sanitized !== catOverride) {
      setCatOverride(sanitized)
    }
  }, [catOverride])

  // Persistence: load from localStorage once
  React.useEffect(() => {
    try {
      const savedMeta = localStorage.getItem('qd_meta')
      const savedMap = localStorage.getItem('qd_colorMap')
      const savedCat = localStorage.getItem('qd_catOverride')
  const savedCatColors = localStorage.getItem('qd_categoryColors')
  const savedCodeMap = localStorage.getItem('qd_codeMap')
      if (savedMeta) setMeta(JSON.parse(savedMeta))
      if (savedMap) setColorMap(sanitizeColorMap(JSON.parse(savedMap)))
      if (savedCat) setCatOverride(sanitizeCategoryOverrides(JSON.parse(savedCat)))
  if (savedCatColors) setCategoryColors(JSON.parse(savedCatColors))
  if (savedCodeMap) setCodeMap(JSON.parse(savedCodeMap))
    } catch {}
  }, [])
  React.useEffect(() => {
    try { localStorage.setItem('qd_meta', JSON.stringify(meta)) } catch {}
  }, [meta])
  React.useEffect(() => {
    try { localStorage.setItem('qd_colorMap', JSON.stringify(colorMap)) } catch {}
  }, [colorMap])
  React.useEffect(() => {
    try { localStorage.setItem('qd_codeMap', JSON.stringify(codeMap)) } catch {}
  }, [codeMap])
  React.useEffect(() => {
    try { localStorage.setItem('qd_catOverride', JSON.stringify(catOverride)) } catch {}
  }, [catOverride])
  React.useEffect(() => {
    try { localStorage.setItem('qd_categoryColors', JSON.stringify(categoryColors)) } catch {}
  }, [categoryColors])

  // Server persistence (SQLite via backend)
  const loadedFromServer = React.useRef(false)
  React.useEffect(() => {
    if (loadedFromServer.current) return
    ;(async () => {
      try {
        const res = await fetch(`${API_BASE}/api/state`)
        if (!res.ok) return
        const st = await res.json()
        if (st && typeof st === 'object') {
          if (st.meta) setMeta((prev) => ({ ...prev, ...st.meta }))
          if (st.colorMap) setColorMap((prev) => sanitizeColorMap({ ...prev, ...st.colorMap }))
          if (st.codeMap) setCodeMap((prev) => ({ ...prev, ...st.codeMap }))
          if (st.categoryColors) setCategoryColors((prev) => ({ ...prev, ...st.categoryColors }))
          if (st.catOverride) setCatOverride((prev) => sanitizeCategoryOverrides({ ...prev, ...st.catOverride }))
        }
        loadedFromServer.current = true
      } catch {
        // ignore offline/backend issues; localStorage still works
      }
    })()
  }, [])

  // Debounced save to server on changes
  const saveTimer = React.useRef<number | undefined>(undefined)
  React.useEffect(() => {
    // avoid posting before initial load attempt
    if (!loadedFromServer.current) return
    if (saveTimer.current) window.clearTimeout(saveTimer.current)
    saveTimer.current = window.setTimeout(() => {
      ;(async () => {
        try {
          await fetch(`${API_BASE}/api/state`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ colorMap, codeMap, categoryColors, catOverride, meta }),
          })
        } catch {
          // ignore
        }
      })()
    }, 600)
    return () => { if (saveTimer.current) window.clearTimeout(saveTimer.current) }
  }, [colorMap, codeMap, categoryColors, catOverride, meta])

  const stats = useMemo(() => {
    if (!data) return [] as { category: string; count: number }[]
    const counts: Record<string, number> = {}
    for (const h of data.highlights) {
      const color = (h.highlight_color||'').toLowerCase()
      const rawCat = colorCategoryFromMap(color, colorMap)
      const cat = color ? (rawCat || colorLabel(color)) : 'non-categorizzato'
      counts[cat] = (counts[cat] || 0) + 1
    }
    return Object.entries(counts).map(([category, count]) => ({ category, count })).sort((a,b)=>b.count-a.count)
  }, [data, colorMap])

  // Initialize categoryColors for any discovered categories (keep existing values)
  React.useEffect(() => {
    const allCats = new Set<string>([...stats.map(s => s.category), 'non-categorizzato'])
    const next: Record<string,string> = { ...categoryColors }
    // default palette
    const palette = ['#ffd43b','#63e6be','#91a7ff','#ff8787','#ffd8a8','#a5d8ff','#b2f2bb','#fcc2d7','#bac8ff','#eebefa']
    let i = 0
    allCats.forEach(cat => {
      if (!next[cat]) {
        next[cat] = palette[i % palette.length]
        i++
      }
    })
    if (JSON.stringify(next) !== JSON.stringify(categoryColors)) setCategoryColors(next)
  }, [stats])

  // build a stable signature for each highlight to link list -> inline span
  function sig(h: Highlight) {
    return `${h.filename}|${h.para_index ?? 'np'}|${h.offset_start}|${h.offset_end}|${h.text}`
  }

  function doScrollTo(id: string) {
    const el = document.getElementById(id)
    if (!el) return false
    el.scrollIntoView({ behavior: 'smooth', block: 'center' })
    const prev = el.getAttribute('data-oldbox')
    const old = (el as HTMLElement).style.boxShadow
    el.setAttribute('data-oldbox', old || '')
    ;(el as HTMLElement).style.boxShadow = '0 0 0 3px rgba(0,0,0,0.35) inset'
    setTimeout(() => {
      if (el instanceof HTMLElement) {
        el.style.boxShadow = prev ?? ''
      }
    }, 650)
    return true
  }

  function scheduleJump(h: Highlight) {
    pendingJump.current = h
    setJumpSeq((seq) => seq + 1)
  }

  function jumpTo(h: Highlight) {
    scheduleJump(h)
    setViewerDoc((prev) => (prev === h.filename ? prev : h.filename))
    if (tab !== 'viewer') setTab('viewer')
  }

  React.useEffect(() => {
    if (!pendingJump.current) return
    let attempts = 0
    let cancelled = false
    const tick = () => {
      if (cancelled) return
      const target = pendingJump.current
      if (!target) return
      // ensure viewer tab + document selected
      if (tab !== 'viewer') {
        if (attempts++ < 20) {
          setTimeout(tick, 80)
        }
        return
      }
      if (viewerDoc && viewerDoc !== target.filename) {
        if (attempts++ < 20) {
          setTimeout(tick, 80)
        }
        return
      }
      const id = `hl-${sig(target)}`
      if (doScrollTo(id)) {
        pendingJump.current = null
        return
      }
      if (attempts++ < 20) {
        setTimeout(tick, 80)
      } else {
        pendingJump.current = null
      }
    }
    tick()
    return () => {
      cancelled = true
    }
  }, [jumpSeq, tab, viewerDoc, data])
  // (Removed stray closing brace that caused build error)

  // Enrich comments by linking to a matching highlight (via quoted text)
  const commentsEnriched = useMemo(() => {
    if (!data) return [] as (CommentItem & { highlight?: Highlight; color?: string | null; category?: string; highlightText?: string })[]
    const norm = (s: string) => (s || '').replace(/\s+/g, ' ').trim().toLowerCase()
    const highsAll = data.highlights
    return data.comments.map(c => {
      const qNorm = norm(c.quoted || '')
      const docHighs = c.filename ? (highlightsByFile[c.filename] || []) : highsAll
      let linked: Highlight | undefined
      if (qNorm && docHighs.length) {
        const candidateParas = new Set<string>()
        if (c.filename) {
          for (const p of paragraphsByFile[c.filename] || []) {
            if (norm(p.text).includes(qNorm)) {
              if (p.para_index !== null && p.para_index !== undefined) {
                candidateParas.add(`idx:${p.para_index}`)
              }
              candidateParas.add(`txt:${norm(p.text)}`)
            }
          }
        }
        const filterMatches = (source: Highlight[]) => source.filter(h => {
          const keyIdx = h.para_index !== null && h.para_index !== undefined ? `idx:${h.para_index}` : null
          const keyTxt = `txt:${norm(h.paragraph)}`
          const inParagraph = candidateParas.size === 0 || candidateParas.has(keyTxt) || (keyIdx ? candidateParas.has(keyIdx) : false)
          if (!inParagraph) return false
          const ht = norm(h.text)
          return ht && (ht.includes(qNorm) || qNorm.includes(ht))
        })
        let candidates = filterMatches(docHighs)
        if (!candidates.length) {
          candidates = docHighs.filter(h => {
            const ht = norm(h.text)
            return ht && (ht.includes(qNorm) || qNorm.includes(ht))
          })
        }
        if (candidates.length) {
          const scored = candidates
            .map(h => {
              const ht = norm(h.text)
              const exact = ht === qNorm
              const contains = ht.includes(qNorm)
              const contained = qNorm.includes(ht)
              const diff = Math.abs(ht.length - qNorm.length)
              const idxScore = h.para_index ?? Number.MAX_SAFE_INTEGER
              return { h, exact, contains, contained, diff, idxScore, textLen: ht.length }
            })
            .sort((a, b) => {
              if (a.exact !== b.exact) return a.exact ? -1 : 1
              if (a.contains !== b.contains) return a.contains ? -1 : 1
              if (a.contained !== b.contained) return a.contained ? -1 : 1
              if (a.diff !== b.diff) return a.diff - b.diff
              if (a.idxScore !== b.idxScore) return a.idxScore - b.idxScore
              return a.textLen - b.textLen
            })
          linked = scored[0]?.h
        }
      }
      const color = linked ? (linked.highlight_color || null) : null
      const colorKey = (color || '').toLowerCase()
      const mapped = colorCategoryFromMap(colorKey, colorMap)
      const defaultCat = colorKey ? (mapped || colorLabel(colorKey)) : 'non-categorizzato'
      const keyOverride = `${c.filename || ''}::${c.id}`
      const effectiveCat = catOverride[keyOverride] || defaultCat
      return { ...c, highlight: linked, color, category: effectiveCat, highlightText: linked?.text || '' }
    })
  }, [data, colorMap, catOverride, highlightsByFile, paragraphsByFile])

  const commentDocs = useMemo(() => {
    const set = new Set<string>()
    commentsEnriched.forEach(c => c.filename && set.add(c.filename))
    return Array.from(set)
  }, [commentsEnriched])

  const commentColors = useMemo(() => {
    const set = new Set<string>()
    commentsEnriched.forEach(c => c.color && set.add((c.color || '').toLowerCase()))
    return Array.from(set)
  }, [commentsEnriched])

  const filteredComments = useMemo(() => {
    let arr = commentsEnriched
    if (fDoc.length) arr = arr.filter(c => fDoc.includes((c.filename || '')))
    if (fCode.length) arr = arr.filter(c => fCode.includes((c.code || '').toLowerCase()))
    if (fColor.length) arr = arr.filter(c => fColor.includes((c.color || '').toLowerCase()))
    if (fCategory.length) arr = arr.filter(c => fCategory.includes((c.category || '').toLowerCase()))
    // metadata-based filters by filename
    if (fIntervistatore.length) arr = arr.filter(c => {
      const fm = c.filename ? meta[c.filename] : undefined
      return fm ? fIntervistatore.map(s=>s.toLowerCase()).includes((fm.intervistatore||'').toLowerCase()) : false
    })
    if (fIntervistato.length) arr = arr.filter(c => {
      const fm = c.filename ? meta[c.filename] : undefined
      return fm ? fIntervistato.map(s=>s.toLowerCase()).includes((fm.intervistato||'').toLowerCase()) : false
    })
    if (fRuolo.length) arr = arr.filter(c => {
      const fm = c.filename ? meta[c.filename] : undefined
      return fm ? fRuolo.map(s=>s.toLowerCase()).includes((fm.ruolo||'').toLowerCase()) : false
    })
    if (fScuola.length) arr = arr.filter(c => {
      const fm = c.filename ? meta[c.filename] : undefined
      return fm ? fScuola.map(s=>s.toLowerCase()).includes((fm.scuola||'').toLowerCase()) : false
    })
    if (fGruppo.length) arr = arr.filter(c => {
      const fm = c.filename ? meta[c.filename] : undefined
      return fm ? fGruppo.map(s=>s.toLowerCase()).includes((fm.gruppo||'').toLowerCase()) : false
    })
    if (fQuery.trim()) {
      const q = fQuery.trim().toLowerCase()
      arr = arr.filter(c =>
        (c.text || '').toLowerCase().includes(q) ||
        (c.quoted || '').toLowerCase().includes(q) ||
        (c.author || '').toLowerCase().includes(q) ||
        ((c.highlightText || c.highlight?.text || '').toLowerCase().includes(q))
      )
    }
    return arr
  }, [commentsEnriched, fDoc, fCode, fColor, fCategory, fQuery, fIntervistatore, fIntervistato, fRuolo, fScuola, fGruppo, meta])

  // Helper to read multiple selected options
  const readMulti = (e: React.ChangeEvent<HTMLSelectElement>) => Array.from(e.target.selectedOptions).map(o => o.value)

  const commentCodes = useMemo(() => {
    const set = new Set<string>()
    commentsEnriched.forEach(c => c.code && set.add((c.code||'').toLowerCase()))
    return Array.from(set)
  }, [commentsEnriched])

  // CSV export (array -> CSV blob)
  function toCSV(rows: any[]): string {
    if (!rows.length) return ''
    const headers = Object.keys(rows[0])
    const esc = (v: any) => {
      const s = v == null ? '' : String(v)
      if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"'
      return s
    }
    return [headers.join(','), ...rows.map(r => headers.map(h => esc(r[h])).join(','))].join('\n')
  }
  function download(filename: string, content: string, type = 'text/csv;charset=utf-8') {
    const blob = new Blob([content], { type })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = filename
    a.click()
    URL.revokeObjectURL(url)
  }
  function exportComments(rows: (typeof commentsEnriched)[number][]) {
    const mapped = rows.map(c => {
      const m = c.filename ? meta[c.filename] : undefined
      return {
        filename: c.filename || '',
        id: c.id,
        author: c.author,
        date: c.date,
        code: c.code || '',
        category: c.category || '',
        color: colorLabel(c.color),
        quoted: c.quoted || '',
        highlighted: c.highlightText || c.highlight?.text || '',
        comment: c.text || '',
        tipo: m?.tipo || '',
        intervistatore: m?.intervistatore || '',
        intervistato: m?.intervistato || '',
        ruolo: m?.ruolo || '',
        gruppo: m?.gruppo || '',
        scuola: m?.scuola || '',
      }
    })
    const csv = toCSV(mapped)
    download(`commenti_${new Date().toISOString().slice(0,10)}.csv`, csv)
  }

  function exportCommentsXLSX(rows: (typeof commentsEnriched)[number][]) {
    const mapped = rows.map(c => {
      const m = c.filename ? meta[c.filename] : undefined
      return {
        File: c.filename || '',
        ID: c.id,
        Autore: c.author,
        Data: c.date,
        Codice: c.code || '',
        Categoria: c.category || '',
        Colore: colorLabel(c.color),
        "Testo selezionato": c.quoted || '',
        "Testo evidenziato": c.highlightText || c.highlight?.text || '',
        Commento: c.text || '',
        Tipo: m?.tipo || '',
        Intervistatore: m?.intervistatore || '',
        Intervistato: m?.intervistato || '',
        Ruolo: m?.ruolo || '',
        Gruppo: m?.gruppo || '',
        Scuola: m?.scuola || '',
      }
    })
    const interviste = mapped.filter(r => r.Tipo === 'Intervista')
    const focus = mapped.filter(r => r.Tipo === 'Focus group')
    const wb = XLSX.utils.book_new()
    const ws1 = XLSX.utils.json_to_sheet(interviste)
    const ws2 = XLSX.utils.json_to_sheet(focus)
    XLSX.utils.book_append_sheet(wb, ws1, 'Interviste')
    XLSX.utils.book_append_sheet(wb, ws2, 'Focus group')
    XLSX.writeFile(wb, `commenti_${new Date().toISOString().slice(0,10)}.xlsx`)
  }

  const optionsFromMeta = useMemo(() => {
    const docs = Object.keys(meta)
    const setI = new Set<string>()
    const setT = new Set<string>()
    const setR = new Set<string>()
    const setS = new Set<string>()
    const setG = new Set<string>()
    for (const d of docs) {
      const fm = meta[d]
      if (fm?.intervistatore) setI.add(fm.intervistatore)
      if (fm?.intervistato) setT.add(fm.intervistato)
      if (fm?.ruolo) setR.add(fm.ruolo)
      if (fm?.scuola) setS.add(fm.scuola)
      if (fm?.gruppo) setG.add(fm.gruppo)
    }
    return {
      intervistatori: Array.from(setI),
      intervistati: Array.from(setT),
      ruoli: Array.from(setR),
      scuole: Array.from(setS),
      gruppi: Array.from(setG),
    }
  }, [meta])

  async function handleUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const files = e.target.files
    if (!files || files.length === 0) return
    setLoading(true)
    setError(null)
    try {
      const form = new FormData()
      for (const f of Array.from(files)) form.append('files', f)
      const res = await fetch(`${API_BASE}/api/upload-multi`, { method: 'POST', body: form })
      if (!res.ok) throw new Error(`HTTP ${res.status}`)
      const json = await res.json()
      setData(json)
      // initialize empty metadata for each filename
      const m: Record<string, FileMeta> = { ...meta }
      const filenames: string[] = Array.from(new Set<string>(json.highlights.map((h: Highlight) => h.filename)))
      filenames.forEach(fn => {
        if (!m[fn]) m[fn] = { tipo: 'Intervista', intervistatore: '', intervistato: '', ruolo: '', scuola: '', gruppo: '', noteGruppo: '' }
      })
      setMeta(m)
    } catch (err: any) {
      setError(err.message || String(err))
    } finally {
      setLoading(false)
    }
  }

  // Load persisted docs on mount
  React.useEffect(() => {
    (async () => {
      try {
        const res = await fetch(`${API_BASE}/api/docs`)
        if (res.ok) {
          const j = await res.json()
          setData(j)
        }
      } catch {}
    })()
  }, [])

  async function handleDeleteFile(fn: string) {
    if (!confirm(`Rimuovere "${fn}"?`)) return
    setLoading(true)
    try {
      const res = await fetch(`${API_BASE}/api/docs/${encodeURIComponent(fn)}`, { method: 'DELETE' })
      if (!res.ok) throw new Error(`HTTP ${res.status}`)
      const j = await res.json()
      setData({ highlights: j.highlights || [], comments: j.comments || [], paragraphs: j.paragraphs || [] })
      setMeta(prev => { const n = { ...prev }; delete n[fn]; return n })
    } catch (err: any) {
      setError(err.message || String(err))
    } finally {
      setLoading(false)
    }
  }

  return (
  <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', width: '100vw', boxSizing: 'border-box', background: '#fafbff', overflowX: 'hidden' }}>
      {/* Top navbar */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr auto 1fr', alignItems: 'center', padding: '10px 16px', background: '#fff', borderBottom: '1px solid #e9ecef', boxShadow: '0 1px 2px rgba(0,0,0,0.03)' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <span style={{ padding: '6px 10px', background: 'linear-gradient(90deg,#6741D9,#22B8CF)', color: 'white', borderRadius: 8, fontWeight: 700 }}>QualiDoc</span>
          <span style={{ color: '#495057' }}>annotazioni, commenti e insight</span>
        </div>
        <div style={{ display: 'flex', gap: 8, justifySelf: 'center' }}>
          <button onClick={() => setTab('viewer')} aria-label="Viewer" title="Viewer" style={{ padding: '6px 10px', borderRadius: 6, border: tab==='viewer' ? '2px solid #444' : '1px solid #ddd', background: tab==='viewer' ? '#f0f0f0' : '#fff', display:'inline-flex', alignItems:'center', justifyContent:'center' }}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M12 5C7 5 3.27 8.11 2 12c1.27 3.89 5 7 10 7s8.73-3.11 10-7c-1.27-3.89-5-7-10-7Zm0 12a5 5 0 1 1 0-10 5 5 0 0 1 0 10Z" fill="currentColor"/></svg>
          </button>
          <button onClick={() => setTab('comments')} aria-label="Commenti" title="Commenti" style={{ padding: '6px 10px', borderRadius: 6, border: tab==='comments' ? '2px solid #444' : '1px solid #ddd', background: tab==='comments' ? '#f0f0f0' : '#fff', display:'inline-flex', alignItems:'center', justifyContent:'center' }}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4 4h16v12H7l-3 3V4Z" stroke="currentColor" strokeWidth="2"/></svg>
          </button>
          <button onClick={() => setTab('kanban')} aria-label="Kanban" title="Kanban" style={{ padding: '6px 10px', borderRadius: 6, border: tab==='kanban' ? '2px solid #444' : '1px solid #ddd', background: tab==='kanban' ? '#f0f0f0' : '#fff', display:'inline-flex', alignItems:'center', justifyContent:'center' }}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4 5h4v14H4V5Zm6 0h4v9h-4V5Zm6 0h4v6h-4V5Z" fill="currentColor"/></svg>
          </button>
          <button onClick={() => setTab('documenti')} aria-label="Documenti" title="Documenti" style={{ padding: '6px 10px', borderRadius: 6, border: tab==='documenti' ? '2px solid #444' : '1px solid #ddd', background: tab==='documenti' ? '#f0f0f0' : '#fff', display:'inline-flex', alignItems:'center', justifyContent:'center' }}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M6 2h9l5 5v15H6V2Zm9 1.5V8h4.5" stroke="currentColor" strokeWidth="2"/></svg>
          </button>
          <button onClick={() => setTab('dashboard')} aria-label="Dashboard" title="Dashboard" style={{ padding: '6px 10px', borderRadius: 6, border: tab==='dashboard' ? '2px solid #444' : '1px solid #ddd', background: tab==='dashboard' ? '#f0f0f0' : '#fff', display:'inline-flex', alignItems:'center', justifyContent:'center' }}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M3 13h8V3H3v10Zm10 8h8V3h-8v18ZM3 21h8v-6H3v6Z" fill="currentColor"/></svg>
          </button>
          <button onClick={() => setTab('impostazioni')} aria-label="Impostazioni" title="Impostazioni" style={{ padding: '6px 10px', borderRadius: 6, border: tab==='impostazioni' ? '2px solid #444' : '1px solid #ddd', background: tab==='impostazioni' ? '#f0f0f0' : '#fff', display:'inline-flex', alignItems:'center', justifyContent:'center' }}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M12 15.5a3.5 3.5 0 1 0 0-7 3.5 3.5 0 0 0 0 7Z" stroke="currentColor" strokeWidth="2"/><path d="M19.4 15a7.96 7.96 0 0 0 .08-1 7.96 7.96 0 0 0-.08-1l2.11-1.65a.5.5 0 0 0 .12-.64l-2-3.46a.5.5 0 0 0-.6-.22l-2.49 1a8.14 8.14 0 0 0-1.73-1L14.5 2.5a.5.5 0 0 0-.5-.5h-4a.5.5 0 0 0-.5.5L9 5.03c-.6.24-1.17.56-1.7.94l-2.5-1a.5.5 0 0 0-.6.22l-2 3.46a.5.5 0 0 0 .12.64L4.4 12a7.96 7.96 0 0 0-.08 1 7.96 7.96 0 0 0 .08 1l-2.11 1.65a.5.5 0 0 0-.12.64l2 3.46a.5.5 0 0 0 .6.22l2.49-1c.53.38 1.1.7 1.7.94L9.5 23.5a.5.5 0 0 0 .5.5h4a.5.5 0 0 0 .5-.5L15 20.97c.6-.24 1.17-.56 1.7-.94l2.49 1a.5.5 0 0 0 .6-.22l2-3.46a.5.5 0 0 0-.12-.64L19.4 15Z" stroke="currentColor" strokeWidth="1.5"/></svg>
          </button>
        </div>
        <div />
      </div>

      {/* Main area */}
  <div style={{ display: 'flex', flex: 1, width: '100%', justifyContent:'center' }}>
        {/* Left pane: page content */}
  <div style={{ flex: 1, padding: '16px 24px', overflowY: 'auto', overflowX: 'auto', borderRight: tab==='viewer' ? '1px solid #eee' : 'none' }}>
          {error && <div style={{ color: 'red' }}>{error}</div>}
          {loading && <div>Caricamento…</div>}

            {tab === 'impostazioni' && (
              <>
                <h3>Impostazioni</h3>
                <h4 style={{ marginBottom: 6 }}>Mappatura: Colore DOCX → Etichetta categoria</h4>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 16 }}>
                  {categories.map(c => (
                    <div key={c} style={{ display: 'flex', alignItems: 'center', gap: 6, background: '#fff', border: '1px solid #eee', borderRadius: 6, padding: 6 }}>
                      <span style={{ width: 12, height: 12, background: c, display: 'inline-block', border: '1px solid #ccc' }} />
                      <span style={{ minWidth: 80 }}>{colorLabel(c)}</span>
                      <input
                        value={colorMap[c] || ''}
                        onChange={e => setColorMap({ ...colorMap, [c]: e.target.value })}
                        placeholder={colorLabel(c)}
                      />
                    </div>
                  ))}
                </div>
                <h4 style={{ marginBottom: 6 }}>Mappatura: XX_X → Etichetta</h4>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 16 }}>
                  {(() => {
                    const tokens = Array.from(new Set((data?.comments || []).flatMap(c => (c.code || '')
                      .split(/[,;\s]+/)
                      .map(t => t.trim())
                      .filter(Boolean))))
                    return tokens.map(t => (
                      <div key={t} style={{ display: 'flex', alignItems: 'center', gap: 6, background: '#fff', border: '1px solid #eee', borderRadius: 6, padding: 6 }}>
                        <span style={{ minWidth: 80 }}>{t}</span>
                        <input
                          value={codeMap[t] || ''}
                          onChange={e => setCodeMap({ ...codeMap, [t]: e.target.value })}
                          placeholder={t}
                        />
                      </div>
                    ))
                  })()}
                </div>
                <h4 style={{ marginBottom: 6 }}>Colori delle etichette categoria</h4>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
                  {Array.from(new Set([
                    ...Object.values(colorMap).filter(Boolean),
                    ...stats.map(s => s.category),
                    'non-categorizzato',
                  ])).map(cat => (
                    <div key={cat} style={{ display: 'flex', alignItems: 'center', gap: 6, background: '#fff', border: '1px solid #eee', borderRadius: 6, padding: 6 }}>
                      <span style={{ width: 12, height: 12, background: categoryColors[cat] || '#ddd', display: 'inline-block', border: '1px solid #ccc' }} />
                      <span style={{ minWidth: 120 }}>{cat}</span>
                      <input type="color" value={categoryColors[cat] || '#dddddd'} onChange={e => setCategoryColors({ ...categoryColors, [cat]: e.target.value })} />
                      <button onClick={() => setCategoryColors(prev => { const n = { ...prev }; delete n[cat]; return n })}>Reset</button>
                    </div>
                  ))}
                </div>
              </>
            )}

            {tab === 'viewer' && (
              <>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <h3 style={{ margin: 0 }}>Documento</h3>
                  {data && (
                    <select value={viewerDoc} onChange={e => setViewerDoc(e.target.value)}>
                      {Array.from(new Set(data.highlights.map(h => h.filename))).map(fn => (
                        <option key={fn} value={fn}>{fn}</option>
                      ))}
                    </select>
                  )}
                </div>
                <div ref={paraRef}>
                  {paragraphs.filter(p => !viewerDoc || p.filename === viewerDoc).map((p, i) => {
                // render paragraph with inline highlights
                const chunks: React.ReactNode[] = []
                let cursor = 0
                for (const h of p.highlights) {
                  const start = Math.max(0, h.offset_start)
                  const end = Math.min(p.text.length, h.offset_end)
                  if (cursor < start) {
                    chunks.push(<span key={`t-${cursor}`}>{p.text.slice(cursor, start)}</span>)
                  }
          const color = (h.highlight_color || '').toLowerCase()
          const rawCat = colorCategoryFromMap(color, colorMap)
          const cat = color ? (rawCat || colorLabel(color)) : 'non-categorizzato'
          const labelColor = categoryColors[cat] || color || 'yellow'
                  chunks.push(
                    <span
                      key={`h-${start}`}
                      id={`hl-${sig(h)}`}
            style={{ background: labelColor }}
            title={`${cat}`}
                    >
                      {p.text.slice(start, end)}
                    </span>
                  )
                  cursor = end
                }
                if (cursor < p.text.length) {
                  chunks.push(<span key={`t-end-${i}`}>{p.text.slice(cursor)}</span>)
                }
                return <p key={i} style={{ lineHeight: 1.7 }}>{chunks}</p>
              })}
                </div>
              </>
            )}

            {tab === 'comments' && (
              <>
                <h3>Tutti i commenti</h3>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 12, alignItems: 'flex-start' }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                  <select multiple value={fDoc} onChange={e => setFDoc(readMulti(e))} size={Math.min(6, Math.max(3, commentDocs.length))}>
                    {commentDocs.map(d => (<option key={d} value={d}>{d}</option>))}
                  </select>
                  <div style={{ display: 'flex', gap: 4 }}>
                    <button onClick={() => setFDoc(commentDocs)}>Tutti</button>
                    <button onClick={() => setFDoc([])}>Nessuno</button>
                  </div>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                  <select multiple value={fCode} onChange={e => setFCode(readMulti(e))} size={Math.min(6, Math.max(3, commentCodes.length))}>
                    {commentCodes.map(c => (<option key={c} value={c}>{c}</option>))}
                  </select>
                  <div style={{ display: 'flex', gap: 4 }}>
                    <button onClick={() => setFCode(commentCodes)}>Tutti</button>
                    <button onClick={() => setFCode([])}>Nessuno</button>
                  </div>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                  <select multiple value={fColor} onChange={e => setFColor(readMulti(e))} size={Math.min(6, Math.max(3, commentColors.length))}>
                    {commentColors.map(c => (<option key={c} value={c}>{colorLabel(c)}</option>))}
                  </select>
                  <div style={{ display: 'flex', gap: 4 }}>
                    <button onClick={() => setFColor(commentColors)}>Tutti</button>
                    <button onClick={() => setFColor([])}>Nessuno</button>
                  </div>
                  </div>
                  {(() => {
                    const cats = Array.from(new Set(commentsEnriched.map(c => (c.category || 'non-categorizzato').toLowerCase())))
                    return (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                        <select multiple value={fCategory} onChange={e => setFCategory(readMulti(e))} size={Math.min(6, Math.max(3, cats.length))}>
                          {cats.map(cat => (<option key={cat} value={cat}>{cat}</option>))}
                        </select>
                        <div style={{ display: 'flex', gap: 4 }}>
                          <button onClick={() => setFCategory(cats)}>Tutti</button>
                          <button onClick={() => setFCategory([])}>Nessuno</button>
                        </div>
                      </div>
                    )
                  })()}
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                    <select multiple value={fIntervistatore} onChange={e => setFIntervistatore(readMulti(e))} size={Math.min(6, Math.max(3, optionsFromMeta.intervistatori.length))}>
                      {optionsFromMeta.intervistatori.map(v => (<option key={v} value={v}>{v}</option>))}
                    </select>
                    <div style={{ display: 'flex', gap: 4 }}>
                      <button onClick={() => setFIntervistatore(optionsFromMeta.intervistatori)}>Tutti</button>
                      <button onClick={() => setFIntervistatore([])}>Nessuno</button>
                    </div>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                    <select multiple value={fIntervistato} onChange={e => setFIntervistato(readMulti(e))} size={Math.min(6, Math.max(3, optionsFromMeta.intervistati.length))}>
                      {optionsFromMeta.intervistati.map(v => (<option key={v} value={v}>{v}</option>))}
                    </select>
                    <div style={{ display: 'flex', gap: 4 }}>
                      <button onClick={() => setFIntervistato(optionsFromMeta.intervistati)}>Tutti</button>
                      <button onClick={() => setFIntervistato([])}>Nessuno</button>
                    </div>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                    <select multiple value={fRuolo} onChange={e => setFRuolo(readMulti(e))} size={Math.min(6, Math.max(3, optionsFromMeta.ruoli.length))}>
                      {optionsFromMeta.ruoli.map(v => (<option key={v} value={v}>{v}</option>))}
                    </select>
                    <div style={{ display: 'flex', gap: 4 }}>
                      <button onClick={() => setFRuolo(optionsFromMeta.ruoli)}>Tutti</button>
                      <button onClick={() => setFRuolo([])}>Nessuno</button>
                    </div>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                    <select multiple value={fScuola} onChange={e => setFScuola(readMulti(e))} size={Math.min(6, Math.max(3, optionsFromMeta.scuole.length))}>
                      {optionsFromMeta.scuole.map(v => (<option key={v} value={v}>{v}</option>))}
                    </select>
                    <div style={{ display: 'flex', gap: 4 }}>
                      <button onClick={() => setFScuola(optionsFromMeta.scuole)}>Tutti</button>
                      <button onClick={() => setFScuola([])}>Nessuno</button>
                    </div>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                    <select multiple value={fGruppo} onChange={e => setFGruppo(readMulti(e))} size={Math.min(6, Math.max(3, optionsFromMeta.gruppi.length))}>
                      {optionsFromMeta.gruppi.map(v => (<option key={v} value={v}>{v}</option>))}
                    </select>
                    <div style={{ display: 'flex', gap: 4 }}>
                      <button onClick={() => setFGruppo(optionsFromMeta.gruppi)}>Tutti</button>
                      <button onClick={() => setFGruppo([])}>Nessuno</button>
                    </div>
                  </div>
                  <input value={fQuery} onChange={e => setFQuery(e.target.value)} placeholder="Cerca nel testo/commento" style={{ flex: 1, minWidth: 180 }} />
                  <button onClick={() => exportComments(filteredComments)}>Esporta CSV (filtrati)</button>
                  <button onClick={() => exportComments(commentsEnriched)}>Esporta CSV (tutti)</button>
                  <button onClick={() => exportCommentsXLSX(filteredComments)}>Esporta XLSX (filtrati)</button>
                  <button onClick={() => exportCommentsXLSX(commentsEnriched)}>Esporta XLSX (tutti)</button>
                </div>

                {filteredComments.length === 0 && <div>Nessun commento.</div>}
                {filteredComments.map((c, idx) => {
                  const color = (c.color || '').toLowerCase()
                  const cat = c.category || '—'
                  const labelColor = categoryColors[cat] || color || '#bbb'
                  const highlightedFull = (c.highlightText || c.highlight?.text || '').trim()
                  const quotedText = (c.quoted || '').trim()
                  const showFullHighlight = highlightedFull && highlightedFull !== quotedText
                  return (
                    <div key={idx} style={{ border: '1px solid #eee', padding: 10, borderLeft: `4px solid ${labelColor}`, marginBottom: 10, background: '#fff' }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', gap: 8 }}>
                        <div style={{ fontSize: 12, color: '#555' }}>{c.filename}</div>
                        <div style={{ fontSize: 12, color: '#777' }}>{c.author} · {c.date}</div>
                      </div>
                      <div style={{ margin: '6px 0', fontWeight: 600 }}>{c.text || '—'}</div>
                      {c.quoted && <div style={{ fontSize: 13, color: '#333' }}>Testo selezionato: <em>“{c.quoted}”</em></div>}
                      {showFullHighlight && (
                        <div style={{ fontSize: 13, color: '#333' }}>Testo evidenziato: <em>“{highlightedFull}”</em></div>
                      )}
                      <div style={{ display: 'flex', gap: 12, alignItems: 'center', marginTop: 6 }}>
                        <span style={{ fontSize: 12 }}>Codice: <strong>{c.code || '—'}</strong></span>
                        <span style={{ fontSize: 12 }}>Etichetta: <strong>{cat}</strong></span>
                        {c.highlight && (
                          <button onClick={() => jumpTo(c.highlight!)} style={{ marginLeft: 'auto' }}>Vai al testo</button>
                        )}
                      </div>
                    </div>
                  )
                })}
              </>
            )}

            {tab === 'kanban' && (
              <>
                <h3>Kanban categorie</h3>
                <div style={{ display: 'flex', gap: 12, alignItems: 'center', marginBottom: 8, flexWrap: 'wrap' }}>
                  <div style={{ display: 'inline-flex', border: '1px solid #dee2e6', borderRadius: 8, overflow: 'hidden' }}>
                    <button onClick={() => setKView('colore')} style={{ padding: '6px 10px', border: 'none', background: kView==='colore' ? '#f1f3f5' : '#fff' }}>Colore</button>
                    <button onClick={() => setKView('xx')} style={{ padding: '6px 10px', borderLeft: '1px solid #dee2e6', background: kView==='xx' ? '#f1f3f5' : '#fff' }}>XX_X</button>
                  </div>
                  {/* DnD removed; helper text hidden */}
                  {/* Filters: mirror Commenti */}
                  <div style={{ display: 'flex', gap: 8, alignItems: 'flex-start', flexWrap:'wrap' }}>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                      <select multiple value={kDoc} onChange={e => setKDoc(readMulti(e))} size={Math.min(6, Math.max(3, commentDocs.length))}>
                        {commentDocs.map(d => (<option key={d} value={d}>{d}</option>))}
                      </select>
                      <div style={{ display: 'flex', gap: 4 }}>
                        <button onClick={() => setKDoc(commentDocs)}>Tutti</button>
                        <button onClick={() => setKDoc([])}>Nessuno</button>
                      </div>
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                      <select multiple value={kCode} onChange={e => setKCode(readMulti(e))} size={Math.min(6, Math.max(3, commentCodes.length))}>
                        {commentCodes.map(c => (<option key={c} value={c}>{c}</option>))}
                      </select>
                      <div style={{ display: 'flex', gap: 4 }}>
                        <button onClick={() => setKCode(commentCodes)}>Tutti</button>
                        <button onClick={() => setKCode([])}>Nessuno</button>
                      </div>
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                      <select multiple value={kColor} onChange={e => setKColor(readMulti(e))} size={Math.min(6, Math.max(3, commentColors.length))}>
                        {commentColors.map(c => (<option key={c} value={c}>{colorLabel(c)}</option>))}
                      </select>
                      <div style={{ display: 'flex', gap: 4 }}>
                        <button onClick={() => setKColor(commentColors)}>Tutti</button>
                        <button onClick={() => setKColor([])}>Nessuno</button>
                      </div>
                    </div>
                    {(() => {
                      const cats = Array.from(new Set(commentsEnriched.map(c => (c.category || 'non-categorizzato').toLowerCase())))
                      return (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                          <select multiple value={kCategory} onChange={e => setKCategory(readMulti(e))} size={Math.min(6, Math.max(3, cats.length))}>
                            {cats.map(cat => (<option key={cat} value={cat}>{cat}</option>))}
                          </select>
                          <div style={{ display: 'flex', gap: 4 }}>
                            <button onClick={() => setKCategory(cats)}>Tutti</button>
                            <button onClick={() => setKCategory([])}>Nessuno</button>
                          </div>
                        </div>
                      )
                    })()}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                      <select multiple value={kIntervistatore} onChange={e => setKIntervistatore(readMulti(e))} size={Math.min(6, Math.max(3, optionsFromMeta.intervistatori.length))}>
                        {optionsFromMeta.intervistatori.map(v => (<option key={v} value={v}>{v}</option>))}
                      </select>
                      <div style={{ display: 'flex', gap: 4 }}>
                        <button onClick={() => setKIntervistatore(optionsFromMeta.intervistatori)}>Tutti</button>
                        <button onClick={() => setKIntervistatore([])}>Nessuno</button>
                      </div>
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                      <select multiple value={kIntervistato} onChange={e => setKIntervistato(readMulti(e))} size={Math.min(6, Math.max(3, optionsFromMeta.intervistati.length))}>
                        {optionsFromMeta.intervistati.map(v => (<option key={v} value={v}>{v}</option>))}
                      </select>
                      <div style={{ display: 'flex', gap: 4 }}>
                        <button onClick={() => setKIntervistato(optionsFromMeta.intervistati)}>Tutti</button>
                        <button onClick={() => setKIntervistato([])}>Nessuno</button>
                      </div>
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                      <select multiple value={kRuolo} onChange={e => setKRuolo(readMulti(e))} size={Math.min(6, Math.max(3, optionsFromMeta.ruoli.length))}>
                        {optionsFromMeta.ruoli.map(v => (<option key={v} value={v}>{v}</option>))}
                      </select>
                      <div style={{ display: 'flex', gap: 4 }}>
                        <button onClick={() => setKRuolo(optionsFromMeta.ruoli)}>Tutti</button>
                        <button onClick={() => setKRuolo([])}>Nessuno</button>
                      </div>
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                      <select multiple value={kGruppo} onChange={e => setKGruppo(readMulti(e))} size={Math.min(6, Math.max(3, optionsFromMeta.gruppi.length))}>
                        {optionsFromMeta.gruppi.map(v => (<option key={v} value={v}>{v}</option>))}
                      </select>
                      <div style={{ display: 'flex', gap: 4 }}>
                        <button onClick={() => setKGruppo(optionsFromMeta.gruppi)}>Tutti</button>
                        <button onClick={() => setKGruppo([])}>Nessuno</button>
                      </div>
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                      <select multiple value={kScuola} onChange={e => setKScuola(readMulti(e))} size={Math.min(6, Math.max(3, optionsFromMeta.scuole.length))}>
                        {optionsFromMeta.scuole.map(v => (<option key={v} value={v}>{v}</option>))}
                      </select>
                      <div style={{ display: 'flex', gap: 4 }}>
                        <button onClick={() => setKScuola(optionsFromMeta.scuole)}>Tutti</button>
                        <button onClick={() => setKScuola([])}>Nessuno</button>
                      </div>
                    </div>
                    <input value={kQuery} onChange={e => setKQuery(e.target.value)} placeholder="Cerca nel testo/commento" style={{ flex: 1, minWidth: 180 }} />
                  </div>
                </div>
                {(() => {
                  // Filter items (mirror Commenti)
                  let items = commentsEnriched
                  if (kDoc.length) items = items.filter(c => kDoc.includes(c.filename || ''))
                  if (kCode.length) items = items.filter(c => kCode.includes((c.code || '').toLowerCase()))
                  if (kColor.length) items = items.filter(c => kColor.includes((c.color || '').toLowerCase()))
                  if (kCategory.length) items = items.filter(c => kCategory.includes((c.category || '').toLowerCase()))
                  if (kIntervistatore.length) items = items.filter(c => {
                    const fm = c.filename ? meta[c.filename] : undefined
                    return fm ? kIntervistatore.map(s=>s.toLowerCase()).includes((fm.intervistatore||'').toLowerCase()) : false
                  })
                  if (kIntervistato.length) items = items.filter(c => {
                    const fm = c.filename ? meta[c.filename] : undefined
                    return fm ? kIntervistato.map(s=>s.toLowerCase()).includes((fm.intervistato||'').toLowerCase()) : false
                  })
                  if (kRuolo.length) items = items.filter(c => {
                    const fm = c.filename ? meta[c.filename] : undefined
                    const r = (fm?.ruolo || '').toLowerCase()
                    return kRuolo.map(s=>s.toLowerCase()).includes(r)
                  })
                  if (kGruppo.length) items = items.filter(c => {
                    const fm = c.filename ? meta[c.filename] : undefined
                    const g = (fm?.gruppo || '').toLowerCase()
                    return kGruppo.map(s=>s.toLowerCase()).includes(g)
                  })
                  if (kScuola.length) items = items.filter(c => {
                    const fm = c.filename ? meta[c.filename] : undefined
                    const s = (fm?.scuola || '').toLowerCase()
                    return kScuola.map(x=>x.toLowerCase()).includes(s)
                  })
                  if (kQuery.trim()) {
                    const q = kQuery.trim().toLowerCase()
                    items = items.filter(c =>
                      (c.text || '').toLowerCase().includes(q) ||
                      (c.quoted || '').toLowerCase().includes(q) ||
                      (c.author || '').toLowerCase().includes(q) ||
                      ((c.highlightText || c.highlight?.text || '').toLowerCase().includes(q))
                    )
                  }

                  // Build columns based on view
                  let columns: string[] = []
                  let groups: Record<string, typeof items> = {}
                  if (kView === 'colore') {
                    const cats = Array.from(new Set([
                      ...items.map(c => c.category || 'non-categorizzato'),
                      'non-categorizzato',
                    ]))
                    columns = cats
                    const byCat: Record<string, typeof items> = {}
                    for (const c of items) {
                      const k = c.category || 'non-categorizzato'
                      ;(byCat[k] ||= []).push(c)
                    }
                    groups = byCat
                  } else {
                    // XX_X view: group by code tokens
                    const tokenSet = new Set<string>()
                    const tokenMap: Record<string, typeof items> = {}
                    const noToken: typeof items = []
                    const split = (s: string) => s.split(/[,;\s]+/).map(t => t.trim()).filter(Boolean)
                    for (const c of items) {
                      const raw = c.code ? split(c.code) : []
                      const tokens = raw.map(t => codeMap[t] || t)
                      if (tokens.length === 0) {
                        noToken.push(c)
                      } else {
                        for (const t of tokens) {
                          tokenSet.add(t)
                          ;(tokenMap[t] ||= []).push(c)
                        }
                      }
                    }
                    columns = [...Array.from(tokenSet).sort(), '—']
                    groups = { ...tokenMap, '—': noToken }
                  }

                  // Drag-and-drop disabled by request

                  return (
                    <div style={{ display: 'flex', gap: 12, alignItems: 'flex-start', overflowX: 'auto' }}>
                      {columns.map(col => (
                        <div key={col}
                          style={{ minWidth: 260, background: '#fff', border: '1px solid #e9ecef', borderRadius: 8, padding: 10 }}>
                          <div style={{ fontWeight: 700, marginBottom: 8 }}>{col} <span style={{ color: '#868e96' }}>({(groups[col]||[]).length})</span></div>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                            {(groups[col] || []).map(c => (
                              <div key={c.id}
                                style={{ border: '1px solid #eee', borderLeft: `4px solid ${categoryColors[c.category || 'non-categorizzato'] || '#bbb'}`, padding: 8, borderRadius: 6, background: '#f8f9fa' }}>
                                <div style={{ fontSize: 12, color: '#555' }}>{c.filename}</div>
                                <div style={{ fontWeight: 600, margin: '4px 0' }}>{c.text || '—'}</div>
                                {c.quoted && <div style={{ fontSize: 12, color: '#495057' }}>“{c.quoted}”</div>}
                                {(() => {
                                  const highlightedFull = (c.highlightText || c.highlight?.text || '').trim()
                                  const quotedText = (c.quoted || '').trim()
                                  if (!highlightedFull || highlightedFull === quotedText) return null
                                  return <div style={{ fontSize: 12, color: '#495057' }}>Testo evidenziato: “{highlightedFull}”</div>
                                })()}
                                <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginTop: 6 }}>
                                  {/* Category badge with category label color */}
                                  <span style={{ fontSize: 11, padding: '2px 6px', borderRadius: 999, background: (categoryColors[c.category || 'non-categorizzato'] || '#e9ecef'), color: '#212529' }}>{c.category || 'non-categorizzato'}</span>
                                  {/* Additional labels from codes like XX_X */}
                                  {(() => {
                                    const tokens = (c.code || '')
                                      .split(/[,;\s]+/)
                                      .map(t => t.trim())
                                      .filter(Boolean)
                                    return tokens.map(t => (
                                      <span key={t} style={{ fontSize: 11, padding: '2px 6px', borderRadius: 999, background: '#fff', border: '1px solid #dee2e6', color: '#343a40' }}>{t}</span>
                                    ))
                                  })()}
                                </div>
                                {c.highlight && <button onClick={() => jumpTo(c.highlight!)} style={{ marginTop: 6 }}>Vai al testo</button>}
                              </div>
                            ))}
                          </div>
                        </div>
                      ))}
                    </div>
                  )
                })()}
              </>
            )}

            {tab === 'documenti' && (
              <>
                <h3>Documenti</h3>
                {(() => {
                  const fileSet = new Set<string>([
                    ...((data ? data.highlights : []) as Highlight[]).map(h => h.filename),
                    ...Object.keys(meta)
                  ])
                  let files = Array.from(fileSet)
                  // Build options from meta
                  const opts = {
                    tipi: Array.from(new Set(Object.values(meta).map(m => m?.tipo).filter(Boolean))) as string[],
                    intervistatori: Array.from(new Set(Object.values(meta).map(m => m?.intervistatore).filter(Boolean))) as string[],
                    intervistati: Array.from(new Set(Object.values(meta).map(m => m?.intervistato).filter(Boolean))) as string[],
                    ruoli: Array.from(new Set(Object.values(meta).map(m => m?.ruolo).filter(Boolean))) as string[],
                    gruppi: Array.from(new Set(Object.values(meta).map(m => m?.gruppo).filter(Boolean))) as string[],
                    scuole: Array.from(new Set(Object.values(meta).map(m => m?.scuola).filter(Boolean))) as string[],
                  }
                  // Apply filters
                  files = files.filter(fn => {
                    const m = meta[fn]
                    if (dTipo.length && !dTipo.includes((m?.tipo || '') as string)) return false
                    if (dIntervistatore.length && !dIntervistatore.map(s=>s.toLowerCase()).includes((m?.intervistatore || '').toLowerCase())) return false
                    if (dIntervistato.length && !dIntervistato.map(s=>s.toLowerCase()).includes((m?.intervistato || '').toLowerCase())) return false
                    if (dRuolo.length && !dRuolo.map(s=>s.toLowerCase()).includes((m?.ruolo || '').toLowerCase())) return false
                    if (dGruppo.length && !dGruppo.map(s=>s.toLowerCase()).includes((m?.gruppo || '').toLowerCase())) return false
                    if (dScuola.length && !dScuola.map(s=>s.toLowerCase()).includes((m?.scuola || '').toLowerCase())) return false
                    if (dQuery && !fn.toLowerCase().includes(dQuery.toLowerCase())) return false
                    return true
                  })
                  
                  const resetFilters = () => {
                    setDTipo([]); setDIntervistatore([]); setDIntervistato([]); setDRuolo([]); setDGruppo([]); setDScuola([]); setDQuery('')
                  }
                  return (
                    <>
                      <div style={{ marginBottom: 12 }}>
                        <input type="file" accept=".docx" multiple onChange={handleUpload} />
                      </div>
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 12, alignItems: 'flex-start' }}>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                          <select multiple value={dTipo} onChange={e => setDTipo(readMulti(e))} size={Math.min(6, Math.max(3, opts.tipi.length || 3))}>
                            {opts.tipi.map(v => (<option key={v} value={v as string}>{v}</option>))}
                          </select>
                          <div style={{ display: 'flex', gap: 4 }}>
                            <button onClick={() => setDTipo(opts.tipi as string[])}>Tutti</button>
                            <button onClick={() => setDTipo([])}>Nessuno</button>
                          </div>
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                          <select multiple value={dIntervistatore} onChange={e => setDIntervistatore(readMulti(e))} size={Math.min(6, Math.max(3, opts.intervistatori.length || 3))}>
                            {opts.intervistatori.map(v => (<option key={v} value={v as string}>{v}</option>))}
                          </select>
                          <div style={{ display: 'flex', gap: 4 }}>
                            <button onClick={() => setDIntervistatore(opts.intervistatori as string[])}>Tutti</button>
                            <button onClick={() => setDIntervistatore([])}>Nessuno</button>
                          </div>
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                          <select multiple value={dIntervistato} onChange={e => setDIntervistato(readMulti(e))} size={Math.min(6, Math.max(3, opts.intervistati.length || 3))}>
                            {opts.intervistati.map(v => (<option key={v} value={v as string}>{v}</option>))}
                          </select>
                          <div style={{ display: 'flex', gap: 4 }}>
                            <button onClick={() => setDIntervistato(opts.intervistati as string[])}>Tutti</button>
                            <button onClick={() => setDIntervistato([])}>Nessuno</button>
                          </div>
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                          <select multiple value={dRuolo} onChange={e => setDRuolo(readMulti(e))} size={Math.min(6, Math.max(3, opts.ruoli.length || 3))}>
                            {opts.ruoli.map(v => (<option key={v} value={v as string}>{v}</option>))}
                          </select>
                          <div style={{ display: 'flex', gap: 4 }}>
                            <button onClick={() => setDRuolo(opts.ruoli as string[])}>Tutti</button>
                            <button onClick={() => setDRuolo([])}>Nessuno</button>
                          </div>
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                          <select multiple value={dGruppo} onChange={e => setDGruppo(readMulti(e))} size={Math.min(6, Math.max(3, opts.gruppi.length || 3))}>
                            {opts.gruppi.map(v => (<option key={v} value={v as string}>{v}</option>))}
                          </select>
                          <div style={{ display: 'flex', gap: 4 }}>
                            <button onClick={() => setDGruppo(opts.gruppi as string[])}>Tutti</button>
                            <button onClick={() => setDGruppo([])}>Nessuno</button>
                          </div>
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                          <select multiple value={dScuola} onChange={e => setDScuola(readMulti(e))} size={Math.min(6, Math.max(3, opts.scuole.length || 3))}>
                            {opts.scuole.map(v => (<option key={v} value={v as string}>{v}</option>))}
                          </select>
                          <div style={{ display: 'flex', gap: 4 }}>
                            <button onClick={() => setDScuola(opts.scuole as string[])}>Tutti</button>
                            <button onClick={() => setDScuola([])}>Nessuno</button>
                          </div>
                        </div>
                        <input value={dQuery} onChange={e => setDQuery(e.target.value)} placeholder="Cerca per filename" style={{ flex: 1, minWidth: 180 }} />
                        <button onClick={resetFilters}>Reset</button>
                      </div>
          <table style={{ width: '100%', borderCollapse: 'collapse', background: '#fff', border: '1px solid #eee', borderRadius: 8 }}>
                        <thead>
                          <tr>
                            <th style={{ textAlign: 'left', borderBottom: '1px solid #ddd', padding: 6 }}>File</th>
                            <th style={{ textAlign: 'left', borderBottom: '1px solid #ddd', padding: 6 }}>Tipo</th>
                            <th style={{ textAlign: 'left', borderBottom: '1px solid #ddd', padding: 6 }}>Intervistatore</th>
                            <th style={{ textAlign: 'left', borderBottom: '1px solid #ddd', padding: 6 }}>Intervistato</th>
                            <th style={{ textAlign: 'left', borderBottom: '1px solid #ddd', padding: 6 }}>Ruolo</th>
                            <th style={{ textAlign: 'left', borderBottom: '1px solid #ddd', padding: 6 }}>Gruppo</th>
                            <th style={{ textAlign: 'left', borderBottom: '1px solid #ddd', padding: 6 }}>Note gruppo</th>
                            <th style={{ textAlign: 'left', borderBottom: '1px solid #ddd', padding: 6 }}>Scuola</th>
                            <th style={{ textAlign: 'right', borderBottom: '1px solid #ddd', padding: 6 }}># Evidenziati</th>
                            <th style={{ textAlign: 'right', borderBottom: '1px solid #ddd', padding: 6 }}># Commenti</th>
            <th style={{ textAlign: 'right', borderBottom: '1px solid #ddd', padding: 6 }}>Azioni</th>
                          </tr>
                        </thead>
                        <tbody>
                          {files.map(fn => {
                          const m = meta[fn] || { tipo: 'Intervista', intervistatore: '', intervistato: '', ruolo: '', scuola: '', gruppo: '', noteGruppo: '' }
                          const hl = data ? data.highlights.filter(h => h.filename === fn).length : 0
                          const cm = data ? data.comments.filter(c => (c.filename || '') === fn).length : 0
                          return (
                            <tr key={fn}>
                              <td style={{ padding: 6, borderBottom: '1px solid #f1f3f5' }}>{fn}</td>
                              <td style={{ padding: 6, borderBottom: '1px solid #f1f3f5' }}>
                                <select value={m.tipo} onChange={e => setMeta({ ...meta, [fn]: { ...m, tipo: e.target.value as FileMeta['tipo'] } })}>
                                  <option>Intervista</option>
                                  <option>Focus group</option>
                                </select>
                              </td>
                              <td style={{ padding: 6, borderBottom: '1px solid #f1f3f5' }}>
                                <input value={m.intervistatore} onChange={e => setMeta({ ...meta, [fn]: { ...m, intervistatore: e.target.value } })} />
                              </td>
                              <td style={{ padding: 6, borderBottom: '1px solid #f1f3f5' }}>
                                <input value={m.intervistato || ''} onChange={e => setMeta({ ...meta, [fn]: { ...m, intervistato: e.target.value } })} />
                              </td>
                              <td style={{ padding: 6, borderBottom: '1px solid #f1f3f5' }}>
                                <input value={m.ruolo || ''} onChange={e => setMeta({ ...meta, [fn]: { ...m, ruolo: e.target.value } })} />
                              </td>
                              <td style={{ padding: 6, borderBottom: '1px solid #f1f3f5' }}>
                                <input value={m.gruppo || ''} onChange={e => setMeta({ ...meta, [fn]: { ...m, gruppo: e.target.value } })} />
                              </td>
                              <td style={{ padding: 6, borderBottom: '1px solid #f1f3f5' }}>
                                <input value={m.noteGruppo || ''} onChange={e => setMeta({ ...meta, [fn]: { ...m, noteGruppo: e.target.value } })} />
                              </td>
                              <td style={{ padding: 6, borderBottom: '1px solid #f1f3f5' }}>
                                <input value={m.scuola} onChange={e => setMeta({ ...meta, [fn]: { ...m, scuola: e.target.value } })} />
                              </td>
                              <td style={{ padding: 6, borderBottom: '1px solid #f1f3f5', textAlign: 'right' }}>{hl}</td>
                              <td style={{ padding: 6, borderBottom: '1px solid #f1f3f5', textAlign: 'right' }}>{cm}</td>
                              <td style={{ padding: 6, borderBottom: '1px solid #f1f3f5', textAlign: 'right' }}>
                                <button onClick={() => handleDeleteFile(fn)} title="Rimuovi">Rimuovi</button>
                              </td>
                            </tr>
                          )
                        })}
                        </tbody>
                      </table>
                    </>
                  )
                })()}
              </>
            )}

            {tab === 'dashboard' && (
              <>
                <h3>Dashboard</h3>
                <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap' }}>
                  <div style={{ flex: '1 1 260px', background: '#fff', border: '1px solid #eee', borderRadius: 8, padding: 12 }}>
                    <h4 style={{ marginTop: 0 }}>Categorie</h4>
                    {stats.length === 0 && <div style={{ color: '#666' }}>Nessun dato</div>}
                    {stats.map(s => (
                      <div key={s.category} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
                        <div style={{ width: 120 }}>{s.category}</div>
                        <div style={{ background: '#91A7FF', height: 10, width: Math.min(220, s.count * 12) }} />
                        <div style={{ marginLeft: 6 }}>{s.count}</div>
                      </div>
                    ))}
                  </div>
                  <div style={{ flex: '1 1 260px', background: '#fff', border: '1px solid #eee', borderRadius: 8, padding: 12 }}>
                    <h4 style={{ marginTop: 0 }}>Per scuola</h4>
                    {(() => {
                      if (!data) return <div style={{ color: '#666' }}>Nessun dato</div>
                      const counts: Record<string, number> = {}
                      for (const h of data.highlights) {
                        const fm = meta[h.filename]
                        const scol = fm?.scuola || '—'
                        counts[scol] = (counts[scol] || 0) + 1
                      }
                      const arr = Object.entries(counts)
                      if (arr.length === 0) return <div style={{ color: '#666' }}>Nessun dato</div>
                      return arr.map(([k,v]) => (
                        <div key={k} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
                          <div style={{ width: 120 }}>{k}</div>
                          <div style={{ background: '#63E6BE', height: 10, width: Math.min(220, v * 12) }} />
                          <div style={{ marginLeft: 6 }}>{v}</div>
                        </div>
                      ))
                    })()}
                  </div>
                </div>
              </>
            )}
      </div>

      {/* Right pane: only in Viewer */}
      {tab === 'viewer' && (
      <div style={{ width: 420, padding: 16, overflow: 'auto' }}>
        <h3>Evidenziati</h3>
        {!data && <div>Nessun dato.</div>}
        {data && data.highlights.filter(h => !viewerDoc || h.filename === viewerDoc).map((h, idx) => {
          const color = (h.highlight_color||'').toLowerCase()
          const rawCat = colorCategoryFromMap(color, colorMap)
          const cat = color ? (rawCat || colorLabel(color)) : 'non-categorizzato'
          return (
            <div
              key={idx}
              onClick={() => jumpTo(h)}
              style={{ cursor: 'pointer', border: '1px solid #eee', padding: 8, borderLeft: `4px solid ${categoryColors[cat] || color || 'yellow'}`, marginBottom: 8 }}
            >
              <div style={{ fontSize: 12, color: '#555' }}>{h.filename}</div>
              <div style={{ fontWeight: 600 }}>{h.text}</div>
              <div style={{ fontSize: 12 }}>Categoria: {cat}</div>
              <div style={{ fontSize: 12, color: '#666' }}>{h.context}</div>
            </div>
          )
        })}
        {(() => {
          const filtered = !data ? [] : data.highlights.filter(h => !viewerDoc || h.filename === viewerDoc)
          const counts: Record<string, number> = {}
          for (const h of filtered) {
            const c = (h.highlight_color||'').toLowerCase()
            const rawCat = colorCategoryFromMap(c, colorMap)
            const cat = c ? (rawCat || colorLabel(c)) : 'non-categorizzato'
            counts[cat] = (counts[cat]||0)+1
          }
          const localStats = Object.entries(counts).map(([category, count]) => ({ category, count })).sort((a,b)=>b.count-a.count)
          return (
            <>
              <h3 style={{ marginTop: 16 }}>Statistiche categorie</h3>
              {localStats.length === 0 && <div>Nessun dato.</div>}
              {localStats.length > 0 && (
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead>
              <tr>
                <th style={{ textAlign: 'left', borderBottom: '1px solid #ddd' }}>Categoria</th>
                <th style={{ textAlign: 'right', borderBottom: '1px solid #ddd' }}>Conteggio</th>
              </tr>
            </thead>
            <tbody>
              {localStats.map(s => (
                <tr key={s.category}>
                  <td style={{ padding: '4px 0' }}>{s.category}</td>
                  <td style={{ padding: '4px 0', textAlign: 'right' }}>{s.count}</td>
                </tr>
              ))}
            </tbody>
          </table>
              )}
            </>
          )
        })()}
      </div>
      )}
  </div>
  </div>
  )
}

export default App
