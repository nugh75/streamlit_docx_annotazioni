import React, { useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'

type MacroKey = 'CE' | 'CS' | 'BE' | 'IN' | 'CC' | 'A'

const MACRO_INFO: Record<MacroKey, { label: string; color: string }> = {
  CE: { label: 'Comunità educante', color: '#d0ebff' },
  CS: { label: 'Cambiamenti scuola', color: '#d3f9d8' },
  BE: { label: 'Benessere', color: '#ffe3e3' },
  IN: { label: 'Inclusione', color: '#ffe8cc' },
  CC: { label: 'Cultura e cittadinanza', color: '#e5dbff' },
  A: { label: 'Aspetti positivi/negativi', color: '#e9ecef' },
}

const CODE_INFO: Record<string, { label: string; macro: MacroKey }> = {
  CE: { label: MACRO_INFO.CE.label, macro: 'CE' },
  CE_T: { label: 'Rapporto con il territorio', macro: 'CE' },
  CE_F: { label: 'Rapporto con le famiglie', macro: 'CE' },
  CE_P: { label: 'Coinvolgimento progettuale (progettazione delle attività e conduzione)', macro: 'CE' },
  CE_Q: { label: 'Qualità dell’esperienza', macro: 'CE' },
  CS: { label: MACRO_INFO.CS.label, macro: 'CS' },
  CS_V: { label: 'Visione di scuola (immagine e ruolo alternativo della scuola)', macro: 'CS' },
  CS_D: { label: 'Impatto sulla didattica per i docenti', macro: 'CS' },
  CS_S: { label: 'Impatto sulla didattica per gli studenti', macro: 'CS' },
  BE: { label: MACRO_INFO.BE.label, macro: 'BE' },
  BE_D: { label: 'Impatto sulle relazioni per i docenti', macro: 'BE' },
  BE_S: { label: 'Impatto sulle relazioni per gli studenti', macro: 'BE' },
  BE_E: { label: 'Impatto emotivo', macro: 'BE' },
  IN: { label: MACRO_INFO.IN.label, macro: 'IN' },
  IN_G: { label: 'Impatto sull’inclusione', macro: 'IN' },
  IN_DS: { label: 'Impatto sulla dispersione scolastica', macro: 'IN' },
  IN_DI: { label: 'Impatto sulla prevenzione delle discriminazioni', macro: 'IN' },
  CC: { label: MACRO_INFO.CC.label, macro: 'CC' },
  CC_CU: { label: 'Impatto sulla cultura', macro: 'CC' },
  CC_CA: { label: 'Impatto sulla cittadinanza attiva', macro: 'CC' },
  A: { label: MACRO_INFO.A.label, macro: 'A' },
  A_P: { label: 'Aspetti positivi', macro: 'A' },
  A_N: { label: 'Aspetti negativi', macro: 'A' },
  A_PS: { label: 'Proposte e suggerimenti', macro: 'A' },
}

const CODE_LIST_BY_TYPE: Record<'Intervista' | 'Focus group', string[]> = {
  Intervista: ['CE_T', 'CE_F', 'CE_P', 'CS_V', 'CS_D', 'CS_S', 'BE_D', 'BE_S', 'BE_E', 'IN_G', 'IN_DS', 'IN_DI', 'CC_CU', 'CC_CA', 'A_P', 'A_N', 'A_PS'],
  'Focus group': ['CE_Q', 'CE_P', 'CE_T', 'CS_V', 'CS_S', 'BE_S', 'BE_E', 'IN_G', 'IN_DS', 'IN_DI', 'CC_CU', 'CC_CA', 'A_P', 'A_N', 'A_PS'],
}

const DEFAULT_CODE_MAP: Record<string, string> = Object.keys(CODE_INFO).reduce((acc, code) => {
  acc[code] = CODE_INFO[code].label
  return acc
}, {} as Record<string, string>)

const LABEL_TO_MACRO: Record<string, MacroKey> = Object.values(CODE_INFO).reduce((acc, info) => {
  acc[info.label] = info.macro
  return acc
}, {} as Record<string, MacroKey>)

const DEFAULT_LABEL_COLORS: Record<string, string> = Object.values(CODE_INFO).reduce((acc, info) => {
  acc[info.label] = MACRO_INFO[info.macro].color
  return acc
}, { 'non-categorizzato': '#f8f9fa' } as Record<string, string>)

const CODE_TOKEN_REGEX = /\b[A-Z]{1,4}(?:_[A-Z]{1,4})?\b/g

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

function normalizeCodeToken(raw: string | null | undefined): string {
  return (raw || '').trim().toUpperCase()
}

function macroFromCode(code: string): MacroKey | null {
  const normalized = normalizeCodeToken(code)
  if (!normalized) return null
  if (CODE_INFO[normalized]) return CODE_INFO[normalized].macro
  const prefix = normalized.split('_')[0]
  return prefix in MACRO_INFO ? (prefix as MacroKey) : null
}

function defaultLabelForCode(code: string): string {
  const normalized = normalizeCodeToken(code)
  if (!normalized) return ''
  return CODE_INFO[normalized]?.label || normalized
}

function isKnownCodeToken(token: string): boolean {
  if (!token) return false
  if (CODE_INFO[token]) return true
  if (token.includes('_')) {
    const prefix = token.split('_')[0]
    return prefix in MACRO_INFO
  }
  return token in MACRO_INFO
}

function defaultColorForLabel(label: string): string {
  return DEFAULT_LABEL_COLORS[label] || '#f8f9fa'
}

function macroFromLabel(label: string): MacroKey | null {
  return LABEL_TO_MACRO[label] || null
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
  codes?: string[]
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
  const [codeMap, setCodeMap] = useState<Record<string, string>>(() => ({ ...DEFAULT_CODE_MAP }))
  const [tab, setTab] = useState<'viewer' | 'comments' | 'dashboard' | 'kanban' | 'documenti' | 'impostazioni' | 'legenda'>('viewer')
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
  const [dbTipo, setDbTipo] = useState<string[]>([])
  const [dbScuola, setDbScuola] = useState<string[]>([])
  const [dbRuolo, setDbRuolo] = useState<string[]>([])
  const [dbGruppo, setDbGruppo] = useState<string[]>([])
  const [dbIntervistatore, setDbIntervistatore] = useState<string[]>([])
  const [dbIntervistato, setDbIntervistato] = useState<string[]>([])
  const [dbQuery, setDbQuery] = useState<string>('')
  const [dbMacroChart, setDbMacroChart] = useState<'istogramma' | 'torta'>('istogramma')
  const [dbCatChart, setDbCatChart] = useState<'istogramma' | 'torta'>('istogramma')
  const [dbCategoryFilter, setDbCategoryFilter] = useState<string[]>([])
  const [dbCompareLayout, setDbCompareLayout] = useState<'stack' | 'columns'>('columns')
  const [compareMode, setCompareMode] = useState<'scuola' | 'ruolo'>('scuola')
  const [compareSelectionA, setCompareSelectionA] = useState<string>('')
  const [compareSelectionB, setCompareSelectionB] = useState<string>('')

  const paraRef = useRef<HTMLDivElement>(null)
  const pendingJump = useRef<Highlight | null>(null)
  const [jumpSeq, setJumpSeq] = useState(0)

  const collectCodes = React.useCallback((comment: CommentItem): string[] => {
    const tokens: string[] = []
    const seen = new Set<string>()
    const pushToken = (token: string | null | undefined) => {
      const normalized = normalizeCodeToken(token)
      if (!normalized || seen.has(normalized)) return
      if (!isKnownCodeToken(normalized)) return
      seen.add(normalized)
      tokens.push(normalized)
    }
    if (Array.isArray(comment.codes)) {
      comment.codes.forEach(pushToken)
    }
    if (comment.code) {
      comment.code
        .split(/[,;\s]+/)
        .map(part => part.trim())
        .forEach(pushToken)
    }
    if (comment.text) {
      for (const match of comment.text.matchAll(CODE_TOKEN_REGEX)) {
        pushToken(match[0])
      }
    }
    return tokens
  }, [])

  const getCategoryColor = React.useCallback((label: string, macroKey?: MacroKey | null) => {
    const override = categoryColors[label]
    if (override) return override
    if (macroKey && MACRO_INFO[macroKey]) return MACRO_INFO[macroKey].color
    return defaultColorForLabel(label)
  }, [categoryColors])

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
  if (savedCodeMap) setCodeMap(prev => ({ ...prev, ...JSON.parse(savedCodeMap) }))
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
    if (!data) return [] as (CommentItem & {
      highlight?: Highlight
      color?: string | null
      category?: string
      highlightText?: string
      macro?: MacroKey | null
      macroLabel?: string | null
      codeTokens: string[]
    })[]
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
      const codeTokens = collectCodes(c)
      const resolved = codeTokens.map(token => ({
        token,
        label: (codeMap[token] || '').trim() || defaultLabelForCode(token),
        macro: macroFromCode(token),
      }))
      const primaryResolved = resolved.find(r => r.label) || resolved[0]
      let defaultCat = primaryResolved ? primaryResolved.label : 'non-categorizzato'
      if (!defaultCat) defaultCat = 'non-categorizzato'
      const keyOverride = `${c.filename || ''}::${c.id}`
      const overrideLabel = (catOverride[keyOverride] || '').trim()
      const effectiveCat = overrideLabel || defaultCat
      const macroKey = macroFromLabel(effectiveCat) || primaryResolved?.macro || null
      const color = getCategoryColor(effectiveCat, macroKey)
      const macroLabel = macroKey ? MACRO_INFO[macroKey].label : null
      return {
        ...c,
        highlight: linked,
        color,
        category: effectiveCat,
        highlightText: linked?.text || '',
        macro: macroKey,
        macroLabel,
        codeTokens,
      }
    })
  }, [data, catOverride, collectCodes, codeMap, getCategoryColor, highlightsByFile, paragraphsByFile])

  const commentDocs = useMemo(() => {
    const set = new Set<string>()
    commentsEnriched.forEach(c => c.filename && set.add(c.filename))
    return Array.from(set)
  }, [commentsEnriched])

  const highlightCategoryMap = useMemo(() => {
    const map = new Map<string, { category: string; color: string; macro: MacroKey | null }>()
    commentsEnriched.forEach(c => {
      if (c.highlight) {
        const key = sig(c.highlight)
        const category = c.category || 'non-categorizzato'
        const macroKey = c.macro || macroFromLabel(category)
        const color = getCategoryColor(category, macroKey)
        map.set(key, { category, color, macro: macroKey })
      }
    })
    return map
  }, [commentsEnriched, getCategoryColor])

  const resolveHighlightCategory = React.useCallback((h: Highlight) => {
    const key = sig(h)
    const mapped = highlightCategoryMap.get(key)
    if (mapped) return mapped
    const colorKey = (h.highlight_color || '').toLowerCase()
    let category = colorKey ? colorCategoryFromMap(colorKey, colorMap) : ''
    if (!category) {
      category = colorKey ? colorLabel(colorKey) : 'non-categorizzato'
    }
    const macroKey = macroFromLabel(category)
    const color = getCategoryColor(category, macroKey)
    return { category, color, macro: macroKey }
  }, [colorMap, getCategoryColor, highlightCategoryMap])

  const stats = useMemo(() => {
    if (!commentsEnriched.length) return [] as { category: string; count: number; color: string }[]
    const counts = new Map<string, { count: number; color: string }>()
    commentsEnriched.forEach(c => {
      const category = (c.category || 'non-categorizzato').trim() || 'non-categorizzato'
      const macroKey = c.macro || macroFromLabel(category)
      const color = getCategoryColor(category, macroKey)
      const current = counts.get(category)
      if (current) {
        current.count += 1
      } else {
        counts.set(category, { count: 1, color })
      }
    })
    return Array.from(counts.entries())
      .map(([category, info]) => ({ category, count: info.count, color: info.color }))
      .sort((a, b) => {
        if (b.count !== a.count) return b.count - a.count
        return a.category.localeCompare(b.category)
      })
  }, [commentsEnriched, getCategoryColor])

  const commentColors = useMemo(() => {
    const set = new Set<string>()
    commentsEnriched.forEach(c => {
      if (c.macro) set.add(c.macro)
    })
    return Array.from(set)
  }, [commentsEnriched])

  const allCategories = useMemo(() => {
    const set = new Set<string>()
    commentsEnriched.forEach(c => {
      const label = (c.category || 'non-categorizzato').trim() || 'non-categorizzato'
      set.add(label)
    })
    if (!set.size) return [] as string[]
    return Array.from(set).sort()
  }, [commentsEnriched])

  const dashboardFiltered = useMemo(() => {
    let arr = commentsEnriched
    if (!arr.length) return [] as typeof commentsEnriched
    if (dbTipo.length) {
      arr = arr.filter(c => {
        const fm = c.filename ? meta[c.filename] : undefined
        return fm ? dbTipo.includes(fm.tipo) : false
      })
    }
    if (dbScuola.length) {
      const targets = dbScuola.map(s => s.toLowerCase())
      arr = arr.filter(c => {
        const fm = c.filename ? meta[c.filename] : undefined
        return fm ? targets.includes((fm.scuola || '').toLowerCase()) : false
      })
    }
    if (dbRuolo.length) {
      const targets = dbRuolo.map(s => s.toLowerCase())
      arr = arr.filter(c => {
        const fm = c.filename ? meta[c.filename] : undefined
        return fm ? targets.includes((fm.ruolo || '').toLowerCase()) : false
      })
    }
    if (dbGruppo.length) {
      const targets = dbGruppo.map(s => s.toLowerCase())
      arr = arr.filter(c => {
        const fm = c.filename ? meta[c.filename] : undefined
        return fm ? targets.includes((fm.gruppo || '').toLowerCase()) : false
      })
    }
    if (dbIntervistatore.length) {
      const targets = dbIntervistatore.map(s => s.toLowerCase())
      arr = arr.filter(c => {
        const fm = c.filename ? meta[c.filename] : undefined
        return fm ? targets.includes((fm.intervistatore || '').toLowerCase()) : false
      })
    }
    if (dbIntervistato.length) {
      const targets = dbIntervistato.map(s => s.toLowerCase())
      arr = arr.filter(c => {
        const fm = c.filename ? meta[c.filename] : undefined
        return fm ? targets.includes((fm.intervistato || '').toLowerCase()) : false
      })
    }
    if (dbQuery.trim()) {
      const q = dbQuery.trim().toLowerCase()
      arr = arr.filter(c =>
        (c.text || '').toLowerCase().includes(q) ||
        (c.quoted || '').toLowerCase().includes(q) ||
        (c.highlightText || '').toLowerCase().includes(q)
      )
    }
    if (dbCategoryFilter.length) {
      const targets = new Set(dbCategoryFilter)
      arr = arr.filter(c => {
        const label = (c.category || 'non-categorizzato').trim() || 'non-categorizzato'
        return targets.has(label)
      })
    }
    return arr
  }, [commentsEnriched, meta, dbTipo, dbScuola, dbRuolo, dbGruppo, dbIntervistatore, dbIntervistato, dbQuery, dbCategoryFilter])

  const dashMacroStats = useMemo(() => {
    if (!dashboardFiltered.length) return [] as { key: string; label: string; count: number; color: string }[]
    const counts = new Map<string, { key: string; label: string; count: number; color: string }>()
    dashboardFiltered.forEach(c => {
      const macroKey = c.macro || macroFromLabel(c.category || '')
      const key = macroKey || 'non-categorizzato'
      const label = macroKey ? `${macroKey} – ${MACRO_INFO[macroKey].label}` : 'non-categorizzato'
      const color = macroKey ? MACRO_INFO[macroKey].color : '#e9ecef'
      const current = counts.get(key)
      if (current) {
        current.count += 1
      } else {
        counts.set(key, { key, label, count: 1, color })
      }
    })
    return Array.from(counts.values()).sort((a, b) => {
      if (b.count !== a.count) return b.count - a.count
      return a.label.localeCompare(b.label)
    })
  }, [dashboardFiltered])

  const dashCategoryStats = useMemo(() => {
    if (!dashboardFiltered.length) return [] as { label: string; count: number; color: string; macroLabel: string }[]
    const counts = new Map<string, { label: string; count: number; color: string; macroLabel: string }>()
    dashboardFiltered.forEach(c => {
      const label = (c.category || 'non-categorizzato').trim() || 'non-categorizzato'
      const macroKey = c.macro || macroFromLabel(label)
      const macroLabel = macroKey ? `${macroKey} – ${MACRO_INFO[macroKey].label}` : '—'
      const color = getCategoryColor(label, macroKey)
      const current = counts.get(label)
      if (current) {
        current.count += 1
      } else {
        counts.set(label, { label, count: 1, color, macroLabel })
      }
    })
    return Array.from(counts.values()).sort((a, b) => {
      if (b.count !== a.count) return b.count - a.count
      return a.label.localeCompare(b.label)
    })
  }, [dashboardFiltered, getCategoryColor])

  type ComparisonSlice = {
    key: string
    label: string
    categories: {
      macroKey: string
      macroLabel: string
      color: string
      count: number
      sub: { label: string; count: number }[]
    }[]
    total: number
  }

  const groupByField = React.useCallback((field: 'scuola' | 'ruolo') => {
    const groups = new Map<string, typeof dashboardFiltered>()
    dashboardFiltered.forEach(c => {
      const fm = c.filename ? meta[c.filename] : undefined
      const raw = fm ? (fm[field] || '—') : '—'
      const key = raw.trim() || '—'
      ;(groups.get(key) || groups.set(key, []).get(key)!)
    })
    return groups
  }, [dashboardFiltered, meta])

  const dashCompareScuole = useMemo(() => {
    if (!dashboardFiltered.length) return [] as ComparisonSlice[]
    const bySchool = new Map<string, typeof dashboardFiltered>()
    dashboardFiltered.forEach(c => {
      const fm = c.filename ? meta[c.filename] : undefined
      const school = (fm?.scuola || '—').trim() || '—'
      if (!bySchool.has(school)) bySchool.set(school, [])
      bySchool.get(school)!.push(c)
    })
    const slices: ComparisonSlice[] = []
    bySchool.forEach((items, school) => {
      const macroMap = new Map<string, { macroKey: string; macroLabel: string; color: string; count: number; sub: Map<string, number> }>()
      items.forEach(item => {
        const macroKey = item.macro || macroFromLabel(item.category || '') || 'NC'
        const macroInfo = macroKey !== 'NC' ? MACRO_INFO[macroKey] : { label: 'non-categorizzato', color: '#e9ecef' }
        if (!macroMap.has(macroKey)) {
          macroMap.set(macroKey, {
            macroKey,
            macroLabel: macroKey !== 'NC' ? `${macroKey} – ${macroInfo.label}` : 'non-categorizzato',
            color: macroInfo.color,
            count: 0,
            sub: new Map<string, number>(),
          })
        }
        const entry = macroMap.get(macroKey)!
        entry.count += 1
        const label = (item.category || 'non-categorizzato').trim() || 'non-categorizzato'
        entry.sub.set(label, (entry.sub.get(label) || 0) + 1)
      })
      const categories = Array.from(macroMap.values()).map(macro => ({
        macroKey: macro.macroKey,
        macroLabel: macro.macroLabel,
        color: macro.color,
        count: macro.count,
        sub: Array.from(macro.sub.entries()).map(([label, count]) => ({ label, count })).sort((a, b) => b.count - a.count),
      })).sort((a, b) => b.count - a.count)
      slices.push({ key: school, label: school, categories, total: items.length })
    })
    return slices.sort((a, b) => b.total - a.total)
  }, [dashboardFiltered, meta])

  const dashCompareRuoli = useMemo(() => {
    if (!dashboardFiltered.length) return [] as ComparisonSlice[]
    const byRole = new Map<string, typeof dashboardFiltered>()
    dashboardFiltered.forEach(c => {
      const fm = c.filename ? meta[c.filename] : undefined
      const role = (fm?.ruolo || '—').trim() || '—'
      if (!byRole.has(role)) byRole.set(role, [])
      byRole.get(role)!.push(c)
    })
    const slices: ComparisonSlice[] = []
    byRole.forEach((items, role) => {
      const macroMap = new Map<string, { macroKey: string; macroLabel: string; color: string; count: number; sub: Map<string, number> }>()
      items.forEach(item => {
        const macroKey = item.macro || macroFromLabel(item.category || '') || 'NC'
        const macroInfo = macroKey !== 'NC' ? MACRO_INFO[macroKey] : { label: 'non-categorizzato', color: '#e9ecef' }
        if (!macroMap.has(macroKey)) {
          macroMap.set(macroKey, {
            macroKey,
            macroLabel: macroKey !== 'NC' ? `${macroKey} – ${macroInfo.label}` : 'non-categorizzato',
            color: macroInfo.color,
            count: 0,
            sub: new Map<string, number>(),
          })
        }
        const entry = macroMap.get(macroKey)!
        entry.count += 1
        const label = (item.category || 'non-categorizzato').trim() || 'non-categorizzato'
        entry.sub.set(label, (entry.sub.get(label) || 0) + 1)
      })
      const categories = Array.from(macroMap.values()).map(macro => ({
        macroKey: macro.macroKey,
        macroLabel: macro.macroLabel,
        color: macro.color,
        count: macro.count,
        sub: Array.from(macro.sub.entries()).map(([label, count]) => ({ label, count })).sort((a, b) => b.count - a.count),
      })).sort((a, b) => b.count - a.count)
      slices.push({ key: role, label: role, categories, total: items.length })
    })
    return slices.sort((a, b) => b.total - a.total)
  }, [dashboardFiltered, meta])

  const compareSlices = useMemo(() => (compareMode === 'scuola' ? dashCompareScuole : dashCompareRuoli), [compareMode, dashCompareScuole, dashCompareRuoli])
  const compareOptions = useMemo(() => compareSlices.map(slice => ({ value: slice.key, label: slice.label })), [compareSlices])

  React.useEffect(() => {
    if (!compareSlices.length) {
      if (compareSelectionA) setCompareSelectionA('')
      if (compareSelectionB) setCompareSelectionB('')
      return
    }
    if (!compareSelectionA || !compareSlices.some(s => s.key === compareSelectionA)) {
      setCompareSelectionA(compareSlices[0].key)
    }
    let desiredB = compareSelectionB
    if (!desiredB || !compareSlices.some(s => s.key === desiredB)) {
      desiredB = compareSlices.find(s => s.key !== compareSlices[0].key)?.key || compareSlices[0].key
    }
    if (compareSelectionA && desiredB === compareSelectionA && compareSlices.length > 1) {
      desiredB = compareSlices.find(s => s.key !== compareSelectionA)?.key || desiredB
    }
    if (desiredB !== compareSelectionB) {
      setCompareSelectionB(desiredB)
    }
  }, [compareSlices, compareSelectionA, compareSelectionB])

  const compareSliceA = useMemo(() => compareSlices.find(s => s.key === compareSelectionA), [compareSlices, compareSelectionA])
  const compareSliceB = useMemo(() => compareSlices.find(s => s.key === compareSelectionB), [compareSlices, compareSelectionB])

  const filteredComments = useMemo(() => {
    let arr = commentsEnriched
    if (fDoc.length) arr = arr.filter(c => fDoc.includes((c.filename || '')))
    if (fCode.length) {
      arr = arr.filter(c => {
        const tokensUpper = (c.codeTokens || []).map(t => t.toUpperCase())
        if (!tokensUpper.length) return false
        return fCode.some(code => tokensUpper.includes(code.toUpperCase()))
      })
    }
    if (fColor.length) {
      arr = arr.filter(c => {
        if (!c.macro) return false
        return fColor.includes(c.macro)
      })
    }
    if (fCategory.length) {
      arr = arr.filter(c => {
        const label = (c.category || 'non-categorizzato').trim() || 'non-categorizzato'
        return fCategory.includes(label)
      })
    }
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
    commentsEnriched.forEach(c => {
      (c.codeTokens || []).forEach(token => set.add(token.toUpperCase()))
    })
    return Array.from(set).sort()
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

  function resetDashboardFilters() {
    setDbTipo([])
    setDbScuola([])
    setDbRuolo([])
    setDbGruppo([])
    setDbIntervistatore([])
    setDbIntervistato([])
    setDbQuery('')
    setDbCategoryFilter([])
  }

  const renderHistogram = (items: { label: string; count: number; color: string; extra?: string }[]) => {
    if (!items.length) return <div style={{ color: '#666' }}>Nessun dato.</div>
    const max = Math.max(...items.map(i => i.count), 1)
    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        {items.map(item => {
          const width = 12 + Math.max(12, Math.round((item.count / max) * 220))
          return (
            <div key={item.label} style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <span style={{ minWidth: 32, padding: '2px 8px', borderRadius: 999, background: '#e9ecef', fontVariantNumeric: 'tabular-nums', textAlign: 'center' }}>{item.count}</span>
                <span>{item.extra ? `${item.extra} · ${item.label}` : item.label}</span>
              </div>
              <div style={{ height: 12, background: '#f1f3f5', borderRadius: 6, overflow: 'hidden' }}>
                <div style={{ width, height: '100%', background: item.color, transition: 'width 0.3s ease' }} />
              </div>
            </div>
          )
        })}
      </div>
    )
  }

  const renderPie = (items: { label: string; count: number; color: string; extra?: string }[]) => {
    if (!items.length) return <div style={{ color: '#666' }}>Nessun dato.</div>
    const total = items.reduce((sum, item) => sum + item.count, 0)
    if (!total) return <div style={{ color: '#666' }}>Nessun dato.</div>
    let acc = 0
    const segments = items.map(item => {
      const start = (acc / total) * 100
      acc += item.count
      const end = (acc / total) * 100
      return `${item.color} ${start.toFixed(2)}% ${end.toFixed(2)}%`
    })
    const gradient = `conic-gradient(${segments.join(', ')})`
    return (
      <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap', alignItems: 'center' }}>
        <div style={{ width: 180, height: 180, borderRadius: '50%', background: gradient, border: '1px solid #dee2e6' }} />
        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
          {items.map(item => (
            <div key={item.label} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ minWidth: 32, padding: '2px 8px', borderRadius: 999, background: '#e9ecef', fontVariantNumeric: 'tabular-nums', textAlign: 'center' }}>{item.count}</span>
              <span style={{ width: 14, height: 14, borderRadius: 4, background: item.color, border: '1px solid #ced4da' }} />
              <span>{item.extra ? `${item.extra} · ${item.label}` : item.label}</span>
            </div>
          ))}
        </div>
      </div>
    )
  }

  const renderChart = (items: { label: string; count: number; color: string; extra?: string }[], chartType: 'istogramma' | 'torta') => {
    return chartType === 'torta' ? renderPie(items) : renderHistogram(items)
  }

  const renderComparisonPanel = (title: string, slice?: ComparisonSlice) => {
    if (!slice) {
      return (
        <div style={{ border: '1px dashed #ced4da', borderRadius: 8, padding: 16, color: '#666', minHeight: 180 }}>
          Seleziona {title.toLowerCase()} da confrontare.
        </div>
      )
    }
    return (
      <div style={{ border: '1px solid #f1f3f5', borderRadius: 8, padding: 16, background: '#fff', boxShadow: dbCompareLayout === 'columns' ? '0 1px 3px rgba(15, 23, 42, 0.08)' : 'none' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 12 }}>
          <span style={{ minWidth: 40, padding: '4px 10px', borderRadius: 999, background: '#dee2e6', fontWeight: 600, textAlign: 'center', fontVariantNumeric: 'tabular-nums' }}>{slice.total}</span>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
            <div style={{ fontWeight: 600, color: '#212529' }}>{slice.label}</div>
            <span style={{ fontSize: 11, color: '#495057' }}>Commenti totali</span>
          </div>
        </div>
        {slice.categories.length === 0 && <div style={{ color: '#666' }}>Nessuna categoria disponibile.</div>}
        {slice.categories.length > 0 && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            {slice.categories.map(cat => (
              <div key={`${slice.key}-${cat.macroKey}`} style={{ border: '1px solid #e9ecef', borderRadius: 6, padding: 10, background: '#fafafb' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: cat.sub.length ? 8 : 0 }}>
                  <span style={{ minWidth: 32, padding: '2px 8px', borderRadius: 999, background: '#e9ecef', fontVariantNumeric: 'tabular-nums', textAlign: 'center' }}>{cat.count}</span>
                  <span style={{ width: 12, height: 12, borderRadius: 4, background: cat.color, border: '1px solid #ced4da' }} />
                  <strong>{cat.macroLabel}</strong>
                </div>
                {cat.sub.length > 0 && (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4, paddingLeft: 12 }}>
                    {cat.sub.map(sub => (
                      <div key={`${slice.key}-${cat.macroKey}-${sub.label}`} style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12 }}>
                        <span style={{ minWidth: 28, padding: '1px 6px', borderRadius: 6, background: '#f1f3f5', fontVariantNumeric: 'tabular-nums', textAlign: 'center' }}>{sub.count}</span>
                        <span>{sub.label}</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            ))}
          </div>
        )}
      </div>
    )
  }
  function exportComments(rows: (typeof commentsEnriched)[number][]) {
    const mapped = rows.map(c => {
      const m = c.filename ? meta[c.filename] : undefined
      const codeString = (c.codeTokens && c.codeTokens.length) ? c.codeTokens.join(', ') : (c.code || '')
      const macroLabel = c.macro ? `${c.macro} – ${MACRO_INFO[c.macro].label}` : ''
      return {
        filename: c.filename || '',
        id: c.id,
        author: c.author,
        date: c.date,
        codes: codeString,
        category: c.category || '',
        macro: macroLabel,
        color_hex: c.color || '',
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
      const codeString = (c.codeTokens && c.codeTokens.length) ? c.codeTokens.join(', ') : (c.code || '')
      const macroLabel = c.macro ? `${c.macro} – ${MACRO_INFO[c.macro].label}` : ''
      return {
        File: c.filename || '',
        ID: c.id,
        Autore: c.author,
        Data: c.date,
        Codici: codeString,
        Categoria: c.category || '',
        Macro: macroLabel,
        "Colore hex": c.color || '',
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
    const setTipo = new Set<string>()
    const setI = new Set<string>()
    const setT = new Set<string>()
    const setR = new Set<string>()
    const setS = new Set<string>()
    const setG = new Set<string>()
    for (const d of docs) {
      const fm = meta[d]
      if (fm?.tipo) setTipo.add(fm.tipo)
      if (fm?.intervistatore) setI.add(fm.intervistatore)
      if (fm?.intervistato) setT.add(fm.intervistato)
      if (fm?.ruolo) setR.add(fm.ruolo)
      if (fm?.scuola) setS.add(fm.scuola)
      if (fm?.gruppo) setG.add(fm.gruppo)
    }
    return {
      tipi: Array.from(setTipo),
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
      const filenamesSet = new Set<string>()
      for (const h of json.highlights as Highlight[]) {
        if (h?.filename) filenamesSet.add(h.filename)
      }
      for (const c of json.comments as CommentItem[]) {
        if (c?.filename) filenamesSet.add(c.filename)
      }
      const filenames: string[] = Array.from(filenamesSet)
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
          <button onClick={() => setTab('legenda')} aria-label="Legenda" title="Legenda" style={{ padding: '6px 10px', borderRadius: 6, border: tab==='legenda' ? '2px solid #444' : '1px solid #ddd', background: tab==='legenda' ? '#f0f0f0' : '#fff', display:'inline-flex', alignItems:'center', justifyContent:'center' }}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M12 3a9 9 0 1 0 9 9 9.01 9.01 0 0 0-9-9Zm0 4a1.2 1.2 0 1 1 0 2.4A1.2 1.2 0 0 1 12 7Zm1 10h-2v-1h1v-3h-1v-1h2v4h1v1Z" fill="currentColor"/></svg>
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

            {tab === 'legenda' && (
              <>
                <h3>Legenda categorie</h3>
                <p style={{ maxWidth: 680, color: '#495057' }}>
                  I colori mostrati nell’interfaccia seguono la macro-categoria associata a ciascun codice. L’etichetta scelta per un commento prevale sul colore originale presente nel file DOCX.
                </p>
                <h4>Macro categorie</h4>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 12, marginBottom: 18 }}>
                  {Object.entries(MACRO_INFO).map(([macro, info]) => (
                    <div key={macro} style={{ border: '1px solid #e9ecef', borderRadius: 8, padding: 12, background: '#fff', display: 'flex', flexDirection: 'column', gap: 6 }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <span style={{ width: 16, height: 16, borderRadius: 4, background: info.color, border: '1px solid #ced4da' }} />
                        <strong>{info.label}</strong>
                      </div>
                      <div style={{ fontSize: 12, color: '#495057' }}>Codice macro: <code>{macro}</code></div>
                      <div style={{ fontSize: 12, color: '#495057' }}>Colore: <code>{info.color}</code></div>
                    </div>
                  ))}
                </div>
                {Object.entries(CODE_LIST_BY_TYPE).map(([tipo, codes]) => (
                  <div key={tipo} style={{ marginBottom: 28 }}>
                    <h4 style={{ marginBottom: 6 }}>Codici {tipo}</h4>
                    <table style={{ width: '100%', borderCollapse: 'collapse', background: '#fff', border: '1px solid #e9ecef', borderRadius: 8, overflow: 'hidden' }}>
                      <thead style={{ background: '#f1f3f5' }}>
                        <tr>
                          <th style={{ textAlign: 'left', padding: 8, borderBottom: '1px solid #dee2e6' }}>Codice</th>
                          <th style={{ textAlign: 'left', padding: 8, borderBottom: '1px solid #dee2e6' }}>Sotto categoria</th>
                          <th style={{ textAlign: 'left', padding: 8, borderBottom: '1px solid #dee2e6' }}>Macro</th>
                          <th style={{ textAlign: 'left', padding: 8, borderBottom: '1px solid #dee2e6' }}>Colore</th>
                        </tr>
                      </thead>
                      <tbody>
                        {codes.map(code => {
                          const normalized = normalizeCodeToken(code)
                          const info = CODE_INFO[normalized]
                          const label = info ? info.label : (codeMap[normalized] || defaultLabelForCode(normalized))
                          const macroKey = info ? info.macro : macroFromCode(normalized)
                          const macroLabel = macroKey ? MACRO_INFO[macroKey].label : '—'
                          const color = getCategoryColor(label, macroKey)
                          return (
                            <tr key={`${tipo}-${code}`}>
                              <td style={{ padding: 8, borderBottom: '1px solid #f1f3f5' }}><code>{normalized}</code></td>
                              <td style={{ padding: 8, borderBottom: '1px solid #f1f3f5' }}>{label}</td>
                              <td style={{ padding: 8, borderBottom: '1px solid #f1f3f5' }}>{macroLabel}</td>
                              <td style={{ padding: 8, borderBottom: '1px solid #f1f3f5', display: 'flex', alignItems: 'center', gap: 6 }}>
                                <span style={{ width: 14, height: 14, borderRadius: 4, border: '1px solid #ced4da', background: color }} />
                                <code>{color}</code>
                              </td>
                            </tr>
                          )
                        })}
                      </tbody>
                    </table>
                  </div>
                ))}
              </>
            )}

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
                    const tokens = Array.from(new Set([...commentCodes, ...Object.keys(DEFAULT_CODE_MAP)])).sort()
                    return tokens.map(t => (
                      <div key={t} style={{ display: 'flex', alignItems: 'center', gap: 6, background: '#fff', border: '1px solid #eee', borderRadius: 6, padding: 6 }}>
                        <span style={{ minWidth: 120 }}>{t}</span>
                        <input
                          value={codeMap[t] || ''}
                          onChange={e => setCodeMap({ ...codeMap, [t]: e.target.value })}
                          placeholder={defaultLabelForCode(t)}
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
                      <span style={{ width: 12, height: 12, background: getCategoryColor(cat, macroFromLabel(cat)), display: 'inline-block', border: '1px solid #ccc' }} />
                      <span style={{ minWidth: 120 }}>{cat}</span>
                      <input type="color" value={categoryColors[cat] || getCategoryColor(cat, macroFromLabel(cat))} onChange={e => setCategoryColors({ ...categoryColors, [cat]: e.target.value })} />
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
          const resolved = resolveHighlightCategory(h)
          const cat = resolved.category || 'non-categorizzato'
          const labelColor = resolved.color || '#f8f9fa'
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
                    {commentCodes.map(c => {
                      const label = (codeMap[c] || '').trim() || defaultLabelForCode(c)
                      return (<option key={c} value={c}>{`${c} – ${label}`}</option>)
                    })}
                  </select>
                  <div style={{ display: 'flex', gap: 4 }}>
                    <button onClick={() => setFCode(commentCodes)}>Tutti</button>
                    <button onClick={() => setFCode([])}>Nessuno</button>
                  </div>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                  <select multiple value={fColor} onChange={e => setFColor(readMulti(e))} size={Math.min(6, Math.max(3, commentColors.length))}>
                    {commentColors.map(c => {
                      const info = (MACRO_INFO as Record<string, { label: string }>)[c]
                      return (<option key={c} value={c}>{info ? `${c} – ${info.label}` : c}</option>)
                    })}
                  </select>
                  <div style={{ display: 'flex', gap: 4 }}>
                    <button onClick={() => setFColor(commentColors)}>Tutti</button>
                    <button onClick={() => setFColor([])}>Nessuno</button>
                  </div>
                  </div>
                  {(() => {
                    const cats = Array.from(new Set(commentsEnriched.map(c => (c.category || 'non-categorizzato').trim() || 'non-categorizzato'))).sort()
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
                  const cat = (c.category || 'non-categorizzato').trim() || 'non-categorizzato'
                  const labelColor = getCategoryColor(cat, c.macro || macroFromLabel(cat))
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
                        <span style={{ fontSize: 12 }}>Codici: <strong>{(c.codeTokens && c.codeTokens.length) ? c.codeTokens.join(', ') : (c.code || '—')}</strong></span>
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
                        {commentCodes.map(c => {
                          const label = (codeMap[c] || '').trim() || defaultLabelForCode(c)
                          return (<option key={c} value={c}>{`${c} – ${label}`}</option>)
                        })}
                      </select>
                      <div style={{ display: 'flex', gap: 4 }}>
                        <button onClick={() => setKCode(commentCodes)}>Tutti</button>
                        <button onClick={() => setKCode([])}>Nessuno</button>
                      </div>
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                      <select multiple value={kColor} onChange={e => setKColor(readMulti(e))} size={Math.min(6, Math.max(3, commentColors.length))}>
                        {commentColors.map(c => {
                          const info = (MACRO_INFO as Record<string, { label: string }>)[c]
                          return (<option key={c} value={c}>{info ? `${c} – ${info.label}` : c}</option>)
                        })}
                      </select>
                      <div style={{ display: 'flex', gap: 4 }}>
                        <button onClick={() => setKColor(commentColors)}>Tutti</button>
                        <button onClick={() => setKColor([])}>Nessuno</button>
                      </div>
                    </div>
                    {(() => {
                      const cats = Array.from(new Set(commentsEnriched.map(c => (c.category || 'non-categorizzato').trim() || 'non-categorizzato'))).sort()
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
                  if (kCode.length) {
                    items = items.filter(c => {
                      const toks = (c.codeTokens || []).map(t => t.toUpperCase())
                      if (!toks.length) return false
                      return kCode.some(code => toks.includes(code.toUpperCase()))
                    })
                  }
                  if (kColor.length) {
                    items = items.filter(c => {
                      if (!c.macro) return false
                      return kColor.includes(c.macro)
                    })
                  }
                  if (kCategory.length) {
                    items = items.filter(c => {
                      const label = (c.category || 'non-categorizzato').trim() || 'non-categorizzato'
                      return kCategory.includes(label)
                    })
                  }
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
                    for (const c of items) {
                      const tokens = c.codeTokens && c.codeTokens.length ? c.codeTokens : []
                      if (tokens.length === 0) {
                        noToken.push(c)
                      } else {
                        for (const token of tokens) {
                          tokenSet.add(token)
                          ;(tokenMap[token] ||= []).push(c)
                        }
                      }
                    }
                    columns = [...Array.from(tokenSet).sort(), '—']
                    groups = { ...tokenMap, '—': noToken }
                  }

                  // Drag-and-drop disabled by request

                  return (
                    <div style={{ display: 'flex', gap: 12, alignItems: 'flex-start', overflowX: 'auto' }}>
                      {columns.map(col => {
                        const header = col === '—'
                          ? '—'
                          : `${col} – ${(codeMap[col] || '').trim() || defaultLabelForCode(col)}`
                        return (
                          <div key={col}
                            style={{ minWidth: 260, background: '#fff', border: '1px solid #e9ecef', borderRadius: 8, padding: 10 }}>
                            <div style={{ fontWeight: 700, marginBottom: 8 }}>{header} <span style={{ color: '#868e96' }}>({(groups[col]||[]).length})</span></div>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                              {(groups[col] || []).map(c => {
                                const categoryLabel = (c.category || 'non-categorizzato').trim() || 'non-categorizzato'
                                const macroKey = c.macro || macroFromLabel(categoryLabel)
                                const badgeColor = getCategoryColor(categoryLabel, macroKey)
                                return (
                                  <div key={c.id}
                                    style={{ border: '1px solid #eee', borderLeft: `4px solid ${badgeColor}`, padding: 8, borderRadius: 6, background: '#f8f9fa' }}>
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
                                      <span style={{ fontSize: 11, padding: '2px 6px', borderRadius: 999, background: badgeColor, color: '#212529' }}>{categoryLabel}</span>
                                      {(c.codeTokens && c.codeTokens.length ? c.codeTokens : (c.code ? [c.code] : [])).map(token => {
                                        const label = (codeMap[token] || '').trim() || defaultLabelForCode(token)
                                        return (
                                          <span key={token} style={{ fontSize: 11, padding: '2px 6px', borderRadius: 999, background: '#fff', border: '1px solid #dee2e6', color: '#343a40' }}>{`${token} – ${label}`}</span>
                                        )
                                      })}
                                    </div>
                                    <div style={{ display: 'flex', gap: 8, marginTop: 6, flexWrap: 'wrap' }}>
                                      {c.highlight && <button onClick={() => jumpTo(c.highlight!)} style={{ fontSize: 12 }}>Vai al testo</button>}
                                      {c.highlight && (
                                        <button onClick={() => scheduleJump({
                                          filename: c.highlight?.filename || c.filename || '',
                                          type: 'highlight',
                                          highlight_color: c.highlight?.highlight_color || null,
                                          text: c.highlight?.text || c.highlightText || '',
                                          context: c.highlight?.context || c.quoted || '',
                                          paragraph: c.highlight?.paragraph || '',
                                          offset_start: c.highlight?.offset_start || 0,
                                          offset_end: c.highlight?.offset_end || 0,
                                          para_index: c.highlight?.para_index || null,
                                        } as Highlight)} style={{ fontSize: 12 }}>Evidenza</button>
                                      )}
                                    </div>
                                  </div>
                                )
                              })}
                            </div>
                          </div>
                        )
                      })}
                    </div>
                  )
                })()}
              </>
            )}

            {tab === 'documenti' && (
              <>
                <h3>Documenti</h3>
                {(() => {
                  const fileSet = new Set<string>()
                  for (const h of (data ? (data.highlights as Highlight[]) : [])) {
                    if (h?.filename) fileSet.add(h.filename)
                  }
                  for (const c of (data ? (data.comments as CommentItem[]) : [])) {
                    if (c?.filename) fileSet.add(c.filename)
                  }
                  for (const key of Object.keys(meta)) fileSet.add(key)
                  
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
                <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', alignItems: 'flex-start', marginBottom: 16 }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                    <label style={{ fontSize: 12, color: '#495057' }}>Tipo documento</label>
                    <select multiple value={dbTipo} onChange={e => setDbTipo(readMulti(e))} size={Math.min(6, Math.max(3, (optionsFromMeta.tipi || []).length || 0))}>
                      {(optionsFromMeta.tipi || []).map(v => (<option key={v} value={v}>{v}</option>))}
                    </select>
                    <div style={{ display: 'flex', gap: 4 }}>
                      <button onClick={() => setDbTipo(optionsFromMeta.tipi || [])}>Tutti</button>
                      <button onClick={() => setDbTipo([])}>Nessuno</button>
                    </div>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                    <label style={{ fontSize: 12, color: '#495057' }}>Scuola</label>
                    <select multiple value={dbScuola} onChange={e => setDbScuola(readMulti(e))} size={Math.min(6, Math.max(3, optionsFromMeta.scuole.length))}>
                      {optionsFromMeta.scuole.map(v => (<option key={v} value={v}>{v}</option>))}
                    </select>
                    <div style={{ display: 'flex', gap: 4 }}>
                      <button onClick={() => setDbScuola(optionsFromMeta.scuole)}>Tutti</button>
                      <button onClick={() => setDbScuola([])}>Nessuno</button>
                    </div>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                    <label style={{ fontSize: 12, color: '#495057' }}>Ruolo</label>
                    <select multiple value={dbRuolo} onChange={e => setDbRuolo(readMulti(e))} size={Math.min(6, Math.max(3, optionsFromMeta.ruoli.length))}>
                      {optionsFromMeta.ruoli.map(v => (<option key={v} value={v}>{v}</option>))}
                    </select>
                    <div style={{ display: 'flex', gap: 4 }}>
                      <button onClick={() => setDbRuolo(optionsFromMeta.ruoli)}>Tutti</button>
                      <button onClick={() => setDbRuolo([])}>Nessuno</button>
                    </div>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                    <label style={{ fontSize: 12, color: '#495057' }}>Gruppo</label>
                    <select multiple value={dbGruppo} onChange={e => setDbGruppo(readMulti(e))} size={Math.min(6, Math.max(3, optionsFromMeta.gruppi.length))}>
                      {optionsFromMeta.gruppi.map(v => (<option key={v} value={v}>{v}</option>))}
                    </select>
                    <div style={{ display: 'flex', gap: 4 }}>
                      <button onClick={() => setDbGruppo(optionsFromMeta.gruppi)}>Tutti</button>
                      <button onClick={() => setDbGruppo([])}>Nessuno</button>
                    </div>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                    <label style={{ fontSize: 12, color: '#495057' }}>Intervistatore</label>
                    <select multiple value={dbIntervistatore} onChange={e => setDbIntervistatore(readMulti(e))} size={Math.min(6, Math.max(3, optionsFromMeta.intervistatori.length))}>
                      {optionsFromMeta.intervistatori.map(v => (<option key={v} value={v}>{v}</option>))}
                    </select>
                    <div style={{ display: 'flex', gap: 4 }}>
                      <button onClick={() => setDbIntervistatore(optionsFromMeta.intervistatori)}>Tutti</button>
                      <button onClick={() => setDbIntervistatore([])}>Nessuno</button>
                    </div>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                    <label style={{ fontSize: 12, color: '#495057' }}>Intervistato</label>
                    <select multiple value={dbIntervistato} onChange={e => setDbIntervistato(readMulti(e))} size={Math.min(6, Math.max(3, optionsFromMeta.intervistati.length))}>
                      {optionsFromMeta.intervistati.map(v => (<option key={v} value={v}>{v}</option>))}
                    </select>
                    <div style={{ display: 'flex', gap: 4 }}>
                      <button onClick={() => setDbIntervistato(optionsFromMeta.intervistati)}>Tutti</button>
                      <button onClick={() => setDbIntervistato([])}>Nessuno</button>
                    </div>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                    <label style={{ fontSize: 12, color: '#495057' }}>Categoria</label>
                    <select
                      multiple
                      value={dbCategoryFilter}
                      onChange={e => setDbCategoryFilter(readMulti(e))}
                      size={Math.min(6, Math.max(3, (allCategories.length || 3)))}
                      disabled={allCategories.length === 0}
                    >
                      {allCategories.length === 0 && <option value="">(Nessuna categoria)</option>}
                      {allCategories.map(cat => (
                        <option key={cat} value={cat}>{cat}</option>
                      ))}
                    </select>
                    <div style={{ display: 'flex', gap: 4 }}>
                      <button onClick={() => setDbCategoryFilter(allCategories)}>Tutte</button>
                      <button onClick={() => setDbCategoryFilter([])}>Nessuna</button>
                    </div>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4, flex: '1 1 220px', minWidth: 200 }}>
                    <label style={{ fontSize: 12, color: '#495057' }}>Ricerca testo</label>
                    <input value={dbQuery} onChange={e => setDbQuery(e.target.value)} placeholder="Cerca in commenti e highlight" />
                    <button onClick={resetDashboardFilters}>Reimposta filtri</button>
                  </div>
                </div>
                <div style={{ marginBottom: 12, color: '#495057' }}>Commenti filtrati: <strong>{dashboardFiltered.length}</strong></div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 24 }}>
                  <div style={{ background: '#fff', border: '1px solid #eee', borderRadius: 8, padding: 16 }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
                      <div>
                        <h4 style={{ margin: 0 }}>Macro categorie</h4>
                        <small style={{ color: '#495057' }}>Distribuzione dei commenti per macro categoria</small>
                      </div>
                      <div style={{ display: 'inline-flex', border: '1px solid #dee2e6', borderRadius: 999, overflow: 'hidden' }}>
                        <button onClick={() => setDbMacroChart('istogramma')} style={{ padding: '6px 12px', border: 'none', background: dbMacroChart==='istogramma' ? '#f1f3f5' : '#fff' }}>Istogramma</button>
                        <button onClick={() => setDbMacroChart('torta')} style={{ padding: '6px 12px', borderLeft: '1px solid #dee2e6', background: dbMacroChart==='torta' ? '#f1f3f5' : '#fff' }}>Torta</button>
                      </div>
                    </div>
                    <div style={{ marginTop: 16 }}>
                      {renderChart(dashMacroStats.map(item => ({ label: item.label, count: item.count, color: item.color })), dbMacroChart)}
                    </div>
                    {dashMacroStats.length > 0 && (
                      <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: 16 }}>
                        <thead>
                          <tr>
                            <th style={{ textAlign: 'left', borderBottom: '1px solid #dee2e6', paddingBottom: 6 }}>Conteggio · Macro categoria</th>
                          </tr>
                        </thead>
                        <tbody>
                          {dashMacroStats.map(item => (
                            <tr key={item.key}>
                              <td style={{ padding: '6px 0' }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                  <span style={{ minWidth: 32, padding: '2px 8px', borderRadius: 999, background: '#e9ecef', fontVariantNumeric: 'tabular-nums', textAlign: 'center' }}>{item.count}</span>
                                  <span>{item.label}</span>
                                </div>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )}
                  </div>
                  <div style={{ background: '#fff', border: '1px solid #eee', borderRadius: 8, padding: 16 }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
                      <div>
                        <h4 style={{ margin: 0 }}>Sotto categorie</h4>
                        <small style={{ color: '#495057' }}>Dettaglio etichette (XX_YY)</small>
                      </div>
                      <div style={{ display: 'inline-flex', border: '1px solid #dee2e6', borderRadius: 999, overflow: 'hidden' }}>
                        <button onClick={() => setDbCatChart('istogramma')} style={{ padding: '6px 12px', border: 'none', background: dbCatChart==='istogramma' ? '#f1f3f5' : '#fff' }}>Istogramma</button>
                        <button onClick={() => setDbCatChart('torta')} style={{ padding: '6px 12px', borderLeft: '1px solid #dee2e6', background: dbCatChart==='torta' ? '#f1f3f5' : '#fff' }}>Torta</button>
                      </div>
                    </div>
                    <div style={{ marginTop: 16 }}>
                      {renderChart(dashCategoryStats.map(item => ({ label: item.label, count: item.count, color: item.color, extra: item.macroLabel !== '—' ? item.macroLabel : undefined })), dbCatChart)}
                    </div>
                    {dashCategoryStats.length > 0 && (
                      <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: 16 }}>
                        <thead>
                          <tr>
                            <th style={{ textAlign: 'left', borderBottom: '1px solid #dee2e6', paddingBottom: 6 }}>Conteggio · Categoria</th>
                          </tr>
                        </thead>
                        <tbody>
                          {dashCategoryStats.map(item => (
                            <tr key={item.label}>
                              <td style={{ padding: '6px 0' }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                  <span style={{ minWidth: 32, padding: '2px 8px', borderRadius: 999, background: '#e9ecef', fontVariantNumeric: 'tabular-nums', textAlign: 'center' }}>{item.count}</span>
                                  <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                                    <span>{item.label}</span>
                                    <span style={{ fontSize: 11, color: '#495057' }}>{item.macroLabel}</span>
                                  </div>
                                </div>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )}
                  </div>
                  <div style={{ background: '#fff', border: '1px solid #eee', borderRadius: 8, padding: 16 }}>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                      <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 12 }}>
                        <h4 style={{ margin: 0 }}>Confronto dedicato</h4>
                        <div>
                          <label style={{ fontSize: 12, color: '#495057', marginRight: 6 }}>Confronta per</label>
                          <select value={compareMode} onChange={e => { const mode = e.target.value === 'ruolo' ? 'ruolo' : 'scuola'; setCompareMode(mode) }}>
                            <option value="scuola">Scuola</option>
                            <option value="ruolo">Ruolo</option>
                          </select>
                        </div>
                        <div style={{ display: 'inline-flex', border: '1px solid #dee2e6', borderRadius: 999, overflow: 'hidden' }}>
                          <button onClick={() => setDbCompareLayout('columns')} style={{ padding: '6px 12px', border: 'none', background: dbCompareLayout==='columns' ? '#f1f3f5' : '#fff' }}>Affiancati</button>
                          <button onClick={() => setDbCompareLayout('stack')} style={{ padding: '6px 12px', borderLeft: '1px solid #dee2e6', background: dbCompareLayout==='stack' ? '#f1f3f5' : '#fff' }}>Verticale</button>
                        </div>
                      </div>
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 12 }}>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                          <label style={{ fontSize: 12, color: '#495057' }}>{compareMode === 'scuola' ? 'Scuola A' : 'Ruolo A'}</label>
                          <select value={compareSelectionA} onChange={e => setCompareSelectionA(e.target.value)}>
                            {compareOptions.length === 0 && <option value="">(Nessun valore)</option>}
                            {compareOptions.map(opt => (
                              <option key={opt.value} value={opt.value}>{opt.label}</option>
                            ))}
                          </select>
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                          <label style={{ fontSize: 12, color: '#495057' }}>{compareMode === 'scuola' ? 'Scuola B' : 'Ruolo B'}</label>
                          <select value={compareSelectionB} onChange={e => setCompareSelectionB(e.target.value)}>
                            {compareOptions.length === 0 && <option value="">(Nessun valore)</option>}
                            {compareOptions.map(opt => (
                              <option key={opt.value} value={opt.value}>{opt.label}</option>
                            ))}
                          </select>
                        </div>
                      </div>
                    </div>
                    <div style={{ marginTop: 20, display: 'flex', flexDirection: dbCompareLayout === 'columns' ? 'row' : 'column', gap: 16 }}>
                      <div style={{ flex: 1, minWidth: dbCompareLayout === 'columns' ? 280 : undefined }}>
                        {renderComparisonPanel(compareMode === 'scuola' ? 'Scuola A' : 'Ruolo A', compareSliceA)}
                      </div>
                      <div style={{ flex: 1, minWidth: dbCompareLayout === 'columns' ? 280 : undefined }}>
                        {renderComparisonPanel(compareMode === 'scuola' ? 'Scuola B' : 'Ruolo B', compareSliceB)}
                      </div>
                    </div>
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
          const resolved = resolveHighlightCategory(h)
          const cat = resolved.category || 'non-categorizzato'
          const col = resolved.color || '#bbb'
          return (
            <div
              key={idx}
              onClick={() => jumpTo(h)}
              style={{ cursor: 'pointer', border: '1px solid #eee', padding: 8, borderLeft: `4px solid ${col}`, marginBottom: 8 }}
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
            const resolved = resolveHighlightCategory(h)
            const cat = resolved.category || 'non-categorizzato'
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
