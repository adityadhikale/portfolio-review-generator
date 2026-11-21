// Typed client for the Python API (crest/api.py).

export type Step = 'portfolio' | 'deck' | 'both'
export type JobStatus = 'queued' | 'running' | 'done' | 'failed'

export interface Check {
  name: string
  status: 'PASS' | 'WARN' | 'FAIL'
  detail: string
  step: string
  items: string[]
}

export interface ExceptionRow {
  severity: 'FAIL' | 'WARN' | 'INFO'
  kind: string
  symbol: string
  detail: string
}

export interface AiSourced {
  symbol: string
  field: string
  value: string
  new: boolean
  period: string
  sources: string
}

export interface ResultFile {
  kind: 'excel' | 'deck' | 'pdf' | 'ai_ratings' | 'report'
  label: string
  name: string
  size: number
}

export interface RunResult {
  status: string
  client: string
  period: string
  asof: string
  holdings_kind: string
  universal: string
  universal_read_at: string
  matched: number
  holdings: number
  checks: Check[]
  exceptions: ExceptionRow[]
  notes: string[]
  ai_sourced: AiSourced[]
  ai_research: { requests: number; searches: number; filled: number; skipped: { symbol: string; reason: string }[] }
  narrative_pending: string[]
  files: ResultFile[]
}

export interface Job {
  id: string
  step: Step
  step_label: string
  filename: string
  status: JobStatus
  stages: { key: string; label: string }[]
  stage: string | null
  created: string
  started: string | null
  finished: string | null
  client: string | null
  period: string | null
  result: RunResult | null
  error: string | null
  queue_position?: number
}

export interface HistoryRow {
  id: string
  step: Step
  step_label: string
  filename: string
  status: JobStatus
  client: string | null
  period: string | null
  created: string
  finished: string | null
  error: string | null
  run_status: string | null
  files: string[]
}

export interface SystemStatus {
  universal: string
  google: boolean
  ai: boolean
  web_search: boolean
  ai_research: string
  libreoffice: boolean
  deck_template: boolean
  queue: number
}

async function json<T>(r: Response): Promise<T> {
  if (!r.ok) {
    let msg = `Request failed (${r.status})`
    try {
      const body = await r.json()
      if (body?.detail) msg = typeof body.detail === 'string' ? body.detail : JSON.stringify(body.detail)
    } catch { /* not JSON */ }
    throw new Error(msg)
  }
  return r.json() as Promise<T>
}

export const api = {
  status: () => fetch('/api/status').then(r => json<SystemStatus>(r)),
  history: () => fetch('/api/runs').then(r => json<HistoryRow[]>(r)),
  job: (id: string) => fetch(`/api/runs/${id}`).then(r => json<Job>(r)),
  templateUrl: '/api/template',
  fileUrl: (id: string, name: string) => `/api/runs/${id}/files/${encodeURIComponent(name)}`,

  start(file: File, step: Step): Promise<Job> {
    const fd = new FormData()
    fd.append('file', file)
    fd.append('step', step)
    return fetch('/api/runs', { method: 'POST', body: fd }).then(r => json<Job>(r))
  },

  /** Live updates for one run; returns a function that stops listening. */
  watch(id: string, onJob: (j: Job) => void, onError?: () => void): () => void {
    const es = new EventSource(`/api/runs/${id}/events`)
    es.onmessage = e => {
      const j = JSON.parse(e.data) as Job
      onJob(j)
      if (j.status === 'done' || j.status === 'failed') es.close()
    }
    es.onerror = () => { es.close(); onError?.() }
    return () => es.close()
  },
}

export function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`
}

export function formatWhen(iso: string | null): string {
  if (!iso) return ''
  const d = new Date(iso)
  return d.toLocaleString('en-GB', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' })
}
