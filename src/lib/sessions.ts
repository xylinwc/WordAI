export interface SessionMeta {
  id: string
  title: string
  createdAt: string
  updatedAt: string
}

export interface SessionState {
  activeId: string | null
  sessions: SessionMeta[]
  max: number
}

const json = { 'Content-Type': 'application/json' }

export async function fetchSessions(docKey: string): Promise<SessionState> {
  const r = await fetch(`/api/agent/sessions?docKey=${encodeURIComponent(docKey)}`)
  if (!r.ok) throw new Error('读取会话列表失败')
  return (await r.json()) as SessionState
}

export async function createSession(docKey: string): Promise<SessionState> {
  const r = await fetch('/api/agent/sessions/create', {
    method: 'POST',
    headers: json,
    body: JSON.stringify({ docKey }),
  })
  const body = await r.json()
  if (!r.ok) throw new Error(String(body?.error ?? '新建会话失败'))
  return body as SessionState
}

export async function switchSessionApi(docKey: string, id: string): Promise<SessionState> {
  const r = await fetch('/api/agent/sessions/switch', {
    method: 'POST',
    headers: json,
    body: JSON.stringify({ docKey, id }),
  })
  if (!r.ok) throw new Error('切换会话失败')
  return (await r.json()) as SessionState
}

/** 多选批量删:一次请求删完,避免中途失败只删一半 */
export async function deleteSessionsApi(docKey: string, ids: string[]): Promise<SessionState> {
  const r = await fetch('/api/agent/sessions/delete', {
    method: 'POST',
    headers: json,
    body: JSON.stringify({ docKey, ids }),
  })
  if (!r.ok) throw new Error('删除会话失败')
  return (await r.json()) as SessionState
}

export async function fetchSessionUi<T>(docKey: string, id: string): Promise<T[]> {
  const r = await fetch(
    `/api/agent/sessions/ui?docKey=${encodeURIComponent(docKey)}&id=${encodeURIComponent(id)}`,
  )
  if (!r.ok) return []
  const body = await r.json()
  return Array.isArray(body?.items) ? (body.items as T[]) : []
}

export async function saveSessionUi<T>(docKey: string, id: string, items: T[]): Promise<void> {
  await fetch('/api/agent/sessions/ui', {
    method: 'POST',
    headers: json,
    body: JSON.stringify({ docKey, id, items }),
  })
}

export async function fetchUndo<T>(docKey: string, id: string): Promise<T[]> {
  const r = await fetch(
    `/api/agent/sessions/undo?docKey=${encodeURIComponent(docKey)}&id=${encodeURIComponent(id)}`,
  )
  if (!r.ok) return []
  const body = await r.json()
  return Array.isArray(body?.entries) ? (body.entries as T[]) : []
}

export async function saveUndo<T>(docKey: string, id: string, entries: T[]): Promise<void> {
  await fetch('/api/agent/sessions/undo', {
    method: 'POST',
    headers: json,
    body: JSON.stringify({ docKey, id, entries }),
  })
}
