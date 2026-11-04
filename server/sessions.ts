import crypto from 'node:crypto'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

/**
 * 按文档隔离的会话存储。
 *
 * 每个 .docx 一个目录,目录内最多 MAX_SESSIONS 个会话;
 * 每个会话有自己独立的 Pi session 目录(pi/)与侧栏对话(ui.json),
 * 所以切回旧会话时 Pi 的上下文也真的跟着切回去。
 */

const ROOT = path.join(os.homedir(), '.wordai', 'sessions')

/** 历史上限:当前会话之外还能归档多少个。不自动汰旧,超了让用户自己删 */
export const MAX_HISTORY = 30

export interface SessionMeta {
  id: string
  title: string
  createdAt: string
  updatedAt: string
}

export interface SessionIndex {
  activeId: string | null
  sessions: SessionMeta[]
}

export interface SessionItem {
  [key: string]: unknown
}

// 'unsaved*' 一律不落盘:未保存的文档没有稳定身份,存了反而容易串稿
const isUnsaved = (docKey: string) => !docKey || docKey.startsWith('unsaved')

const docHash = (docKey: string) => crypto.createHash('sha1').update(docKey).digest('hex').slice(0, 16)

export const docDir = (docKey: string) => path.join(ROOT, docHash(docKey))

const indexFile = (docKey: string) => path.join(docDir(docKey), 'index.json')

/** 该会话的 Pi session 目录(SessionManager 的 sessionDir) */
export const piDirOf = (docKey: string, id: string) => path.join(docDir(docKey), `s-${id}`, 'pi')

const uiFile = (docKey: string, id: string) => path.join(docDir(docKey), `s-${id}`, 'ui.json')
const undoFile = (docKey: string, id: string) => path.join(docDir(docKey), `s-${id}`, 'undo.json')

function atomicWrite(file: string, data: string) {
  fs.mkdirSync(path.dirname(file), { recursive: true })
  const tmp = `${file}.${process.pid}.tmp`
  fs.writeFileSync(tmp, data, 'utf8')
  fs.renameSync(tmp, file)
}

function newId(): string {
  return `${Date.now().toString(36)}${crypto.randomBytes(3).toString('hex')}`
}

function makeMeta(): SessionMeta {
  const now = new Date().toISOString()
  return { id: newId(), title: '新对话', createdAt: now, updatedAt: now }
}

export function readIndex(docKey: string): SessionIndex {
  if (isUnsaved(docKey)) return { activeId: null, sessions: [] }
  try {
    const raw = fs.readFileSync(indexFile(docKey), 'utf8')
    const parsed = JSON.parse(raw) as SessionIndex
    const sessions = Array.isArray(parsed.sessions) ? parsed.sessions.filter((s) => s?.id) : []
    return { activeId: parsed.activeId ?? sessions[0]?.id ?? null, sessions }
  } catch {
    return { activeId: null, sessions: [] }
  }
}

function writeIndex(docKey: string, index: SessionIndex) {
  if (isUnsaved(docKey)) return
  atomicWrite(indexFile(docKey), JSON.stringify(index, null, 2))
  // 顺手记下原始路径,排查时能对上号
  atomicWrite(path.join(docDir(docKey), 'doc.json'), JSON.stringify({ docKey }, null, 2))
}

/**
 * 确保至少有一个会话,返回索引。
 * 首次打开某文档时自动建第一个会话。
 */
export function ensureIndex(docKey: string): SessionIndex {
  // 未保存文档不落盘,也就没有"会话"可言:直接给空索引,不能每次生成新 id
  if (isUnsaved(docKey)) return { activeId: null, sessions: [] }

  const idx = readIndex(docKey)
  if (idx.sessions.length) {
    const activeId = idx.sessions.some((s) => s.id === idx.activeId) ? idx.activeId : idx.sessions[0].id
    if (activeId !== idx.activeId) {
      idx.activeId = activeId
      writeIndex(docKey, idx)
    }
    return idx
  }
  const meta = makeMeta()
  const next: SessionIndex = { activeId: meta.id, sessions: [meta] }
  writeIndex(docKey, next)
  return next
}

export function createSession(docKey: string): { ok: boolean; index: SessionIndex; error?: string } {
  const idx = ensureIndex(docKey)
  // sessions 里含当前会话,所以历史条数是 length - 1
  if (idx.sessions.length - 1 >= MAX_HISTORY) {
    return { ok: false, index: idx, error: `历史会话最多 ${MAX_HISTORY} 个,请先在历史里删掉一些` }
  }
  const meta = makeMeta()
  const next: SessionIndex = { activeId: meta.id, sessions: [meta, ...idx.sessions] }
  writeIndex(docKey, next)
  return { ok: true, index: next }
}

export function switchSession(docKey: string, id: string): SessionIndex {
  const idx = ensureIndex(docKey)
  if (!idx.sessions.some((s) => s.id === id)) return idx
  const next = { ...idx, activeId: id }
  writeIndex(docKey, next)
  return next
}

export function deleteSession(docKey: string, id: string): SessionIndex {
  return deleteSessions(docKey, [id])
}

/**
 * 批量删除会话。一次 rm -rf 多个 s-<id>/(ui / undo / Pi 上下文 / 附件全在里面),
 * 删到当前会话就把 activeId 换成剩下的第一个;一个不剩就建个新的。
 */
export function deleteSessions(docKey: string, ids: string[]): SessionIndex {
  const idx = ensureIndex(docKey)
  const kill = new Set(ids.filter(Boolean))
  if (!kill.size) return idx

  for (const id of kill) {
    try {
      fs.rmSync(path.join(docDir(docKey), `s-${id}`), { recursive: true, force: true })
    } catch {
      /* 删不掉文件不影响索引一致 */
    }
  }

  const sessions = idx.sessions.filter((s) => !kill.has(s.id))
  if (!sessions.length) {
    const meta = makeMeta()
    const next: SessionIndex = { activeId: meta.id, sessions: [meta] }
    writeIndex(docKey, next)
    return next
  }

  const activeId = idx.activeId && !kill.has(idx.activeId) ? idx.activeId : sessions[0].id
  const next: SessionIndex = { activeId, sessions }
  writeIndex(docKey, next)
  return next
}

export function readUi(docKey: string, id: string): SessionItem[] {
  if (isUnsaved(docKey)) return []
  try {
    const parsed = JSON.parse(fs.readFileSync(uiFile(docKey, id), 'utf8'))
    return Array.isArray(parsed) ? (parsed as SessionItem[]) : []
  } catch {
    return []
  }
}

export function writeUi(docKey: string, id: string, items: SessionItem[]): SessionIndex {
  const idx = ensureIndex(docKey)
  if (!isUnsaved(docKey) && idx.sessions.some((s) => s.id === id)) {
    atomicWrite(uiFile(docKey, id), JSON.stringify(items))
    // 标题取首条用户提问,方便在标签上分辨
    const firstUser = items.find((it) => it?.kind === 'user' && typeof it.text === 'string') as
      | { text: string }
      | undefined
    const title = firstUser ? firstUser.text.replace(/\s+/g, ' ').trim().slice(0, 24) : '新对话'
    idx.sessions = idx.sessions.map((s) =>
      s.id === id ? { ...s, title: title || '新对话', updatedAt: new Date().toISOString() } : s,
    )
    writeIndex(docKey, idx)
  }
  return idx
}

export function readUndo(docKey: string, id: string): SessionItem[] {
  if (isUnsaved(docKey)) return []
  try {
    const parsed = JSON.parse(fs.readFileSync(undoFile(docKey, id), 'utf8'))
    return Array.isArray(parsed) ? (parsed as SessionItem[]) : []
  } catch {
    return []
  }
}

export function writeUndo(docKey: string, id: string, entries: SessionItem[]): void {
  if (isUnsaved(docKey)) return
  const idx = ensureIndex(docKey)
  if (!idx.sessions.some((s) => s.id === id)) return
  atomicWrite(undoFile(docKey, id), JSON.stringify(entries))
}
