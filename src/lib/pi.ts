import { getOutline, readParagraphs, searchDocument } from '../office'
import type { Anchor } from './anchor'

export const UI_REV = 'r24.6'

export type { Anchor }

/** 依据链:这条草稿参考了哪份附件的哪一页 */
export interface Basis {
  name: string
  page?: number
  note?: string
}

/**
 * 前端 ↔ Pi 内核桥。
 * 模型与规划在服务端(Pi),Word 工具在本机执行:服务端下发 tool 事件,
 * 这里跑 Office.js 并把结果回传。
 */

export interface EditItem {
  id: string
  paragraphIndex: number | null
  find: string | null
  replace: string
  reason: string
  status: 'pending' | 'applied' | 'skipped' | 'failed' | 'commented' | 'inserted'
  error?: string
}

export interface Proposal {
  /** paragraphIndex → 该段原文,卡片用来做上下文聚焦 */
  context?: Record<number, string>
  summary: string
  selectionText: string
  /** 发送那一刻冻结的选区锚点,三个动作共用 */
  anchor?: Anchor
  basis?: Basis[]
  edits: EditItem[]
}

export interface CommentItem {
  id: string
  paragraphIndex: number | null
  find: string | null
  text: string
  severity: 'info' | 'warn'
  status: 'pending' | 'applied' | 'skipped' | 'failed'
  error?: string
}

export interface InsertBlock {
  id: string
  kind: 'heading' | 'paragraph' | 'list' | 'table'
  level?: number
  text?: string
  items?: string[]
  rows?: string[][]
  status: 'pending' | 'inserted' | 'skipped' | 'failed'
  error?: string
}

export interface InsertBundle {
  summary: string
  anchorIndex: number | null
  anchor?: Anchor
  blocks: InsertBlock[]
}

export interface CommentBundle {
  summary: string
  selectionText: string
  anchor?: Anchor
  basis?: Basis[]
  comments: CommentItem[]
}

/** 全文审查:一条待完善项 */
export interface AuditIssue {
  id: string
  category: string
  severity: 'priority' | 'normal'
  paragraphIndex: number | null
  find: string | null
  title: string
  detail: string
  suggestion: string
  status: 'pending' | 'commented' | 'skipped' | 'failed'
  error?: string
}

export interface AuditBundle {
  summary: string
  anchor?: Anchor
  basis?: Basis[]
  issues: AuditIssue[]
}

export interface PiHandlers {
  selectionText: string
  /** 发送那一刻冻结的选区锚点 */
  anchor?: Anchor
  onText: (delta: string) => void
  onTool: (name: string) => void
  onProposal: (proposal: Proposal) => void
  onComments: (bundle: CommentBundle) => void
  onInsert: (bundle: InsertBundle) => void
  onAudit: (bundle: AuditBundle) => void
  onToolError: (name: string, message: string) => void
  onDone: () => void
  /** 用户主动中止(而不是出错) */
  onAborted?: () => void
  onError: (message: string) => void
}

const uid = () => Math.random().toString(36).slice(2, 10)

const post = async (url: string, body: unknown) => {
  const r = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  })
  return r.json().catch(() => ({}))
}

/** 依据链归一化:只收有文件名的,同名同页去重 */
function normalizeBasis(raw: unknown): Basis[] {
  const list = Array.isArray(raw) ? raw : []
  const out: Basis[] = []
  const seen = new Set<string>()

  for (const item of list) {
    const b = item as Record<string, unknown>
    const name = typeof b.name === 'string' ? b.name.trim() : ''
    if (!name) continue
    const page = Number(b.page)
    const key = `${name}#${Number.isFinite(page) ? Math.floor(page) : ''}`
    if (seen.has(key)) continue
    seen.add(key)
    out.push({
      name,
      page: Number.isFinite(page) && page > 0 ? Math.floor(page) : undefined,
      note: typeof b.note === 'string' && b.note.trim() ? b.note.trim() : undefined,
    })
  }

  return out.slice(0, 6)
}

function normalizeEdits(
  raw: unknown,
  selectionText: string,
  summary: string,
  anchor?: Anchor,
  basisRaw?: unknown,
): Proposal {
  const list = Array.isArray(raw) ? raw : []
  const edits: EditItem[] = []
  const basis = normalizeBasis(basisRaw)

  for (const item of list) {
    const e = item as Record<string, unknown>
    const replace = typeof e.replace === 'string' ? e.replace : ''
    if (!replace.trim()) continue

    // 负数不是段落号:模型常拿 -1 表示「作用于当前选区」,这里统一归成 null
    const raw = Number(e.paragraphIndex)
    const hasPara = Number.isFinite(raw) && raw >= 0
    const find = typeof e.find === 'string' && e.find.trim() ? e.find : null

    edits.push({
      id: uid(),
      paragraphIndex: hasPara ? Number(e.paragraphIndex) : null,
      find: hasPara ? find : null,
      replace,
      reason: typeof e.reason === 'string' ? e.reason.trim() : '',
      status: 'pending',
    })
  }

  // 模型也可能把依据写在单条 edit 上,一并收上来展示
  for (const item of list) basis.push(...normalizeBasis((item as Record<string, unknown>).basis))
  const merged = normalizeBasis(basis)

  return { summary, selectionText, anchor, basis: merged, edits }
}

function normalizeComments(
  args: Record<string, unknown>,
  selectionText: string,
  anchor?: Anchor,
): CommentBundle {
  const list = Array.isArray(args.comments) ? args.comments : []
  const comments: CommentItem[] = []
  const summary = String(args.summary ?? '审阅建议')

  for (const item of list) {
    const c = item as Record<string, unknown>
    const text = typeof c.text === 'string' ? c.text.trim() : ''
    if (!text) continue

    const hasPara = Number.isFinite(Number(c.paragraphIndex))
    const find = typeof c.find === 'string' && c.find.trim() ? c.find : null

    comments.push({
      id: uid(),
      paragraphIndex: hasPara ? Number(c.paragraphIndex) : null,
      find: hasPara ? find : null,
      text,
      severity: c.severity === 'warn' ? 'warn' : 'info',
      status: 'pending',
    })
  }

  return { summary, selectionText, anchor, basis: normalizeBasis(args.basis), comments }
}

function normalizeAudit(args: Record<string, unknown>, anchor?: Anchor): AuditBundle {
  const list = Array.isArray(args.issues) ? args.issues : []
  const issues: AuditIssue[] = []

  for (const item of list) {
    const a = item as Record<string, unknown>
    const title = typeof a.title === 'string' ? a.title.trim() : ''
    const detail = typeof a.detail === 'string' ? a.detail.trim() : ''
    if (!title && !detail) continue

    const raw = Number(a.paragraphIndex)
    const hasPara = Number.isFinite(raw) && raw >= 0
    const find = hasPara && typeof a.find === 'string' && a.find.trim() ? a.find : null

    issues.push({
      id: uid(),
      category: typeof a.category === 'string' && a.category.trim() ? a.category.trim().slice(0, 12) : '结构',
      severity: a.severity === 'priority' ? 'priority' : 'normal',
      paragraphIndex: hasPara ? Math.floor(raw) : null,
      find,
      title: title.slice(0, 60),
      detail,
      suggestion: typeof a.suggestion === 'string' ? a.suggestion.trim() : '',
      status: 'pending',
    })
  }

  return {
    summary: String(args.summary ?? '全文审查'),
    anchor,
    basis: normalizeBasis(args.basis),
    issues,
  }
}

function normalizeInsert(raw: unknown, anchor?: Anchor): InsertBundle {
  const args = raw as Record<string, unknown>
  const list = Array.isArray(args.blocks) ? args.blocks : []
  const blocks: InsertBlock[] = []

  for (const item of list) {
    const b = item as Record<string, unknown>
    const kind = String(b.kind ?? '').trim()
    if (!['heading', 'paragraph', 'list', 'table'].includes(kind)) continue

    const text = typeof b.text === 'string' ? b.text : undefined
    const items = Array.isArray(b.items) ? b.items.map(String).filter((t) => t.trim()) : undefined
    const rows = Array.isArray(b.rows)
      ? (b.rows as unknown[][]).map((r) => (Array.isArray(r) ? r.map((c) => String(c ?? '')) : []))
      : undefined

    const hasContent =
      (kind === 'table' && rows?.length) ||
      (kind === 'list' && items?.length) ||
      ((kind === 'heading' || kind === 'paragraph') && text?.trim())
    if (!hasContent) continue

    blocks.push({
      id: uid(),
      kind: kind as InsertBlock['kind'],
      level: Number.isFinite(Number(b.level)) ? Number(b.level) : undefined,
      text,
      items,
      rows,
      status: 'pending',
    })
  }

  const rawAnchor = args.anchorParagraphIndex
  return {
    summary: String(args.summary ?? '新增内容草稿'),
    anchorIndex: Number.isFinite(Number(rawAnchor)) ? Number(rawAnchor) : null,
    anchor,
    blocks,
  }
}

async function execLocalTool(name: string, args: Record<string, unknown>, h: PiHandlers): Promise<unknown> {
  switch (name) {
    case 'doc_outline':
      return await getOutline()

    case 'doc_read':
      return await readParagraphs(Number(args.from), Number(args.to))

    case 'doc_search':
      return await searchDocument(String(args.query ?? ''), Number(args.limit) || 8)

    case 'propose_insert': {
      const bundle = normalizeInsert(args, h.anchor)
      if (!bundle.blocks.length) throw new Error('blocks 为空或没有可插入内容')
      h.onInsert(bundle)
      return { ok: true, count: bundle.blocks.length, note: '已展示给用户,等待勾选确认' }
    }

    case 'propose_comment': {
      const bundle = normalizeComments(args, h.selectionText, h.anchor)
      if (!bundle.comments.length) throw new Error('comments 为空或缺少 text')
      h.onComments(bundle)
      return { ok: true, count: bundle.comments.length, note: '已展示给用户,等待逐条确认' }
    }

    case 'propose_edit': {
      const proposal = normalizeEdits(
        args.edits,
        h.selectionText,
        String(args.summary ?? '文档改动草稿'),
        h.anchor,
        args.basis,
      )
      if (!proposal.edits.length) throw new Error('edits 为空或缺少 replace')
      h.onProposal(proposal)
      return { ok: true, count: proposal.edits.length, note: '已展示给用户,等待勾选确认' }
    }

    case 'propose_audit': {
      const bundle = normalizeAudit(args, h.anchor)
      if (!bundle.issues.length) throw new Error('issues 为空')
      h.onAudit(bundle)
      return { ok: true, count: bundle.issues.length, note: '已展示审查清单,等待逐项处理' }
    }

    default:
      throw new Error(`未知工具 ${name}`)
  }
}

export function runPiAgent(
  prompt: string,
  docKey: string,
  h: PiHandlers,
  options: { signal?: AbortSignal } = {},
): Promise<void> {
  return new Promise((resolve) => {
    let finished = false
    let controller: AbortController | null = null

    const finish = () => {
      if (finished) return
      finished = true
      controller?.abort()
      options.signal?.removeEventListener('abort', onAbort)
      resolve()
    }

    /** 用户中止:必须同时通知服务端,只断流的话模型会继续跑完并烧额度 */
    const onAbort = () => {
      void post('/api/agent/abort', {}).catch(() => {})
      finish()
    }

    if (options.signal) {
      if (options.signal.aborted) {
        finish()
        return
      }
      options.signal.addEventListener('abort', onAbort, { once: true })
    }

    const handle = async (msg: Record<string, unknown>) => {
      if (finished) return
      const type = msg.type as string

      if (type === 'text') {
        h.onText(String(msg.delta ?? ''))
        return
      }

      if (type === 'tool') {
        const name = String(msg.name ?? '')
        h.onTool(name)
        const callId = String(msg.callId ?? '')
        const args = (msg.args ?? {}) as Record<string, unknown>
        try {
          const result = await execLocalTool(name, args, h)
          await post('/api/agent/tool-result', { callId, ok: true, result })
        } catch (e) {
          const message = (e as Error).message
          try {
            h.onToolError(name, message)
          } catch {
            /* ignore */
          }
          await post('/api/agent/tool-result', { callId, ok: false, error: message })
        }
        return
      }

      if (type === 'done') {
        h.onDone()
        finish()
        return
      }

      if (type === 'aborted') {
        h.onAborted?.()
        finish()
        return
      }

      if (type === 'error') {
        h.onError(String(msg.message ?? '未知错误'))
        finish()
      }
    }

    // 用 fetch + ReadableStream 读 SSE,而不是 EventSource:
    // WKWebView 的 EventSource 会把事件攒到连接关闭再派发,看起来就不流式了。
    const readStream = async () => {
      controller = new AbortController()
      const res = await fetch(`/api/agent/stream?t=${Date.now()}`, {
        signal: controller.signal,
        headers: { Accept: 'text/event-stream' },
      })
      if (!res.body) throw new Error('无法读取事件流')

      const reader = res.body.getReader()
      const dec = new TextDecoder()
      let buf = ''

      while (!finished) {
        const { value, done } = await reader.read()
        if (done) break
        buf += dec.decode(value, { stream: true })

        // 事件之间以空行分隔;最后一段可能不完整,留到下一轮
        const parts = buf.split('\n\n')
        buf = parts.pop() ?? ''

        for (const part of parts) {
          const line = part.split('\n').find((l) => l.startsWith('data: '))
          if (!line) continue
          let msg: Record<string, unknown>
          try {
            msg = JSON.parse(line.slice(6))
          } catch {
            continue
          }
          await handle(msg)
          if (finished) break
        }
      }
    }

    // 先启动任务,再连流:服务端会回放本次 run 的缓冲事件,不会丢开头
    post('/api/agent/run', { prompt, docKey, rev: UI_REV })
      .then(async (r) => {
        if (r?.error) {
          h.onError(String(r.error))
          finish()
          return
        }
        await readStream()
        finish()
      })
      .catch((e) => {
        if ((e as Error).name === 'AbortError') return
        h.onError((e as Error).message)
        finish()
      })

    // 兜底:避免界面永远停在「工作中」
    setTimeout(finish, 10 * 60 * 1000)
  })
}

export async function resetPiSession(): Promise<void> {
  await post('/api/agent/reset', {})
}
