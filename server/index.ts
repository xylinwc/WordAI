import 'dotenv/config'
import express from 'express'
import cors from 'cors'
import crypto from 'node:crypto'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import {
  abortRun,
  bindDoc,
  clearBuffer,
  currentModelInfo,
  isRunning,
  listAvailableModels,
  resetSession,
  resolveToolCall,
  runPrompt,
  selectModel,
  sessionInfo,
  setSessionId,
  subscribe,
} from './pi-runner.ts'
import {
  createSession,
  deleteSession,
  deleteSessions,
  ensureIndex,
  MAX_HISTORY,
  readUi,
  readUndo,
  switchSession,
  writeUi,
  writeUndo,
  type SessionItem,
} from './sessions.ts'
import { attachFiles, listRefs, pickFiles, removeRef } from './refs.ts'

export const BUILD_ID = 'b-' + String(Math.floor(Date.now() / 1000)).slice(-6)

const app = express()
app.use(cors())
app.use(express.json({ limit: '16mb' }))

const PORT = Number(process.env.API_PORT ?? 8787)

/** 模型与凭证都由 Pi 自己管理,这里只做只读展示 */
function piModelInfo(): { model: string; provider: string } {
  try {
    const raw = fs.readFileSync(path.join(os.homedir(), '.pi', 'agent', 'settings.json'), 'utf8')
    const s = JSON.parse(raw)
    return { model: s.defaultModel ?? 'pi', provider: s.defaultProvider ?? '' }
  } catch {
    return { model: 'pi', provider: '' }
  }
}

app.get('/api/health', (_req, res) => {
  res.json({ ok: true, ...piModelInfo() })
})

/** 运行状态:前端用它判断要不要禁用发送 */
app.get('/api/agent/info', (_req, res) => {
  res.json(sessionInfo())
})

/** 当前模型:WordAI 选择优先,其次活动 session,最后 Pi 默认 */
app.get('/api/agent/model', async (_req, res) => {
  const cur = await currentModelInfo().catch(() => null)
  if (cur) {
    res.json({ model: cur.id, provider: cur.provider, name: cur.name })
    return
  }
  res.json(piModelInfo())
})

/** 可切换的模型列表(只列有可用凭证的) */
app.get('/api/agent/models', async (_req, res) => {
  try {
    res.json(await listAvailableModels())
  } catch (e) {
    res.json({ current: null, models: [], error: (e as Error).message })
  }
})

/** 切换 WordAI 全局模型:运行中拒绝,避免中途换模型 */
app.post('/api/agent/model', async (req, res) => {
  const provider = String(req.body?.provider ?? '')
  const id = String(req.body?.id ?? '')
  try {
    const current = await selectModel(provider, id)
    res.json({ ok: true, current })
  } catch (e) {
    const message = (e as Error).message
    const status = isRunning() && message.includes('正在生成') ? 409 : 400
    res.status(status).json({ error: message })
  }
})

app.post('/api/agent/run', async (req, res) => {
  console.log(`[run] rev=${(req.body ?? {}).rev ?? '(旧前端,未上报)'} docKey=${(req.body ?? {}).docKey ?? '-'} len=${String((req.body ?? {}).prompt ?? '').length}`)
  const prompt = String(req.body?.prompt ?? '').trim()
  if (!prompt) {
    res.status(400).json({ error: 'prompt 必填' })
    return
  }
  if (isRunning()) {
    res.status(409).json({ error: '上一个任务还在运行' })
    return
  }

  // 文档变了就重建 session,避免读到的还是上一份文档
  const docKey = String(req.body?.docKey ?? '')
  let switched = false
  try {
    switched = await bindDoc(docKey)
  } catch {
    /* 绑定失败不阻塞本次请求 */
  }

  clearBuffer()
  res.json({ ok: true, reset: switched })
  // 先响应再后台跑,前端通过 SSE 收事件
  void runPrompt(prompt)
})

app.get('/api/agent/stream', (req, res) => {
  res.writeHead(200, {
    'Content-Type': 'text/event-stream; charset=utf-8',
    'Cache-Control': 'no-cache, no-transform',
    Connection: 'keep-alive',
    'X-Accel-Buffering': 'no',
  })
  res.write(': connected\n\n')

  const unsubscribe = subscribe((e) => {
    res.write(`data: ${JSON.stringify(e)}\n\n`)
  })

  const ping = setInterval(() => res.write(': ping\n\n'), 15_000)
  req.on('close', () => {
    clearInterval(ping)
    unsubscribe()
  })
})

/* ---------------- 会话管理:1 个当前 + 历史池(最多 30) ---------------- */

app.get('/api/agent/sessions', (req, res) => {
  const docKey = String(req.query.docKey ?? '')
  const idx = ensureIndex(docKey)
  res.json({ activeId: idx.activeId, sessions: idx.sessions, max: MAX_HISTORY })
})

app.post('/api/agent/sessions/create', async (req, res) => {
  const docKey = String(req.body?.docKey ?? '')
  const r = createSession(docKey)
  if (!r.ok) {
    res.status(409).json({ error: r.error, ...r.index, max: MAX_HISTORY })
    return
  }
  await setSessionId(r.index.activeId)
  res.json({ ...r.index, max: MAX_HISTORY })
})

app.post('/api/agent/sessions/switch', async (req, res) => {
  const docKey = String(req.body?.docKey ?? '')
  const id = String(req.body?.id ?? '')
  const idx = switchSession(docKey, id)
  await setSessionId(idx.activeId)
  res.json({ ...idx, max: MAX_HISTORY })
})

/** 单删与多选删共用一个路由:传 ids 走批量 */
app.post('/api/agent/sessions/delete', async (req, res) => {
  const docKey = String(req.body?.docKey ?? '')
  const ids = Array.isArray(req.body?.ids) ? (req.body.ids as unknown[]).map(String) : []
  const single = String(req.body?.id ?? '')
  const idx = ids.length ? deleteSessions(docKey, ids) : deleteSession(docKey, single)
  await setSessionId(idx.activeId)
  res.json({ ...idx, max: MAX_HISTORY })
})

app.get('/api/agent/sessions/ui', (req, res) => {
  const docKey = String(req.query.docKey ?? '')
  const id = String(req.query.id ?? '')
  res.json({ items: readUi(docKey, id) })
})

app.post('/api/agent/sessions/ui', (req, res) => {
  const docKey = String(req.body?.docKey ?? '')
  const id = String(req.body?.id ?? '')
  const items = Array.isArray(req.body?.items) ? (req.body.items as SessionItem[]) : []
  const idx = writeUi(docKey, id, items)
  res.json({ ok: true, sessions: idx.sessions, activeId: idx.activeId })
})

app.get('/api/agent/sessions/undo', (req, res) => {
  const docKey = String(req.query.docKey ?? '')
  const id = String(req.query.id ?? '')
  res.json({ entries: readUndo(docKey, id) })
})

app.post('/api/agent/sessions/undo', (req, res) => {
  const docKey = String(req.body?.docKey ?? '')
  const id = String(req.body?.id ?? '')
  const entries = Array.isArray(req.body?.entries) ? (req.body.entries as SessionItem[]) : []
  writeUndo(docKey, id, entries)
  res.json({ ok: true })
})

/** 附件作用域 = 文档 + 会话;前端带上自己的 activeId,保证和 Pi 内核看到的是同一份 */
const refScopeOf = (body: Record<string, unknown>) => ({
  docKey: String(body.docKey ?? ''),
  id: String(body.id ?? '') || sessionInfo().sessionId || 'default',
})

app.get('/api/refs', (req, res) => {
  const docKey = String(req.query.docKey ?? '')
  const id = String(req.query.id ?? '') || sessionInfo().sessionId || 'default'
  res.json({ refs: listRefs(docKey, id) })
})

/**
 * 附加文献。不传 path 就弹 macOS 原生选文件窗口 —— 侧栏是 WKWebView,
 * 用 <input type=file> 历史上不稳,而原生对话框能直接拿绝对路径,省掉上传与临时文件。
 */
app.post('/api/refs/attach', async (req, res) => {
  const { docKey, id } = refScopeOf(req.body ?? {})
  try {
    const given = Array.isArray(req.body?.paths) ? (req.body.paths as unknown[]).map(String) : []
    const single = typeof req.body?.path === 'string' && req.body.path ? [String(req.body.path)] : []
    const files = given.length || single.length ? [...single, ...given] : await pickFiles()
    if (!files.length) {
      res.json({ cancelled: true, added: [], refs: listRefs(docKey, id) })
      return
    }
    const result = await attachFiles(docKey, id, files)
    console.log(
      `[refs] attach docKey=${docKey} id=${id} added=${result.added.map((r) => `${r.name}(${r.chunks}块)`).join(',') || '-'} errors=${result.errors.length}`,
    )
    res.json({ ...result, refs: listRefs(docKey, id) })
  } catch (e) {
    res.status(500).json({ error: (e as Error).message, refs: listRefs(docKey, id) })
  }
})

app.post('/api/refs/remove', (req, res) => {
  const { docKey, id } = refScopeOf(req.body ?? {})
  const refId = String(req.body?.refId ?? '')
  res.json({ ok: true, refs: removeRef(docKey, id, refId) })
})

/**
 * 拖拽上传:WKWebView 里拿不到被拖文件的绝对路径,只能传字节。
 * 先落临时文件再走同一套抽取流程,抽取完就删——服务端不留原文件。
 */
app.post('/api/refs/attach-upload', express.raw({ type: '*/*', limit: '100mb' }), async (req, res) => {
  const { docKey, id } = refScopeOf(req.query as Record<string, unknown>)
  const raw = String(req.query.name ?? 'attachment')
  const name = path.basename(raw).replace(/[/\\:*?"<>|]/g, '_').slice(0, 120)
  const ext = path.extname(name).slice(1).toLowerCase()
  const body = req.body as unknown

  if (!Buffer.isBuffer(body) || !body.length) {
    res.status(400).json({ error: '上传内容为空', refs: listRefs(docKey, id) })
    return
  }

  const tmp = path.join(os.tmpdir(), `wordai-up-${crypto.randomBytes(4).toString('hex')}.${ext || 'bin'}`)
  try {
    fs.writeFileSync(tmp, body)
    const result = await attachFiles(docKey, id, [{ path: tmp, name }])
    console.log(`[refs] upload docKey=${docKey} name=${name} bytes=${body.length} added=${result.added.length}`)
    res.json({ ...result, refs: listRefs(docKey, id) })
  } catch (e) {
    res.status(500).json({ error: (e as Error).message, refs: listRefs(docKey, id) })
  } finally {
    fs.rmSync(tmp, { force: true })
  }
})

/** 仅供自测与调试:单独探测文档切换是否触发 session 重建 */
app.post('/api/agent/bind', async (req, res) => {
  const key = String(req.body?.docKey ?? '')
  res.json({ ok: true, reset: await bindDoc(key) })
})

app.post('/api/debug/log', (req, res) => {
  const line = (req.body ?? {}).line ?? ''
  console.log(`[js] ${line}`)
  res.json({ ok: true })
})

app.post('/api/agent/tool-result', (req, res) => {
  // 诊断:确认前端是否认领工具(旧版前端会返回「未知工具」)
  const { callId, ok, result, error } = req.body ?? {}
  if (!callId) {
    res.status(400).json({ error: 'callId 必填' })
    return
  }
  const summary = ok ? 'ok' : `FAILED: ${error ?? '未知'}`
  console.log(`[tool-result] ${String(callId).slice(0, 18)} → ${summary}`)
  res.json({ ok: resolveToolCall(String(callId), Boolean(ok), result, error) })
})

app.post('/api/agent/reset', async (_req, res) => {
  await resetSession()
  res.json({ ok: true })
})

/** 中止当前生成:前端点中止键或按 Esc 时调用 */
app.post('/api/agent/abort', async (_req, res) => {
  const ok = await abortRun()
  console.log(`[abort] ${ok ? '已中止' : '当前没有运行中的任务'}`)
  res.json({ ok })
})

// 生产模式下同一个服务托管前端静态资源
if (process.env.NODE_ENV === 'production') {
  const dist = path.resolve(process.cwd(), 'dist')
  if (fs.existsSync(dist)) {
    app.use(express.static(dist))
    app.get(/^\/(?!api).*/, (_req, res) => res.sendFile(path.join(dist, 'index.html')))
  }
}

app.listen(PORT, () => {
  const m = piModelInfo()
  console.log(`[WordAI] build=${BUILD_ID}`)
  console.log(`[WordAI] API 代理已启动 http://localhost:${PORT}  模型=${m.model}  provider=${m.provider || '-'}`)
})
