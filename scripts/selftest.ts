/**
 * 自测:验证本地服务各接口可用。不调用模型,不消耗额度。
 * 用法:npm run selftest
 *
 * 真链路(模型 + 工具 RPC)需要真实文档与凭证,见 scripts/e2e-agent.mjs。
 */
import { spawn } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { attachFiles, listRefs, readRefBlocks, removeRef, searchRefs } from '../server/refs.ts'
import { docDir } from '../server/sessions.ts'
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const PORT = 8799
const BASE = `http://localhost:${PORT}`

let failed = 0
const check = (label: string, pass: boolean, extra?: unknown) => {
  console.log(`${pass ? '✓' : '✗'} ${label}${pass ? '' : `  ← ${JSON.stringify(extra)}`}`)
  if (!pass) failed++
}

const waitFor = async (url: string, tries = 60) => {
  for (let i = 0; i < tries; i++) {
    try {
      const r = await fetch(url)
      if (r.ok) return true
    } catch {
      /* retry */
    }
    await new Promise((r) => setTimeout(r, 250))
  }
  return false
}

const api = spawn('npx', ['tsx', 'server/index.ts'], {
  cwd: ROOT,
  env: {
    ...process.env,
    API_PORT: String(PORT),
  },
  stdio: 'ignore',
})

try {
  check('服务启动', await waitFor(`${BASE}/api/health`))

  const health = await (await fetch(`${BASE}/api/health`)).json()
  check('health 返回模型名', typeof health.model === 'string' && health.model.length > 0, health)

  // 会话:每文档 1 个当前 + 最多 30 个历史
  const sess = await (await fetch(`${BASE}/api/agent/sessions?docKey=unsaved-x.docx`)).json()
  check('会话接口可用', Array.isArray(sess.sessions) && sess.max === 30, sess)

  const agent = await (await fetch(`${BASE}/api/agent/info`)).json()
  check('agent 空闲', agent.running === false, agent)

  const model = await (await fetch(`${BASE}/api/agent/model`)).json()
  check('读取 Pi 模型配置', typeof model.model === 'string' && model.model.length > 0, model)

  const badRun = await fetch(`${BASE}/api/agent/run`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({}),
  })
  check('空 prompt 被拒绝', badRun.status === 400, badRun.status)

  // 文档切换必须重建 session,否则会读到上一份文档
  const bind = async (docKey: string) =>
    (await (await fetch(`${BASE}/api/agent/bind`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ docKey }),
    })).json()) as { reset: boolean }

  // 用 unsaved* 前缀:既能测出「换文档要重置」,又不会在真实文档目录里建会话
  const b1 = await bind('unsaved-a.docx')
  const b2 = await bind('unsaved-a.docx')
  const b3 = await bind('unsaved-b.docx')
  const b4 = await bind('unsaved-b.docx')
  check('首次绑定不重置', b1.reset === false, b1)
  check('同文档不重置', b2.reset === false, b2)
  check('换文档触发重置', b3.reset === true, b3)
  check('重置后同文档稳定', b4.reset === false, b4)

  const badTool = await fetch(`${BASE}/api/agent/tool-result`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ callId: 'not-exist', ok: true }),
  })
  const badToolBody = await badTool.json()
  check('未知 callId 返回 false', badToolBody.ok === false, badToolBody)

  const idleAbort = await (
    await fetch(`${BASE}/api/agent/abort`, { method: 'POST' })
  ).json()
  check('空闲时中止返回 false', idleAbort.ok === false, idleAbort)

  // 附件:用临时 txt 走真实抽取链路(不弹文件对话框)
  const tmp = path.join(os.tmpdir(), `wordai-ref-${Date.now()}.txt`)
  fs.writeFileSync(
    tmp,
    [
      'Runx2 是成骨相关的转录因子,在糖尿病主动脉纤维化中表达上调。',
      '研究显示胶原沉积增加与主动脉僵硬相关,提示潜在的干预靶点。',
      'The transcription factor Runx2 promotes aortic fibrosis and stiffness.',
    ].join('\n\n'),
    'utf8',
  )

  const refDocKey = 'unsaved-refs-selftest'
  const attachRes = await (
    await fetch(`${BASE}/api/refs/attach`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ docKey: refDocKey, id: 'r1', path: tmp }),
    })
  ).json()
  check('附加文献成功', attachRes.added?.length === 1 && attachRes.refs?.length === 1, attachRes)

  const listRes = await (await fetch(`${BASE}/api/refs?docKey=${refDocKey}&id=r1`)).json()
  check('附件列表可读', listRes.refs?.length === 1, listRes)

  const refId = attachRes.added?.[0]?.id as string

  // HTTP 走的是另一个进程;检索/读取直接在本地进程验证,否则内存里的未保存文档附件对不上
  const localScope = 'unsaved-refs-local'
  const local = await attachFiles(localScope, 'r1', [tmp])
  const localId = local.added[0]?.id
  const hits = searchRefs(localScope, 'r1', '主动脉纤维化 Runx2', 3)
  check('附件检索命中', hits.hits.length > 0 && hits.hits[0].snippet.includes('Runx2'), hits.hits[0])

  const read = localId ? readRefBlocks(localScope, 'r1', localId, 0, 0) : { items: [] as { text: string }[] }
  check('附件按块读取', read.items.length === 1 && read.items[0].text.length > 0, read.items.length)

  const rmRes = await (
    await fetch(`${BASE}/api/refs/remove`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ docKey: refDocKey, id: 'r1', refId }),
    })
  ).json()
  check('移除附件生效', rmRes.refs?.length === 0, rmRes)

  if (localId) removeRef(localScope, 'r1', localId)
  check('误删后清理', listRefs(localScope, 'r1').length === 0, listRefs(localScope, 'r1'))

  // 历史会话:新建 -> 当前归档;多选批量删
  const histDoc = path.join(os.tmpdir(), `wordai-hist-${Date.now()}.docx`)
  const post = async (url: string, body: unknown) =>
    (await fetch(`${BASE}${url}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    })).json()

  const first = await (await fetch(`${BASE}/api/agent/sessions?docKey=${encodeURIComponent(histDoc)}`)).json()
  const c1 = await post('/api/agent/sessions/create', { docKey: histDoc })
  const c2 = await post('/api/agent/sessions/create', { docKey: histDoc })
  check('历史会话逐次累加', c1.sessions.length === 2 && c2.sessions.length === 3, [c1.sessions.length, c2.sessions.length])

  const victims = c2.sessions.filter((s: { id: string }) => s.id !== c2.activeId).map((s: { id: string }) => s.id)
  const deleted = await post('/api/agent/sessions/delete', { docKey: histDoc, ids: victims })
  check('多选批量删除', deleted.sessions.length === 1 && deleted.activeId === c2.activeId, deleted.sessions.length)

  const afterWatch = await (await fetch(`${BASE}/api/agent/sessions/delete`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ docKey: histDoc, ids: [deleted.activeId] }),
  })).json()
  check('删光后自动补一个会话', afterWatch.sessions.length === 1 && afterWatch.activeId !== deleted.activeId, afterWatch.activeId)

  fs.rmSync(docDir(histDoc), { recursive: true, force: true })
  void first

  fs.rmSync(tmp, { force: true })
} finally {
  api.kill('SIGTERM')
}

console.log(failed ? `\n${failed} 项失败` : '\n自测全部通过')
process.exit(failed ? 1 : 0)
