import crypto from 'node:crypto'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import {
  createAgentSession,
  DefaultResourceLoader,
  defineTool,
  getAgentDir,
  ModelRuntime,
  SessionManager,
  type AgentSession,
} from '@earendil-works/pi-coding-agent'
import { Type } from 'typebox'
import { ensureIndex, piDirOf } from './sessions.ts'
import { listRefs, pdfToolDepsOk, readRefBlocks, refsForPrompt, searchRefs } from './refs.ts'

/**
 * Pi 内核桥接。
 *
 * Word 侧栏(Office.js)不能跑在 Node 里,所以这里把 Pi 当成 agent 内核跑在服务端,
 * 把「读文档 / 提改动」注册成 Pi 的自定义工具;工具真正执行时通过 SSE 下发到前端,
 * 等前端用 POST 回传结果。模型、规划、多轮、技能全部由 Pi 负责。
 */

export type PiEvent =
  | { type: 'text'; delta: string }
  | { type: 'tool'; callId: string; name: string; args: unknown }
  | { type: 'thinking' }
  | { type: 'done' }
  | { type: 'aborted' }
  | { type: 'error'; message: string }

interface Pending {
  resolve: (value: unknown) => void
  reject: (err: Error) => void
}

const WORK_DIR = path.join(os.homedir(), '.wordai', 'workspace')

const listeners = new Set<(e: PiEvent) => void>()
const buffer: PiEvent[] = []
const pending = new Map<string, Pending>()

let session: AgentSession | null = null
let starting: Promise<AgentSession> | null = null
let running = false
/** 用户点了中止:prompt 返回后要看这个标记决定发 done 还是 aborted */
let aborting = false
/** 当前 session 绑定的文档标识;换文档必须重建,否则会串上下文 */
let sessionDocKey: string | null = null
/** 当前会话 id;同一文档下切换会话也要重建,否则两个对话会共用上下文 */
let sessionId: string | null = null

export const isRunning = () => running

function emit(e: PiEvent) {
  // 全部事件进 buffer:前端是先 POST run 再连 SSE,靠回放补齐开头
  buffer.push(e)
  if (buffer.length > 3000) buffer.splice(0, buffer.length - 3000)
  for (const fn of listeners) fn(e)
}

export function subscribe(fn: (e: PiEvent) => void): () => void {
  listeners.add(fn)
  for (const e of buffer) fn(e)
  return () => listeners.delete(fn)
}

export function clearBuffer() {
  buffer.length = 0
}

/** 前端把工具执行结果回传到这里 */
export function resolveToolCall(callId: string, ok: boolean, result: unknown, error?: string) {
  const p = pending.get(callId)
  if (!p) return false
  pending.delete(callId)
  if (ok) p.resolve(result)
  else p.reject(new Error(error || '前端执行失败'))
  return true
}

const TOOL_TIMEOUT = 90_000

function callClient(name: string, args: unknown, signal?: AbortSignal): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const callId = crypto.randomUUID()
    const timer = setTimeout(() => {
      pending.delete(callId)
      reject(new Error(`前端 ${TOOL_TIMEOUT / 1000}s 内未响应工具 ${name}`))
    }, TOOL_TIMEOUT)

    pending.set(callId, {
      resolve: (v) => {
        clearTimeout(timer)
        resolve(v)
      },
      reject: (e) => {
        clearTimeout(timer)
        reject(e)
      },
    })

    emit({ type: 'tool', callId, name, args })

    signal?.addEventListener('abort', () => {
      if (!pending.has(callId)) return
      clearTimeout(timer)
      pending.delete(callId)
      reject(new Error('已取消'))
    })
  })
}

/** 文献工具的作用域:跟随当前文档 + 当前会话,和 Word 工具各论各的 */
function refScope(): { docKey: string; id: string } {
  return { docKey: sessionDocKey ?? 'unknown', id: sessionId ?? 'default' }
}

/* ------------------------------------------------------------------ */
/* 模型选择:WordAI 全局,存 ~/.wordai/model.json,不动 Pi 默认           */
/* ------------------------------------------------------------------ */

export interface ModelInfo {
  provider: string
  id: string
  name: string
  reasoning?: boolean
  contextWindow?: number
}

type ChatModel = ReturnType<ModelRuntime['getModels']>[number]

interface StoredModel {
  provider: string
  id: string
  updatedAt: string
}

const MODEL_STORE = path.join(os.homedir(), '.wordai', 'model.json')
const MODEL_CACHE_TTL = 30_000

let modelRuntime: ModelRuntime | null = null
let modelRuntimeStarting: Promise<ModelRuntime> | null = null
let modelListCache: { at: number; models: ChatModel[] } | null = null
let storedModel: StoredModel | null | undefined

function ensureModelRuntime(): Promise<ModelRuntime> {
  if (modelRuntime) return Promise.resolve(modelRuntime)
  if (!modelRuntimeStarting) {
    modelRuntimeStarting = ModelRuntime.create()
      .then((rt) => {
        modelRuntime = rt
        return rt
      })
      .finally(() => {
        modelRuntimeStarting = null
      })
  }
  return modelRuntimeStarting
}

async function availableModels(force = false): Promise<ChatModel[]> {
  const rt = await ensureModelRuntime()
  if (!force && modelListCache && Date.now() - modelListCache.at < MODEL_CACHE_TTL) {
    return modelListCache.models
  }
  const models = [...(await rt.getAvailable())]
  modelListCache = { at: Date.now(), models }
  return models
}

function readStoredModel(): StoredModel | null {
  try {
    const raw = JSON.parse(fs.readFileSync(MODEL_STORE, 'utf8')) as Record<string, unknown>
    const provider = typeof raw.provider === 'string' ? raw.provider.trim() : ''
    const id = typeof raw.id === 'string' ? raw.id.trim() : ''
    if (!provider || !id) return null
    return { provider, id, updatedAt: String(raw.updatedAt ?? '') }
  } catch {
    return null
  }
}

function writeStoredModel(provider: string, id: string): void {
  try {
    fs.mkdirSync(path.dirname(MODEL_STORE), { recursive: true })
    const tmp = `${MODEL_STORE}.tmp`
    fs.writeFileSync(tmp, JSON.stringify({ provider, id, updatedAt: new Date().toISOString() }, null, 2))
    fs.renameSync(tmp, MODEL_STORE)
  } catch (e) {
    // 写不了就算内存里记着,不阻塞本次切换
    console.log(`[model] 保存失败: ${(e as Error).message}`)
  }
}

/**
 * 读取并校验 WordAI 保存的模型。
 * 列表非空但不包含它 = 凭证失效或模型下架 → 当作没选,回落 Pi 默认。
 * 列表为空(无凭证)或取列表失败 → 保留选择,免得临时故障把用户设置抹掉。
 */
async function loadStoredModel(): Promise<StoredModel | null> {
  if (storedModel !== undefined) return storedModel
  const raw = readStoredModel()
  if (!raw) {
    storedModel = null
    return null
  }

  const list = await availableModels().catch(() => null)
  if (list && list.length && !list.some((m) => m.provider === raw.provider && m.id === raw.id)) {
    console.log(`[model] 保存的模型不可用,回落 Pi 默认: ${raw.provider}/${raw.id}`)
    storedModel = null
    return null
  }

  storedModel = raw
  return raw
}

function toModelInfo(m: ChatModel): ModelInfo {
  return {
    provider: m.provider,
    id: m.id,
    name: m.name || m.id,
    reasoning: m.reasoning,
    contextWindow: m.contextWindow,
  }
}

/** 读 Pi 的 defaultModel,只在没有 WordAI 选择时做展示与兜底 */
function settingsDefaultModel(): ModelInfo | null {
  try {
    const raw = JSON.parse(fs.readFileSync(path.join(getAgentDir(), 'settings.json'), 'utf8')) as Record<string, unknown>
    const provider = typeof raw.defaultProvider === 'string' ? raw.defaultProvider : ''
    const id = typeof raw.defaultModel === 'string' ? raw.defaultModel : ''
    if (!provider || !id) return null
    return { provider, id, name: id }
  } catch {
    return null
  }
}

/** 当前模型:活动 session 优先,其次 WordAI 保存的选择,最后 Pi 默认 */
export async function currentModelInfo(): Promise<ModelInfo | null> {
  if (session?.model) return toModelInfo(session.model)

  const stored = await loadStoredModel().catch(() => null)
  if (stored) {
    const rt = await ensureModelRuntime().catch(() => null)
    const m = rt?.getModel(stored.provider, stored.id)
    return m ? toModelInfo(m) : { provider: stored.provider, id: stored.id, name: stored.id }
  }

  const fallback = settingsDefaultModel()
  if (fallback) {
    const rt = await ensureModelRuntime().catch(() => null)
    const m = rt?.getModel(fallback.provider, fallback.id)
    if (m) return toModelInfo(m)
  }
  return fallback
}

/** 列可用模型 + 当前模型;列失败也返回 200,让界面能显示原因 */
export async function listAvailableModels(): Promise<{
  current: ModelInfo | null
  models: ModelInfo[]
  error?: string
}> {
  let models: ModelInfo[] = []
  let error: string | undefined
  try {
    models = (await availableModels()).map(toModelInfo)
  } catch (e) {
    error = (e as Error).message
  }
  const current = await currentModelInfo().catch(() => null)
  return { current, models, error }
}

/** 切换模型:校验可用 → 应用到当前 session(如果有)→ 落盘 WordAI 全局 */
export async function selectModel(provider: string, modelId: string): Promise<ModelInfo> {
  if (running) throw new Error('正在生成,等这一轮结束再切换模型')
  const providerId = provider.trim()
  const id = modelId.trim()
  if (!providerId || !id) throw new Error('provider / id 不能为空')

  let model = (await availableModels()).find((m) => m.provider === providerId && m.id === id)
  if (!model) {
    // 缓存可能过期:强制刷新一次再判
    model = (await availableModels(true)).find((m) => m.provider === providerId && m.id === id)
  }
  if (!model) throw new Error('该模型当前不可用(缺少凭证或不在目录中)')

  if (session) await session.setModel(model)
  writeStoredModel(providerId, id)
  storedModel = { provider: providerId, id, updatedAt: new Date().toISOString() }
  console.log(`[model] 已切换为 ${providerId}/${id}`)
  return toModelInfo(model)
}

const textResult = (value: unknown) => ({
  content: [{ type: 'text' as const, text: JSON.stringify(value) }],
  details: undefined,
})

const BASIS_SCHEMA = Type.Optional(
  Type.Array(
    Type.Object({
      name: Type.String({ description: '附件文件名,必须与请求开头的「已附加参考文献」清单里的名字一致' }),
      page: Type.Optional(Type.Integer({ description: '页码(1 起)' })),
      note: Type.Optional(Type.String({ description: '这一页用来支撑什么,一句话' })),
    }),
    { description: '改写依据:这条草稿参考了哪份附件的哪一页。凡是依据附件改写/批注就必须填' },
  ),
)

const WORD_PROMPT = `你是 Word 文档写作助手,运行在 Word 右侧侧栏里,通过工具直接读写用户当前打开的那份文档。

你的能力(全部通过工具完成,不要臆测文档内容):
- doc_outline:拿标题结构与总段落数、总字数(不含正文)
- doc_read:按段落号读正文,单次上限约 6000 字
- doc_search:全文检索,返回命中段落号与前后段
- propose_edit:提交改动草稿,由用户在界面上逐处勾选确认后才写入
- propose_comment:提交批注草稿,以 Word 原生批注标在原文上,不改正文
- propose_insert:提交待插入的新内容(标题 / 段落 / 列表 / 表格),用于新增章节或表格
- propose_audit:提交全文审查清单(结构/声明/引文/逻辑),每条带定位与建议,用户可逐项转成批注

工作规则:
1. 不确定文档内容时先调 doc_outline / doc_search,不要凭空编写。
2. 要改某段落前,必须先用 doc_read 看到该段原文,find 必须是该段内的连续片段并逐字复制(含标点)。
3. 用户只是要改他自己选中的文字时,**省略** paragraphIndex 与 find(不要传 -1、0 或 null 字符串),
   replace 写完整新文本。paragraphIndex 必须是 >= 0 的整数,负数会被判为无效。
4. 改动一律通过 propose_edit 提交,不要在回复里贴改写后的全文;提交前用一两句话说明你发现了什么、准备怎么改。
5. 提交后用一句话提醒用户勾选并点「应用」才会真正写入。
6. 保持原文语言、语气、专业术语与引用一致;不杜撰事实、数据或文献;改动尽量小,不顺手改无关内容。
7. 用户要求「审阅 / 检查 / 提意见 / 挑毛病 / 这里是投稿前把关」时,用 propose_comment 而不是 propose_edit:
   批注要指出具体问题并给出建议改法,一句话说清,尽量不超过 80 字;不要泛泛而谈「建议润色」。
8. 用户要求「写一段 / 加一节 / 生成表格 / 补大纲 / 起草」这类**新增**内容时,用 propose_insert;
   propose_edit 只能替换已有文字,无法新增段落。锚点省略就接在当前光标段落之后。
9. 工具里的 paragraphIndex 从 0 开始;用户口头说的「第 N 段」通常指第 N 个自然段(1 起),换算时减 1。拿不准就先用 doc_read 确认。
10. 工具报错或返回 ok:false 时,必须如实说明失败原因,禁止声称「已提交成功」。用户侧看不到你的工具调用,只能看你的话。
11. 用户只是提问、与文档无关时,直接简短回答,不要调用工具。
12. 对话开头若出现「已附加参考文献」清单,你可以用 ref_list / ref_search / ref_read 读它们的原文。
    需要文献内容时必须先 ref_search 定位真实段落,再 ref_read 取原文核对,原文优先于你的记忆。
13. 用户要求「参照 / 参考这篇文献的某段来改写」时:先 ref_search 找到那段 → ref_read 看到原文 →
    再 propose_edit。改写要贴合文献的措辞、术语与论证口径,不要引入文献里没有的数据或结论。
14. 附件里检索不到用户说的内容,就直说「附件中没有检索到相关内容」,禁止假装引用或编造文献。
15. 需要引用出处时,按文献本身的引用格式写(如 (Raaz et al., 2019) 或 [1])。
16. **依据链是硬要求**:凡是依据附件改写或批注,必须在 propose_edit / propose_comment 的 basis 里填
    [{name: 附件文件名, page: 页码, note: 用来支撑什么}]。只写在 reason 里不算,界面靠 basis 展示「引用自」并核对页码。
    没有依据附件时省略 basis,禁止编造附件名或页码。
17. 用户要求「全文检查逻辑 / 结构审查 / 找逻辑漏洞 / 投稿前通读」时,用 propose_audit 一次性给清单,
    不要用 propose_comment 逐条堆。每条必须给 paragraphIndex(看过原文后填)与 find(可定位的原文片段),
    category 用「声明空缺 / 章节待填 / 引文格式 / 逻辑断层 / 术语不一」这类短标签;priority 只给真正阻断投稿的问题。
18. 已附加的文献是 PDF 时,你也可以直接用 pdf_info / pdf_extract_text / pdf_extract_tables 读它的某几页或表格;
    路径只用上面清单给出的绝对路径。`

function buildTools(runIdRef: { current: string }) {
  void runIdRef
  return [
    defineTool({
      name: 'doc_outline',
      label: '文档结构',
      description: '获取当前 Word 文档的标题结构与总段落数、总字数。不含正文。长文档请先调用。',
      parameters: Type.Object({}),
      execute: async (_id, _params, signal) => textResult(await callClient('doc_outline', {}, signal)),
    }),

    defineTool({
      name: 'doc_read',
      label: '读取正文',
      description: '读取指定段落范围的正文。段落索引来自 doc_outline 或 doc_search。单次最多约 6000 字。',
      parameters: Type.Object({
        from: Type.Integer({ description: '起始段落索引(含)' }),
        to: Type.Integer({ description: '结束段落索引(不含)' }),
      }),
      execute: async (_id, params, signal) => textResult(await callClient('doc_read', params, signal)),
    }),

    defineTool({
      name: 'doc_search',
      label: '检索文档',
      description: '在全文检索字符串,返回命中所在的段落索引与前后段落。',
      parameters: Type.Object({
        query: Type.String({ description: '要检索的原文片段' }),
        limit: Type.Optional(Type.Integer({ description: '最多返回多少条,默认 8' })),
      }),
      execute: async (_id, params, signal) => textResult(await callClient('doc_search', params, signal)),
    }),

    defineTool({
      name: 'propose_insert',
      label: '提交插入内容',
      description:
        '提交待插入的新内容草稿(标题 / 段落 / 列表 / 表格),用户勾选确认后才写入文档。用于新增章节、补写段落、生成表格。',
      parameters: Type.Object({
        summary: Type.String({ description: '一句话说明要插入什么' }),
        anchorParagraphIndex: Type.Optional(
          Type.Integer({ description: '插入到该段落之后(0 起);省略则接在当前光标所在段落之后' }),
        ),
        blocks: Type.Array(
          Type.Object({
            kind: Type.String({ description: 'heading | paragraph | list | table' }),
            level: Type.Optional(Type.Integer({ description: 'heading 层级 1-4,默认 2' })),
            text: Type.Optional(Type.String({ description: 'heading / paragraph 的正文' })),
            items: Type.Optional(Type.Array(Type.String(), { description: 'list 的条目' })),
            rows: Type.Optional(
              Type.Array(Type.Array(Type.String()), { description: 'table 的二维数组,第一行作表头' }),
            ),
          }),
        ),
      }),
      execute: async (_id, params, signal) => textResult(await callClient('propose_insert', params, signal)),
    }),

    defineTool({
      name: 'propose_comment',
      label: '提交批注',
      description:
        '提交批注草稿,以 Word 原生批注标在原文上,不修改正文。用户逐条勾选后才会写入。用于审阅、找问题、提建议。',
      parameters: Type.Object({
        summary: Type.String({ description: '一句话总结本次审阅关注点' }),
        basis: BASIS_SCHEMA,
        comments: Type.Array(
          Type.Object({
            paragraphIndex: Type.Optional(Type.Integer({ description: '目标段落索引;省略表示用户当前选中的文字' })),
            find: Type.Optional(Type.String({ description: '该段落内要批注的连续原文片段(逐字复制);省略表示整段' })),
            text: Type.String({ description: '批注内容:指出问题并给建议,尽量不超过 80 字' }),
            severity: Type.Optional(Type.String({ description: 'info 或 warn,默认 info' })),
          }),
        ),
      }),
      execute: async (_id, params, signal) => textResult(await callClient('propose_comment', params, signal)),
    }),

    defineTool({
      name: 'propose_edit',
      label: '提交改动',
      description: '提交改动草稿,由用户在界面上逐处勾选确认后才会写入文档。此工具本身不修改文档。',
      parameters: Type.Object({
        summary: Type.String({ description: '一句话总结本次改动' }),
        basis: BASIS_SCHEMA,
        edits: Type.Array(
          Type.Object({
            paragraphIndex: Type.Optional(Type.Integer({ description: '目标段落索引;省略表示用户当前选中的文字' })),
            find: Type.Optional(Type.String({ description: '该段落内要被替换的连续原文片段(逐字复制);省略表示整段替换' })),
            replace: Type.String({ description: '替换后的文本' }),
            reason: Type.Optional(Type.String({ description: '这一处为什么要改' })),
          }),
        ),
      }),
      execute: async (_id, params, signal) => textResult(await callClient('propose_edit', params, signal)),
    }),

    defineTool({
      name: 'propose_audit',
      label: '提交全文审查',
      description:
        '提交全文审查清单(结构 / 声明 / 引文 / 逻辑)。每条带段落定位与修改建议,用户在界面上跳转查看并逐项转成 Word 批注。不修改正文。',
      parameters: Type.Object({
        summary: Type.String({ description: '一句话总结本次审查范围与结论' }),
        basis: BASIS_SCHEMA,
        issues: Type.Array(
          Type.Object({
            category: Type.Optional(Type.String({ description: '短标签,如「声明空缺」「引文格式」「逻辑断层」' })),
            severity: Type.Optional(Type.String({ description: 'priority 或 normal;priority 只给阻断投稿的问题' })),
            paragraphIndex: Type.Optional(Type.Integer({ description: '问题所在段落索引(0 起)' })),
            find: Type.Optional(Type.String({ description: '该段内可定位的原文片段(逐字复制)' })),
            title: Type.String({ description: '一句话题目,不超过 30 字' }),
            detail: Type.String({ description: '问题是什么、为什么是问题' }),
            suggestion: Type.Optional(Type.String({ description: '建议怎么补/怎么改' })),
          }),
        ),
      }),
      execute: async (_id, params, signal) => textResult(await callClient('propose_audit', params, signal)),
    }),

    /* ---------------- 参考文献:在 Node 本地执行,不走前端 RPC ---------------- */

    defineTool({
      name: 'ref_list',
      label: '列出附件',
      description: '列出本次对话已附加的参考文献(名称、页数、块数、refId)。',
      parameters: Type.Object({}),
      execute: async () => {
        const { docKey, id } = refScope()
        const refs = listRefs(docKey, id)
        return textResult({
          count: refs.length,
          refs: refs.map((r) => ({
            refId: r.id,
            name: r.name,
            kind: r.kind,
            pages: r.pages,
            blocks: r.chunks,
            chars: r.chars,
            path: r.path,
          })),
        })
      },
    }),

    defineTool({
      name: 'ref_search',
      label: '检索附件',
      description:
        '在已附加的参考文献里检索,返回命中的段落块(含页码、块号与原文片段)。要看完整段落再用 ref_read。',
      parameters: Type.Object({
        query: Type.String({ description: '要检索的原文片段或关键词;中英文都行' }),
        limit: Type.Optional(Type.Integer({ description: '最多返回几条,默认 6' })),
        refId: Type.Optional(Type.String({ description: '只搜某一份文献;省略表示搜全部附件' })),
      }),
      execute: async (_id, params) => {
        const { docKey, id } = refScope()
        const limit = Number(params.limit)
        const { hits, refs, scanned } = searchRefs(
          docKey,
          id,
          String(params.query ?? ''),
          Number.isFinite(limit) && limit > 0 ? Math.min(limit, 12) : 6,
          typeof params.refId === 'string' && params.refId.trim() ? params.refId.trim() : undefined,
        )
        return textResult({
          attachments: refs.length,
          scannedBlocks: scanned,
          hits: hits.map((h) => ({
            refId: h.refId,
            name: h.name,
            page: h.page,
            block: h.block,
            snippet: h.snippet,
            score: h.score,
          })),
          note: hits.length
            ? 'snippet 只是片段;要完整原文请用 ref_read(refId, from=block, to=block)'
            : '没有命中,可换关键词再搜,或如实告知用户附件里没有相关内容',
        })
      },
    }),

    defineTool({
      name: 'ref_read',
      label: '读取附件原文',
      description: '按块号或页码读取参考文献原文,单次约 6000 字。块号来自 ref_search 的 block 字段。',
      parameters: Type.Object({
        refId: Type.Optional(Type.String({ description: '文献 id;只有一份附件时可省略' })),
        from: Type.Optional(Type.Integer({ description: '起始块号(含),默认 0' })),
        to: Type.Optional(Type.Integer({ description: '结束块号(含);省略表示从 from 起往后读' })),
        page: Type.Optional(Type.Integer({ description: '只读某一页(1 起);与 from/to 二选一' })),
      }),
      execute: async (_id, params) => {
        const { docKey, id } = refScope()
        const refs = listRefs(docKey, id)
        if (!refs.length) throw new Error('本次对话还没有附加任何参考文献')

        const wanted = typeof params.refId === 'string' && params.refId.trim() ? params.refId.trim() : null
        const target = wanted ? refs.find((r) => r.id === wanted) : refs.length === 1 ? refs[0] : null
        if (!target) throw new Error(wanted ? `没有 refId=${wanted} 的附件` : '有多份附件,请指定 refId')

        const from = Number.isFinite(Number(params.from)) ? Number(params.from) : 0
        const to = Number.isFinite(Number(params.to)) ? Number(params.to) : Number.POSITIVE_INFINITY
        const page = Number.isFinite(Number(params.page)) && Number(params.page) > 0 ? Number(params.page) : undefined
        return textResult(readRefBlocks(docKey, id, target.id, from, to, page))
      },
    }),
  ]
}

async function ensureSession(): Promise<AgentSession> {
  if (session) return session
  if (starting) return starting

  starting = (async () => {
    fs.mkdirSync(WORK_DIR, { recursive: true })

    const loader = new DefaultResourceLoader({
      cwd: WORK_DIR,
      agentDir: getAgentDir(),
      systemPromptOverride: () => WORD_PROMPT,
      // 不加载 ~/.pi 或项目里的 APPEND_SYSTEM.md / 上下文文件
      appendSystemPromptOverride: () => [],
      extensionFactories: [
        (pi) => {
          for (const tool of buildTools({ current: '' })) pi.registerTool(tool as never)
        },
      ],
    })
    await loader.reload()

    // 每个会话一个独立目录:重开面板/切回旧会话时 Pi 能接着原来的上下文聊下去
    const persist = sessionDocKey && sessionDocKey !== 'unsaved' && sessionId
    const sessionManager = persist
      ? SessionManager.continueRecent(WORK_DIR, piDirOf(sessionDocKey!, sessionId!))
      : SessionManager.inMemory(WORK_DIR)

    // PDF 直读工具来自 pi-pdf 扩展(靠 python3 的 pypdf/pdfplumber);缺依赖就不挂,免得模型调了必失败
    const pdfTools = (await pdfToolDepsOk()) ? ['pdf_info', 'pdf_extract_text', 'pdf_extract_tables'] : []

    // WordAI 全局模型:创建/恢复会话都强制用它,避免旧 transcript 把模型切回去
    const runtime = await ensureModelRuntime().catch(() => null)
    const stored = await loadStoredModel().catch(() => null)
    const pinned = runtime && stored ? runtime.getModel(stored.provider, stored.id) : undefined

    const created = await createAgentSession({
      resourceLoader: loader,
      sessionManager,
      modelRuntime: runtime ?? undefined,
      model: pinned,
      // 白名单:只保留与 Word 相关的工具,不暴露 bash / 文件读写
      tools: [
        'doc_outline',
        'doc_read',
        'doc_search',
        'propose_edit',
        'propose_comment',
        'propose_insert',
        'propose_audit',
        'ref_list',
        'ref_search',
        'ref_read',
        ...pdfTools,
      ],
    })

    created.session.subscribe((event) => {
      if (event.type === 'message_update' && event.assistantMessageEvent.type === 'text_delta') {
        emit({ type: 'text', delta: event.assistantMessageEvent.delta })
      }
    })

    // 恢复旧会话时 Pi 会按 transcript 恢复当时记录的模型,这里再覆盖一次
    if (pinned) {
      const current = created.session.model
      if (!current || current.provider !== pinned.provider || current.id !== pinned.id) {
        try {
          await created.session.setModel(pinned)
        } catch (e) {
          console.log(`[model] 覆盖会话模型失败: ${(e as Error).message}`)
        }
      }
    }

    session = created.session
    return created.session
  })()

  try {
    return await starting
  } finally {
    starting = null
  }
}

/**
 * 绑定文档 + 会话。任一变化都丢弃内存中的 session,下次按新目录重新加载。
 * 返回 true 表示发生了切换。
 */
export async function bindSession(docKey: string, wantedId?: string | null): Promise<boolean> {
  const key = docKey || 'unknown'
  const idx = key.startsWith('unsaved') ? null : ensureIndex(key)
  const id = wantedId && idx?.sessions.some((s) => s.id === wantedId) ? wantedId : (idx?.activeId ?? null)

  if (sessionDocKey === key && sessionId === id) return false

  const had = sessionDocKey !== null || sessionId !== null
  await resetSession()
  sessionDocKey = key
  sessionId = id
  return had
}

/** 兼容旧调用 */
export const bindDoc = (docKey: string) => bindSession(docKey)

/** 切换会话:丢弃内存 session,下次按新目录加载 */
export async function setSessionId(id: string | null): Promise<void> {
  if (sessionId === id) return
  await resetSession()
  sessionId = id
}

export async function runPrompt(prompt: string): Promise<void> {
  if (running) throw new Error('上一个任务还在运行')
  const s = await ensureSession()
  running = true
  aborting = false
  buffer.length = 0
  try {
    // 附件清单拼在提示词最前面:模型不调工具也知道手上有哪些文献、路径是什么
    const { docKey, id } = refScope()
    const refsBlock = refsForPrompt(docKey, id)
    await s.prompt(refsBlock ? `${refsBlock}\n\n${prompt}` : prompt)
    emit({ type: aborting ? 'aborted' : 'done' })
  } catch (e) {
    if (aborting) emit({ type: 'aborted' })
    else emit({ type: 'error', message: (e as Error).message })
  } finally {
    aborting = false
    running = false
  }
}

/**
 * 中止当前生成。两件事必须同时做:
 * 1) 挂起的工具调用就地失败 —— 前端已经断流了,再等就是白等 90s 超时;
 * 2) 让 Pi 自己停 —— 只断前端流的话,模型会继续跑完并烧额度。
 */
export async function abortRun(): Promise<boolean> {
  if (!running) return false
  aborting = true
  for (const [, p] of pending) p.reject(new Error('已中止'))
  pending.clear()
  try {
    await session?.abort()
  } catch {
    /* 已经跑完了就算了 */
  }
  return true
}

export async function resetSession(): Promise<void> {
  sessionDocKey = null
  sessionId = null
  for (const [, p] of pending) p.reject(new Error('会话已重置'))
  pending.clear()
  buffer.length = 0
  const old = session
  session = null
  old?.dispose()
}

export function sessionInfo() {
  return {
    running,
    hasSession: Boolean(session),
    docKey: sessionDocKey,
    sessionId,
    workDir: WORK_DIR,
  }
}
