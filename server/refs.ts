import crypto from 'node:crypto'
import { execFile } from 'node:child_process'
import fs from 'node:fs'
import path from 'node:path'
import { promisify } from 'node:util'
import { docDir } from './sessions.ts'

/**
 * 参考文献(附件)库。
 *
 * 附件不塞进提示词:一份论文几万字,塞进去既爆 token 又诱导模型编造。
 * 这里把文件抽成文本 → 切块 → 落盘,模型通过 ref_search / ref_read 按需取原文。
 * 与 Word 的 6 个工具不同,文献工具直接在 Node 里执行,不走前端 RPC,没有 90s 超时风险。
 */

const run = promisify(execFile)

export const MAX_REFS = 6
const MAX_FILE_BYTES = 80 * 1024 * 1024
const MAX_CHARS = 400_000
const CHUNK_MAX = 900
const CHUNK_MIN = 320
const TEXT_DEPTS = new Set(['txt', 'md', 'markdown', 'text', 'csv', 'tsv'])
const TEXTLIKE_DEPTS = new Set(['docx', 'doc', 'rtf', 'rtfd', 'html', 'htm', 'webarchive'])

export interface RefChunk {
  i: number
  page?: number
  text: string
}

export interface RefMeta {
  id: string
  name: string
  path: string
  kind: string
  pages: number
  chars: number
  chunks: number
  addedAt: string
}

export interface RefDoc extends RefMeta {
  items: RefChunk[]
}

/** 未保存文档/无会话时只放内存:没有稳定身份,落盘反而会串稿 */
const memIndex = new Map<string, RefMeta[]>()
const memDocs = new Map<string, RefDoc>()
/** 检索用的分词缓存,按 scope/refId 存;移除附件时一起清掉 */
const tokenCache = new Map<string, Map<string, number>[]>()

const isUnsaved = (docKey: string) => !docKey || docKey.startsWith('unsaved')
const scopeKey = (docKey: string, id: string) => `${docKey}::${id || 'default'}`
const refsDir = (docKey: string, id: string) => path.join(docDir(docKey), `s-${id}`, 'refs')
const indexFile = (docKey: string, id: string) => path.join(refsDir(docKey, id), 'index.json')
const docFile = (docKey: string, id: string, refId: string) => path.join(refsDir(docKey, id), `${refId}.json`)

function atomicWrite(file: string, data: string) {
  fs.mkdirSync(path.dirname(file), { recursive: true })
  const tmp = `${file}.${process.pid}.tmp`
  fs.writeFileSync(tmp, data, 'utf8')
  fs.renameSync(tmp, file)
}

const persist = (docKey: string, id: string) => !isUnsaved(docKey) && Boolean(id)

/* ------------------------------ 文件选择 ------------------------------ */

/**
 * 弹 macOS 原生文件对话框。
 * 不用 <input type=file>:Office 侧栏是 WKWebView,历史上对 file input 支持不稳,
 * 而且原生对话框能直接拿到绝对路径,省掉上传与临时文件。
 */
export async function pickFiles(): Promise<string[]> {
  const script = [
    'tell application "Finder" to activate',
    'set fs to (choose file with prompt "选择参考文献(可多选)" with multiple selections allowed)',
    'set out to ""',
    'repeat with f in fs',
    'set out to out & (POSIX path of f) & linefeed',
    'end repeat',
    'return out',
  ]
  let stdout: string
  try {
    const r = await run('osascript', script.flatMap((line) => ['-e', line]), { timeout: 180_000 })
    stdout = r.stdout
  } catch (e) {
    const msg = String((e as { stderr?: string; message?: string }).stderr ?? (e as Error).message ?? '')
    if (/User canceled|-128/.test(msg)) return []
    throw new Error(`打开文件对话框失败:${msg.trim().slice(0, 200)}`)
  }
  return stdout
    .split('\n')
    .map((s) => s.trim())
    .filter(Boolean)
}

/* ------------------------------ 文本抽取 ------------------------------ */

const PY_PDF = `
import json, sys
try:
    import pymupdf as fitz
except ImportError:
    import fitz

doc = fitz.open(sys.argv[1])
pages = []
for page in doc:
    paras = []
    for b in page.get_text("blocks"):
        if len(b) >= 7 and b[6] != 0:  # 0 = 文本块,其余是图片
            continue
        t = (b[4] or "").strip()
        if t:
            paras.append(t)
    if not paras:
        t = page.get_text().strip()
        if t:
            paras = [t]
    pages.append(paras)
print(json.dumps({"pages": doc.page_count, "pages_text": pages}, ensure_ascii=False))
`

/** 去掉软连字符、把段内换行接回去、压缩多余空白 */
function tidy(text: string): string {
  return text
    .replace(/\u00ad/g, '')
    .replace(/\r\n?/g, '\n')
    .replace(/[ \t\u00a0]+/g, ' ')
    .replace(/ *\n */g, '\n')
    .trim()
}

/** 长段按句末切开,尽量不切断句子 */
function splitLong(text: string, max: number): string[] {
  if (text.length <= max) return [text]
  const out: string[] = []
  let rest = text
  while (rest.length > max) {
    const window = rest.slice(0, max)
    const cut = Math.max(
      window.lastIndexOf('。'),
      window.lastIndexOf('. '),
      window.lastIndexOf('!'),
      window.lastIndexOf('?'),
      window.lastIndexOf(';'),
      window.lastIndexOf('；'),
      window.lastIndexOf('\n'),
    )
    const at = cut > max * 0.5 ? cut + 1 : max
    out.push(rest.slice(0, at).trim())
    rest = rest.slice(at).trim()
  }
  if (rest) out.push(rest)
  return out
}

/** 把段落合并成 320–900 字的块:太碎检索没上下文,太长一次读回来浪费 token */
function chunkParagraphs(paras: { text: string; page?: number }[]): RefChunk[] {
  const items: RefChunk[] = []
  let buf = ''
  let bufPage: number | undefined

  const flush = () => {
    const text = buf.trim()
    if (text) items.push({ i: items.length, page: bufPage, text })
    buf = ''
    bufPage = undefined
  }

  for (const p of paras) {
    const text = tidy(p.text)
    if (!text) continue
    if (buf && (bufPage !== p.page || buf.length + text.length > CHUNK_MAX)) flush()

    for (const piece of splitLong(text, CHUNK_MAX)) {
      if (buf && bufPage !== p.page) flush()
      if (!buf) bufPage = p.page
      buf = buf ? `${buf}\n${piece}` : piece
      if (buf.length >= CHUNK_MIN && buf.length + piece.length > CHUNK_MAX) flush()
    }
  }
  flush()
  return items
}

/** 每页重复出现的页眉页脚:短文本 + 出现在过半页面上就丢掉 */
function dropRunningHeads(pagesText: string[][]): string[][] {
  if (pagesText.length < 4) return pagesText
  const seen = new Map<string, number>()
  for (const page of pagesText) {
    for (const p of page) {
      const key = p.slice(0, 80)
      if (key.length <= 80) seen.set(key, (seen.get(key) ?? 0) + 1)
    }
  }
  const limit = pagesText.length / 2
  return pagesText.map((page) =>
    page.filter((p) => {
      const key = p.slice(0, 80)
      return key.length > 80 || (seen.get(key) ?? 0) <= limit
    }),
  )
}

async function extractPdf(file: string): Promise<{ kind: string; pages: number; items: RefChunk[] }> {
  const { stdout } = await run('python3', ['-c', PY_PDF, file], { timeout: 120_000, maxBuffer: 64 * 1024 * 1024 })
  const parsed = JSON.parse(stdout) as { pages: number; pages_text: string[][] }
  const pagesText = dropRunningHeads(parsed.pages_text ?? [])
  const paras: { text: string; page?: number }[] = []
  pagesText.forEach((page, idx) => {
    for (const text of page) paras.push({ text, page: idx + 1 })
  })
  return { kind: 'pdf', pages: parsed.pages ?? pagesText.length, items: chunkParagraphs(paras) }
}

async function extractTextLike(file: string): Promise<{ kind: string; pages: number; items: RefChunk[] }> {
  const { stdout } = await run('textutil', ['-convert', 'txt', '-stdout', file], {
    timeout: 60_000,
    maxBuffer: 64 * 1024 * 1024,
  })
  const paras = stdout
    .split(/\n{1,}/)
    .map((t) => ({ text: t }))
    .filter((p) => p.text.trim())
  return { kind: 'text', pages: 0, items: chunkParagraphs(paras) }
}

async function extractPlain(file: string): Promise<{ kind: string; pages: number; items: RefChunk[] }> {
  const raw = fs.readFileSync(file, 'utf8')
  const paras = raw
    .split(/\n{1,}/)
    .map((t) => ({ text: t }))
    .filter((p) => p.text.trim())
  return { kind: 'text', pages: 0, items: chunkParagraphs(paras) }
}

async function extractFile(file: string): Promise<{ kind: string; pages: number; items: RefChunk[] }> {
  const ext = path.extname(file).slice(1).toLowerCase()
  if (ext === 'pdf') return extractPdf(file)
  if (TEXT_DEPTS.has(ext)) return extractPlain(file)
  if (TEXTLIKE_DEPTS.has(ext)) return extractTextLike(file)
  throw new Error(`不支持的格式 .${ext || '?'}(可用:pdf / docx / doc / rtf / html / txt / md)`)
}

/* ------------------------------ 增删查 ------------------------------ */

export function listRefs(docKey: string, id: string): RefMeta[] {
  const key = scopeKey(docKey, id)
  if (!persist(docKey, id)) return memIndex.get(key) ?? []
  try {
    const parsed = JSON.parse(fs.readFileSync(indexFile(docKey, id), 'utf8'))
    return Array.isArray(parsed) ? (parsed as RefMeta[]) : []
  } catch {
    return memIndex.get(key) ?? []
  }
}

function saveIndex(docKey: string, id: string, refs: RefMeta[]) {
  const key = scopeKey(docKey, id)
  memIndex.set(key, refs)
  if (!persist(docKey, id)) return
  atomicWrite(indexFile(docKey, id), JSON.stringify(refs, null, 2))
}

function loadDoc(docKey: string, id: string, refId: string): RefDoc | null {
  const key = scopeKey(docKey, id)
  const hit = memDocs.get(`${key}/${refId}`)
  if (hit) return hit
  if (!persist(docKey, id)) return null
  try {
    return JSON.parse(fs.readFileSync(docFile(docKey, id, refId), 'utf8')) as RefDoc
  } catch {
    return null
  }
}

export interface AttachResult {
  added: RefMeta[]
  skipped: string[]
  errors: { file: string; error: string }[]
}

/** 拖拽上传的文件没有本地路径,只带一个展示用的名字 */
export interface AttachInput {
  path: string
  name?: string
}

export async function attachFiles(
  docKey: string,
  id: string,
  files: Array<string | AttachInput>,
): Promise<AttachResult> {
  const added: RefMeta[] = []
  const skipped: string[] = []
  const errors: { file: string; error: string }[] = []
  const refs = listRefs(docKey, id)

  for (const input of files) {
    const file = typeof input === 'string' ? input : input.path
    const fromUpload = typeof input !== 'string' && Boolean(input.name)
    const name = (typeof input === 'string' ? '' : (input.name ?? '')) || path.basename(file)
    try {
      if (refs.length >= MAX_REFS) throw new Error(`最多附加 ${MAX_REFS} 份文献,请先移除一份`)
      if (!fs.existsSync(file)) throw new Error('文件不存在')
      const stat = fs.statSync(file)
      if (stat.isDirectory()) throw new Error('这是个文件夹,请选择文件')
      if (stat.size > MAX_FILE_BYTES) throw new Error('文件超过 80MB')

      const dup = fromUpload ? refs.find((r) => r.name === name) : refs.find((r) => r.path === file)
      if (dup) {
        skipped.push(name)
        continue
      }

      const { kind, pages, items } = await extractFile(file)
      if (!items.length) throw new Error('抽不出文本(可能是扫描件,需先 OCR)')

      // 超大文献截断:宁可能搜到前半部分,也不要一次塞爆内存
      let chars = 0
      const kept: RefChunk[] = []
      for (const it of items) {
        if (chars + it.text.length > MAX_CHARS) break
        kept.push({ ...it, i: kept.length })
        chars += it.text.length
      }

      const meta: RefMeta = {
        id: crypto.randomBytes(4).toString('hex'),
        name,
        // 拖拽上传的文件没有本地路径,留空——提示词里也不会让模型去读它
        path: fromUpload ? '' : file,
        kind,
        pages,
        chars,
        chunks: kept.length,
        addedAt: new Date().toISOString(),
      }
      const doc: RefDoc = { ...meta, items: kept }
      memDocs.set(`${scopeKey(docKey, id)}/${meta.id}`, doc)
      if (persist(docKey, id)) atomicWrite(docFile(docKey, id, meta.id), JSON.stringify(doc))

      refs.push(meta)
      added.push(meta)
    } catch (e) {
      errors.push({ file: name, error: (e as Error).message })
    }
  }

  if (added.length) saveIndex(docKey, id, refs)
  return { added, skipped, errors }
}

export function removeRef(docKey: string, id: string, refId: string): RefMeta[] {
  const key = scopeKey(docKey, id)
  const refs = listRefs(docKey, id).filter((r) => r.id !== refId)
  saveIndex(docKey, id, refs)
  memDocs.delete(`${key}/${refId}`)
  tokenCache.delete(`${key}/${refId}`)
  if (persist(docKey, id)) fs.rmSync(docFile(docKey, id, refId), { force: true })
  return refs
}

/* ------------------------------ 检索 ------------------------------ */

/** 英文按词、中文按 bigram。中文没有空格,不切二元组就只能整段匹配 */
function tokenize(text: string): string[] {
  const out: string[] = []
  for (const m of text.toLowerCase().matchAll(/[a-z0-9][a-z0-9._-]{1,}/g)) out.push(m[0])
  for (const m of text.matchAll(/[\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF]+/g)) {
    const run = m[0]
    if (run.length === 1) out.push(run)
    else for (let i = 0; i < run.length - 1; i++) out.push(run.slice(i, i + 2))
  }
  return out
}

function chunkCounts(docKey: string, id: string, doc: RefDoc): Map<string, number>[] {
  const key = `${scopeKey(docKey, id)}/${doc.id}`
  const hit = tokenCache.get(key)
  if (hit) return hit
  const counts = doc.items.map((it) => {
    const m = new Map<string, number>()
    for (const t of tokenize(it.text)) m.set(t, (m.get(t) ?? 0) + 1)
    return m
  })
  tokenCache.set(key, counts)
  return counts
}

export interface RefHit {
  refId: string
  name: string
  page: number | null
  block: number
  snippet: string
  score: number
}

export function searchRefs(
  docKey: string,
  id: string,
  query: string,
  limit = 6,
  onlyRefId?: string,
): { hits: RefHit[]; refs: RefMeta[]; scanned: number } {
  const refs = listRefs(docKey, id).filter((r) => !onlyRefId || r.id === onlyRefId)
  const docs = refs.map((r) => loadDoc(docKey, id, r.id)).filter((d): d is RefDoc => Boolean(d))
  const q = [...new Set(tokenize(query))]
  if (!q.length || !docs.length) return { hits: [], refs, scanned: 0 }

  // df 用来看哪个词稀有:稀有词命中比「的」「the」这种泛词值钱得多
  const df = new Map<string, number>()
  const perDoc = docs.map((doc) => ({ doc, counts: chunkCounts(docKey, id, doc) }))
  let total = 0
  for (const { counts } of perDoc) {
    for (const c of counts) {
      total += 1
      for (const t of q) if (c.has(t)) df.set(t, (df.get(t) ?? 0) + 1)
    }
  }

  const hits: RefHit[] = []
  const loweredQuery = query.trim().toLowerCase()
  for (const { doc, counts } of perDoc) {
    counts.forEach((c, idx) => {
      let score = 0
      for (const t of q) {
        const n = c.get(t)
        if (!n) continue
        const idf = Math.log(1 + total / (1 + (df.get(t) ?? 0)))
        score += idf * Math.min(n, 3)
      }
      if (score <= 0) return
      const text = doc.items[idx].text
      if (loweredQuery.length >= 6 && text.toLowerCase().includes(loweredQuery)) score += 6
      hits.push({
        refId: doc.id,
        name: doc.name,
        page: doc.items[idx].page ?? null,
        block: doc.items[idx].i,
        snippet: snippetOf(text, q),
        score: Math.round(score * 100) / 100,
      })
    })
  }

  hits.sort((a, b) => b.score - a.score)
  return { hits: hits.slice(0, Math.max(1, limit)), refs, scanned: total }
}

const CJK = /[\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF]/

function snippetOf(text: string, tokens: string[]): string {
  const lower = text.toLowerCase()
  let at = -1
  for (const t of tokens) {
    if (CJK.test(t) && t.length < 2) continue
    const i = lower.indexOf(t)
    if (i >= 0 && (at < 0 || i < at)) at = i
  }
  if (at < 0) return text.slice(0, 240)
  const start = Math.max(0, at - 120)
  const end = Math.min(text.length, at + 200)
  return `${start > 0 ? '…' : ''}${text.slice(start, end)}${end < text.length ? '…' : ''}`
}

export interface RefReadResult {
  ref: Pick<RefMeta, 'id' | 'name' | 'kind' | 'pages' | 'chunks' | 'chars'>
  from: number
  to: number
  items: { i: number; page: number | null; text: string }[]
}

export function readRefBlocks(
  docKey: string,
  id: string,
  refId: string,
  from = 0,
  to = Number.POSITIVE_INFINITY,
  page?: number,
): RefReadResult {
  const meta = listRefs(docKey, id).find((r) => r.id === refId)
  if (!meta) throw new Error(`没有 id 为 ${refId} 的附件`)
  const doc = loadDoc(docKey, id, refId)
  if (!doc) throw new Error(`附件 ${meta.name} 的正文已丢失,请重新附加`)

  const start = Math.max(0, Math.floor(from))
  const items: RefReadResult['items'] = []
  let chars = 0
  for (const it of doc.items) {
    if (page !== undefined && it.page !== page) continue
    if (it.i < start || it.i > to) continue
    if (chars + it.text.length > 6000 && items.length) break
    items.push({ i: it.i, page: it.page ?? null, text: it.text })
    chars += it.text.length
  }
  return {
    ref: { id: meta.id, name: meta.name, kind: meta.kind, pages: meta.pages, chunks: meta.chunks, chars: meta.chars },
    from: start,
    to: items.length ? items[items.length - 1].i : start,
    items,
  }
}

/* ------------------------------ 给提示词用的清单 ------------------------------ */

export function refsForPrompt(docKey: string, id: string): string {
  const refs = listRefs(docKey, id)
  if (!refs.length) return ''
  const lines = refs.map((r, i) => {
    const unit = r.pages ? `${r.pages} 页 / ${r.chunks} 段` : `${r.chunks} 段`
    const head = `${i + 1}. ${r.name}(${r.kind.toUpperCase()},${unit},refId=${r.id})`
    return r.path ? `${head}\n   ${r.path}` : `${head}\n   (拖拽附加,无本地路径,只能靠 ref_search / ref_read 读)`
  })
  return `【已附加参考文献】\n${lines.join('\n')}`
}

/* ------------------------------ 依赖探测 ------------------------------ */

let pyDeps: boolean | null = null

/** skill 的 pdf_* 工具靠 python3 的 pypdf/pdfplumber,缺了就别把它们挂进白名单 */
export async function pdfToolDepsOk(): Promise<boolean> {
  if (pyDeps !== null) return pyDeps
  try {
    await run('python3', ['-c', 'import pypdf, pdfplumber'], { timeout: 8000 })
    pyDeps = true
  } catch {
    pyDeps = false
  }
  return pyDeps
}
