/// <reference types="office-js" />

import { locateUniqueParagraph, type Anchor } from './lib/anchor'
import type { UndoOp, UndoResult } from './lib/undo'

/**
 * Office.js 封装。
 * 原则:AI 只能读;所有写入都由用户在 UI 上点击触发,且写入前做文本比对,
 * 防止索引漂移把改动落到错位置。
 */

let trackedRange: Word.Range | null = null
let trackedText = ''

/** 段落文本快照缓存(写操作后清空),避免每次工具调用都全量读文档 */
let snapshot: { at: number; texts: string[] } | null = null
const SNAPSHOT_TTL = 1500

export function initOffice(): Promise<void> {
  return new Promise((resolve) => {
    if (typeof Office !== 'undefined' && typeof Office.onReady === 'function') {
      Office.onReady(() => resolve())
    } else {
      resolve()
    }
  })
}

export const inWord = (): boolean => typeof Word !== 'undefined'

export function clearSnapshot() {
  snapshot = null
}

/* ------------------------------------------------------------------ */
/* 选区                                                                */
/* ------------------------------------------------------------------ */

export interface SelectionInfo {
  text: string
  empty: boolean
}

export async function captureSelection(): Promise<SelectionInfo> {
  if (!inWord()) return { text: '', empty: true }

  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection()
    sel.load('text')
    await ctx.sync()

    if (trackedRange) {
      try {
        ctx.trackedObjects.remove(trackedRange)
      } catch {
        /* 已失效 */
      }
      trackedRange = null
    }

    trackedRange = sel
    trackedText = sel.text ?? ''
    ctx.trackedObjects.add(sel)
    await ctx.sync()

    return { text: trackedText, empty: !trackedText.trim() }
  })
}

/**
 * 冻结当前选区为锚点。**必须在发送那一刻调用**,
 * 之后无论用户点到哪里,写入都只认这个锚点。
 */
export async function captureAnchor(selectionText?: string): Promise<Anchor> {
  const text = (selectionText ?? trackedText ?? '').trim()
  if (!inWord() || !text) return { paragraphIndex: null, text, chars: text.length }

  try {
    // 不用 loadTexts 的缓存:锚点必须对应当前文档,宁可多读一次
    const texts = await Word.run(async (ctx) => {
      const paras = ctx.document.body.paragraphs
      paras.load('items/text')
      await ctx.sync()
      return paras.items.map((p) => p.text ?? '')
    })
    return { paragraphIndex: locateUniqueParagraph(texts, text), text, chars: text.length }
  } catch {
    return { paragraphIndex: null, text, chars: text.length }
  }
}

export interface ApplyResult {
  ok: boolean
  message?: string
}

export async function applyReplacement(
  revised: string,
  expectedOriginal: string,
  anchor?: Anchor,
): Promise<ApplyResult> {
  if (!inWord()) return { ok: false, message: '当前不在 Word 环境中' }
  // 没有原文就绝不动手:空选区插文字是最难排查的「静默写错」
  if (!expectedOriginal.trim()) {
    return { ok: false, message: '缺少原文锚点:请在文档中选中要改的文字,或让 AI 指定段落与 find。' }
  }

  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection()
    sel.load('text')
    await ctx.sync()

    const selText = sel.text ?? ''
    let target: Word.Range | null = null
    let via = ''

    // 1) 冻结锚点优先:段落号 + 原文片段,完全不依赖「此刻选区」
    //    search 有 255 字上限,超过 200 字就交给后面的 locateByText(逐字校验)
    if (anchor && anchor.paragraphIndex !== null && anchor.text.trim() && anchor.text.length <= 200) {
      const paras = ctx.document.body.paragraphs
      paras.load('items/text')
      await ctx.sync()
      const p = paras.items[anchor.paragraphIndex]
      if (p) {
        const hits = p.search(anchor.text, { matchCase: true })
        hits.load('items')
        await ctx.sync()
        if (hits.items.length) {
          target = hits.items[0]
          via = '冻结锚点'
        }
      }
    }

    // 2) 锚点失效才退回选区;注意这是兜底,不是主路径
    if (!target && selText === expectedOriginal) {
      target = sel
      via = '当前选区'
    } else if (!target && trackedRange) {
      try {
        ctx.trackedObjects.add(trackedRange)
        trackedRange.load('text')
        await ctx.sync()
        if ((trackedRange.text ?? '') === expectedOriginal) {
          target = trackedRange
          via = '保活选区'
        }
      } catch {
        /* 对象已失效,走失败分支 */
      }
    }

    // 选区没了(比如撤销之后)就靠文本重新定位一次
    if (!target) {
      const found = await locateByText(ctx, expectedOriginal)
      if (found) {
        target = found
        via = '文本定位'
      }
    }

    // 3) 最后兜底:按全文唯一文本定位(逐字校验)
    if (!target) {
      return {
        ok: false,
        message: `找不到原文(锚点 ${anchor?.text.length ?? 0} 字,当前选区 ${selText.length} 字,期望 ${expectedOriginal.length} 字)。可能已被手动修改,请重新选中后再试。`,
      }
    }

    target.insertText(revised, Word.InsertLocation.replace)
    await ctx.sync()

    // 回读校验:不验证就报成功 = 骗用户
    const body = ctx.document.body
    body.load('text')
    await ctx.sync()

    const probe = revised.replace(/\s+/g, '').slice(0, 20)
    const docText = (body.text ?? '').replace(/\s+/g, '')
    const written = probe.length > 0 && docText.includes(probe)

    if (!written) {
      return {
        ok: false,
        message: `写入未生效(经${via},选区 ${target.text?.length ?? 0} 字)。文档内容没有变化,请重试。`,
      }
    }

    if (trackedRange) {
      try {
        ctx.trackedObjects.remove(trackedRange)
      } catch {
        /* ignore */
      }
      trackedRange = null
      trackedText = ''
    }
    clearSnapshot()

    return { ok: true }
  })
}

/* ------------------------------------------------------------------ */
/* 只读:文档结构 / 正文 / 检索                                          */
/* ------------------------------------------------------------------ */

export interface OutlineItem {
  index: number
  level: number
  text: string
}

export interface Outline {
  totalParagraphs: number
  totalChars: number
  headings: OutlineItem[]
  truncated: boolean
}

const headingLevel = (style: string, builtin: string): number | null => {
  const s = `${style} ${builtin}`
  const m = s.match(/Heading\s*([1-9])/i) ?? s.match(/标题\s*([1-9])/)
  if (m) return Number(m[1])
  return null
}

const MAX_HEADINGS = 120

export async function getOutline(): Promise<Outline> {
  if (!inWord()) throw new Error('当前不在 Word 环境中')

  return Word.run(async (ctx) => {
    const paras = ctx.document.body.paragraphs
    let texts: string[] = []
    let headings: OutlineItem[] = []

    try {
      // 优先拿样式,才能识别标题层级
      paras.load('text/style/styleBuiltIn')
      await ctx.sync()

      const items = paras.items
      texts = items.map((p) => p.text ?? '')

      for (let i = 0; i < items.length; i++) {
        const level = headingLevel(items[i].style ?? '', String(items[i].styleBuiltIn ?? ''))
        if (level === null) continue
        const text = texts[i].trim()
        if (!text) continue
        headings.push({ index: i, level, text })
        if (headings.length >= MAX_HEADINGS) break
      }
    } catch {
      // 旧版 Word 不支持 styleBuiltIn 时降级:只给段落数与正文
      paras.load('text')
      await ctx.sync()
      texts = paras.items.map((p) => p.text ?? '')
    }

    snapshot = { at: Date.now(), texts }

    return {
      totalParagraphs: texts.length,
      totalChars: texts.reduce((n, t) => n + t.length, 0),
      headings,
      truncated: headings.length >= MAX_HEADINGS,
    }
  })
}

async function loadTexts(): Promise<string[]> {
  if (!inWord()) throw new Error('当前不在 Word 环境中')
  if (snapshot && Date.now() - snapshot.at < SNAPSHOT_TTL) return snapshot.texts

  const texts = await Word.run(async (ctx) => {
    const paras = ctx.document.body.paragraphs
    paras.load('text')
    await ctx.sync()
    return paras.items.map((p) => p.text ?? '')
  })

  snapshot = { at: Date.now(), texts }
  return texts
}

export interface ReadResult {
  total: number
  from: number
  to: number
  chars: number
  truncated: boolean
  paragraphs: { index: number; text: string }[]
}

export async function readParagraphs(from: number, to: number, maxChars = 6000): Promise<ReadResult> {
  const texts = await loadTexts()
  const total = texts.length
  const start = Math.max(0, Math.min(Math.floor(from), total - 1))
  const end = Math.max(start + 1, Math.min(Math.floor(to), total))

  const paragraphs: { index: number; text: string }[] = []
  let chars = 0
  let truncated = false

  for (let i = start; i < end; i++) {
    const text = texts[i]
    if (chars + text.length > maxChars) {
      truncated = true
      break
    }
    chars += text.length
    paragraphs.push({ index: i, text })
  }

  return { total, from: start, to: start + paragraphs.length, chars, truncated, paragraphs }
}

export interface SearchHit {
  index: number
  text: string
  before?: { index: number; text: string }
  after?: { index: number; text: string }
}

export interface SearchResult {
  total: number
  query: string
  hits: SearchHit[]
}

export async function searchDocument(query: string, limit = 8): Promise<SearchResult> {
  const texts = await loadTexts()
  const q = query.toLowerCase()
  const hits: SearchHit[] = []

  for (let i = 0; i < texts.length && hits.length < limit; i++) {
    const text = texts[i]
    if (!text || !text.toLowerCase().includes(q)) continue
    hits.push({
      index: i,
      text: text.length > 400 ? `${text.slice(0, 400)}…` : text,
      before: i > 0 ? { index: i - 1, text: texts[i - 1].slice(0, 160) } : undefined,
      after: i < texts.length - 1 ? { index: i + 1, text: texts[i + 1].slice(0, 160) } : undefined,
    })
  }

  return { total: texts.length, query, hits }
}

/* ------------------------------------------------------------------ */
/* 写入:按段落定位                                                     */
/* ------------------------------------------------------------------ */

export interface DocEdit {
  paragraphIndex: number
  find: string | null
  replace: string
}

export interface EditOutcome {
  at: number
  ok: boolean
  reason?: string
  /** 写入前的整段原文,撤销时用来还原 */
  before?: string
  /** 写入后的整段文本,撤销时用来定位 */
  after?: string
}

/**
 * 在一次 Word.run 里完成多段改写。
 * 每处都先校验:段落存在 + find 能在该段落内找到 → 才动手。
 */
export async function applyDocEdits(
  edits: DocEdit[],
): Promise<{ applied: number; outcomes: EditOutcome[] }> {
  if (!inWord()) throw new Error('当前不在 Word 环境中')
  if (!edits.length) return { applied: 0, outcomes: [] }

  const outcomes: EditOutcome[] = []
  let applied = 0

  await Word.run(async (ctx) => {
    const paras = ctx.document.body.paragraphs
    paras.load('items')
    await ctx.sync()

    const items = paras.items

    for (let k = 0; k < edits.length; k++) {
      const e = edits[k]

      if (!Number.isInteger(e.paragraphIndex) || e.paragraphIndex < 0) {
        outcomes.push({ at: k, ok: false, reason: `段落号无效(${e.paragraphIndex})` })
        continue
      }

      const p = items[e.paragraphIndex]

      if (!p) {
        outcomes.push({ at: k, ok: false, reason: `第 ${e.paragraphIndex} 段不存在` })
        continue
      }

      try {
        p.load('text')
        await ctx.sync()
        const text = p.text ?? ''

        if (e.find) {
          if (!text.includes(e.find)) {
            outcomes.push({ at: k, ok: false, reason: '原文片段未找到,段落可能已被修改' })
            continue
          }
          const hits = p.search(e.find, { matchCase: true })
          hits.load('items')
          await ctx.sync()
          if (!hits.items.length) {
            outcomes.push({ at: k, ok: false, reason: 'Word 定位失败(可能含特殊字符)' })
            continue
          }
          hits.items[0].insertText(e.replace, Word.InsertLocation.replace)
        } else {
          p.insertText(e.replace, Word.InsertLocation.replace)
        }

        await ctx.sync()

        // 回读校验:Word 有时会静默拒绝(表格内、受保护区域),不能盲信
        p.load('text')
        await ctx.sync()
        const probe = e.replace.replace(/\s+/g, '').slice(0, 20)
        const afterText = p.text ?? ''
        const after = afterText.replace(/\s+/g, '')
        if (probe && !after.includes(probe)) {
          outcomes.push({ at: k, ok: false, reason: '写入后回读未发现新文本,Word 可能拒绝了该位置' })
          continue
        }

        applied++
        outcomes.push({ at: k, ok: true, before: text, after: afterText })
      } catch (err) {
        outcomes.push({ at: k, ok: false, reason: (err as Error).message })
      }
    }
  })

  clearSnapshot()
  return { applied, outcomes }
}


/* ------------------------------------------------------------------ */
/* 撤销:按反向操作还原,每条都先校验再动手                              */
/* ------------------------------------------------------------------ */

const norm = (t: string) => (t ?? '').replace(/\s+/g, ' ').trim()

/**
 * 撤销后把还原的那一段重新选中。
 * 否则用户想「撤销后再改一版」时,文档里已经没有选区,再点应用会报「选区已变化」。
 */
async function reselect(ctx: Word.RequestContext, text: string): Promise<void> {
  try {
    const range = await locateByText(ctx, text)
    if (range) range.select()
  } catch {
    /* 选不中就选不中,不影响还原本身 */
  }
}

/** 插入段落撤销时最多扫描多少段,避免异常文档把主线程拖住 */
const INSERT_UNDO_LIMIT = 200

/**
 * 插入段落的兜底定位。
 *
 * 常规路径是「在 body.paragraphs 里找连续文本序列」,但真实文档里会踩两种情况:
 *  1. 段落落在表格单元格里,body.paragraphs 未必包含它;
 *  2. Word 把插入文本里的换行拆成了多段。
 * 这里改用全文检索探针 + 段落文本核对,只在**唯一命中且整段都是插入内容**时才删。
 */
async function deleteInsertedParagraphs(ctx: Word.RequestContext, texts: string[]): Promise<boolean> {
  const list = texts
    .map((raw) => ({ raw, wanted: norm(raw) }))
    .filter((x) => x.wanted)
  if (!list.length || list.length !== texts.length) return false

  const paras = ctx.document.body.paragraphs
  paras.load('items/text')
  await ctx.sync()

  const targets: Word.Paragraph[] = []

  for (const { raw, wanted } of list) {
    const exact = paras.items.filter((p) => norm(p.text) === wanted)
    if (exact.length === 1) {
      targets.push(exact[0])
      continue
    }
    if (exact.length > 1) return false

    // 段落集合里没有(表格内 / 被合并):用全文检索定位
    const probe = raw.replace(/[\r\n\v]+/g, ' ').trim().slice(0, 24)
    if (probe.length < 6) return false

    let ranges: Word.Range[] = []
    const found = ctx.document.body.search(probe, { matchCase: true })
    found.load('items')
    await ctx.sync()
    ranges = found.items

    // body.search 在部分版本里搜不到表格内容,再逐表搜一次
    if (!ranges.length) {
      const tables = ctx.document.body.tables
      tables.load('items')
      await ctx.sync()
      for (const t of tables.items.slice(0, 50)) {
        try {
          const tr = t.getRange(Word.RangeLocation.whole)
          const hits = tr.search(probe, { matchCase: true })
          hits.load('items')
          await ctx.sync()
          ranges.push(...hits.items)
        } catch {
          /* 该版本不支持表格范围搜索,跳过 */
        }
      }
    }

    if (!ranges.length || ranges.length > 20) return false

    // 探针可能命中多处,只认「整段完全等于插入内容」的那一个
    const hits: Word.Paragraph[] = []
    for (const range of ranges) {
      const ps = range.paragraphs
      ps.load('items/text')
      await ctx.sync()
      for (const p of ps.items) if (norm(p.text) === wanted) hits.push(p)
    }
    if (hits.length !== 1) return false
    targets.push(hits[0])
  }

  // 同一段被命中两次就放弃,宁可撤不掉也不乱删
  if (new Set(targets).size !== targets.length) return false

  for (const p of targets.reverse()) p.delete()
  await ctx.sync()
  return true
}

export async function undoOps(ops: UndoOp[]): Promise<UndoResult> {
  if (!inWord()) return { restored: 0, failures: ['当前不在 Word 环境中'] }

  const failures: string[] = []
  let restored = 0

  await Word.run(async (ctx) => {
    const body = ctx.document.body
    const paras = body.paragraphs
    paras.load('items/text')
    await ctx.sync()

    // 倒序执行:最后写的先撤,避免前面的定位被后面的改动干扰
    for (let i = ops.length - 1; i >= 0; i--) {
      const op = ops[i]
      try {
        if (op.kind === 'paragraph') {
          const A = norm(op.after)
          const B = norm(op.before)
          const hits = paras.items.filter((p) => norm(p.text).includes(A))

          if (hits.length !== 1) {
            failures.push(
              hits.length === 0
                ? '找不到改后的文字,可能已被手动修改'
                : '改后的文字在多处出现,无法确定该还原哪一处',
            )
            continue
          }

          const p = hits[0]
          const wholeParagraph = norm(p.text) === A

          if (wholeParagraph) {
            p.insertText(op.before, Word.InsertLocation.replace)
            await ctx.sync()
            p.load('text')
            await ctx.sync()
            if (norm(p.text) === B) {
              restored++
              await reselect(ctx, op.before)
            } else failures.push('还原后回读不一致')
            continue
          }

          // 只改了一部分:用 search 精确回退。注意 Office.js 的 search 串有长度上限,超了就放弃
          if (A.length > 200) {
            failures.push('改动片段过长(超过 200 字),无法安全定位,已跳过')
            continue
          }

          const found = p.search(op.after, { matchCase: true })
          found.load('items')
          await ctx.sync()
          if (!found.items.length) {
            failures.push('在该段中找不到改后的片段')
            continue
          }

          found.items[0].insertText(op.before, Word.InsertLocation.replace)
          await ctx.sync()
          p.load('text')
          await ctx.sync()
          if (norm(p.text).includes(B)) {
            restored++
            await reselect(ctx, op.before)
          } else failures.push('还原后回读不一致')
          continue
        }

        if (op.kind === 'comment') {
          const cs = body.getComments()
          cs.load('items/id')
          await ctx.sync()

          const hit = cs.items.find((c) => c.id === op.commentId)
          if (!hit) {
            failures.push('批注已不存在,可能已被手动删除')
            continue
          }

          // 集合批量 load 拿不到 content(部分 Word 版本报「属性 content 不可用」),
          // 必须拿到具体对象后再单独 load 一次。
          hit.load('content')
          await ctx.sync()

          // 部分 Word 版本读不到 content。这时退一步用批注锚定的正文做校验,
          // 两条路都对不上就绝不删 —— 宁可撤不掉,也不能乱删。
          if (typeof hit.content !== 'string') {
            diag(`undo comment ${op.commentId} content unreadable`)
            let anchorText = ''
            try {
              const r = hit.getRange()
              r.load('text')
              await ctx.sync()
              anchorText = r.text ?? ''
            } catch {
              /* 连范围也读不到 */
            }
            const want = norm(op.anchorText ?? '')
            const got = norm(anchorText)
            const matched = want && got && (want.includes(got) || got.includes(want))
            if (!matched) {
              failures.push('读不到批注内容,为免误删已跳过')
              continue
            }
            diag(`undo comment ${op.commentId} verified by anchored range`)
          } else if (norm(hit.content) !== norm(op.text)) {
            diag(`undo comment mismatch: "${hit.content.slice(0, 40)}" vs "${op.text.slice(0, 40)}"`)
            failures.push('批注内容已改变,为免误删已跳过')
            continue
          }

          hit.delete()
          await ctx.sync()
          restored++
          continue
        }

        if (op.kind === 'insertParagraph') {
          const rawTexts = op.texts.filter((t) => t.trim())
          if (!rawTexts.length) {
            failures.push('插入记录为空')
            continue
          }

          const all = paras.items
          const findSeq = (seq: string[]) => {
            let start = -1
            for (let j = 0; j + seq.length <= all.length && j < INSERT_UNDO_LIMIT; j++) {
              let ok = true
              for (let w = 0; w < seq.length; w++) {
                if (norm(all[j + w].text) !== seq[w]) {
                  ok = false
                  break
                }
              }
              if (!ok) continue
              if (start !== -1) return { start: -2, dup: true }
              start = j
            }
            return { start, dup: false }
          }

          let seq = rawTexts.map(norm)
          let res = findSeq(seq)

          // Word 可能把段内换行拆成了多段:按行再试一次
          if (res.start < 0 && !res.dup) {
            const parts = rawTexts
              .flatMap((t) => t.split(/[\r\n\v]+/))
              .map(norm)
              .filter(Boolean)
            if (parts.length > 1 && parts.every((x) => x.length >= 4)) {
              const r2 = findSeq(parts)
              if (r2.dup) res = r2
              else if (r2.start >= 0) {
                seq = parts
                res = r2
              }
            }
          }

          if (res.dup) {
            failures.push('插入的段落出现多处,无法确定该删哪一段')
            continue
          }

          if (res.start >= 0) {
            for (let w = seq.length - 1; w >= 0; w--) all[res.start + w].delete()
            await ctx.sync()

            paras.load('items/text')
            await ctx.sync()
            restored++
            continue
          }

          // 段落集合里找不到:全文检索兜底(表格内 / 段落被合并)
          const ok = await deleteInsertedParagraphs(ctx, rawTexts)
          if (ok) {
            paras.load('items/text')
            await ctx.sync()
            restored++
            continue
          }

          const head = norm(rawTexts[0]).slice(0, 12)
          const near = head ? all.find((p) => norm(p.text).includes(head)) : undefined
          diag(
            `undo insertParagraph miss: paras=${all.length} wanted="${norm(rawTexts[0]).slice(0, 40)}" ` +
              `near="${near ? norm(near.text).slice(0, 40) : '-'}"`,
          )
          failures.push('找不到插入的段落(可能已被手动修改或删除)')
          continue
        }

        if (op.kind === 'insertTable') {
          const tables = body.tables
          tables.load('items')
          await ctx.sync()

          let target: Word.Table | null = null
          for (const t of tables.items) {
            const r = t.getRange(Word.RangeLocation.whole)
            r.load('text')
            await ctx.sync()
            const text = norm(r.text)
            if (op.firstCell && text.startsWith(norm(op.firstCell))) {
              if (target) {
                target = null
                failures.push('首格内容相同的表格不止一个,无法确定该删哪一个')
                break
              }
              target = t
            }
          }

          if (!target) {
            if (!failures.length) failures.push('找不到插入的表格')
            continue
          }

          target.delete()
          await ctx.sync()
          restored++
        }
      } catch (e) {
        failures.push((e as Error).message)
      }
    }
  })

  clearSnapshot()
  return { restored, failures }
}

/** 一次性读出若干段落的原文,给卡片做上下文聚焦 */
export async function readParagraphContexts(indices: number[]): Promise<Record<number, string>> {
  const uniq = [...new Set(indices.filter((i) => Number.isFinite(i)))]
  if (!inWord() || !uniq.length) return {}

  return Word.run(async (ctx) => {
    const paras = ctx.document.body.paragraphs
    paras.load('items/text')
    await ctx.sync()

    const out: Record<number, string> = {}
    for (const i of uniq) {
      const p = paras.items[i]
      if (p) out[i] = p.text ?? ''
    }
    return out
  })
}

/** 把文档滚动到某处并选中目标,让用户看见 AI 说的是哪一句 */
export async function locateInDocument(paragraphIndex: number, find: string | null): Promise<boolean> {
  if (!inWord()) return false

  return Word.run(async (ctx) => {
    const paras = ctx.document.body.paragraphs
    paras.load('items')
    await ctx.sync()

    const p = paras.items[paragraphIndex]
    if (!p) return false

    if (find) {
      const hits = p.search(find, { matchCase: true })
      hits.load('items')
      await ctx.sync()
      if (hits.items.length) {
        hits.items[0].select()
        return true
      }
    }

    p.select()
    return true
  })
}


/**
 * 按文本在全文里定位一段范围。
 *
 * 用途有两个:
 *  1. 选区被撤销/编辑后失效时,靠文本重新找到目标再写
 *  2. 撤销还原后,把这段重新选中,让用户能直接再应用一次
 *
 * 用首尾片段锚定起止段落,再用 expandTo 拼成完整范围,最后**逐字校验**;
 * 对不上就返回 null,绝不猜。
 */
async function locateByText(ctx: Word.RequestContext, expected: string): Promise<Word.Range | null> {
  const target = norm(expected)
  if (!target) return null

  const paras = ctx.document.body.paragraphs
  paras.load('items/text')
  await ctx.sync()

  const segs = expected
    .split(/[\r\n\v]+/)
    .map((x) => x.trim())
    .filter(Boolean)
  const headSeg = segs[0] ?? expected.trim()
  const tailSeg = segs[segs.length - 1] ?? expected.trim()

  const headProbe = headSeg.slice(0, 40)
  const tailProbe = tailSeg.slice(-40)
  if (!headProbe || !tailProbe) return null

  let startIdx = -1
  let endIdx = -1
  for (let i = 0; i < paras.items.length; i++) {
    const t = norm(paras.items[i].text)
    if (startIdx === -1 && t.includes(headProbe)) startIdx = i
    if (t.includes(tailProbe)) endIdx = i
  }
  if (startIdx === -1 || endIdx === -1 || endIdx < startIdx) return null

  const hs = paras.items[startIdx].search(headProbe, { matchCase: true })
  const ts = paras.items[endIdx].search(tailProbe, { matchCase: true })
  hs.load('items')
  ts.load('items')
  await ctx.sync()
  if (!hs.items.length || !ts.items.length) return null

  const range = hs.items[0].expandTo(ts.items[ts.items.length - 1])
  range.load('text')
  await ctx.sync()

  return norm(range.text) === target ? range : null
}

/** 选区摘要 */
export function summarize(text: string): string {
  const clean = text.replace(/\s+/g, ' ').trim()
  if (!clean) return '未选中内容'
  return clean.length > 40 ? `${clean.slice(0, 40)}… (${clean.length} 字)` : `${clean} (${clean.length} 字)`
}

/* ------------------------------------------------------------------ */
/* 选区实时监听                                                         */
/* ------------------------------------------------------------------ */

async function readSelectionText(): Promise<string> {
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection()
    sel.load('text')
    await ctx.sync()
    return sel.text ?? ''
  })
}

/**
 * 监听 Word 里的选区变化,让侧栏随时知道用户选了什么。
 * 拖选会连续触发,所以 debounce 350ms;这里只读不保活,
 * 真正的定位仍由 captureSelection 在发送时完成。
 */
export function onSelectionChanged(cb: (text: string) => void): () => void {
  if (!inWord() || typeof Office === 'undefined' || !Office.context?.document) return () => {}

  let timer: ReturnType<typeof setTimeout> | null = null
  let disposed = false

  const handler = () => {
    if (timer) clearTimeout(timer)
    timer = setTimeout(async () => {
      if (disposed) return
      try {
        cb(await readSelectionText())
      } catch {
        /* 选区短暂不可用,忽略 */
      }
    }, 350)
  }

  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler)

  return () => {
    disposed = true
    if (timer) clearTimeout(timer)
    try {
      Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler })
    } catch {
      /* ignore */
    }
  }
}

/* ------------------------------------------------------------------ */
/* 文档身份:用于在切换文档时重置会话                                     */
/* ------------------------------------------------------------------ */

/**
 * 返回当前文档的稳定标识。
 * 已保存文档用文件路径;未保存的新文档没有 url,统一归为 'unsaved'
 * (两个都未保存的文档无法区分,靠「页面加载即视为新文档」兜底)。
 */
export async function getDocKey(): Promise<string> {
  if (!inWord() || typeof Office === 'undefined' || !Office.context?.document?.getFilePropertiesAsync) {
    return 'unknown'
  }

  return new Promise((resolve) => {
    try {
      Office.context.document.getFilePropertiesAsync((result) => {
        const url = result.status === Office.AsyncResultStatus.Succeeded ? (result.value?.url ?? '') : ''
        resolve(url || 'unsaved')
      })
    } catch {
      resolve('unknown')
    }
  })
}

/* ------------------------------------------------------------------ */
/* 批注:AI 标问题,不改正文                                             */
/* ------------------------------------------------------------------ */

export interface DocComment {
  paragraphIndex: number | null
  find: string | null
  text: string
  /** 冻结锚点:paragraphIndex 为 null 时靠它定位,绝不退回「此刻选区」 */
  anchor?: Anchor
}

export interface CommentOutcome {
  at: number
  ok: boolean
  reason?: string
  /** 撤销时按 id 精确删除 */
  commentId?: string
}

export const supportsComments = (): boolean => {
  try {
    return Boolean(Office?.context?.requirements?.isSetSupported?.('WordApi', '1.4'))
  } catch {
    return false
  }
}

/**
 * 把批注写入 Word。每条先定位(段落 + find),再 insertComment,
 * 最后回读 content 校验 —— 和正文写入一样,不盲信成功。
 */
export async function applyComments(
  list: DocComment[],
): Promise<{ applied: number; outcomes: CommentOutcome[] }> {
  if (!inWord()) throw new Error('当前不在 Word 环境中')
  if (!supportsComments()) throw new Error('当前 Word 版本不支持批注 API(需 WordApi 1.4 / Word 2021+)')
  if (!list.length) return { applied: 0, outcomes: [] }

  const outcomes: CommentOutcome[] = []
  let applied = 0

  await Word.run(async (ctx) => {
    const paras = ctx.document.body.paragraphs
    paras.load('items')
    await ctx.sync()

    for (let k = 0; k < list.length; k++) {
      const c = list[k]
      try {
        let range: Word.Range

        if (c.paragraphIndex === null) {
          const at = c.anchor
          if (at && at.paragraphIndex !== null && at.text.trim() && at.text.length <= 200) {
            const p = paras.items[at.paragraphIndex]
            if (!p) {
              outcomes.push({ at: k, ok: false, reason: `锚点段落 ${at.paragraphIndex} 已不存在` })
              continue
            }
            const hits = p.search(at.text, { matchCase: true })
            hits.load('items')
            await ctx.sync()
            if (!hits.items.length) {
              outcomes.push({ at: k, ok: false, reason: '锚点原文未找到,可能已被修改' })
              continue
            }
            range = hits.items[0]
          } else if (at && at.text.trim()) {
            const found = await locateByText(ctx, at.text)
            if (!found) {
              outcomes.push({ at: k, ok: false, reason: '锚点原文在全文找不到唯一匹配' })
              continue
            }
            range = found
          } else {
            outcomes.push({ at: k, ok: false, reason: '没有冻结的选区锚点,请重新选中后再说一次' })
            continue
          }
        } else {
          const p = paras.items[c.paragraphIndex]
          if (!p) {
            outcomes.push({ at: k, ok: false, reason: `第 ${c.paragraphIndex} 段不存在` })
            continue
          }
          if (c.find) {
            const hits = p.search(c.find, { matchCase: true })
            hits.load('items')
            await ctx.sync()
            if (!hits.items.length) {
              outcomes.push({ at: k, ok: false, reason: '原文片段未找到' })
              continue
            }
            range = hits.items[0]
          } else {
            range = p.getRange()
          }
        }

        const comment = range.insertComment(c.text)
        comment.load('content')
        await ctx.sync()

        if ((comment.content ?? '').trim() !== c.text.trim()) {
          outcomes.push({ at: k, ok: false, reason: '批注写入后回读不一致' })
          continue
        }

        applied++
        outcomes.push({ at: k, ok: true, commentId: comment.id })
      } catch (e) {
        diag(`comment ${k} ERROR: ${(e as Error).message}`)
        outcomes.push({ at: k, ok: false, reason: (e as Error).message })
      }
    }
  })

  return { applied, outcomes }
}


/* ------------------------------------------------------------------ */
/* 结构化生成:插入标题/段落/列表/表格                                   */
/* ------------------------------------------------------------------ */

export interface InsertBlock {
  kind: 'heading' | 'paragraph' | 'list' | 'table'
  level?: number
  text?: string
  items?: string[]
  rows?: string[][]
}

const diag = (line: string) => {
  try {
    void fetch('/api/debug/log', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ line }),
    })
  } catch {
    /* ignore */
  }
}

export interface InsertOutcome {
  at: number
  ok: boolean
  reason?: string
  /** 撤销用:插入的段落文本 */
  texts?: string[]
  /** 撤销用:插入的表格特征 */
  table?: { firstCell: string; rows: number }
}

type HeadingStyleName = 'Heading1' | 'Heading2' | 'Heading3' | 'Heading4'

const headingStyle = (level?: number): HeadingStyleName => {
  if (level === 1) return 'Heading1'
  if (level === 3) return 'Heading3'
  if (level && level >= 4) return 'Heading4'
  return 'Heading2'
}

/**
 * 在锚点段落之后依次插入 blocks。
 * 锚点:anchorIndex 指定的段落;为 null 时用当前光标所在段落。
 */
export async function applyInsert(
  anchorIndex: number | null,
  blocks: InsertBlock[],
): Promise<{ inserted: number; outcomes: InsertOutcome[] }> {
  if (!inWord()) throw new Error('当前不在 Word 环境中')
  if (!blocks.length) return { inserted: 0, outcomes: [] }

  const outcomes: InsertOutcome[] = []
  let inserted = 0
  diag(`applyInsert start anchor=${anchorIndex} blocks=${blocks.length}`)

  await Word.run(async (ctx) => {
    let cursor: Word.Paragraph | null = null

    if (anchorIndex !== null) {
      const paras = ctx.document.body.paragraphs
      paras.load('items')
      await ctx.sync()
      cursor = paras.items[anchorIndex] ?? null
    } else {
      const sel = ctx.document.getSelection()
      const ps = sel.paragraphs
      ps.load('items')
      await ctx.sync()
      const arr = ps.items
      cursor = arr.length ? arr[arr.length - 1] : null
    }

    diag(`cursor=${cursor ? 'found' : 'NULL'} text=${(cursor?.text ?? '').slice(0, 40)}`)

    if (!cursor) {
      for (let k = 0; k < blocks.length; k++) {
        outcomes.push({ at: k, ok: false, reason: '找不到插入位置:请先把光标放到目标段落' })
      }
      return
    }

    for (let k = 0; k < blocks.length; k++) {
      const b = blocks[k]
      try {
        if (b.kind === 'table') {
          const rows = (b.rows ?? []).filter((r) => Array.isArray(r) && r.length)
          if (!rows.length) {
            outcomes.push({ at: k, ok: false, reason: '表格没有内容' })
            continue
          }
          const cols = Math.max(...rows.map((r) => r.length))
          const values = rows.map((r) => {
            const copy = [...r]
            while (copy.length < cols) copy.push('')
            return copy
          })

          // rowCount 在插入后可能读不回真实值,校验改用表格内容回读
          const cur: Word.Paragraph = cursor
          const anchorRange: Word.Range = cur.getRange(Word.RangeLocation.after)
          diag(`table: ${values.length}x${cols} anchorRangeOk`)

          const table: Word.Table = anchorRange.insertTable(
            values.length,
            cols,
            Word.InsertLocation.after,
            values,
          )
          diag('table: insertTable returned')

          const tRange: Word.Range = table.getRange(Word.RangeLocation.whole)
          tRange.load('text')
          await ctx.sync()
          diag(`table: readback len=${(tRange.text ?? '').length} head=${(tRange.text ?? '').slice(0, 40)}`)

          // 回读校验:表格里必须能找到首格内容,否则视为没写进去
          const probe = (values[0]?.[0] ?? '').trim()
          if (probe && !(tRange.text ?? '').includes(probe)) {
            outcomes.push({ at: k, ok: false, reason: '表格回读不到内容,写入未生效' })
            continue
          }

          // 表格后必须有一个段落承接后续块,没有就补一个
          const tailPs: Word.ParagraphCollection = table.getRange(Word.RangeLocation.after).paragraphs
          tailPs.load('items')
          await ctx.sync()
          cursor =
            tailPs.items[0] ??
            table.getRange(Word.RangeLocation.after).insertParagraph('', Word.InsertLocation.after)

          inserted++
          outcomes.push({
            at: k,
            ok: true,
            table: { firstCell: (values[0]?.[0] ?? '').trim(), rows: values.length },
          })
          continue
        }

        const texts = b.kind === 'list' ? (b.items ?? []).filter((t) => t.trim()) : [b.text ?? '']
        if (!texts.length || !texts[0].trim()) {
          outcomes.push({ at: k, ok: false, reason: '内容为空' })
          continue
        }

        let allOk = true
        let failReason: string | undefined
        const written: string[] = []

        for (const t of texts) {
          const p = (cursor as Word.Paragraph).insertParagraph(t, Word.InsertLocation.after)
          if (b.kind === 'heading') p.styleBuiltIn = headingStyle(b.level)
          else if (b.kind === 'list') p.styleBuiltIn = 'ListParagraph'
          p.load('text')
          await ctx.sync()

          const actual = p.text ?? ''
          if (actual.trim() !== t.trim()) {
            allOk = false
            failReason = '插入后回读不一致,Word 可能改写了内容'
          } else {
            written.push(actual)
          }
          cursor = p
        }

        if (allOk) inserted++
        // 撤销比对用**回读到的实际文本**,不用入参:Word 会改空白与换行
        outcomes.push({ at: k, ok: allOk, reason: failReason, texts: allOk ? written : undefined })
      } catch (e) {
        diag(`block ${k} (${b.kind}) ERROR: ${(e as Error).message}`)
        outcomes.push({ at: k, ok: false, reason: (e as Error).message })
      }
    }
  })

  clearSnapshot()
  return { inserted, outcomes }
}
