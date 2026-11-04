import { useEffect, useMemo, useState } from 'react'
import { diffChars, diffWords } from 'diff'
import type { Anchor, Basis, CommentBundle, CommentItem, EditItem, InsertBlock, InsertBundle, Proposal } from '../lib/pi'
import type { RefMeta } from '../lib/refs'

export type DraftItem =
  | { kind: 'proposal'; proposal: Proposal }
  | { kind: 'comments'; bundle: CommentBundle }
  | { kind: 'insert'; bundle: InsertBundle }

interface Props {
  item: DraftItem
  refs: RefMeta[]
  busy: boolean
  onApplyEdits: (chosen: EditItem[]) => void
  onApplyComments: (chosen: CommentItem[]) => void
  onApplyInsert: (chosen: InsertBlock[]) => void
  /** 同一份改动换算成批注(不改正文) */
  onCommentEdits: (chosen: EditItem[]) => void
  /** 同一份改动插入到锚点段落下方(保留原文) */
  onInsertEdits: (chosen: EditItem[]) => void
  onSkipAll: () => void
  onLocate: (paragraphIndex: number, find: string | null) => void
  onCopy: (text: string) => void
}

const isCJK = (s: string) => /[\u4e00-\u9fff]/.test(s)

const units = (s: string) => {
  const cjk = (s.match(/[\u4e00-\u9fff]/g) ?? []).length
  const words = (s.replace(/[\u4e00-\u9fff]/g, ' ').trim().match(/[A-Za-z0-9][A-Za-z0-9'-]*/g) ?? []).length
  return cjk + words
}

/** 「-4w / +4w」那种改动量统计 */
function diffStat(pairs: { before: string; after: string }[]) {
  let removed = 0
  let added = 0
  let cjk = false
  for (const { before, after } of pairs) {
    if (isCJK(before + after)) cjk = true
    const parts = isCJK(before + after) ? diffChars(before, after) : diffWords(before, after)
    for (const p of parts) {
      if (p.added) added += units(p.value)
      else if (p.removed) removed += units(p.value)
    }
  }
  if (!removed && !added) return ''
  return `${removed ? `-${removed}` : '±0'}${cjk ? '字' : 'w'} / +${added}${cjk ? '字' : 'w'}`
}

function FocusLine({ text, find }: { text: string; find: string | null }) {
  if (!text.trim()) return null

  const RADIUS = 26
  const at = find ? text.indexOf(find) : -1

  if (!find || at < 0) {
    const clipped = text.replace(/\s+/g, ' ').trim()
    return <div className="focus-line">{clipped.length > 90 ? `${clipped.slice(0, 90)}…` : clipped}</div>
  }

  const head = text.slice(Math.max(0, at - RADIUS), at)
  const tail = text.slice(at + find.length, at + find.length + RADIUS)

  return (
    <div className="focus-line">
      {at > RADIUS && <span className="focus-ell">…</span>}
      {head}
      <mark className="focus-hit">{find}</mark>
      {tail}
      {at + find.length + RADIUS < text.length && <span className="focus-ell">…</span>}
    </div>
  )
}

function DiffText({ before, after }: { before: string; after: string }) {
  const [expanded, setExpanded] = useState(false)
  const parts = useMemo(
    () => (isCJK(before + after) ? diffChars(before, after) : diffWords(before, after)),
    [before, after],
  )
  const tooLong = before.length + after.length > 400

  const changed = parts.reduce((n, p) => n + (p.added || p.removed ? p.value.length : 0), 0)
  const total = before.length + after.length
  const heavy = total > 120 && changed / total > 0.5

  if (heavy) {
    return (
      <>
        <div className={`diff-pair ${tooLong && !expanded ? 'clamped' : ''}`}>
          <div className="pair-before">{before}</div>
          <div className="pair-after">{after}</div>
        </div>
        {tooLong && (
          <button className="edit-expand" onClick={() => setExpanded((v) => !v)}>
            {expanded ? '收起' : `展开全文 · ${before.length} → ${after.length} 字`}
          </button>
        )}
      </>
    )
  }

  return (
    <>
      <div className={`edit-diff ${tooLong && !expanded ? 'clamped' : ''}`}>
        <p>
          {parts.map((p, i) =>
            p.added ? (
              <span key={i} className="ins">
                {p.value}
              </span>
            ) : p.removed ? (
              <span key={i} className="del">
                {p.value}
              </span>
            ) : (
              <span key={i}>{p.value}</span>
            ),
          )}
        </p>
      </div>
      {tooLong && (
        <button className="edit-expand" onClick={() => setExpanded((v) => !v)}>
          {expanded ? '收起' : `展开全文 · ${before.length} → ${after.length} 字`}
        </button>
      )}
    </>
  )
}

const extOf = (name: string) => (name.split('.').pop() ?? '').toLowerCase().slice(0, 4)

/** 「引用自 ref.pdf p3」依据链 */
function BasisRow({ basis, refs }: { basis?: Basis[]; refs: RefMeta[] }) {
  if (!basis?.length) return null

  return (
    <div className="attach-basis">
      <span>引用自</span>
      {basis.map((b, i) => {
        const known = refs.find((r) => r.name === b.name)
        const ext = extOf(b.name)
        const kind = ext === 'pdf' ? 'pdf' : ['png', 'jpg', 'jpeg', 'webp'].includes(ext) ? 'img' : ''
        return (
          <span key={`${b.name}-${i}`} className="attach-pill" title={b.note ?? b.name}>
            <span className={`attach-ext ${kind}`}>{ext.toUpperCase() || 'FILE'}</span>
            <span className="attach-name">{b.name}</span>
            {b.page ? <span className="attach-meta">p{b.page}</span> : null}
            {!b.page && known?.pages ? <span className="attach-meta">{known.pages}p</span> : null}
            {known ? <span className="attach-ok">✓</span> : null}
          </span>
        )
      })}
    </div>
  )
}

const locateLabel = (paragraphIndex: number | null, anchor?: Anchor) => {
  if (paragraphIndex !== null) return `第 ${paragraphIndex + 1} 段`
  return anchor?.text ? '当前选区' : '选区'
}

export function DraftCard({
  item,
  refs,
  busy,
  onApplyEdits,
  onApplyComments,
  onApplyInsert,
  onCommentEdits,
  onInsertEdits,
  onSkipAll,
  onLocate,
  onCopy,
}: Props) {
  const rows =
    item.kind === 'proposal'
      ? item.proposal.edits
      : item.kind === 'comments'
        ? item.bundle.comments
        : item.bundle.blocks

  const [selected, setSelected] = useState<Set<string>>(() => new Set(rows.map((r) => r.id)))

  // 应用后变 applied、撤销后回 pending,选中态都要留着。
  // 依赖用 id 串而不是 rows:后者每次 render 都是新数组,会死循环。
  const idKey = rows.map((r) => r.id).join(',')
  useEffect(() => {
    setSelected((prev) => {
      const next = new Set<string>()
      for (const r of rows) if (prev.has(r.id)) next.add(r.id)
      return next.size === prev.size ? prev : next
    })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [idKey])

  const pending = rows.filter((r) => r.status === 'pending')
  const chosenIds = pending.filter((r) => selected.has(r.id)).map((r) => r.id)
  const toggle = (id: string) =>
    setSelected((prev) => {
      const next = new Set(prev)
      if (next.has(id)) next.delete(id)
      else next.add(id)
      return next
    })
  const editParagraphs = [...new Set(item.kind === 'proposal' ? item.proposal.edits.map((e) => e.paragraphIndex).filter((i): i is number => i !== null) : [])]
  const head =
    item.kind === 'proposal'
      ? {
          title: `就地改写 · ${
            editParagraphs.length === 0
              ? '当前选区'
              : editParagraphs.length === 1
                ? `第 ${editParagraphs[0] + 1} 段`
                : `${editParagraphs.length} 段 · ${item.proposal.edits.length} 处`
          }`,
          sub: item.proposal.summary,
          count: `${item.proposal.edits.length} 处`,
          stat: diffStat(
            item.proposal.edits.map((e) => ({ before: e.find ?? item.proposal.anchor?.text ?? '', after: e.replace })),
          ),
        }
      : item.kind === 'comments'
        ? {
            title: '批注审阅',
            sub: item.bundle.summary,
            count: `${item.bundle.comments.length} 条`,
            stat: `${item.bundle.comments.length} 条`,
          }
        : {
            title: '结构化插入',
            sub: item.bundle.summary,
            count: `${item.bundle.blocks.length} 块`,
            stat: `${item.bundle.blocks.length} 块`,
          }

  const done = rows.filter((r) => r.status === 'applied' || r.status === 'inserted' || r.status === 'commented').length
  const failed = rows.filter((r) => r.status === 'failed').length
  const basis = item.kind === 'proposal' ? item.proposal.basis : item.kind === 'comments' ? item.bundle.basis : undefined

  const chosenEdits = () =>
    item.kind === 'proposal' ? item.proposal.edits.filter((e) => chosenIds.includes(e.id)) : []
  const chosenComments = () =>
    item.kind === 'comments' ? item.bundle.comments.filter((c) => chosenIds.includes(c.id)) : []
  const chosenBlocks = () => (item.kind === 'insert' ? item.bundle.blocks.filter((b) => chosenIds.includes(b.id)) : [])

  const copyText = () => {
    const text =
      item.kind === 'proposal'
        ? chosenEdits()
            .map((e) => e.replace)
            .join('\n\n') || item.proposal.edits.map((e) => e.replace).join('\n\n')
        : item.kind === 'comments'
          ? item.bundle.comments.map((c) => c.text).join('\n')
          : item.bundle.blocks
              .map((b) => b.text ?? b.items?.join('\n') ?? b.rows?.map((r) => r.join(' | ')).join('\n') ?? '')
              .join('\n\n')
    onCopy(text)
  }

  return (
    <div className="artifact-card">
      <div className="art-head">
        <div className="art-title-wrap">
          <span className="art-title">{head.title}</span>
          <span className="art-sub">{head.sub}</span>
        </div>
        <div className="art-head-right">
          <span className="art-stat">{head.stat || head.count}</span>
          <button className="btn ghost icon" disabled={busy} onClick={copyText} title="复制新文本">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round">
              <rect x="9" y="9" width="11" height="11" rx="2" />
              <path d="M15 5H6a2 2 0 0 0-2 2v9" />
            </svg>
          </button>
        </div>
      </div>

      <ul className="art-list">
        {item.kind === 'proposal' &&
          item.proposal.edits.map((e) => {
            const before = e.find ?? (e.paragraphIndex === null ? item.proposal.anchor?.text ?? item.proposal.selectionText : '')
            return (
              <li key={e.id} className={`art-row ${e.status}`}>
                <input
                  type="checkbox"
                  checked={selected.has(e.id)}
                  disabled={busy || e.status !== 'pending'}
                  onChange={() => toggle(e.id)}
                />
                <div className="art-main">
                  <div className="art-meta">
                    {e.paragraphIndex !== null ? (
                      <button className="loc-chip" title="在文档中定位" onClick={() => onLocate(e.paragraphIndex!, e.find)}>
                        {locateLabel(e.paragraphIndex, item.proposal.anchor)}
                      </button>
                    ) : (
                      <span className="loc-chip static">{locateLabel(null, item.proposal.anchor)}</span>
                    )}
                    {e.reason && <span className="art-reason">{e.reason}</span>}
                    <StatusBadge status={e.status} />
                  </div>
                  {e.paragraphIndex !== null && item.proposal.context?.[e.paragraphIndex] && (
                    <FocusLine text={item.proposal.context[e.paragraphIndex]} find={e.find} />
                  )}
                  {before ? (
                    <DiffText before={before} after={e.replace} />
                  ) : (
                    <div className="edit-diff">
                      <p className="plain">{e.replace}</p>
                    </div>
                  )}
                  {e.error && <div className="edit-err">{e.error}</div>}
                </div>
              </li>
            )
          })}

        {item.kind === 'comments' &&
          item.bundle.comments.map((c) => (
            <li key={c.id} className={`art-row ${c.status}`}>
              <input
                type="checkbox"
                checked={selected.has(c.id)}
                disabled={busy || c.status !== 'pending'}
                onChange={() => toggle(c.id)}
              />
              <div className="art-main">
                <div className="art-meta">
                  {c.paragraphIndex !== null ? (
                    <button className="loc-chip" onClick={() => onLocate(c.paragraphIndex!, c.find)}>
                      {locateLabel(c.paragraphIndex, item.bundle.anchor)}
                    </button>
                  ) : (
                    <span className="loc-chip static">{locateLabel(null, item.bundle.anchor)}</span>
                  )}
                  {c.severity === 'warn' && <span className="art-badge warn">需注意</span>}
                  <StatusBadge status={c.status} />
                </div>
                <div className="comment-text">{c.text}</div>
                {c.find && (
                  <div className="comment-quote" title={c.find}>
                    「{c.find.length > 60 ? `${c.find.slice(0, 60)}…` : c.find}」
                  </div>
                )}
                {c.error && <div className="edit-err">{c.error}</div>}
              </div>
            </li>
          ))}

        {item.kind === 'insert' &&
          item.bundle.blocks.map((b) => (
            <li key={b.id} className={`art-row ${b.status}`}>
              <input type="checkbox" checked={selected.has(b.id)} disabled={busy} onChange={() => toggle(b.id)} />
              <div className="art-main">
                <div className="art-meta">
                  <span className="loc-chip static">{blockLabel(b)}</span>
                  <StatusBadge status={b.status} />
                </div>
                <div className="insert-preview">
                  <BlockPreview b={b} />
                </div>
                {b.error && <div className="edit-err">{b.error}</div>}
              </div>
            </li>
          ))}
      </ul>

      <BasisRow basis={basis} refs={refs} />

      <div className="art-actions">
        {pending.length === 0 ? (
          <span className="art-state">
            已处理{done > 0 ? ` · 写入 ${done} 处` : ''}
            {failed > 0 ? ` · 失败 ${failed} 处` : ''}
          </span>
        ) : (
          <>
            {item.kind === 'proposal' && (
              <>
                <button className="btn clay" disabled={busy || !chosenIds.length} onClick={() => onApplyEdits(chosenEdits())}>
                  {busy ? '写入中…' : `✓ 替换 ${chosenIds.length} 处`}
                </button>
                <button
                  className="btn"
                  disabled={busy || !chosenIds.length}
                  onClick={() => onCommentEdits(chosenEdits())}
                  title="不修改正文,把改动建议写成 Word 批注"
                >
                  批注
                </button>
                <button
                  className="btn"
                  disabled={busy || !chosenIds.length}
                  onClick={() => onInsertEdits(chosenEdits())}
                  title="保留原文,把新写法插入到这段下方"
                >
                  插入下方
                </button>
              </>
            )}
            {item.kind === 'comments' && (
              <button className="btn clay" disabled={busy || !chosenIds.length} onClick={() => onApplyComments(chosenComments())}>
                {busy ? '写入中…' : `✓ 写入 ${chosenIds.length} 条批注`}
              </button>
            )}
            {item.kind === 'insert' && (
              <button className="btn clay" disabled={busy || !chosenIds.length} onClick={() => onApplyInsert(chosenBlocks())}>
                {busy ? '插入中…' : `✓ 插入 ${chosenIds.length} 块`}
              </button>
            )}
            <div className="grow" />
            <button className="btn" disabled={busy} onClick={onSkipAll} title="这份草稿一条都不执行">
              放弃
            </button>
          </>
        )}
      </div>
    </div>
  )
}

function StatusBadge({ status }: { status: string }) {
  if (status === 'applied') return <span className="art-badge ok">已写入</span>
  if (status === 'inserted') return <span className="art-badge ok">已插入</span>
  if (status === 'commented') return <span className="art-badge ok">已批注</span>
  if (status === 'skipped') return <span className="art-badge">已跳过</span>
  if (status === 'failed') return <span className="art-badge err">失败</span>
  return null
}

const blockLabel = (b: InsertBlock) => {
  if (b.kind === 'heading') return `标题 ${b.level ?? 2} 级`
  if (b.kind === 'paragraph') return '段落'
  if (b.kind === 'list') return `列表 ${b.items?.length ?? 0} 项`
  return `表格 ${b.rows?.length ?? 0}×${b.rows?.[0]?.length ?? 0}`
}

function BlockPreview({ b }: { b: InsertBlock }) {
  if (b.kind === 'heading') {
    const level = Math.min(4, Math.max(1, b.level ?? 2))
    const Tag = `h${level}` as 'h1' | 'h2' | 'h3' | 'h4'
    return <Tag className={`pv-heading lv${level}`}>{b.text}</Tag>
  }

  if (b.kind === 'paragraph') return <p className="pv-para">{b.text}</p>

  if (b.kind === 'list') {
    return (
      <ul className="pv-list">
        {(b.items ?? []).map((t, i) => (
          <li key={i}>{t}</li>
        ))}
      </ul>
    )
  }

  const rows = b.rows ?? []
  return (
    <div className="pv-table-wrap">
      <table className="pv-table">
        <tbody>
          {rows.map((r, i) => (
            <tr key={i}>
              {r.map((c, j) => (i === 0 ? <th key={j}>{c}</th> : <td key={j}>{c}</td>))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
