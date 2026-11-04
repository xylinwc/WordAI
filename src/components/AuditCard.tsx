import type { AuditBundle, AuditIssue, Basis } from '../lib/pi'
import type { RefMeta } from '../lib/refs'

interface Props {
  bundle: AuditBundle
  refs: RefMeta[]
  busy: boolean
  /** 把若干条审查项转成 Word 批注 */
  onComment: (issues: AuditIssue[]) => void
  onSkipAll: () => void
  onLocate: (paragraphIndex: number, find: string | null) => void
  /** 让 AI 针对这一条起草修复内容(再走一轮) */
  onDraftFix: (issue: AuditIssue) => void
}

const extOf = (name: string) => (name.split('.').pop() ?? '').toLowerCase().slice(0, 4)

function BasisRow({ basis, refs }: { basis?: Basis[]; refs: RefMeta[] }) {
  if (!basis?.length) return null
  return (
    <div className="attach-basis">
      <span>引用自</span>
      {basis.map((b, i) => {
        const ext = extOf(b.name)
        const kind = ext === 'pdf' ? 'pdf' : ['png', 'jpg', 'jpeg', 'webp'].includes(ext) ? 'img' : ''
        return (
          <span key={`${b.name}-${i}`} className="attach-pill" title={b.note ?? b.name}>
            <span className={`attach-ext ${kind}`}>{ext.toUpperCase() || 'FILE'}</span>
            <span className="attach-name">{b.name}</span>
            {b.page ? <span className="attach-meta">p{b.page}</span> : null}
            {refs.some((r) => r.name === b.name) ? <span className="attach-ok">✓</span> : null}
          </span>
        )
      })}
    </div>
  )
}

/**
 * 全文审查清单(视图 04)。
 * 逐项可跳转 / 可单条转批注;底部一键全部转批注,仍然走 applyComments 的逐条校验。
 */
export function AuditCard({ bundle, refs, busy, onComment, onSkipAll, onLocate, onDraftFix }: Props) {
  const pending = bundle.issues.filter((i) => i.status === 'pending')
  const done = bundle.issues.filter((i) => i.status === 'commented').length
  const failed = bundle.issues.filter((i) => i.status === 'failed').length

  return (
    <div className="audit-card">
      <div className="audit-head">
        <div className="audit-title-wrap">
          <span className="audit-title">全文结构与逻辑审查</span>
          <span className="audit-sub">{bundle.summary}</span>
        </div>
        <span className="audit-badge">
          {pending.length > 0 ? `${pending.length} 处待完善` : done > 0 ? `已成批注 ${done} 处` : '已处理'}
        </span>
      </div>

      {bundle.issues.map((issue) => (
        <div key={issue.id} className={`issue-card ${issue.severity === 'priority' ? 'priority' : ''} ${issue.status}`}>
          <div className="issue-top">
            <span className={`issue-cat ${issue.severity === 'priority' ? '' : 'secondary'}`}>
              {issue.severity === 'priority' ? '●' : '○'} {issue.category}
              {issue.paragraphIndex !== null ? ` · 第 ${issue.paragraphIndex + 1} 段` : ''}
            </span>
            {issue.paragraphIndex !== null ? (
              <button className="jump" onClick={() => onLocate(issue.paragraphIndex!, issue.find)}>
                跳转 ↗
              </button>
            ) : null}
          </div>

          <div className="issue-title">{issue.title}</div>
          <div className="issue-desc">{issue.detail}</div>
          {issue.suggestion && <div className="issue-desc suggestion">建议:{issue.suggestion}</div>}
          {issue.error && <div className="edit-err">{issue.error}</div>}

          {issue.status === 'pending' ? (
            <div className="issue-actions">
              <button
                className={`btn sm ${issue.severity === 'priority' ? 'clay' : ''}`}
                disabled={busy}
                onClick={() => onComment([issue])}
              >
                转批注
              </button>
              <button className="btn sm" disabled={busy} onClick={() => onDraftFix(issue)}>
                起草修复
              </button>
              {issue.severity === 'priority' ? <span className="issue-flag">优先</span> : null}
            </div>
          ) : (
            <div className="issue-actions">
              <span className={`art-badge ${issue.status === 'commented' ? 'ok' : issue.status === 'failed' ? 'err' : ''}`}>
                {issue.status === 'commented' ? '已成批注' : issue.status === 'failed' ? '失败' : '已跳过'}
              </span>
            </div>
          )}
        </div>
      ))}

      <BasisRow basis={bundle.basis} refs={refs} />

      <div className="audit-foot">
        {pending.length > 0 ? (
          <>
            <button className="btn clay block" disabled={busy} onClick={() => onComment(pending)}>
              {busy ? '写入中…' : `✓ 将全部 ${pending.length} 项审查建议转为 Word 批注`}
            </button>
            <button className="btn ghost" disabled={busy} onClick={onSkipAll} title="这份审查清单一条都不处理">
              放弃
            </button>
          </>
        ) : (
          <span className="art-state">
            已处理{done > 0 ? ` · 批注 ${done} 处` : ''}
            {failed > 0 ? ` · 失败 ${failed} 处` : ''}
          </span>
        )}
      </div>
    </div>
  )
}
