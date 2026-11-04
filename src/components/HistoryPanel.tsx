import { useState } from 'react'
import type { SessionMeta } from '../lib/sessions'

interface Props {
  sessions: SessionMeta[]
  activeId: string | null
  busy: boolean
  onClose: () => void
  onOpen: (id: string) => void
  onDelete: (ids: string[]) => void
}

const fmt = (iso: string) => {
  const d = new Date(iso)
  if (Number.isNaN(d.getTime())) return ''
  const p = (n: number) => String(n).padStart(2, '0')
  return `${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`
}

/**
 * 历史会话抽屉。
 * 当前会话不计入历史,也不能在这里删;其余按最近更新倒序,支持多选批量删。
 */
export function HistoryPanel({ sessions, activeId, busy, onClose, onOpen, onDelete }: Props) {
  const [selected, setSelected] = useState<string[]>([])
  const [confirming, setConfirming] = useState(false)

  const active = sessions.find((s) => s.id === activeId) ?? null
  const history = sessions
    .filter((s) => s.id !== activeId)
    .sort((a, b) => (b.updatedAt || b.createdAt).localeCompare(a.updatedAt || a.createdAt))

  const toggle = (id: string) =>
    setSelected((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]))

  // 当前会话也可以勾选删除;删它时服务端会把 activeId 换成剩下的第一个
  const selectable = active ? [active.id, ...history.map((s) => s.id)] : history.map((s) => s.id)
  const allOn = selectable.length > 0 && selected.length === selectable.length
  const deletingActive = Boolean(active && selected.includes(active.id))

  function doDelete() {
    if (!selected.length) return
    onDelete(selected)
    setSelected([])
    setConfirming(false)
  }

  return (
    <div className="sheet">
      <header className="sheet-head">
        <button className="mini" onClick={onClose} title="返回对话">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.9" strokeLinecap="round" strokeLinejoin="round">
            <path d="M15 5l-7 7 7 7" />
          </svg>
        </button>
        <span className="sheet-title">历史会话</span>
        <div className="grow" />
        <span className="hist-count">
          {history.length} 个{active ? ' · 1 个当前' : ''}
        </span>
      </header>

      <div className="sheet-body hist-body">
        {active && (
          <div className={`hist-row current ${selected.includes(active.id) ? 'on' : ''}`}>
            <button
              className={`hist-check ${selected.includes(active.id) ? 'on' : ''}`}
              disabled={busy}
              onClick={() => toggle(active.id)}
              title={selected.includes(active.id) ? '取消选中当前会话' : '选中当前会话'}
            >
              {selected.includes(active.id) && (
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M5 12.5l4.5 4.5L19 7.5" />
                </svg>
              )}
            </button>
            <div className="hist-main">
              <div className="hist-title" title={active.title}>
                {active.title}
              </div>
              <div className="hist-meta">当前会话 · 更新于 {fmt(active.updatedAt || active.createdAt)}</div>
            </div>
          </div>
        )}

        {history.length === 0 ? (
          <div className="hist-empty">
            还没有历史会话。
            <br />
            点右上角的「+」开新对话,当前这段就会归档到这里。
          </div>
        ) : (
          history.map((s) => {
            const on = selected.includes(s.id)
            return (
              <div
                key={s.id}
                className={`hist-row ${on ? 'on' : ''}`}
                onClick={() => !busy && onOpen(s.id)}
                role="button"
                title="点这一行切回该会话(原当前会话会归档)"
              >
                <button
                  className={`hist-check ${on ? 'on' : ''}`}
                  disabled={busy}
                  onClick={(e) => {
                    e.stopPropagation()
                    toggle(s.id)
                  }}
                  title={on ? '取消选中' : '选中'}
                >
                  {on && (
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round">
                      <path d="M5 12.5l4.5 4.5L19 7.5" />
                    </svg>
                  )}
                </button>
                <div className="hist-main">
                  <div className="hist-title" title={s.title}>
                    {s.title}
                  </div>
                  <div className="hist-meta">更新于 {fmt(s.updatedAt || s.createdAt)}</div>
                </div>
              </div>
            )
          })
        )}
      </div>

      <div className="sheet-foot">
        {confirming ? (
          <>
            <button className="text" onClick={() => setConfirming(false)}>
              取消
            </button>
            <div className="grow" />
            <span className="hist-warn">
              {deletingActive ? '含当前会话,删除后自动切到剩下的第一个' : '删除后 Pi 上下文、撤销栈与附件一起消失'}
            </span>
            <button className="solid danger" disabled={busy} onClick={doDelete}>
              确认删除 {selected.length} 个
            </button>
          </>
        ) : (
          <>
            <button
              className="text"
              disabled={busy || !selectable.length}
              onClick={() => setSelected(allOn ? [] : selectable)}
            >
              {allOn ? '取消全选' : '全选'}
            </button>
            <div className="grow" />
            <button
              className="danger-text"
              disabled={busy || !selected.length}
              onClick={() => setConfirming(true)}
            >
              删除选中{selected.length ? ` (${selected.length})` : ''}
            </button>
          </>
        )}
      </div>
    </div>
  )
}
