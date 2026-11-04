import { useState } from 'react'
import type { UndoEntry } from '../lib/undo'

interface Props {
  stack: UndoEntry[]
  busy: boolean
  onUndo: () => void
}

const clock = (iso: string) => {
  const d = new Date(iso)
  if (Number.isNaN(d.getTime())) return ''
  return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`
}

export function UndoBar({ stack, busy, onUndo }: Props) {
  const [open, setOpen] = useState(false)

  if (!stack.length) return null

  const latest = stack[stack.length - 1]

  return (
    <div className="undo-bar">
      <button className="undo-toggle" onClick={() => setOpen((v) => !v)} title="查看撤销历史">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.9" strokeLinecap="round" strokeLinejoin="round">
          <path d="M9 14 4 9l5-5" />
          <path d="M4 9h9a7 7 0 0 1 0 14h-3" />
        </svg>
        <span className="undo-latest">
          {clock(latest.at)} {latest.label}
        </span>
      </button>

      <button className="undo-do" disabled={busy} onClick={onUndo}>
        {busy ? '撤销中…' : '撤销'}
      </button>

      {open && stack.length > 1 && (
        <ul className="undo-list">
          {[...stack].reverse().map((e, i) => (
            <li key={e.id} className={i === 0 ? 'now' : ''}>
              <span className="undo-time">{clock(e.at)}</span>
              <span className="undo-label">{e.label}</span>
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}
