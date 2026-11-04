import { useEffect, useState } from 'react'
import { UI_REV, resetPiSession } from '../lib/pi'
import { DEFAULT_CMDS, normalizeSlug, validateCmds, type QuickCmd } from '../lib/commands'

interface Props {
  onClose: () => void
  cmds: QuickCmd[]
  onChangeCmds: (list: QuickCmd[]) => void
}

const EMPTY: QuickCmd = { cmd: '/', title: '', sub: '', text: '' }

export function Settings({ onClose, cmds, onChangeCmds }: Props) {
  const [model, setModel] = useState('pi')
  const [provider, setProvider] = useState('')
  const [running, setRunning] = useState(false)
  const [msg, setMsg] = useState<{ ok: boolean; text: string } | null>(null)

  const [editing, setEditing] = useState<number | 'new' | null>(null)
  const [draft, setDraft] = useState<QuickCmd>(EMPTY)
  const [err, setErr] = useState('')
  const [confirmReset, setConfirmReset] = useState(false)

  useEffect(() => {
    fetch('/api/agent/model')
      .then((r) => r.json())
      .then((j) => {
        setModel(j.model ?? 'pi')
        setProvider(j.provider ?? '')
      })
      .catch(() => {})
    fetch('/api/agent/info')
      .then((r) => r.json())
      .then((j) => setRunning(Boolean(j.running)))
      .catch(() => {})
  }, [])

  async function doReset() {
    try {
      await resetPiSession()
      setMsg({ ok: true, text: '会话已重置,下一条消息将重新建立上下文' })
    } catch (e) {
      setMsg({ ok: false, text: (e as Error).message })
    }
  }

  function startEdit(i: number) {
    setEditing(i)
    setDraft({ ...cmds[i] })
    setErr('')
  }

  function startNew() {
    setEditing('new')
    setDraft({ ...EMPTY })
    setErr('')
  }

  function commit() {
    const next: QuickCmd = {
      cmd: normalizeSlug(draft.cmd),
      title: draft.title.trim(),
      sub: draft.sub.trim(),
      text: draft.text.trim(),
    }
    const others = editing === 'new' || editing === null ? cmds : cmds.filter((_, i) => i !== editing)
    const problem = validateCmds([...others, next])
    if (problem) {
      setErr(problem)
      return
    }
    if (editing === 'new') onChangeCmds([...cmds, next])
    else if (editing !== null) onChangeCmds(cmds.map((c, i) => (i === editing ? next : c)))
    setEditing(null)
    setErr('')
  }

  function remove(i: number) {
    onChangeCmds(cmds.filter((_, x) => x !== i))
    if (editing === i) setEditing(null)
  }

  function restoreDefaults() {
    onChangeCmds(DEFAULT_CMDS.map((c) => ({ ...c })))
    setEditing(null)
    setErr('')
    setConfirmReset(false)
  }

  const form = (
    <div className="cmdf">
      <div className="cmdf-row">
        <input
          className="cmdf-slug"
          value={draft.cmd}
          placeholder="/fast"
          spellCheck={false}
          onChange={(e) => setDraft({ ...draft, cmd: e.target.value })}
        />
        <input
          className="grow"
          value={draft.title}
          placeholder="标题,如 润色得更正式"
          onChange={(e) => setDraft({ ...draft, title: e.target.value })}
        />
      </div>
      <input
        value={draft.sub}
        placeholder="副标题(可空),如 学术语态"
        onChange={(e) => setDraft({ ...draft, sub: e.target.value })}
      />
      <textarea
        rows={3}
        value={draft.text}
        placeholder="点一下会填进输入框的完整指令…"
        onChange={(e) => setDraft({ ...draft, text: e.target.value })}
      />
      {err && <div className="cmdf-err">{err}</div>}
      <div className="cmdf-row">
        <button className="btn sm clay" onClick={commit}>
          保存
        </button>
        <button
          className="btn sm ghost"
          onClick={() => {
            setEditing(null)
            setErr('')
          }}
        >
          取消
        </button>
      </div>
    </div>
  )

  return (
    <div className="sheet">
      <header className="sheet-head">
        <button className="mini" onClick={onClose} title="返回">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M15 18l-6-6 6-6" />
          </svg>
        </button>
        <span className="sheet-title">设置</span>
      </header>

      <div className="sheet-body">
        <div className="info-card">
          <div className="info-row">
            <span className="info-k">内核</span>
            <span className="info-v">Pi · 本地 Node 服务</span>
          </div>
          <div className="info-row">
            <span className="info-k">模型</span>
            <span className="info-v">
              {model}
              {provider ? <em className="info-sub">{provider}</em> : null}
            </span>
          </div>
          <div className="info-row">
            <span className="info-k">状态</span>
            <span className="info-v">{running ? '任务运行中' : '空闲'}</span>
          </div>
          <div className="info-row">
            <span className="info-k">界面</span>
            <span className="info-v">{UI_REV}</span>
          </div>
        </div>

        <div className="set-sec">
          <div className="set-sec-head">
            <span>快捷指令</span>
            <em>输入框打 / 触发,点一下只填这一条</em>
          </div>

          {cmds.length === 0 && editing !== 'new' && (
            <div className="set-empty">一条都没有了,点「恢复默认」拿回内置的 6 条。</div>
          )}

          {cmds.map((c, i) =>
            editing === i ? (
              <div key={`edit-${c.cmd}`}>{form}</div>
            ) : (
              <div key={c.cmd} className="cmd-row">
                <span className="cmd-tag">{c.cmd}</span>
                <span className="cmd-row-name">{c.title}</span>
                <span className="cmd-row-sub">{c.sub}</span>
                <button className="btn sm ghost" onClick={() => startEdit(i)} title="编辑">
                  改
                </button>
                <button className="cmd-del" onClick={() => remove(i)} title={`删除 ${c.cmd}`}>
                  ×
                </button>
              </div>
            ),
          )}

          {editing === 'new' && form}

          <div className="set-sec-foot">
            {editing === null && (
              <button className="btn sm" onClick={startNew}>
                + 新增指令
              </button>
            )}
            <div className="grow" />
            {confirmReset ? (
              <>
                <button className="btn sm ghost" onClick={() => setConfirmReset(false)}>
                  取消
                </button>
                <button className="btn sm clay" onClick={restoreDefaults}>
                  确认覆盖
                </button>
              </>
            ) : (
              <button className="btn sm ghost" onClick={() => setConfirmReset(true)} title="用内置的 6 条替换当前列表">
                恢复默认
              </button>
            )}
          </div>
        </div>

        <div className="sheet-note">
          模型与凭证由 Pi 自己管理。模型可在输入坞切换,只对 WordAI 生效(存
          <code>~/.wordai/model.json</code>),不改 Pi 的默认;文档内容通过 Office.js 在本机读取,
          模型只会收到你实际读过的段落。
        </div>

        {msg && <div className={`hint ${msg.ok ? 'ok' : 'err'}`}>{msg.text}</div>}
      </div>

      <footer className="sheet-foot">
        <button onClick={() => void doReset()}>重置会话</button>
        <div className="grow" />
        <button className="solid" onClick={onClose}>
          完成
        </button>
      </footer>
    </div>
  )
}
