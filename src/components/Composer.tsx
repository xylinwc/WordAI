import { useRef, useState } from 'react'
import { modelKey, type ModelInfo } from '../lib/models'
import type { QuickCmd, Rule } from '../lib/commands'
import type { RefMeta } from '../lib/refs'

interface Props {
  value: string
  onChange: (v: string) => void
  onSend: () => void
  onStop: () => void
  busy: boolean
  picking: boolean
  refs: RefMeta[]
  selChars: number
  selIgnored: boolean
  onToggleSelection: () => void
  onAddRefs: () => void
  onRemoveRef: (r: RefMeta) => void
  cmds: QuickCmd[]
  models: ModelInfo[]
  currentModel: ModelInfo | null
  modelsLoading: boolean
  modelsError: string
  modelBusy: boolean
  onSelectModel: (m: ModelInfo) => void
  rules: Rule[]
  onAddRule: (text: string) => void
  onRemoveRule: (id: string) => void
}

const extOf = (r: RefMeta) => {
  const k = (r.kind || r.name.split('.').pop() || '').toLowerCase()
  return k === 'markdown' ? 'md' : k.slice(0, 4)
}

const kindClass = (ext: string) => (ext === 'pdf' ? 'pdf' : ['png', 'jpg', 'jpeg', 'webp', 'gif'].includes(ext) ? 'img' : '')

/**
 * 输入坞:附件胶囊 + 上下文胶囊 + `/`指令与 `#`规则浮层。
 * 发送键在忙碌时变中止键(Esc 同效)。
 */
export function Composer({
  value,
  onChange,
  onSend,
  onStop,
  busy,
  picking,
  refs,
  selChars,
  selIgnored,
  onToggleSelection,
  onAddRefs,
  onRemoveRef,
  cmds,
  models,
  currentModel,
  modelsLoading,
  modelsError,
  modelBusy,
  onSelectModel,
  rules,
  onAddRule,
  onRemoveRule,
}: Props) {
  const [menu, setMenu] = useState<{ kind: 'cmd' | 'rule'; q: string } | null>(null)
  const [ruleDraft, setRuleDraft] = useState('')
  const taRef = useRef<HTMLTextAreaElement | null>(null)

  // 按 provider 分组;optgroup 能让一长串模型还能扫读
  const modelGroups = new Map<string, ModelInfo[]>()
  for (const m of models) {
    const list = modelGroups.get(m.provider) ?? []
    list.push(m)
    modelGroups.set(m.provider, list)
  }
  const currentKey = currentModel ? modelKey(currentModel) : ''
  const currentListed = currentModel ? models.some((m) => modelKey(m) === currentKey) : false

  function update(next: string) {
    onChange(next)

    // 输入 / 或 # 时弹浮层;只在开头或空格之后触发,免得路径、话题标签误伤
    const cmd = /(?:^|\s)\/([\w-]*)$/.exec(next)
    if (cmd) {
      setMenu({ kind: 'cmd', q: cmd[1] })
      return
    }
    if (/(?:^|\s)#$/.test(next)) {
      setMenu({ kind: 'rule', q: '' })
      return
    }
    setMenu(null)
  }

  function pickCmd(text: string) {
    const next = value.replace(/(^|\s)\/[\w-]*$/, (_m, p1: string) => `${p1}${text}`)
    onChange(next)
    setMenu(null)
    taRef.current?.focus()
  }

  function commitRule() {
    const text = ruleDraft.trim()
    if (text) onAddRule(text)
    setRuleDraft('')
    onChange(value.replace(/(^|\s)#$/, '$1'))
    setMenu(null)
    taRef.current?.focus()
  }

  const cmdList = cmds.filter((c) => {
    const q = menu?.kind === 'cmd' ? menu.q.toLowerCase() : ''
    return !q || c.cmd.slice(1).includes(q) || c.title.includes(q)
  })

  return (
    <div className="dock">
      <div className="session-row">
        {selChars > 0 || selIgnored ? (
          <button
            className={`ctx-pill ${selIgnored ? '' : 'clay'}`}
            onClick={onToggleSelection}
            title={selIgnored ? '点一下恢复:把选区一起发给 AI' : '点一下忽略选区,只按指令处理全文'}
          >
            {selIgnored ? '已忽略选区' : `已选 ${selChars} 字`} ×
          </button>
        ) : (
          <span className="ctx-pill">未选中文字 · 可读全文</span>
        )}
        {refs.length > 0 && (
          <span className="ctx-pill sand" title="已附加的附件">
            附件 ×{refs.length}
          </span>
        )}
      </div>

      <div className="composer">
        {refs.length > 0 && (
          <div className="attach-row">
            {refs.map((r) => {
              const ext = extOf(r)
              return (
                <span key={r.id} className="attach-pill">
                  <span className={`attach-ext ${kindClass(ext)}`}>{ext.toUpperCase()}</span>
                  <span className="attach-name">{r.name}</span>
                  <span className="attach-meta">{r.pages ? `${r.pages}p` : `${r.chunks} 段`}</span>
                  <span className="attach-ok">✓</span>
                  <button className="attach-del" disabled={busy || picking} onClick={() => onRemoveRef(r)} title="移除这份附件">
                    ×
                  </button>
                </span>
              )
            })}
          </div>
        )}

        <textarea
          ref={taRef}
          value={value}
          title="/ 指令   # 规则"
          rows={3}
          placeholder="请求修改、描述风格,或直接说改哪一段…"
          onChange={(e) => update(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Escape' && menu) {
              e.preventDefault()
              setMenu(null)
              return
            }
            if (e.key === 'Enter' && !e.shiftKey) {
              e.preventDefault()
              // 浮层开着时 Enter 只用来选指令,不要发半截话
              if (menu?.kind === 'cmd' && cmdList.length) {
                pickCmd(cmdList[0].text)
                return
              }
              if (menu) return
              if (!busy) onSend()
            }
          }}
        />

        <div className="composer-foot">
          <div className="foot-left">
            <span
              className={`preset-select model-picker ${models.length ? '' : 'off'}`}
              title={
                currentModel
                  ? `${currentModel.provider} / ${currentModel.id}`
                  : modelsError || (modelsLoading ? '正在读取模型列表…' : '没有可用模型')
              }
            >
              <select
                value={currentKey}
                disabled={busy || picking || modelBusy || modelsLoading || models.length === 0}
                onChange={(e) => {
                  const m = models.find((x) => modelKey(x) === e.target.value)
                  if (m) onSelectModel(m)
                }}
                aria-label="选择模型"
              >
                {!currentModel && (
                  <option value="">{modelsLoading ? '加载模型…' : modelsError ? '模型不可用' : '选择模型'}</option>
                )}
                {currentModel && !currentListed && (
                  <option value={currentKey}>{currentModel.name || currentModel.id}</option>
                )}
                {[...modelGroups].map(([provider, list]) => (
                  <optgroup key={provider} label={provider}>
                    {list.map((m) => (
                      <option key={modelKey(m)} value={modelKey(m)} title={`${m.provider}/${m.id}`}>
                        {m.name || m.id}
                      </option>
                    ))}
                  </optgroup>
                ))}
              </select>
              <span className="preset-caret">▾</span>
            </span>

            {rules.length > 0 && (
              <span className="preset-tag" title={rules.map((r) => `# ${r.text}`).join('\n')}>
                #{rules.length} 规则
              </span>
            )}
          </div>

          <div className="foot-right">
            <button
              className="attach-btn"
              disabled={busy || picking}
              onClick={onAddRefs}
              title="附加附件(PDF / DOCX / TXT / MD)"
              aria-label="附加附件"
            >
              {picking ? (
                <span className="spinner" />
              ) : (
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M21.4 11.1l-8.5 8.5a5.5 5.5 0 0 1-7.8-7.8l8.5-8.5a3.5 3.5 0 0 1 4.9 4.9l-8.5 8.5a1.5 1.5 0 0 1-2.1-2.1l7.8-7.8" />
                </svg>
              )}
            </button>

            <button
              className={`send-btn ${busy ? 'stop' : ''}`}
              disabled={!busy && !value.trim()}
              onClick={() => (busy ? onStop() : onSend())}
              title={busy ? '中止生成(Esc)' : '发送(Enter)'}
            >
              {busy ? (
                <svg width="12" height="12" viewBox="0 0 24 24">
                  <rect x="6" y="6" width="12" height="12" rx="2" fill="currentColor" />
                </svg>
              ) : (
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M12 19V5M6 11l6-6 6 6" />
                </svg>
              )}
            </button>
          </div>
        </div>
      </div>

      {menu?.kind === 'cmd' && (
        <div className="popover cmd-popover">
          <div className="pop-sec">指令 · 点击填入输入框</div>
          {cmdList.length === 0 ? (
            <div className="pop-empty">没有匹配的指令</div>
          ) : (
            cmdList.map((c) => (
              <button key={c.cmd} className="pop-item" onClick={() => pickCmd(c.text)}>
                <span className="pop-left">
                  <span className="cmd-tag">{c.cmd}</span>
                  <span className="pop-name">{c.title}</span>
                </span>
                <span className="pop-meta">{c.sub}</span>
              </button>
            ))
          )}
        </div>
      )}

      {menu?.kind === 'rule' && (
        <div className="popover cmd-popover">
          <div className="pop-sec">写作规则 · 每轮都会拼进提示词</div>
          {rules.map((r) => (
            <div key={r.id} className="pop-rule">
              <span className="pop-prefix">#rule</span>
              <span className="pop-name">{r.text}</span>
              <button className="attach-del" onClick={() => onRemoveRule(r.id)} title="删除这条规则">
                ×
              </button>
            </div>
          ))}
          <div className="pop-input">
            <input
              autoFocus
              value={ruleDraft}
              placeholder="如:保持基因符号斜体(Klf4、Oct4)"
              onChange={(e) => setRuleDraft(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.preventDefault()
                  commitRule()
                }
                if (e.key === 'Escape') setMenu(null)
              }}
            />
            <button className="btn sm clay" onClick={commitRule} disabled={!ruleDraft.trim()}>
              添加
            </button>
          </div>
        </div>
      )}
    </div>
  )
}
