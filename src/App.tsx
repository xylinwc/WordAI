import { useCallback, useEffect, useRef, useState, type DragEvent } from 'react'
import {
  applyComments,
  applyInsert,
  captureAnchor,
  locateInDocument,
  readParagraphContexts,
  undoOps,
  applyDocEdits,
  applyReplacement,
  captureSelection,
  getDocKey,
  inWord,
  initOffice,
  onSelectionChanged,
} from './office'
import {
  resetPiSession,
  runPiAgent,
  type Anchor,
  type AuditBundle,
  type AuditIssue,
  type CommentBundle,
  type CommentItem,
  type EditItem,
  type InsertBlock,
  type InsertBundle,
  type Proposal,
} from './lib/pi'
import {
  createSession,
  deleteSessionsApi,
  fetchSessionUi,
  fetchSessions,
  fetchUndo,
  saveSessionUi,
  saveUndo,
  switchSessionApi,
  type SessionMeta,
} from './lib/sessions'
import { UNDO_LIMIT, type CardRevert, type UndoEntry, type UndoOp } from './lib/undo'
import { attachRefs, fetchRefs, removeRef, uploadRefs, type AttachResult, type RefMeta } from './lib/refs'
import { dropLegacyPreset, loadCmds, rulesPrompt, saveCmds, type QuickCmd, type Rule } from './lib/commands'
import { fetchModels, switchModel as switchModelApi, type ModelInfo } from './lib/models'
import { Composer } from './components/Composer'
import { AuditCard } from './components/AuditCard'
import { issueCommentText } from './lib/audit'
import { DraftCard } from './components/DraftCard'
import { HistoryPanel } from './components/HistoryPanel'
import { UndoBar } from './components/UndoBar'
import { Markdown } from './components/Markdown'
import { Settings } from './components/Settings'

type Item =
  | { id: string; kind: 'user'; text: string }
  | { id: string; kind: 'assistant'; text: string; tool?: string; durationMs?: number }
  | { id: string; kind: 'proposal'; proposal: Proposal }
  | { id: string; kind: 'comments'; bundle: CommentBundle }
  | { id: string; kind: 'insert'; bundle: InsertBundle }
  | { id: string; kind: 'audit'; bundle: AuditBundle }
  | { id: string; kind: 'error'; text: string }
  | { id: string; kind: 'notice'; text: string }

const uid = () => Math.random().toString(36).slice(2, 10)

/** 预览态的历史列表,只为 `?demo=1` 下能看面板样子 */
const DEMO_SESSIONS: SessionMeta[] = [
  { id: 'd1', title: '审阅引言', createdAt: '2026-01-01T09:12:00Z', updatedAt: '2026-01-02T10:20:00Z' },
  { id: 'd2', title: '统一术语「方法」→「干预策略」', createdAt: '2026-01-02T14:03:00Z', updatedAt: '2026-01-03T09:41:00Z' },
  { id: 'd3', title: '补方法学与统计描述', createdAt: '2026-01-03T16:20:00Z', updatedAt: '2026-01-04T11:05:00Z' },
]

/** 拖拽时可接受的附件格式(与 server/refs.ts 的抽取能力对齐) */
const REF_EXT = new Set(['pdf', 'docx', 'doc', 'rtf', 'rtfd', 'html', 'htm', 'txt', 'md', 'markdown', 'csv'])
const REF_MAX = 80 * 1024 * 1024

const DEMO: Item[] = [
  { id: 'd1', kind: 'user', text: '把「方法」统一叫「干预策略」,全文检查一遍' },
  {
    id: 'd2',
    kind: 'assistant',
    durationMs: 12800,
    text: '全文检索到 **3 处**使用「方法」的表述,已统一为「干预策略」。确认后点「应用选中」才会写入:',
  },
  {
    id: 'd3',
    kind: 'proposal',
    proposal: {
      summary: '统一术语「方法」→「干预策略」',
      selectionText: '',
      basis: [{ name: 'Raaz 2019 — Runx2 and aortic fibrosis.pdf', page: 3, note: '干预策略的术语口径' }],
      context: {
        11: 'Chronic kidney disease (CKD) is a global health threat, which increases the risk of cardiovascular disease (CVD) morbidity and mortality. In the meanwhile, cardiovascular rather than end-stage kidney disease is the leading cause of death in this high-risk population.',
        17: '在实验设计上,该方法的优势在于成本低、周期短,且不依赖特殊设备,因此适合在基层实验室推广。',
        30: '因此,本研究在方法学层面提出了一个可复现的框架,后续工作将围绕该框架展开验证。',
      },
      edits: [
        {
          id: 'e1',
          paragraphIndex: 11,
          find: 'Chronic kidney disease (CKD) is a global health threat, which increases the risk of cardiovascular disease (CVD) morbidity and mortality. In the meanwhile, cardiovascular rather than end-stage kidney disease (CKD stage 5) is the leading cause of death in this high-risk population.',
          replace: 'Chronic kidney disease (CKD) constitutes a global health burden and is associated with increased cardiovascular morbidity and mortality. Moreover, cardiovascular events, rather than end-stage kidney disease (CKD stage 5), represent the leading cause of death in this high-risk population.',
          reason: '术语统一',
          status: 'pending',
        },
        {
          id: 'e2',
          paragraphIndex: 17,
          find: '该方法的优势',
          replace: '该干预策略的优势',
          reason: '术语统一',
          status: 'pending',
        },
        {
          id: 'e3',
          paragraphIndex: 30,
          find: '方法学层面',
          replace: '干预策略层面',
          reason: '术语统一',
          status: 'pending',
        },
      ],
    },
  },
]

const demoMode = () => (import.meta.env.DEV ? new URLSearchParams(location.search).get('demo') : null)
const isDemo = () => demoMode() !== null

const DEMO_INSERT: Item[] = [
  { id: 'i1', kind: 'user', text: '在方法部分补一个样本信息表,再加一段统计方法' },
  {
    id: 'i2',
    kind: 'assistant',
    durationMs: 21400,
    text: '按文档现有小四、Times New Roman 的格式起草了 **3 个块**:一张三线表、一个二级标题、一段统计描述。插入位置默认为光标所在段落之后。',
  },
  {
    id: 'i3',
    kind: 'insert',
    bundle: {
      summary: '方法部分:样本信息表 + 统计方法',
      anchorIndex: 31,
      blocks: [
        {
          id: 'b1',
          kind: 'heading',
          level: 2,
          text: '2.5 样本信息与统计方法',
          status: 'inserted',
        },
        {
          id: 'b2',
          kind: 'table',
          rows: [
            ['组别', 'n', '年龄(岁)', 'Scr(μmol/L)'],
            ['Control', '12', '22.4 ± 1.8', '31.2 ± 4.1'],
            ['CKD 4w', '12', '23.1 ± 2.0', '68.5 ± 7.3'],
            ['CKD 8w', '12', '22.8 ± 1.9', '102.7 ± 9.6'],
          ],
          status: 'pending',
        },
        {
          id: 'b3',
          kind: 'paragraph',
          text: '数据以均值 ± 标准差表示,组间比较采用单因素方差分析,事后检验使用 Tukey 法;P < 0.05 视为差异有统计学意义。所有分析在 GraphPad Prism 9 中完成。',
          status: 'pending',
        },
      ],
    },
  },
]

const DEMO_COMMENTS: Item[] = [
  { id: 'x1', kind: 'user', text: '帮我审阅引言,指出问题' },
  {
    id: 'x2',
    kind: 'assistant',
    durationMs: 18200,
    text: '读了引言共 3 段,发现 **3 处**值得修改:用词比喻化、被动语态缺引用、以及一处结论没交代统计方法。已提交为批注:',
  },
  {
    id: 'x3',
    kind: 'comments',
    bundle: {
      summary: '审阅引言:用词、引用与统计交代',
      selectionText: '',
      comments: [
        {
          id: 'c1',
          paragraphIndex: 11,
          find: 'vascular terrain',
          text: '「vascular terrain」是比喻性口语表达,学术写作建议改为 vasculature 或 vascular architecture。',
          severity: 'warn',
          status: 'pending',
        },
        {
          id: 'c2',
          paragraphIndex: 11,
          find: 'has been employed',
          text: '被动完成时未交代由谁开展、依据何在。建议改主动句式并补文献引用。',
          severity: 'warn',
          status: 'pending',
        },
        {
          id: 'c3',
          paragraphIndex: 26,
          find: 'observed in CKD',
          text: '此处结论需标注样本量与统计方法,否则读者无法判断可信度。',
          severity: 'info',
          status: 'pending',
        },
      ],
    },
  },
]

const DEMO_AUDIT: Item[] = [
  { id: 'a0', kind: 'user', text: '投稿前帮我通读一遍,找结构上的漏洞' },
  {
    id: 'a1',
    kind: 'assistant',
    durationMs: 24300,
    text: '通读了全文 **162 段**,发现 **3 处**需要处理:一处声明空缺、两处章节待填。已生成审查清单:',
  },
  {
    id: 'a2',
    kind: 'audit',
    bundle: {
      summary: '全文结构与逻辑审查',
      basis: [{ name: 'ref.pdf', page: 3 }],
      issues: [
        {
          id: 'i1',
          category: '声明空缺',
          severity: 'priority',
          paragraphIndex: 151,
          find: 'Data Availability Statement',
          title: 'Data Availability Statement 下为空行',
          detail: '标题下目前为空行,缺少单细胞测序数据存放的 GEO 编号与公开访问声明。',
          suggestion: '补充 GSE 编号与「data are available upon reasonable request」声明。',
          status: 'pending',
        },
        {
          id: 'i2',
          category: '章节待填',
          severity: 'normal',
          paragraphIndex: 148,
          find: 'Author Contributions',
          title: '作者贡献与基金声明未填正文',
          detail: 'Author Contributions 与 Funding 标题下尚未填入正文,可按 CRediT 规范补充。',
          suggestion: '插入作者贡献占位段与基金编号。',
          status: 'pending',
        },
        {
          id: 'i3',
          category: '引文格式',
          severity: 'normal',
          paragraphIndex: 158,
          find: '2045–59',
          title: '引文页码使用了缩写格式',
          detail: '引文 1 与 2 的页码写成 2045-59 与 575-90,与期刊要求的完整起止页不一致。',
          suggestion: '统一为完整起止页码(2045–2059 / 575–590)。',
          status: 'pending',
        },
      ],
    },
  },
]

const fmtDur = (ms: number) => {
  const s = Math.max(1, Math.round(ms / 1000))
  return `${String(Math.floor(s / 60)).padStart(2, '0')}:${String(s % 60).padStart(2, '0')}`
}

/** 把一处改动写成批注时的一句话 */
const commentFromEdit = (e: EditItem) => {
  const body = e.replace.replace(/\s+/g, ' ').trim()
  const head = e.reason ? `${e.reason}:` : '建议改为:'
  return `${head}${body.length > 140 ? `${body.slice(0, 140)}…` : body}`.slice(0, 240)
}

function toolLabel(name: string) {
  switch (name) {
    case 'doc_outline':
      return '读取文档结构'
    case 'doc_read':
      return '读取正文'
    case 'doc_search':
      return '检索文档'
    case 'propose_edit':
      return '生成改动草稿'
    case 'propose_comment':
      return '生成批注草稿'
    case 'propose_insert':
      return '生成插入草稿'
    case 'propose_audit':
      return '生成全文审查清单'
    case 'ref_list':
      return '查看参考文献'
    case 'ref_search':
      return '检索参考文献'
    case 'ref_read':
      return '读取参考文献原文'
    case 'pdf_info':
    case 'pdf_extract_text':
    case 'pdf_extract_tables':
      return '翻阅 PDF 原文'
    default:
      return name
  }
}

export default function App() {
  const [items, setItems] = useState<Item[]>(() => {
    const m = demoMode()
    if (m === 'comments') return DEMO_COMMENTS
    if (m === 'insert') return DEMO_INSERT
    if (m === 'audit') return DEMO_AUDIT
    return isDemo() ? DEMO : []
  })
  const [input, setInput] = useState('')
  const [busy, setBusy] = useState(false)
  const [selText, setSelText] = useState('')
  /** 用户主动忽略当前选区(只按指令处理全文) */
  const [selIgnored, setSelIgnored] = useState(false)
  const [cmds, setCmds] = useState<QuickCmd[]>(() => loadCmds())
  const [rules, setRules] = useState<Rule[]>(() => {
    try {
      const raw = JSON.parse(localStorage.getItem('wordai.rules') ?? '[]')
      return Array.isArray(raw) ? raw.filter((r): r is Rule => typeof r?.id === 'string' && typeof r?.text === 'string') : []
    } catch {
      return []
    }
  })
  const [ready, setReady] = useState(false)
  const [showSettings, setShowSettings] = useState(false)
  const [models, setModels] = useState<ModelInfo[]>([])
  const [currentModel, setCurrentModel] = useState<ModelInfo | null>(null)
  const [modelsLoading, setModelsLoading] = useState(true)
  const [modelsError, setModelsError] = useState('')
  const [modelSaving, setModelSaving] = useState(false)
  const bottomRef = useRef<HTMLDivElement | null>(null)
  const docKeyRef = useRef<string | null>(null)
  /** 发送那一刻冻结的选区锚点:本轮所有草稿共用它,绝不退回「此刻选区」 */
  const anchorRef = useRef<Anchor | null>(null)
  /** 当前这一轮生成的中止句柄:发送键在忙碌时就是中止键 */
  const runRef = useRef<AbortController | null>(null)

  const [sessions, setSessions] = useState<SessionMeta[]>([])
  const [activeId, setActiveId] = useState<string | null>(null)
  const [maxSessions, setMaxSessions] = useState(5)
  const [loaded, setLoaded] = useState(false)
  const [undoStack, setUndoStack] = useState<UndoEntry[]>([])
  const [refs, setRefs] = useState<RefMeta[]>(() =>
    isDemo()
      ? [
          {
            id: 'demo',
            name: 'Raaz 2019 — Runx2 and aortic fibrosis.pdf',
            path: '~/Downloads/nihms-712452.pdf',
            kind: 'pdf',
            pages: 23,
            chars: 58420,
            chunks: 121,
            addedAt: '2026-01-01T00:00:00Z',
          },
        ]
      : [],
  )
  const [historyOpen, setHistoryOpen] = useState(false)
  const [picking, setPicking] = useState(false)
  const [dragging, setDragging] = useState(false)
  /** 进入/离开会在子元素上反复触发,用计数防抖 */
  const dragDepth = useRef(0)
  const undoRef = useRef<{ docKey: string; id: string } | null>(null)
  // 保存时要知道「这批 items 属于哪个会话」,不能指望 state 及时更新
  const saveRef = useRef<{ docKey: string; id: string } | null>(null)
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null)

  const refreshSelection = useCallback(async () => {
    try {
      const sel = await captureSelection()
      setSelText(sel.text)
      return sel
    } catch {
      return { text: '', empty: true }
    }
  }, [])

  useEffect(() => {
    initOffice().then(async () => {
      setReady(true)
      if (!inWord()) return
      await refreshSelection()

      if (isDemo()) {
        setLoaded(true)
        return
      }

      // 面板重新加载 = 回到该文档上次的会话,而不是从零开始
      try {
        await resetPiSession()
        const docKey = await getDocKey()
        docKeyRef.current = docKey
        const st = await fetchSessions(docKey)
        setSessions(st.sessions)
        setActiveId(st.activeId)
        setMaxSessions(st.max ?? 5)
        if (st.activeId) {
          const items = await fetchSessionUi<Item>(docKey, st.activeId)
          saveRef.current = { docKey, id: st.activeId }
          setItems(items)
          undoRef.current = { docKey, id: st.activeId }
          setUndoStack(await fetchUndo<UndoEntry>(docKey, st.activeId))
        }
        setRefs(await fetchRefs(docKey, st.activeId ?? ''))
      } catch {
        /* 读不到历史就当新会话,不阻塞使用 */
      } finally {
        setLoaded(true)
      }
    })
  }, [refreshSelection])

  // 对话落到磁盘:切会话/关面板都不会丢
  useEffect(() => {
    if (!loaded || isDemo()) return
    const target = saveRef.current
    if (!target) return
    if (timerRef.current) clearTimeout(timerRef.current)
    timerRef.current = setTimeout(() => {
      void saveSessionUi(target.docKey, target.id, items).catch(() => {})
    }, 800)
    return () => {
      if (timerRef.current) clearTimeout(timerRef.current)
    }
  }, [items, loaded])

  useEffect(() => {
    if (!ready || !inWord()) return
    // 实时跟随文档选区,不用手动点「取选区」
    return onSelectionChanged((text) => {
      setSelText(text)
      // 用户重新选了别的东西,就不再沿用「忽略」
      if (text.trim()) setSelIgnored(false)
    })
  }, [ready])

  useEffect(() => {
    let alive = true
    fetchModels()
      .then((c) => {
        if (!alive) return
        setModels(c.models)
        setCurrentModel(c.current)
        setModelsError(c.error ?? '')
      })
      .catch((e) => {
        if (alive) setModelsError((e as Error).message)
      })
      .finally(() => {
        if (alive) setModelsLoading(false)
      })
    return () => {
      alive = false
    }
  }, [])

  // 旧版「写作预设」已经下线,清掉它的残留键,免得以后有人翻到半截配置
  useEffect(() => {
    dropLegacyPreset()
  }, [])

  useEffect(() => {
    saveCmds(cmds)
  }, [cmds])

  useEffect(() => {
    try {
      localStorage.setItem('wordai.rules', JSON.stringify(rules))
    } catch {
      /* ignore */
    }
  }, [rules])

  // 流式期间用瞬时滚动:平滑动画每来一个 token 就被打断,看起来像卡住
  const streaming = items.some((it) => it.kind === 'assistant' && !it.durationMs)

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: streaming ? 'auto' : 'smooth', block: 'end' })
  }, [items, streaming])

  const patch = (id: string, fn: (it: Item) => Item) =>
    setItems((prev) => prev.map((it) => (it.id === id ? fn(it) : it)))

  /** 切到某个会话:换 Pi 上下文 + 换侧栏对话 */
  async function switchTo(id: string) {
    if (busy || isDemo() || id === activeId) return
    const docKey = saveRef.current?.docKey ?? docKeyRef.current
    if (!docKey) return
    setBusy(true)
    try {
      // 切走前先把当前对话存下来,不然最后一轮白丢
      const prev = saveRef.current
      if (prev) await saveSessionUi(prev.docKey, prev.id, items).catch(() => {})

      // 切回历史里的会话:它变当前,原当前自动归档(无需额外标记)
      const st = await switchSessionApi(docKey, id)
      await loadActive(st, docKey)
      setHistoryOpen(false)
      await resetPiSession().catch(() => {})
    } catch (e) {
      setItems((prev) => [...prev, { id: uid(), kind: 'error', text: (e as Error).message }])
    } finally {
      setBusy(false)
    }
  }

  /** 把一份会话索引落到界面上:标题栏、对话、撤销栈、附件一起换 */
  async function loadActive(st: { activeId: string | null; sessions: SessionMeta[]; max?: number }, docKey: string) {
    setSessions(st.sessions)
    setActiveId(st.activeId)
    setMaxSessions(st.max ?? 30)

    const id = st.activeId
    if (!id) {
      saveRef.current = null
      undoRef.current = null
      setItems([])
      setUndoStack([])
      setRefs([])
      return
    }

    saveRef.current = { docKey, id }
    undoRef.current = { docKey, id }
    setItems(await fetchSessionUi<Item>(docKey, id))
    setUndoStack(await fetchUndo<UndoEntry>(docKey, id))
    setRefs(await fetchRefs(docKey, id))
  }

  /** 新建会话 = 当前会话归档进历史 */
  async function newSession() {
    if (busy || isDemo()) return
    const docKey = saveRef.current?.docKey ?? docKeyRef.current
    if (!docKey) return
    setBusy(true)
    try {
      const prev = saveRef.current
      if (prev) await saveSessionUi(prev.docKey, prev.id, items).catch(() => {})

      const st = await createSession(docKey)
      await loadActive(st, docKey)
      setHistoryOpen(false)
      await resetPiSession().catch(() => {})
    } catch (e) {
      setItems((prev) => [...prev, { id: uid(), kind: 'error', text: (e as Error).message }])
    } finally {
      setBusy(false)
    }
  }

  /** 多选批量删会话:目录一起删,删到当前会话就换到剩下的第一个 */
  async function deleteSessions(ids: string[]) {
    if (busy || isDemo() || !ids.length) return
    const docKey = saveRef.current?.docKey ?? docKeyRef.current
    if (!docKey) return
    setBusy(true)
    try {
      // 只删历史的场合先落一次盘,免得当前对话丢掉;删的要是当前会话就没必要了
      const prev = saveRef.current
      if (prev && !ids.includes(prev.id)) await saveSessionUi(prev.docKey, prev.id, items).catch(() => {})

      const st = await deleteSessionsApi(docKey, ids)
      await loadActive(st, docKey)
      await resetPiSession().catch(() => {})
    } catch (e) {
      setItems((prev) => [...prev, { id: uid(), kind: 'error', text: (e as Error).message }])
    } finally {
      setBusy(false)
    }
  }

  /** 记一次可撤销的写入 */
  function pushUndo(label: string, ops: UndoOp[], revert?: CardRevert) {
    if (!ops.length) return
    const entry: UndoEntry = {
      id: uid(),
      at: new Date().toISOString(),
      label,
      ops,
      revert: revert && revert.ids.length ? [revert] : undefined,
    }
    setUndoStack((prev) => {
      const next = [...prev, entry].slice(-UNDO_LIMIT)
      const target = undoRef.current
      if (target) void saveUndo(target.docKey, target.id, next).catch(() => {})
      return next
    })
  }

  /** 把卡片上被撤销的子项打回「待应用」 */
  function applyCardRevert(list?: CardRevert[]) {
    if (!list?.length) return
    setItems((prev) =>
      prev.map((it) => {
        const hit = list.find((r) => r.itemId === it.id)
        if (!hit) return it

        if (hit.kind === 'proposal' && it.kind === 'proposal') {
          return {
            ...it,
            proposal: {
              ...it.proposal,
              edits: it.proposal.edits.map((e) =>
                hit.ids.includes(e.id) ? { ...e, status: 'pending' as const, error: undefined } : e,
              ),
            },
          }
        }

        if (hit.kind === 'comments' && it.kind === 'comments') {
          return {
            ...it,
            bundle: {
              ...it.bundle,
              comments: it.bundle.comments.map((c) =>
                hit.ids.includes(c.id) ? { ...c, status: 'pending' as const, error: undefined } : c,
              ),
            },
          }
        }

        if (hit.kind === 'insert' && it.kind === 'insert') {
          return {
            ...it,
            bundle: {
              ...it.bundle,
              blocks: it.bundle.blocks.map((b) =>
                hit.ids.includes(b.id) ? { ...b, status: 'pending' as const, error: undefined } : b,
              ),
            },
          }
        }

        if (hit.kind === 'audit' && it.kind === 'audit') {
          return {
            ...it,
            bundle: {
              ...it.bundle,
              issues: it.bundle.issues.map((i) =>
                hit.ids.includes(i.id) ? { ...i, status: 'pending' as const, error: undefined } : i,
              ),
            },
          }
        }

        return it
      }),
    )
  }

  /**
   * 附加参考文献。服务端弹 macOS 原生选文件窗口,这个请求会一直挂着等用户选完。
   * 文献内容不进提示词:附件只在服务端建索引,模型用 ref_search / ref_read 按需取。
   */
  async function addRefs() {
    if (busy || picking || isDemo()) return
    const docKey = saveRef.current?.docKey ?? docKeyRef.current
    if (!docKey) {
      setItems((prev) => [...prev, { id: uid(), kind: 'error', text: '还没读到文档路径,稍等一秒再点。' }])
      return
    }

    setPicking(true)
    try {
      const r = await attachRefs(docKey, saveRef.current?.id ?? '')
      if (r.error) throw new Error(r.error)
      reportAttach(r)
    } catch (e) {
      setItems((prev) => [...prev, { id: uid(), kind: 'error', text: (e as Error).message }])
    } finally {
      setPicking(false)
    }
  }

  /** 对话框与拖拽两个入口共用一套汇报 */
  function reportAttach(r: AttachResult) {
    setRefs(r.refs)

    const added = r.added ?? []
    if (added.length) {
      const list = added.map((a) => `${a.name}(${a.pages ? `${a.pages} 页` : `${a.chunks} 段`})`).join('、')
      setItems((prev) => [
        ...prev,
        {
          id: uid(),
          kind: 'notice',
          text: `已附加 ${added.length} 份参考文献:${list}。可以直接说「参照这份文献第 N 页那段改写我选中的文字」。`,
        },
      ])
    }

    const errs = r.errors ?? []
    if (errs.length) {
      setItems((prev) => [
        ...prev,
        { id: uid(), kind: 'error', text: `没能附加:${errs.map((e) => `${e.file}(${e.error})`).join(';')}` },
      ])
    }
  }

  /** 拖进来的文件拿不到路径,只能上传字节;先在前端拦掉格式/体积不符的 */
  async function dropFiles(files: File[]) {
    if (busy || picking || isDemo() || !files.length) return
    const docKey = saveRef.current?.docKey ?? docKeyRef.current
    if (!docKey) {
      setItems((prev) => [...prev, { id: uid(), kind: 'error', text: '还没读到文档路径,稍等一秒再拖。' }])
      return
    }

    const ok: File[] = []
    const bad: { file: string; error: string }[] = []
    for (const f of files) {
      const ext = f.name.split('.').pop()?.toLowerCase() ?? ''
      if (!REF_EXT.has(ext)) bad.push({ file: f.name, error: `不支持的格式 .${ext || '?'}` })
      else if (f.size > REF_MAX) bad.push({ file: f.name, error: '超过 80MB' })
      else ok.push(f)
    }
    if (bad.length) {
      setItems((prev) => [
        ...prev,
        { id: uid(), kind: 'error', text: `没能附加:${bad.map((b) => `${b.file}(${b.error})`).join(';')}` },
      ])
    }
    if (!ok.length) return

    setPicking(true)
    try {
      const r = await uploadRefs(docKey, saveRef.current?.id ?? '', ok)
      reportAttach(r)
    } catch (e) {
      setItems((prev) => [...prev, { id: uid(), kind: 'error', text: (e as Error).message }])
    } finally {
      setPicking(false)
    }
  }

  async function dropRef(ref: RefMeta) {
    if (busy || isDemo()) return
    const docKey = saveRef.current?.docKey ?? docKeyRef.current
    if (!docKey) return
    try {
      setRefs(await removeRef(docKey, saveRef.current?.id ?? '', ref.id))
    } catch (e) {
      setItems((prev) => [...prev, { id: uid(), kind: 'error', text: (e as Error).message }])
    }
  }

  async function undoLast() {
    if (busy || isDemo()) return
    const entry = undoStack[undoStack.length - 1]
    if (!entry) return

    setBusy(true)
    try {
      const { restored, failures } = await undoOps(entry.ops)

      // 文档还原后,卡片也回到「待应用」;部分失败时只回退确实还原的那些不现实,
      // 因为 ops 与子项不是一一对应,所以按整条回退,失败的会重新标成待应用让用户决定
      if (restored > 0) applyCardRevert(entry.revert)

      // 部分成功必须出栈(半成品再撤一次只会更乱);
      // 但一处都没还原时文档根本没变,保留记录让用户处理完还能再撤。
      if (restored > 0 || failures.length === 0) {
        setUndoStack((prev) => {
          const next = prev.filter((e) => e.id !== entry.id)
          const target = undoRef.current
          if (target) void saveUndo(target.docKey, target.id, next).catch(() => {})
          return next
        })
      }

      if (failures.length) {
        setItems((prev) => [
          ...prev,
          {
            id: uid(),
            kind: 'error',
            text: `撤销「${entry.label}」:还原 ${restored} 处,${failures.length} 处未撤销(${failures[0]})。记录已保留,确认文档后可再点撤销。`,
          },
        ])
      } else {
        setItems((prev) => [
          ...prev,
          {
            id: uid(),
            kind: 'notice',
            text: `已撤销「${entry.label}」,文档还原 ${restored} 处,卡片已回到待应用状态(原文已重新选中)`,
          },
        ])
      }
      await refreshSelection()
    } catch (e) {
      setItems((prev) => [...prev, { id: uid(), kind: 'error', text: (e as Error).message }])
    } finally {
      setBusy(false)
    }
  }

  async function send(text?: string) {
    const instruction = (text ?? input).trim()
    if (!instruction || busy || isDemo()) return

    setInput('')
    setBusy(true)
    const ctl = new AbortController()
    runRef.current = ctl

    // 文档在页面存活期间被换掉(同窗口切换)时,连同 UI 一起重置
    const docKey = await getDocKey()
    const switched = docKeyRef.current !== null && docKeyRef.current !== docKey
    docKeyRef.current = docKey

    // 同窗口换了文档:对话要重置,落盘目标也必须跟着换,否则会把新文档的对话写进旧会话
    if (switched) {
      try {
        const st = await fetchSessions(docKey)
        setSessions(st.sessions)
        setActiveId(st.activeId)
        setMaxSessions(st.max ?? 5)
        saveRef.current = st.activeId ? { docKey, id: st.activeId } : null
        setRefs(await fetchRefs(docKey, st.activeId ?? ''))
      } catch {
        saveRef.current = null
        setRefs([])
      }
    } else if (!saveRef.current) {
      try {
        const st = await fetchSessions(docKey)
        setSessions(st.sessions)
        setActiveId(st.activeId)
        setMaxSessions(st.max ?? 5)
        saveRef.current = st.activeId ? { docKey, id: st.activeId } : null
        setRefs(await fetchRefs(docKey, st.activeId ?? ''))
      } catch {
        /* 忽略 */
      }
    }

    setItems((prev) => [
      ...(switched ? [] : prev),
      { id: uid(), kind: 'user', text: instruction },
    ])

    const sel = await refreshSelection()
    if (!inWord()) {
      setItems((prev) => [...prev, { id: uid(), kind: 'error', text: '当前不在 Word 中,无法读取文档。' }])
      setBusy(false)
      return
    }

    // 锚点在这里冻结:之后用户点到哪里,写入都只认它
    const effectiveSel = selIgnored ? '' : sel.text
    anchorRef.current = await captureAnchor(effectiveSel)

    const assistantId = uid()
    const startedAt = Date.now()
    setItems((prev) => [...prev, { id: assistantId, kind: 'assistant', text: '' }])

    let assistantText = ''
    let stopped = false
    const base = effectiveSel.trim()
      ? `【用户当前选中的文字】\n<<<\n${effectiveSel}\n>>>\n\n【指令】\n${instruction}`
      : `【用户没有选中文字】\n\n【指令】\n${instruction}`
    const prefix = rulesPrompt(rules)
    const prompt = prefix ? `${prefix}\n\n${base}` : base

    await runPiAgent(prompt, docKey, {
      selectionText: effectiveSel,
      anchor: anchorRef.current ?? undefined,
      onText: (delta) => {
        assistantText += delta
        patch(assistantId, (it) => (it.kind === 'assistant' ? { ...it, text: assistantText } : it))
      },
      onTool: (name) => {
        patch(assistantId, (it) => (it.kind === 'assistant' ? { ...it, tool: name } : it))
      },
      onProposal: (proposal) => {
        const itemId = uid()
        setItems((prev) => [
          ...prev.map((it) =>
            it.id === assistantId && it.kind === 'assistant' ? { ...it, tool: undefined } : it,
          ),
          { id: itemId, kind: 'proposal', proposal },
        ])

        // 补上各段原文,卡片才能显示「改的这句在什么上下文里」
        const indices = proposal.edits
          .map((e) => e.paragraphIndex)
          .filter((i): i is number => i !== null)
        if (indices.length) {
          void readParagraphContexts(indices)
            .then((context) => {
              if (!Object.keys(context).length) return
              setItems((prev) =>
                prev.map((it) =>
                  it.id === itemId && it.kind === 'proposal'
                    ? { ...it, proposal: { ...it.proposal, context } }
                    : it,
                ),
              )
            })
            .catch(() => {})
        }
      },
      onComments: (bundle) => {
        setItems((prev) => [
          ...prev.map((it) =>
            it.id === assistantId && it.kind === 'assistant' ? { ...it, tool: undefined } : it,
          ),
          { id: uid(), kind: 'comments', bundle },
        ])
      },
      onInsert: (bundle) => {
        // 模型没给锚点就用发送时冻结的段落,别落到「此刻光标」
        const fixed =
          bundle.anchorIndex === null && bundle.anchor?.paragraphIndex !== null && bundle.anchor !== undefined
            ? { ...bundle, anchorIndex: bundle.anchor.paragraphIndex }
            : bundle
        setItems((prev) => [
          ...prev.map((it) =>
            it.id === assistantId && it.kind === 'assistant' ? { ...it, tool: undefined } : it,
          ),
          { id: uid(), kind: 'insert', bundle: fixed },
        ])
      },
      onAudit: (bundle) => {
        setItems((prev) => [
          ...prev.map((it) =>
            it.id === assistantId && it.kind === 'assistant' ? { ...it, tool: undefined } : it,
          ),
          { id: uid(), kind: 'audit', bundle },
        ])
      },
      onToolError: (name, message) => {
        const hint = message.includes('未知工具')
          ? `侧栏版本较旧,不支持「${name}」。请关闭侧栏(Cmd+W 或右上角 ×)后重新打开。`
          : `工具 ${name} 失败:${message}`
        setItems((prev) => [...prev, { id: uid(), kind: 'error', text: hint }])
      },
      onDone: () => {},
      onAborted: () => {
        stopped = true
      },
      onError: (message) => {
        setItems((prev) => [...prev, { id: uid(), kind: 'error', text: message }])
      },
    }, { signal: ctl.signal })

    runRef.current = null
    const ms = Date.now() - startedAt
    patch(assistantId, (it) =>
      it.kind === 'assistant'
        ? { ...it, durationMs: ms, tool: undefined, text: assistantText || it.text }
        : it,
    )

    if (stopped) {
      // 一个字都没生成时不要留一个空的「工作了 00:03」气泡
      setItems((prev) => [
        ...prev.filter((it) => !(it.id === assistantId && it.kind === 'assistant' && !it.text.trim())),
        { id: uid(), kind: 'notice', text: '已中止本次生成,已写入的内容不受影响' },
      ])
    }
    setBusy(false)
  }

  async function applyChosen(itemId: string, proposal: Proposal, chosen: EditItem[]) {
    setBusy(true)
    const results = new Map<string, { ok: boolean; error?: string }>()

    try {
      const paraEdits = chosen.filter((e) => e.paragraphIndex !== null)
      const selEdits = chosen.filter((e) => e.paragraphIndex === null)

      let paraOutcomes: Awaited<ReturnType<typeof applyDocEdits>>['outcomes'] = []
      if (paraEdits.length) {
        const { outcomes } = await applyDocEdits(
          paraEdits.map((e) => ({ paragraphIndex: e.paragraphIndex!, find: e.find, replace: e.replace })),
        )
        paraOutcomes = outcomes
        for (const o of outcomes) {
          const target = paraEdits[o.at]
          if (target) results.set(target.id, { ok: o.ok, error: o.reason })
        }
      }

      for (const e of selEdits) {
        const expected = proposal.anchor?.text ?? proposal.selectionText
        if (!expected.trim()) {
          results.set(e.id, { ok: false, error: '没有可用的选区锚点,请先选中要改的文字' })
          continue
        }
        const r = await applyReplacement(e.replace, expected, proposal.anchor)
        results.set(e.id, { ok: r.ok, error: r.message })
      }

      setItems((prev) =>
        prev.map((it) => {
          if (it.id !== itemId || it.kind !== 'proposal') return it
          const edits = it.proposal.edits.map((e) => {
            const r = results.get(e.id)
            if (!r) return e
            return { ...e, status: r.ok ? ('applied' as const) : ('failed' as const), error: r.error }
          })
          return { ...it, proposal: { ...it.proposal, edits } }
        }),
      )

      // 记录可撤销的反向操作(只记真正写进去的)
      const ops: UndoOp[] = []
      for (const o of paraOutcomes) {
        if (o.ok && o.before && o.after) ops.push({ kind: 'paragraph', before: o.before, after: o.after })
      }
      for (const e of selEdits) {
        const r = results.get(e.id)
        if (r?.ok) ops.push({ kind: 'paragraph', before: proposal.anchor?.text ?? proposal.selectionText, after: e.replace })
      }
      pushUndo(proposal.summary || '改写', ops, {
        itemId,
        kind: 'proposal',
        ids: [...results.entries()].filter(([, r]) => r.ok).map(([id]) => id),
      })

      const failed = [...results.values()].filter((r) => !r.ok)
      if (failed.length) {
        setItems((prev) => [
          ...prev,
          { id: uid(), kind: 'error', text: `${failed.length} 处未能写入:${failed[0].error ?? '未知原因'}` },
        ])
      }
      await refreshSelection()
    } catch (e) {
      setItems((prev) => [...prev, { id: uid(), kind: 'error', text: (e as Error).message }])
    } finally {
      setBusy(false)
    }
  }

  async function applyInsertChosen(itemId: string, bundle: InsertBundle, chosen: InsertBlock[]) {
    setBusy(true)
    const results = new Map<string, { ok: boolean; error?: string }>()

    try {
      const { outcomes } = await applyInsert(
        bundle.anchorIndex,
        chosen.map((b) => ({
          kind: b.kind,
          level: b.level,
          text: b.text,
          items: b.items,
          rows: b.rows,
        })),
      )
      const ops: UndoOp[] = []
      for (const o of outcomes) {
        const target = chosen[o.at]
        if (target) results.set(target.id, { ok: o.ok, error: o.reason })
        if (!o.ok) continue
        if (o.table) ops.push({ kind: 'insertTable', firstCell: o.table.firstCell, rows: o.table.rows })
        else if (o.texts?.length) ops.push({ kind: 'insertParagraph', texts: o.texts })
      }
      if (ops.length) {
        pushUndo(`插入 ${ops.length} 块`, ops, {
          itemId,
          kind: 'insert',
          ids: [...results.entries()].filter(([, r]) => r.ok).map(([id]) => id),
        })
      }

      setItems((prev) =>
        prev.map((it) => {
          if (it.id !== itemId || it.kind !== 'insert') return it
          const blocks = it.bundle.blocks.map((b) => {
            const r = results.get(b.id)
            if (!r) return b
            return { ...b, status: r.ok ? ('inserted' as const) : ('failed' as const), error: r.error }
          })
          return { ...it, bundle: { ...it.bundle, blocks } }
        }),
      )

      const failedList = [...results.values()].filter((r) => !r.ok)
      if (failedList.length) {
        setItems((prev) => [
          ...prev,
          { id: uid(), kind: 'error', text: `${failedList.length} 块未能插入:${failedList[0].error ?? '未知原因'}` },
        ])
      }
    } catch (e) {
      setItems((prev) => [...prev, { id: uid(), kind: 'error', text: (e as Error).message }])
    } finally {
      setBusy(false)
    }
  }

  function skipAllInsert(itemId: string) {
    setItems((prev) =>
      prev.map((it) => {
        if (it.id !== itemId || it.kind !== 'insert') return it
        return {
          ...it,
          bundle: {
            ...it.bundle,
            blocks: it.bundle.blocks.map((b) => (b.status === 'pending' ? { ...b, status: 'skipped' as const } : b)),
          },
        }
      }),
    )
  }

  async function applyCommentChosen(itemId: string, bundle: CommentBundle, chosen: CommentItem[]) {
    setBusy(true)
    const results = new Map<string, { ok: boolean; error?: string }>()

    try {
      const { outcomes } = await applyComments(
        chosen.map((c) => ({
          paragraphIndex: c.paragraphIndex,
          find: c.find,
          text: c.text,
          anchor: c.paragraphIndex === null ? bundle.anchor : undefined,
        })),
      )
      const ops: UndoOp[] = []
      for (const o of outcomes) {
        const target = chosen[o.at]
        if (target) results.set(target.id, { ok: o.ok, error: o.reason })
        if (o.ok && o.commentId) {
          ops.push({
            kind: 'comment',
            commentId: o.commentId,
            text: target?.text ?? '',
            anchorText: target?.find ?? bundle.anchor?.text ?? '',
          })
        }
      }
      pushUndo(`批注 ${ops.length} 条`, ops, {
        itemId,
        kind: 'comments',
        ids: [...results.entries()].filter(([, r]) => r.ok).map(([id]) => id),
      })

      setItems((prev) =>
        prev.map((it) => {
          if (it.id !== itemId || it.kind !== 'comments') return it
          const comments = it.bundle.comments.map((c) => {
            const r = results.get(c.id)
            if (!r) return c
            return { ...c, status: r.ok ? ('applied' as const) : ('failed' as const), error: r.error }
          })
          return { ...it, bundle: { ...it.bundle, comments } }
        }),
      )

      const failedList = [...results.values()].filter((r) => !r.ok)
      if (failedList.length) {
        setItems((prev) => [
          ...prev,
          { id: uid(), kind: 'error', text: `${failedList.length} 条批注未能写入:${failedList[0].error ?? '未知原因'}` },
        ])
      }
    } catch (e) {
      setItems((prev) => [...prev, { id: uid(), kind: 'error', text: (e as Error).message }])
    } finally {
      setBusy(false)
    }
  }

  function skipAllComments(itemId: string) {
    setItems((prev) =>
      prev.map((it) => {
        if (it.id !== itemId || it.kind !== 'comments') return it
        return {
          ...it,
          bundle: {
            ...it.bundle,
            comments: it.bundle.comments.map((c) =>
              c.status === 'pending' ? { ...c, status: 'skipped' as const } : c,
            ),
          },
        }
      }),
    )
  }

  function skipAll(itemId: string) {
    setItems((prev) =>
      prev.map((it) => {
        if (it.id !== itemId || it.kind !== 'proposal') return it
        return {
          ...it,
          proposal: {
            ...it.proposal,
            edits: it.proposal.edits.map((e) => (e.status === 'pending' ? { ...e, status: 'skipped' as const } : e)),
          },
        }
      }),
    )
  }

  /** 不改正文,把同一份改动写成 Word 原生批注 */
  async function applyAsComment(itemId: string, proposal: Proposal, chosen: EditItem[]) {
    if (!chosen.length) return
    setBusy(true)
    const results = new Map<string, { ok: boolean; error?: string }>()

    try {
      const list = chosen.map((e) => ({
        paragraphIndex: e.paragraphIndex,
        find: e.find,
        text: commentFromEdit(e),
        anchor: e.paragraphIndex === null ? proposal.anchor : undefined,
      }))
      const { outcomes } = await applyComments(list)

      const ops: UndoOp[] = []
      for (const o of outcomes) {
        const target = chosen[o.at]
        if (target) results.set(target.id, { ok: o.ok, error: o.reason })
        if (o.ok && o.commentId && target) {
          ops.push({
            kind: 'comment',
            commentId: o.commentId,
            text: commentFromEdit(target),
            anchorText: target.find ?? proposal.anchor?.text ?? '',
          })
        }
      }

      setItems((prev) =>
        prev.map((it) => {
          if (it.id !== itemId || it.kind !== 'proposal') return it
          const edits = it.proposal.edits.map((e) => {
            const r = results.get(e.id)
            if (!r) return e
            return { ...e, status: r.ok ? ('commented' as const) : ('failed' as const), error: r.error }
          })
          return { ...it, proposal: { ...it.proposal, edits } }
        }),
      )

      pushUndo(`批注 ${ops.length} 条`, ops, {
        itemId,
        kind: 'proposal',
        ids: [...results.entries()].filter(([, r]) => r.ok).map(([id]) => id),
      })

      const failed = [...results.values()].filter((r) => !r.ok)
      if (failed.length) {
        setItems((prev) => [
          ...prev,
          { id: uid(), kind: 'error', text: `${failed.length} 条批注未能写入:${failed[0].error ?? '未知原因'}` },
        ])
      }
    } catch (e) {
      setItems((prev) => [...prev, { id: uid(), kind: 'error', text: (e as Error).message }])
    } finally {
      setBusy(false)
    }
  }

  /** 保留原文,把新写法插入到锚点段落下方 */
  async function applyAsInsert(itemId: string, proposal: Proposal, chosen: EditItem[]) {
    if (!chosen.length) return
    setBusy(true)
    const results = new Map<string, { ok: boolean; error?: string }>()

    try {
      // 每处改动插到自己锚点段之后;选区类改用冻结锚点的段落
      const groups = new Map<number | null, EditItem[]>()
      for (const e of chosen) {
        const at = e.paragraphIndex ?? proposal.anchor?.paragraphIndex ?? null
        const list = groups.get(at) ?? []
        list.push(e)
        groups.set(at, list)
      }

      const ops: UndoOp[] = []
      for (const [at, list] of groups) {
        if (at === null) {
          for (const e of list) results.set(e.id, { ok: false, error: '没有可用的段落锚点,请先把光标放到目标段落' })
          continue
        }

        const { outcomes } = await applyInsert(
          at,
          list.map((e) => ({ kind: 'paragraph' as const, text: e.replace })),
        )
        for (const o of outcomes) {
          const target = list[o.at]
          if (target) results.set(target.id, { ok: o.ok, error: o.reason })
          if (o.ok && o.texts?.length) ops.push({ kind: 'insertParagraph', texts: o.texts })
        }
      }

      setItems((prev) =>
        prev.map((it) => {
          if (it.id !== itemId || it.kind !== 'proposal') return it
          const edits = it.proposal.edits.map((e) => {
            const r = results.get(e.id)
            if (!r) return e
            return { ...e, status: r.ok ? ('inserted' as const) : ('failed' as const), error: r.error }
          })
          return { ...it, proposal: { ...it.proposal, edits } }
        }),
      )

      pushUndo(`插入 ${ops.length} 块`, ops, {
        itemId,
        kind: 'proposal',
        ids: [...results.entries()].filter(([, r]) => r.ok).map(([id]) => id),
      })

      const failed = [...results.values()].filter((r) => !r.ok)
      if (failed.length) {
        setItems((prev) => [
          ...prev,
          { id: uid(), kind: 'error', text: `${failed.length} 处未能插入:${failed[0].error ?? '未知原因'}` },
        ])
      }
    } catch (e) {
      setItems((prev) => [...prev, { id: uid(), kind: 'error', text: (e as Error).message }])
    } finally {
      setBusy(false)
    }
  }

  /** 审查项 → Word 批注 */
  async function applyAuditComments(itemId: string, bundle: AuditBundle, issues: AuditIssue[]) {
    if (!issues.length) return
    setBusy(true)
    const results = new Map<string, { ok: boolean; error?: string }>()

    try {
      const list = issues.map((i) => ({
        paragraphIndex: i.paragraphIndex,
        find: i.find,
        text: issueCommentText(i),
        anchor: i.paragraphIndex === null ? bundle.anchor : undefined,
      }))
      const { outcomes } = await applyComments(list)

      const ops: UndoOp[] = []
      for (const o of outcomes) {
        const target = issues[o.at]
        if (target) results.set(target.id, { ok: o.ok, error: o.reason })
        if (o.ok && o.commentId && target) {
          ops.push({
            kind: 'comment',
            commentId: o.commentId,
            text: issueCommentText(target),
            anchorText: target.find ?? bundle.anchor?.text ?? '',
          })
        }
      }

      setItems((prev) =>
        prev.map((it) => {
          if (it.id !== itemId || it.kind !== 'audit') return it
          const next = it.bundle.issues.map((i) => {
            const r = results.get(i.id)
            if (!r) return i
            return { ...i, status: r.ok ? ('commented' as const) : ('failed' as const), error: r.error }
          })
          return { ...it, bundle: { ...it.bundle, issues: next } }
        }),
      )

      pushUndo(`批注 ${ops.length} 条`, ops, {
        itemId,
        kind: 'audit',
        ids: [...results.entries()].filter(([, r]) => r.ok).map(([id]) => id),
      })

      const failed = [...results.values()].filter((r) => !r.ok)
      if (failed.length) {
        setItems((prev) => [
          ...prev,
          { id: uid(), kind: 'error', text: `${failed.length} 条审查建议未能写入:${failed[0].error ?? '未知原因'}` },
        ])
      }
    } catch (e) {
      setItems((prev) => [...prev, { id: uid(), kind: 'error', text: (e as Error).message }])
    } finally {
      setBusy(false)
    }
  }

  function skipAllIssues(itemId: string) {
    setItems((prev) =>
      prev.map((it) => {
        if (it.id !== itemId || it.kind !== 'audit') return it
        return {
          ...it,
          bundle: {
            ...it.bundle,
            issues: it.bundle.issues.map((i) => (i.status === 'pending' ? { ...i, status: 'skipped' as const } : i)),
          },
        }
      }),
    )
  }

  /** 针对一条审查项,让 AI 再跑一轮起草修复内容 */
  function draftFix(issue: AuditIssue) {
    if (busy) return
    const where = issue.paragraphIndex !== null ? `第 ${issue.paragraphIndex + 1} 段` : '当前选区'
    void send(
      `针对${where}的审查项「${issue.title}」:${issue.detail}${issue.suggestion ? ` 建议:${issue.suggestion}` : ''}。` +
        '先 doc_read 看原文,再用 propose_insert 起草修复内容;如果能直接改原句就用 propose_edit。',
    )
  }

  async function copyText(text: string) {
    const clean = text.trim()
    if (!clean) return
    try {
      await navigator.clipboard.writeText(clean)
    } catch {
      // WKWebView 的剪贴板权限时灵时不灵,留一条老路
      const ta = document.createElement('textarea')
      ta.value = clean
      ta.style.position = 'fixed'
      ta.style.opacity = '0'
      document.body.appendChild(ta)
      ta.select()
      try {
        document.execCommand('copy')
      } catch {
        /* 复制失败不影响主流程 */
      }
      ta.remove()
    }
    setItems((prev) => [...prev, { id: uid(), kind: 'notice', text: '已复制到剪贴板' }])
  }

  function addRule(text: string) {
    const clean = text.trim().slice(0, 120)
    if (!clean) return
    setRules((prev) => (prev.some((r) => r.text === clean) ? prev : [...prev, { id: uid(), text: clean }].slice(-12)))
  }

  function removeRule(id: string) {
    setRules((prev) => prev.filter((r) => r.id !== id))
  }

  /** 切模型:先乐观更新,失败回滚并提示;生成中不允许 */
  async function changeModel(m: ModelInfo) {
    if (busy || modelSaving) return
    const prev = currentModel
    setCurrentModel(m)
    setModelSaving(true)
    try {
      setCurrentModel(await switchModelApi(m.provider, m.id))
    } catch (e) {
      setCurrentModel(prev)
      setItems((items) => [
        ...items,
        { id: uid(), kind: 'error', text: `切换模型失败:${(e as Error).message}` },
      ])
    } finally {
      setModelSaving(false)
    }
  }

  const hasThread = items.length > 0

  // 预览态用假数据,其余情况就是服务端那份索引;当前会话不算历史
  const uiSessions = isDemo() ? DEMO_SESSIONS : sessions
  const uiActiveId = isDemo() ? DEMO_SESSIONS[DEMO_SESSIONS.length - 1].id : activeId
  const historyCount = uiSessions.filter((s) => s.id !== uiActiveId).length

  /** 中止:交给 runPiAgent 里的 signal,它会同时断开 SSE 与通知服务端 */
  function stop() {
    runRef.current?.abort()
  }

  /* ---------------- 拖拽附加:WKWebView 拿不到路径,只能传字节 ---------------- */

  const isFileDrag = (e: DragEvent) => Array.from(e.dataTransfer?.types ?? []).includes('Files')

  function onDragEnter(e: DragEvent) {
    if (!isFileDrag(e)) return
    e.preventDefault()
    dragDepth.current += 1
    setDragging(true)
  }

  function onDragOver(e: DragEvent) {
    if (!isFileDrag(e)) return
    // 不 preventDefault 浏览器就吞掉 drop,这是 HTML5 拖放的老规矩
    e.preventDefault()
    e.dataTransfer.dropEffect = 'copy'
  }

  function onDragLeave(e: DragEvent) {
    if (!isFileDrag(e)) return
    dragDepth.current = Math.max(0, dragDepth.current - 1)
    if (!dragDepth.current) setDragging(false)
  }

  function onDrop(e: DragEvent) {
    if (!isFileDrag(e)) return
    e.preventDefault()
    dragDepth.current = 0
    setDragging(false)
    void dropFiles(Array.from(e.dataTransfer.files ?? []))
  }

  // 忙碌时按 Esc 也能中止;绑在 window 上,焦点在输入框里时同样生效
  useEffect(() => {
    if (!busy) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Escape') return
      e.preventDefault()
      stop()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [busy])
  return (
    <div
      className="app"
      onDragEnter={onDragEnter}
      onDragOver={onDragOver}
      onDragLeave={onDragLeave}
      onDrop={onDrop}
    >
      <header className="topbar">
        <div className="brand">
          <span className={`dot ${ready ? 'on' : ''}`} title={ready ? 'Word 已连接' : '连接中'} />
          WordAI
        </div>
        <div className="grow" />
        <button className="icon-btn hist-btn" onClick={() => setHistoryOpen(true)} title="历史会话">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round">
            <path d="M4 7h16M4 12h16M4 17h10" />
          </svg>
          {historyCount > 0 && <span className="hist-badge">{historyCount}</span>}
        </button>
        <button
          className="icon-btn"
          disabled={busy || isDemo() || sessions.length - 1 >= maxSessions}
          onClick={() => void newSession()}
          title={
            sessions.length - 1 >= maxSessions
              ? `历史会话最多 ${maxSessions} 个,先在历史里删掉一些`
              : '新增对话(当前会话归档进历史)'
          }
        >
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round">
            <path d="M12 5v14M5 12h14" />
          </svg>
        </button>
        <button className="icon-btn" onClick={() => setShowSettings(true)} title="设置">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round">
            <path d="M4 8h16M4 16h16" />
            <circle cx="9" cy="8" r="2.4" />
            <circle cx="15" cy="16" r="2.4" />
          </svg>
        </button>
      </header>

      <main className="thread">
        {!hasThread ? (
          <div className="empty">
            <div className="greeting">
              <h2>有什么要改的?</h2>
              <p>选中一段文字可以让 AI 就地改写;不选也能让它读全文、按结构找问题。</p>
            </div>
            <div className="prompt-list">
              {cmds.slice(0, 4).map((c) => (
                <button key={c.cmd} className="prompt-card" disabled={busy} onClick={() => void send(c.text)}>
                  <span className="prompt-left">
                    <span className="cmd-tag">{c.cmd}</span>
                    <span className="prompt-title">{c.title}</span>
                  </span>
                  <span className="prompt-sub">{c.sub}</span>
                </button>
              ))}
            </div>
          </div>
        ) : null}

        {items.map((it) => {
          if (it.kind === 'user')
            return (
              <div key={it.id} className="row-end">
                <div className="bubble">{it.text}</div>
              </div>
            )

          if (it.kind === 'error')
            return (
              <div key={it.id} className="notice">
                {it.text}
              </div>
            )

          if (it.kind === 'notice')
            return (
              <div key={it.id} className="notice ok">
                {it.text}
              </div>
            )

          if (it.kind === 'assistant') {
            const working = !it.durationMs
            const label = it.tool ? toolLabel(it.tool) : working ? '思考中' : '已完成思考'
            return (
              <div key={it.id} className="turn">
                <div className={`think-bar ${working ? 'working-bar' : ''}`}>
                  <span>
                    {working ? '▸' : '▸'} {label}
                  </span>
                  <span className="mono">{working ? '···' : fmtDur(it.durationMs!)}</span>
                </div>
                {it.text ? <Markdown text={it.text} /> : null}
                {working && <span className="caret" />}
              </div>
            )
          }

          if (it.kind === 'audit') {
            return (
              <AuditCard
                key={it.id}
                bundle={it.bundle}
                refs={refs}
                busy={busy}
                onComment={(issues) => void applyAuditComments(it.id, it.bundle, issues)}
                onSkipAll={() => skipAllIssues(it.id)}
                onLocate={(idx, find) => void locateInDocument(idx, find).catch(() => {})}
                onDraftFix={draftFix}
              />
            )
          }

          if (it.kind === 'proposal' || it.kind === 'comments' || it.kind === 'insert') {
            return (
              <DraftCard
                key={it.id}
                item={it}
                refs={refs}
                busy={busy}
                onApplyEdits={(chosen) => {
                  if (it.kind === 'proposal') void applyChosen(it.id, it.proposal, chosen)
                }}
                onCommentEdits={(chosen) => {
                  if (it.kind === 'proposal') void applyAsComment(it.id, it.proposal, chosen)
                }}
                onInsertEdits={(chosen) => {
                  if (it.kind === 'proposal') void applyAsInsert(it.id, it.proposal, chosen)
                }}
                onApplyComments={(chosen) => {
                  if (it.kind === 'comments') void applyCommentChosen(it.id, it.bundle, chosen)
                }}
                onApplyInsert={(chosen) => {
                  if (it.kind === 'insert') void applyInsertChosen(it.id, it.bundle, chosen)
                }}
                onSkipAll={() => {
                  if (it.kind === 'proposal') skipAll(it.id)
                  else if (it.kind === 'comments') skipAllComments(it.id)
                  else skipAllInsert(it.id)
                }}
                onLocate={(idx, find) => void locateInDocument(idx, find).catch(() => {})}
                onCopy={copyText}
              />
            )
          }

          return null
        })}
        <div ref={bottomRef} />
      </main>

      <footer className="dock-wrap">
        {(isDemo() || inWord()) && (
          <UndoBar
            stack={
              isDemo()
                ? [
                    {
                      id: 'u1',
                      at: new Date(Date.now() - 95000).toISOString(),
                      label: '统一术语「方法」→「干预策略」',
                      ops: [],
                    },
                  ]
                : undoStack
            }
            busy={busy}
            onUndo={() => void undoLast()}
          />
        )}

        <Composer
          value={input}
          onChange={setInput}
          onSend={() => void send()}
          onStop={stop}
          busy={busy}
          picking={picking}
          refs={refs}
          selChars={selText.trim().length}
          selIgnored={selIgnored}
          onToggleSelection={() => setSelIgnored((v) => !v)}
          onAddRefs={() => void addRefs()}
          onRemoveRef={(r) => void dropRef(r)}
          cmds={cmds}
          models={models}
          currentModel={currentModel}
          modelsLoading={modelsLoading}
          modelsError={modelsError}
          modelBusy={busy || modelSaving}
          onSelectModel={(m) => void changeModel(m)}
          rules={rules}
          onAddRule={addRule}
          onRemoveRule={removeRule}
        />
      </footer>

      {dragging && (
        <div className="drop-mask">
          <div className="drop-card">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round">
              <path d="M20 11.5l-8.2 8.2a4.6 4.6 0 0 1-6.5-6.5l8.6-8.6a3.1 3.1 0 0 1 4.4 4.4l-8.5 8.5a1.6 1.6 0 0 1-2.3-2.3l7.4-7.4" />
            </svg>
            <div className="drop-title">松开即可附加为改写依据</div>
            <div className="drop-sub">PDF / DOCX / RTF / TXT / MD · 单份 ≤ 80MB</div>
          </div>
        </div>
      )}

      {historyOpen && (
        <HistoryPanel
          sessions={uiSessions}
          activeId={uiActiveId}
          busy={busy}
          onClose={() => setHistoryOpen(false)}
          onOpen={(id) => void switchTo(id)}
          onDelete={(ids) => void deleteSessions(ids)}
        />
      )}

      {showSettings && (
        <Settings onClose={() => setShowSettings(false)} cmds={cmds} onChangeCmds={setCmds} />
      )}
    </div>
  )
}
