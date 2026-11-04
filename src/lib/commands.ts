/**
 * 快捷指令与写作规则。
 * 指令是用户点一下填进输入框的一段话(一次性);规则是每轮都拼进提示词的前缀(长期)。
 * 两者都不动模型与工具。
 */

export interface QuickCmd {
  /** 触发词,必须 / 开头、小写,全局唯一 */
  cmd: string
  title: string
  sub: string
  /** 点进去后填进输入框的完整指令 */
  text: string
}

export const DEFAULT_CMDS: QuickCmd[] = [
  {
    cmd: '/polish',
    title: '润色得更正式',
    sub: '学术语态',
    text: '把选中的文字润色得更正式:保持原意、术语与引用不变,用克制、书面化的学术语态。',
  },
  {
    cmd: '/rewrite',
    title: '深度改写',
    sub: '允许重排结构',
    text: '深度改写选中的文字:允许重排句子结构与段落逻辑,不只替换词;所有事实、数据与引用必须原样保留。',
  },
  {
    cmd: '/tighten',
    title: '精简冗余表达',
    sub: '压缩约 20%',
    text: '精简选中的文字:删掉冗余与重复,压缩约 20%,事实、数据与引用不能丢。',
  },
  {
    cmd: '/proof',
    title: '修正错别字与标点',
    sub: '拼写 / 斜体',
    text: '校对选中的文字:修正错别字、标点与全半角,统一基因符号与术语的斜体、大小写。',
  },
  {
    cmd: '/audit',
    title: '检查全文的逻辑漏洞',
    sub: '结构审阅',
    text: '通读全文,检查结构与逻辑漏洞(声明空缺、章节待填、引文格式、前后矛盾),用 propose_audit 给出审查清单。',
  },
  {
    cmd: '/transition',
    title: '补写段落过渡句',
    sub: '承上启下',
    text: '参照附件与上下文,为当前段落补写一句承上启下的过渡句。',
  },
]

const CMDS_KEY = 'wordai.cmds'

/** 补上前导 / 、转小写、去掉非法字符 */
export function normalizeSlug(raw: string): string {
  const body = raw
    .trim()
    .replace(/^\/+/, '')
    .toLowerCase()
    .replace(/[^a-z0-9_-]/g, '')
  return body ? `/${body}` : ''
}

/** 合法 slug:至少一个字符的身体 */
export function isValidSlug(raw: string): boolean {
  return /^\/[a-z0-9_-]+$/.test(raw)
}

function clean(raw: unknown): QuickCmd | null {
  if (!raw || typeof raw !== 'object') return null
  const o = raw as Record<string, unknown>
  if (typeof o.cmd !== 'string' || typeof o.title !== 'string' || typeof o.text !== 'string') return null
  const cmd = normalizeSlug(o.cmd)
  if (!isValidSlug(cmd)) return null
  const title = o.title.trim()
  const text = o.text.trim()
  if (!title || !text) return null
  return { cmd, title, sub: typeof o.sub === 'string' ? o.sub.trim() : '', text }
}

/**
 * 读用户自己的指令。
 * 键不存在 = 从没改过 → 用默认;键存在但解析失败 = 数据坏了 → 也回落默认;
 * 键存在且是空数组 = 用户主动删光 → 尊重,返回空。
 */
export function loadCmds(): QuickCmd[] {
  try {
    const raw = localStorage.getItem(CMDS_KEY)
    if (raw === null) return DEFAULT_CMDS
    const parsed: unknown = JSON.parse(raw)
    if (!Array.isArray(parsed)) return DEFAULT_CMDS
    const seen = new Set<string>()
    const list: QuickCmd[] = []
    for (const item of parsed) {
      const c = clean(item)
      if (!c || seen.has(c.cmd)) continue
      seen.add(c.cmd)
      list.push(c)
    }
    return list
  } catch {
    return DEFAULT_CMDS
  }
}

export function saveCmds(list: QuickCmd[]): void {
  try {
    localStorage.setItem(CMDS_KEY, JSON.stringify(list))
  } catch {
    /* 隐私模式写不了就算了 */
  }
}

/** 清掉旧版「写作预设」留下的键,避免残值 */
export function dropLegacyPreset(): void {
  try {
    localStorage.removeItem('wordai.preset')
  } catch {
    /* ignore */
  }
}

/** 校验一组指令,返回第一条错误文案;没问题返回空串 */
export function validateCmds(list: QuickCmd[], selfIndex?: number): string {
  const seen = new Set<string>()
  for (let i = 0; i < list.length; i++) {
    const c = list[i]
    if (!isValidSlug(c.cmd)) return `第 ${i + 1} 条的触发词不合法:只能小写字母、数字、- 和 _,且以 / 开头`
    if (i !== selfIndex && seen.has(c.cmd)) return `触发词重复:${c.cmd}`
    seen.add(c.cmd)
    if (!c.title.trim()) return `第 ${i + 1} 条缺标题`
    if (!c.text.trim()) return `第 ${i + 1} 条缺指令正文`
  }
  return ''
}

export interface Rule {
  id: string
  text: string
}

export function rulesPrompt(rules: Rule[]): string {
  const list = rules.map((r) => r.text.trim()).filter(Boolean)
  if (!list.length) return ''
  return `【写作规则:必须遵守】\n${list.map((t) => `- ${t}`).join('\n')}`
}
