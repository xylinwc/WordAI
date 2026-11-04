/**
 * 写入锚点。
 *
 * 教训(r21 回滚):草稿到达时若不冻结锚点,三个动作就会各自去读「此刻选区」,
 * 用户点一下别处,批注就挂到空选区上。所以锚点在**发送那一刻**冻结,
 * 之后所有写入只认它,当前选区只当兜底。
 */
export interface Anchor {
  /** 选区所在段落;跨段或无法唯一确定时为 null(写入时改用文本定位) */
  paragraphIndex: number | null
  /** 冻结时的原文(逐字) */
  text: string
  /** 字数,界面展示用 */
  chars: number
}

export const normSpace = (t: string) => (t ?? '').replace(/\s+/g, ' ').trim()

/**
 * 用选区文本反查它落在哪一段。
 * 唯一命中才算,出现两处相同文本就返回 null —— 宁可走文本定位,也不猜。
 */
export function locateUniqueParagraph(texts: string[], selection: string): number | null {
  const target = normSpace(selection)
  if (!target) return null

  let hit = -1
  for (let i = 0; i < texts.length; i++) {
    if (normSpace(texts[i]).includes(target)) {
      if (hit !== -1) return null
      hit = i
    }
  }
  return hit === -1 ? null : hit
}
