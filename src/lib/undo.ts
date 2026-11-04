/**
 * 撤销栈。
 *
 * 每次「应用」写入成功后,记录一组反向操作;撤销时倒序执行。
 * 关键设计:每条反向操作都自带**校验依据**,对不上就拒绝,绝不盲删。
 */

export type UndoOp =
  /** 整段还原:按 after 文本在文档里唯一定位,再写回 before */
  | { kind: 'paragraph'; before: string; after: string }
  /** 删除批注:按 commentId 精确删除 */
  | { kind: 'comment'; commentId: string; text: string; anchorText?: string }
  /** 删除插入的段落:按文本序列连续匹配 */
  | { kind: 'insertParagraph'; texts: string[] }
  /** 删除插入的表格:按首格文本 + 行数匹配 */
  | { kind: 'insertTable'; firstCell: string; rows: number }

/** 撤销后要把卡片上的哪些子项恢复成「待应用」 */
export interface CardRevert {
  itemId: string
  kind: 'proposal' | 'comments' | 'insert' | 'audit'
  ids: string[]
}

export interface UndoEntry {
  id: string
  /** ISO 时间 */
  at: string
  /** 展示用,如「第 12 段 术语统一」 */
  label: string
  ops: UndoOp[]
  /** 文档还原后,把卡片的这些子项打回 pending */
  revert?: CardRevert[]
}

export interface UndoResult {
  restored: number
  failures: string[]
}

export const UNDO_LIMIT = 20
