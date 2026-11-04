import type { AuditIssue } from './pi'

/** 审查项转批注时的正文(统一口径,避免各处拼得不一样) */
export function issueCommentText(issue: AuditIssue): string {
  const head = issue.title ? `${issue.title}。` : ''
  const body = issue.detail ?? ''
  const tail = issue.suggestion ? ` 建议:${issue.suggestion}` : ''
  return `${head}${body}${tail}`.slice(0, 240)
}
