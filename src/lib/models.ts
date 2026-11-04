/** 模型选择:列表与切换都走 Pi 内核,不碰凭证 */

export interface ModelInfo {
  provider: string
  id: string
  name: string
  reasoning?: boolean
  contextWindow?: number
}

export interface ModelCatalog {
  current: ModelInfo | null
  models: ModelInfo[]
  error?: string
}

/** option 的 value:provider 与 id 之间用不可见控制符分隔,避免被 id 里的符号撞车 */
export const modelKey = (m: Pick<ModelInfo, 'provider' | 'id'>) => `${m.provider}\u0001${m.id}`

export async function fetchModels(): Promise<ModelCatalog> {
  const r = await fetch('/api/agent/models')
  const j = (await r.json().catch(() => ({}))) as Partial<ModelCatalog> & { error?: string }
  if (!r.ok) throw new Error(j.error || `读取模型列表失败(${r.status})`)
  return {
    current: j.current ?? null,
    models: Array.isArray(j.models) ? j.models : [],
    error: typeof j.error === 'string' ? j.error : undefined,
  }
}

export async function switchModel(provider: string, id: string): Promise<ModelInfo> {
  const r = await fetch('/api/agent/model', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ provider, id }),
  })
  const j = (await r.json().catch(() => ({}))) as { current?: ModelInfo; error?: string }
  if (!r.ok || !j.current) throw new Error(j.error || `切换模型失败(${r.status})`)
  return j.current
}
