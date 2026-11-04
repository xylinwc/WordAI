/** 参考文献(附件)的前端封装:抽取与检索都在服务端,这里只做增删查 */

export interface RefMeta {
  id: string
  name: string
  path: string
  kind: string
  pages: number
  chars: number
  chunks: number
  addedAt: string
}

export interface AttachResult {
  added?: RefMeta[]
  skipped?: string[]
  errors?: { file: string; error: string }[]
  refs: RefMeta[]
  cancelled?: boolean
  error?: string
}

const post = async (url: string, body: unknown) => {
  const r = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  })
  return (await r.json().catch(() => ({}))) as Record<string, unknown>
}

export async function fetchRefs(docKey: string, id: string): Promise<RefMeta[]> {
  const q = new URLSearchParams({ docKey, id })
  const r = await fetch(`/api/refs?${q}`)
  const j = (await r.json().catch(() => ({}))) as { refs?: RefMeta[] }
  return Array.isArray(j.refs) ? j.refs : []
}

/** 不传 paths 时服务端会弹 macOS 原生选文件窗口,期间这个请求会一直挂着等用户 */
export async function attachRefs(docKey: string, id: string, paths?: string[]): Promise<AttachResult> {
  const j = await post('/api/refs/attach', paths?.length ? { docKey, id, paths } : { docKey, id })
  return {
    ...(j as unknown as AttachResult),
    refs: Array.isArray(j.refs) ? (j.refs as RefMeta[]) : [],
  }
}

export async function removeRef(docKey: string, id: string, refId: string): Promise<RefMeta[]> {
  const j = await post('/api/refs/remove', { docKey, id, refId })
  return Array.isArray(j.refs) ? (j.refs as RefMeta[]) : []
}

export async function uploadRefs(docKey: string, id: string, files: File[]): Promise<AttachResult> {
  const added: RefMeta[] = []
  const errors: { file: string; error: string }[] = []
  let refs: RefMeta[] = []

  for (const f of files) {
    const q = new URLSearchParams({ docKey, id, name: f.name })
    try {
      const r = await fetch(`/api/refs/attach-upload?${q}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/octet-stream' },
        body: f,
      })
      const j = (await r.json().catch(() => ({}))) as Record<string, unknown>
      if (j.error) errors.push({ file: f.name, error: String(j.error) })
      if (Array.isArray(j.errors)) errors.push(...(j.errors as { file: string; error: string }[]))
      if (Array.isArray(j.added)) added.push(...(j.added as RefMeta[]))
      if (Array.isArray(j.refs)) refs = j.refs as RefMeta[]
    } catch (e) {
      errors.push({ file: f.name, error: (e as Error).message })
    }
  }

  return { added, errors, refs }
}
