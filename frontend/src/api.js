async function send(path, options) {
  let response
  try {
    response = await fetch(path, options)
  } catch (error) {
    throw new Error('无法连接后端，请稍后重试。', { cause: error })
  }
  const payload = await response.json().catch(() => null)
  if (!response.ok) {
    throw Object.assign(
      new Error(typeof payload?.detail === 'string' ? payload.detail : '请求参数或服务响应有误。'),
      { status: response.status },
    )
  }
  if (!payload) throw new Error('后端返回的数据格式有误。')
  const source = response.headers.get('X-A23-Source-SHA256')
  const result = response.headers.get('X-A23-Result-ID')
  const seconds = response.headers.get('X-A23-Fill-Seconds')
  if (source) payload.source_sha256 = source
  if (result) payload.result_id = result
  if (seconds !== null) payload.fill_seconds = Number(seconds)
  if (response.headers.get('X-A23-Score-Status') === 'failed') payload.score_failed = true
  return payload
}

function upload(path, file, metrics) {
  const body = new FormData()
  body.append('file', file)
  if (metrics) body.append('metrics', metrics)
  return send(path, { method: 'POST', body })
}

function operationUpload(path, file, fields) {
  const body = new FormData()
  body.append('file', file)
  for (const [key, value] of Object.entries(fields)) body.append(key, value)
  return send(path, { method: 'POST', body })
}

function templateUpload(path, source, template, planId) {
  const body = new FormData()
  body.append('source', source)
  body.append('template', template)
  if (planId) body.append('plan_id', planId)
  return send(path, { method: 'POST', body })
}

function params(values) {
  return new URLSearchParams(values).toString()
}

export const api = {
  docs: '/docs',
  health: () => send('/health'),
  parse: (file) => upload('/parse', file),
  ingest: (file) => upload('/ingest', file),
  extract: (file, metrics) => upload('/extract', file, metrics),
  fill: (doc, template, sourceSha256) => send('/fill', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ doc, template, source_sha256: sourceSha256 }),
  }),
  reports: (doc, file, resultId) => send('/reports?' + params({ doc, file, ...(resultId ? { result_id: resultId } : {}) })),
  cell: (doc, file, row, column, resultId) => send('/cells?' + params({ doc, file, row, column, result_id: resultId })),
  download: (name, resultId) => '/files?' + params({ name, result_id: resultId }),
  operationPreview: (file, instruction) => operationUpload('/operations/preview', file, { instruction }),
  operationExecute: (file, planId) => operationUpload('/operations', file, { plan_id: planId }),
  operationDownload: (resultId) => `/operations/files/${encodeURIComponent(resultId)}`,
  templateInspect: (source, template) => templateUpload('/templates/inspect', source, template),
  templateFill: (source, template, planId) => templateUpload('/templates/fill', source, template, planId),
  xlsxTemplateInspect: (source, template) => templateUpload('/templates/xlsx/inspect', source, template),
  xlsxTemplateFill: (source, template, planId, sheet, mapping) => {
    const body = new FormData()
    body.append('source', source)
    body.append('template', template)
    body.append('plan_id', planId)
    body.append('sheet', sheet)
    body.append('mapping', JSON.stringify(mapping))
    return send('/templates/xlsx/fill', { method: 'POST', body })
  },
}
