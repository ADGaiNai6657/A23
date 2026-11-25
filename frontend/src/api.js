async function send(path, options) {
  const response = await fetch(path, options)
  const payload = await response.json()
  if (!response.ok) {
    throw Object.assign(
      new Error(typeof payload.detail === 'string' ? payload.detail : '请求参数或服务响应有误'),
      { status: response.status },
    )
  }
  return payload
}

function upload(path, file, metrics) {
  const body = new FormData()
  body.append('file', file)
  if (metrics) body.append('metrics', metrics)
  return send(path, { method: 'POST', body })
}

function params(values) {
  return new URLSearchParams(values).toString()
}

export const api = {
  health: () => send('/health'),
  parse: (file) => upload('/parse', file),
  ingest: (file) => upload('/ingest', file),
  extract: (file, metrics) => upload('/extract', file, metrics),
  fill: (doc, template) => send('/fill', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ doc, template }),
  }),
  reports: (doc, file) => send('/reports?' + params({ doc, file })),
  cell: (doc, file, row, column) => send('/cells?' + params({ doc, file, row, column })),
  download: (name) => '/files?' + params({ name }),
}
