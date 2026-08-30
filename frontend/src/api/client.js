const BASE = '/api'

async function get(path) {
  const res = await fetch(`${BASE}${path}`)
  if (!res.ok) throw new Error(`API ${path} failed (${res.status})`)
  return res.json()
}

export const api = {
  summary: () => get('/summary'),
  history: (months = 36) => get(`/history?months=${months}`),
  forecast: (component = 'all_in') => get(`/forecast?component=${component}`),
  drivers: () => get('/drivers'),
  importance: () => get('/drivers/importance'),
  modelMetrics: () => get('/model/metrics'),
  validation: () => get('/validation'),
  upload: async (file) => {
    const fd = new FormData()
    fd.append('file', file)
    const res = await fetch(`${BASE}/upload`, { method: 'POST', body: fd })
    const data = await res.json()
    if (!res.ok) throw new Error(data.detail || 'Upload failed')
    return data
  },
}
