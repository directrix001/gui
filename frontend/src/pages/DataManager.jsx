import { useEffect, useMemo, useState } from 'react'

export default function DataManager() {
  const [drivers, setDrivers] = useState(null)
  const [q, setQ] = useState('')
  const [err, setErr] = useState(null)

  useEffect(() => {
    fetch('/api/real/drivers').then(r => r.json())
      .then(j => setDrivers(j.drivers)).catch(e => setErr(e.message))
  }, [])

  const rows = useMemo(() => {
    if (!drivers) return []
    const months = {}
    for (const d of drivers) {
      for (const p of d.series) {
        months[p.month] = months[p.month] || { month: p.month }
        months[p.month][d.key] = p.value
      }
    }
    let out = Object.values(months).sort((a, b) => b.month.localeCompare(a.month))
    const query = q.trim().toLowerCase()
    if (query) out = out.filter(r => r.month.toLowerCase().includes(query))
    return out
  }, [drivers, q])

  if (err) return <div className="loading">⚠️ {err}</div>
  if (!drivers) return <div className="loading">Loading data…</div>

  const fmt = (v, digits = 4) =>
    v == null ? '—' : v.toLocaleString('en-US', { maximumFractionDigits: digits })

  return (
    <div className="page">
      <div className="card">
        <div className="card-head" style={{ flexWrap: 'wrap', gap: 10 }}>
          <h3>Factor Data</h3>
          <input className="data-search" placeholder="Search month… e.g. 2026 or 2026-03"
            value={q} onChange={e => setQ(e.target.value)} />
        </div>
        <div className="card-body">
          <div style={{ overflowX: 'auto' }}>
            <table className="news-table data-table">
              <thead>
                <tr>
                  <th>Month</th>
                  <th>LME ($/lb)</th>
                  <th>Midwest Premium ($/lb)</th>
                  <th>CNG ($/lb)</th>
                  <th>PPI Index</th>
                </tr>
              </thead>
              <tbody>
                {rows.map(r => (
                  <tr key={r.month}>
                    <td className="mono">{r.month}</td>
                    <td className="mono">{fmt(r.lme)}</td>
                    <td className="mono">{fmt(r.midwest_premium)}</td>
                    <td className="mono">{fmt(r.gas)}</td>
                    <td className="mono">{fmt(r.labour, 3)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            {rows.length === 0 && (
              <div className="loading" style={{ padding: 24 }}>No months match "{q}".</div>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}
