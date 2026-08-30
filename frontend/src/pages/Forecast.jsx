import { useEffect, useMemo, useRef, useState } from 'react'
import {
  ResponsiveContainer, ComposedChart, Line, XAxis, YAxis,
  CartesianGrid, Tooltip, ReferenceLine,
} from 'recharts'
import { COLORS, tooltipStyle } from '../components/ui.jsx'

export default function Forecast() {
  const [parts, setParts] = useState([])
  const [selected, setSelected] = useState('')
  const [data, setData] = useState(null)
  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState(null)
  const [showDetails, setShowDetails] = useState(false)
  const [batchBusy, setBatchBusy] = useState(false)
  const [batchMsg, setBatchMsg] = useState(null)
  const [batchResult, setBatchResult] = useState(null)
  const [activeSheet, setActiveSheet] = useState(0)
  const fileRef = useRef(null)
  const tier1Width = 160

  useEffect(() => {
    fetch('/api/v1/forecast-excel/parts').then(r => r.json())
      .then(j => {
        setParts(j.parts || [])
        if (j.parts?.length) setSelected('0')
      })
      .catch(e => setErr(e.message))
  }, [])

  const part = parts[Number(selected)] || null

  async function runForecast() {
    if (!part) return
    setBusy(true); setErr(null); setData(null)
    try {
      const r = await fetch('/api/v1/forecast-excel', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ part_number: part.part_number, tier_1: part.tier_1 }),
      })
      const j = await r.json()
      if (!r.ok) throw new Error(j.detail || 'Forecast failed')
      setData(j)
    } catch (e) { setErr(e.message) }
    finally { setBusy(false) }
  }

  useEffect(() => { if (part) runForecast() }, [selected, parts.length])

  const chart = useMemo(() => {
    if (!data) return []
    const base = {
      month: data.base_year_month, label: 'Base (actual)',
      price: data.base_price, kind: 'actual',
    }
    const rest = data.forecasts.map(f => ({
      month: f.year_month, label: f.month_label,
      price: f.predicted_price, kind: 'forecast',
      quarter: f.quarter_context.quarter_label,
    }))
    return [base, ...rest]
  }, [data])

  async function runBatch(f) {
    if (!f) return
    setBatchBusy(true); setBatchMsg(null); setBatchResult(null); setActiveSheet(0)
    const fd = new FormData()
    fd.append('file', f)
    try {
      const r = await fetch('/api/v1/forecast-batch/preview', { method: 'POST', body: fd })
      const j = await r.json().catch(() => ({}))
      if (!r.ok) throw new Error(j.detail || `Batch failed (${r.status})`)
      setBatchResult(j)
      setBatchMsg({
        ok: true,
        text: `${j.parts_count} part(s) · ${j.sheets.length} sheets generated. Preview below — download when ready.`,
      })
    } catch (e) { setBatchMsg({ ok: false, text: e.message }) }
    finally { setBatchBusy(false); if (fileRef.current) fileRef.current.value = '' }
  }

  function downloadBatch() {
    if (!batchResult) return
    const bytes = Uint8Array.from(atob(batchResult.workbook_b64), c => c.charCodeAt(0))
    const blob = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = batchResult.filename || 'forecast.xlsx'
    a.click()
    URL.revokeObjectURL(url)
  }

  const sheet = batchResult?.sheets?.[activeSheet] || null

  return (
    <div className="page">
      <section>
        <div className="section-label">12-month part price forecast · quarterly formula engine</div>
        <div className="card">
          <div className="card-head" style={{ flexWrap: 'wrap', gap: 10 }}>
            <h3>
              Part:&nbsp;
              <select className="part-select" value={selected}
                onChange={e => setSelected(e.target.value)}>
                {parts.map((p, i) => (
                  <option key={`${p.part_number}|${p.tier_1}`} value={i}>
                    {p.part_number} — {p.tier_1} ({p.weight_lbs} lbs)
                  </option>
                ))}
              </select>
            </h3>
          </div>
          <div className="card-body">
            {err && <div className="calc-alert err" style={{ marginBottom: 10 }}>{err}</div>}
            {busy && <div className="loading">Running forecast…</div>}
            {data && !busy && (
              <>
                {/* Chart + compact monthly table side by side to cut scrolling */}
                <div className="forecast-split">
                  <div className="forecast-chart">
                    <ResponsiveContainer>
                      <ComposedChart data={chart} margin={{ top: 8, right: 8, bottom: 0, left: 0 }}>
                        <CartesianGrid stroke={COLORS.grid} vertical={false} />
                        <XAxis dataKey="month" tick={{ fontSize: 10.5 }} />
                        <YAxis tick={{ fontSize: 10.5 }} width={64}
                          domain={['auto', 'auto']}
                          tickFormatter={v => '$' + v.toLocaleString('en-US')} />
                        <Tooltip {...tooltipStyle}
                          formatter={(v, n, p) => [
                            '$' + v.toLocaleString('en-US', { minimumFractionDigits: 4 }),
                            p.payload.kind === 'actual' ? 'Actual base' : `Predicted (${p.payload.quarter})`,
                          ]} />
                        <ReferenceLine x={chart[0]?.month} stroke={COLORS.navy}
                          strokeDasharray="4 4"
                          label={{ value: 'today', position: 'top', fontSize: 10.5, fill: COLORS.navy }} />
                        <Line dataKey="price" stroke={COLORS.coral} strokeWidth={2.6}
                          dot={{ r: 3, fill: COLORS.coral }} activeDot={{ r: 5 }} />
                      </ComposedChart>
                    </ResponsiveContainer>
                  </div>

                  <div className="forecast-mini">
                    <div className="sql-table-wrap" style={{ maxHeight: 300, border: '1px solid var(--gp-line)', borderRadius: 10 }}>
                      <table className="sql-table">
                        <thead><tr>
                          <th>Month</th><th>Qtr</th><th style={{ textAlign: 'right' }}>Predicted</th>
                        </tr></thead>
                        <tbody>
                          {data.forecasts.map(f => (
                            <tr key={f.year_month}>
                              <td>{f.month_label}</td>
                              <td>{f.quarter_context.quarter_label}</td>
                              <td style={{ textAlign: 'right', fontWeight: 700 }}>
                                ${f.predicted_price.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>

                <button className="btn ghost" style={{ marginTop: 12, padding: '6px 13px', fontSize: 12 }}
                  onClick={() => setShowDetails(s => !s)}>
                  {showDetails ? 'Hide' : 'Show'} full formula variables per month
                </button>
                {showDetails && (
                  <div className="sql-table-wrap" style={{ marginTop: 10, maxHeight: 320, border: '1px solid var(--gp-line)', borderRadius: 10 }}>
                    <table className="sql-table">
                      <thead><tr>
                        {['Month', 'Qtr', 'Prev Qtr', 'MC_Q', 'MC_Q-1', 'PPI_Q', 'PPI_Q-1',
                          'PPI Factor', 'CNG_Q', 'AMS_Q', 'AMS_Q-1', 'Δ AMS', 'Base used', 'Predicted'].map(h =>
                          <th key={h}>{h}</th>)}
                      </tr></thead>
                      <tbody>
                        {data.forecasts.map(f => {
                          const q = f.quarter_context
                          return (
                            <tr key={f.year_month}>
                              <td>{f.year_month}</td>
                              <td>{q.quarter_label}</td>
                              <td>{q.prev_quarter_label}</td>
                              <td>{q.mc_q.toFixed(4)}</td>
                              <td>{q.mc_q_prev.toFixed(4)}</td>
                              <td>{q.ppi_q.toFixed(2)}</td>
                              <td>{q.ppi_q_prev.toFixed(2)}</td>
                              <td>{(q.ppi_factor * 100).toFixed(4)}%</td>
                              <td>{q.cng_q.toFixed(4)}</td>
                              <td>{q.ams_q.toFixed(4)}</td>
                              <td>{q.ams_q_prev.toFixed(4)}</td>
                              <td>{q.ams_delta.toFixed(4)}</td>
                              <td>${f.base_price_used.toFixed(4)}</td>
                              <td style={{ fontWeight: 700 }}>${f.predicted_price.toFixed(4)}</td>
                            </tr>
                          )
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </>
            )}
          </div>
        </div>
      </section>

      <section>
        <div className="section-label">Batch forecast — Excel in, workbook out</div>
        <div className="card">
          <div className="card-body">
            <input ref={fileRef} type="file" accept=".xlsx"
              onChange={e => runBatch(e.target.files?.[0])} disabled={batchBusy} />
            {batchBusy && <div className="loading" style={{ padding: 12 }}>Processing batch…</div>}
            {batchMsg && (
              <div className={`calc-alert ${batchMsg.ok ? 'ok' : 'err'}`} style={{ marginTop: 10 }}>
                {batchMsg.text}
              </div>
            )}

            {batchResult && (
              <div style={{ marginTop: 14 }}>
                <div className="sheet-toolbar">
                  <div className="sheet-tabs">
                    {batchResult.sheets.map((s, i) => (
                      <button key={s.name}
                        className={`sheet-tab ${i === activeSheet ? 'active' : ''}`}
                        onClick={() => setActiveSheet(i)}>
                        {s.name}
                      </button>
                    ))}
                  </div>
                  <button className="btn" style={{ padding: '6px 14px', fontSize: 12.5, flexShrink: 0 }}
                    onClick={downloadBatch}>
                    ↓ Download .xlsx
                  </button>
                </div>

                {sheet && (() => {
                  const tier1Idx = sheet.columns.findIndex(c => /tier\s*1/i.test(String(c)))
                  const tier1Cell = (idx) => idx === tier1Idx
                    ? { maxWidth: tier1Width, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }
                    : null
                  return (
                    <div className="sql-table-wrap" style={{ marginTop: 10, maxHeight: 420, border: '1px solid var(--gp-line)', borderRadius: 10 }}>
                      <table className="sql-table">
                        <thead><tr>
                          {sheet.columns.map((c, i) => (
                            <th key={i} style={{ textAlign: i < 2 ? 'left' : 'right', ...tier1Cell(i) }}>{c}</th>
                          ))}
                        </tr></thead>
                        <tbody>
                          {sheet.rows.map((row, ri) => {
                            const isErr = row.some(v => typeof v === 'string' && v.startsWith('ERROR'))
                            return (
                              <tr key={ri} className={isErr ? 'row-error' : ''}>
                                {row.map((v, ci) => (
                                  <td key={ci}
                                    title={ci === tier1Idx ? String(v) : undefined}
                                    style={{ textAlign: ci < 2 ? 'left' : 'right', ...tier1Cell(ci) }}>{v}</td>
                                ))}
                              </tr>
                            )
                          })}
                        </tbody>
                      </table>
                    </div>
                  )
                })()}
              </div>
            )}
          </div>
        </div>
      </section>
    </div>
  )
}
