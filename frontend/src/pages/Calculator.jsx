import { useRef, useState } from 'react'
import { COLORS } from '../components/ui.jsx'

const FIELDS = [
  ['weight', 'Part Weight (PWt)', 3.6],
  ['current_price', 'Current Price (P_Current)', 47.09],
  ['ppi_q', 'PPI — Current Qtr', 281.07],
  ['ppi_q1', 'PPI — Previous Qtr', 263.8],
  ['drauss_factor', 'Drauss Factor (DF_c)', 1.44],
  ['mc_q', 'Metal Cost — Current Qtr', 3.26],
  ['mc_q_1', 'Metal Cost — Previous Qtr', 2.90],
  ['cng_q', 'CNG — Current Qtr', 0.97],
  ['cng_q_1', 'CNG — Previous Qtr', 0.72],
]

const SAMPLE = Object.fromEntries(FIELDS.map(([k, , d]) => [k, d]))

const SAMPLE_ROWS = [
  ['Part-A101', 3.6, 47.09, 281.07, 263.8, 1.44, 3.26, 2.90, 0.97, 0.72],
  ['Part-B202', 2.1, 30.00, 281.07, 263.8, 1.44, 3.26, 2.90, 0.97, 0.72],
  ['Part-C303', 5.4, 61.75, 281.07, 263.8, 1.44, 3.26, 2.90, 0.97, 0.72],
  ['Part-D404', 1.8, 22.40, 281.07, 263.8, 1.44, 3.26, 2.90, 0.97, 0.72],
]

const num = (v, d = 4) =>
  v == null || Number.isNaN(Number(v)) ? '—' : Number(v).toLocaleString('en-US', { maximumFractionDigits: d })

function downloadBlob(text, filename, type = 'text/csv') {
  const url = URL.createObjectURL(new Blob([text], { type }))
  const a = document.createElement('a')
  a.href = url; a.download = filename; a.click()
  URL.revokeObjectURL(url)
}

export default function Calculator() {
  const [form, setForm] = useState({ ...SAMPLE })
  const [result, setResult] = useState(null)
  const [calcErr, setCalcErr] = useState(null)
  const [busy, setBusy] = useState(false)
  const [showFormula, setShowFormula] = useState(false)

  const [drag, setDrag] = useState(false)
  const [batch, setBatch] = useState(null)
  const [batchErr, setBatchErr] = useState(null)
  const [batchBusy, setBatchBusy] = useState(false)
  const inputRef = useRef()

  async function calculate() {
    setBusy(true); setCalcErr(null)
    try {
      const payload = Object.fromEntries(
        Object.entries(form).map(([k, v]) => [k, parseFloat(v)]))
      if (Object.values(payload).some(v => Number.isNaN(v)))
        throw new Error('All fields must be valid numbers.')
      const res = await fetch('/api/calculator/single', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      })
      const data = await res.json()
      if (!res.ok) throw new Error(data.detail || 'Calculation failed')
      setResult(data)
    } catch (e) { setCalcErr(e.message); setResult(null) }
    finally { setBusy(false) }
  }

  async function handleFile(file) {
    if (!file) return
    setBatchBusy(true); setBatchErr(null); setBatch(null)
    try {
      const fd = new FormData()
      fd.append('file', file)
      const res = await fetch('/api/calculator/batch', { method: 'POST', body: fd })
      const data = await res.json()
      if (!res.ok) throw new Error(data.detail || 'Upload failed')
      setBatch(data)
    } catch (e) { setBatchErr(e.message) }
    finally { setBatchBusy(false) }
  }

  function downloadSample() {
    const header = 'Part_ID,PWt,P_Current,PPI_Q,PPI_Q-1,DF_c,MC_Q,MC_Q-1,CNG_Q,CNG_Q-1'
    downloadBlob([header, ...SAMPLE_ROWS.map(r => r.join(','))].join('\n'),
      'sample_parts.csv')
  }

  function downloadResults() {
    if (!batch?.results?.length) return
    const cols = Object.keys(batch.results[0])
    const esc = v => {
      const s = v == null ? '' : String(v)
      return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s
    }
    downloadBlob(
      [cols.join(','), ...batch.results.map(r => cols.map(c => esc(r[c])).join(','))].join('\n'),
      'new_price_results.csv')
  }

  const previewCols = batch?.results?.length ? Object.keys(batch.results[0]) : []

  return (
    <div className="page">
      {/* ── manual: inputs left, live result right — no scrolling ── */}
      <div className="card">
        <div className="card-head">
          <h3>Manual Calculation</h3>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
            <button className="btn ghost" style={{ padding: '6px 12px', fontSize: 12 }}
              onClick={() => setShowFormula(f => !f)}>
              {showFormula ? 'Hide formula' : 'View formula'}
            </button>
            <button className="btn ghost" style={{ padding: '6px 12px', fontSize: 12 }}
              onClick={() => { setForm({ ...SAMPLE }); setCalcErr(null) }}>
              Load sample values
            </button>
          </div>
        </div>
        <div className="card-body">
          <div className="calc-layout">
            <div>
              <div className="calc-grid three">
                {FIELDS.map(([key, label]) => (
                  <label key={key} className="calc-field">
                    <span>{label}</span>
                    <input type="number" step="any" value={form[key]}
                      onChange={e => setForm(f => ({ ...f, [key]: e.target.value }))} />
                  </label>
                ))}
              </div>
              <button className="btn" style={{ marginTop: 14 }} onClick={calculate} disabled={busy}>
                {busy ? 'Calculating…' : 'Calculate New Price'}
              </button>
              {calcErr && <div className="calc-alert err">{calcErr}</div>}
            </div>

            <div className="calc-result-panel">
              <div className="calc-main">
                <span>New Price (P_New)</span>
                <strong>{result ? num(result.new_price) : '—'}</strong>
              </div>
              <div className="calc-sub">
                <div><span>AMS_Q</span><strong>{result ? num(result.ams_q) : '—'}</strong></div>
                <div><span>AMS_Q-1</span><strong>{result ? num(result.ams_q_1) : '—'}</strong></div>
                <div><span>PPI Factor</span><strong>{result ? num(result.ppi_factor, 6) : '—'}</strong></div>
              </div>
              {showFormula && (
                <div className="calc-formula">
                  P_New = P_Current + [(AMS_Q − AMS_Q-1) × PWt] + (PPI_Factor × P_Current)<br />
                  AMS_Q = (MC_Q × DF_c) + CNG_Q<br />
                  AMS_Q-1 = (MC_Q-1 × DF_c) + CNG_Q-1<br />
                  PPI_Factor = (PPI_Q − PPI_Q-1) / PPI_Q-1
                </div>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* ── excel batch below ── */}
      <div className="card">
        <div className="card-head">
          <h3>Excel / CSV Batch</h3>
          <button className="btn ghost" style={{ padding: '6px 12px', fontSize: 12 }}
            onClick={downloadSample}>
            Download sample file
          </button>
        </div>
        <div className="card-body">
          <div
            className={`dropzone${drag ? ' drag' : ''}`}
            style={{ padding: 26 }}
            onDragOver={e => { e.preventDefault(); setDrag(true) }}
            onDragLeave={() => setDrag(false)}
            onDrop={e => { e.preventDefault(); setDrag(false); handleFile(e.dataTransfer.files[0]) }}
          >
            <h4>{batchBusy ? 'Computing…' : 'Drop your Excel or CSV here'}</h4>
            <p style={{ marginBottom: 14, fontSize: 12.5 }}>
              Columns: PWt, P_Current, PPI_Q, PPI_Q-1, MC_Q, MC_Q-1, CNG_Q, CNG_Q-1
              (DF_c optional → 1.44). Naming is flexible; extra columns like Part_ID pass through.
            </p>
            <button className="btn" onClick={() => inputRef.current.click()} disabled={batchBusy}>
              Choose file
            </button>
            <input ref={inputRef} type="file" accept=".csv,.xlsx,.xls" hidden
              onChange={e => { handleFile(e.target.files[0]); e.target.value = '' }} />
          </div>

          {batchErr && <div className="calc-alert err">{batchErr}</div>}
          {batch && (
            <div style={{ marginTop: 14 }}>
              <div className="calc-alert ok">
                <strong>{batch.filename}</strong> — {batch.count} rows calculated
                {batch.errors.length > 0 && `, ${batch.errors.length} row(s) skipped`}.
                <button className="btn" style={{ marginLeft: 12, padding: '6px 14px', fontSize: 12 }}
                  onClick={downloadResults}>
                  Download results CSV
                </button>
              </div>
              {batch.errors.length > 0 && (
                <div className="calc-alert warn">
                  {batch.errors.slice(0, 5).map(e => `Row ${e.row}: ${e.error}`).join(' · ')}
                  {batch.errors.length > 5 && ` · +${batch.errors.length - 5} more`}
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      {batch?.results?.length > 0 && (
        <div className="card">
          <div className="card-head">
            <h3>Results Preview</h3>
            <span className="hint">
              First {Math.min(batch.results.length, 50)} of {batch.count} rows — new_price appended
            </span>
          </div>
          <div className="card-body" style={{ overflowX: 'auto' }}>
            <table>
              <thead>
                <tr>{previewCols.map(c => (
                  <th key={c} style={c === 'new_price' ? { color: COLORS.coralDark } : undefined}>{c}</th>
                ))}</tr>
              </thead>
              <tbody>
                {batch.results.slice(0, 50).map((r, i) => (
                  <tr key={i}>
                    {previewCols.map(c => (
                      <td key={c} className="mono"
                        style={c === 'new_price' ? { fontWeight: 700, color: COLORS.coralDark } : undefined}>
                        {typeof r[c] === 'number' ? num(r[c]) : String(r[c] ?? '—')}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  )
}
