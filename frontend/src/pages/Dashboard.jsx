import { useEffect, useMemo, useState } from 'react'
import {
  ResponsiveContainer, LineChart, Line, XAxis, YAxis,
  CartesianGrid, Tooltip, Legend, ReferenceLine,
} from 'recharts'
import { COLORS, tooltipStyle } from '../components/ui.jsx'

const NOW = `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}`

function splitHistFcst(rows, keys) {
  return rows.map(r => {
    const o = { month: r.month }
    for (const k of keys) {
      o[`${k}_hist`] = r.month <= NOW ? r[k] : null
      o[`${k}_fcst`] = r.month >= NOW ? r[k] : null
    }
    return o
  })
}

function ChartCard({ title, range, data, lines, unit }) {
  return (
    <div className="card">
      <div className="card-head">
        <h3>{title}</h3>
        <span className="hint">{range}</span>
      </div>
      <div className="card-body" style={{ height: 250 }}>
        <ResponsiveContainer>
          <LineChart data={data}>
            <CartesianGrid stroke={COLORS.grid} vertical={false} />
            <XAxis dataKey="month" tick={{ fontSize: 10 }} minTickGap={28} />
            <YAxis tick={{ fontSize: 10.5 }} width={54} domain={['auto', 'auto']}
              tickFormatter={v => v.toLocaleString('en-US', { maximumFractionDigits: 2 })} />
            <Tooltip {...tooltipStyle}
              formatter={(v, name) => [
                v?.toLocaleString('en-US', { maximumFractionDigits: 4 }) + (unit ? ` ${unit}` : ''),
                name,
              ]} />
            {lines.length > 1 && <Legend wrapperStyle={{ fontSize: 11.5 }} />}
            <ReferenceLine x={NOW} stroke={COLORS.navy} strokeDasharray="4 4"
              label={{ value: 'today', position: 'top', fontSize: 10, fill: COLORS.navy }} />
            {lines.map(l => (
              <Line key={l.key + '_hist'} name={l.name} dataKey={`${l.key}_hist`}
                stroke={l.color} strokeWidth={2.2} dot={false} connectNulls={false} />
            ))}
            {lines.map(l => (
              <Line key={l.key + '_fcst'} name={`${l.name} (forward)`} dataKey={`${l.key}_fcst`}
                stroke={l.color} strokeWidth={2.2} strokeDasharray="5 4" dot={false}
                legendType="none" connectNulls={false} />
            ))}
          </LineChart>
        </ResponsiveContainer>
      </div>
    </div>
  )
}

export default function Dashboard() {
  const [drivers, setDrivers] = useState(null)
  const [err, setErr] = useState(null)

  useEffect(() => {
    fetch('/api/real/drivers').then(r => r.json())
      .then(j => setDrivers(j.drivers)).catch(e => setErr(e.message))
  }, [])

  const byKey = useMemo(() => {
    if (!drivers) return {}
    return Object.fromEntries(drivers.map(d => [d.key, d]))
  }, [drivers])

  const merged = useMemo(() => {
    if (!drivers) return []
    const months = {}
    for (const d of drivers) {
      for (const p of d.series) {
        months[p.month] = months[p.month] || { month: p.month }
        months[p.month][d.key] = p.value
      }
    }
    return Object.values(months).sort((a, b) => a.month.localeCompare(b.month))
      .map(r => ({ ...r, total: r.lme != null && r.midwest_premium != null
        ? +(r.lme + r.midwest_premium).toFixed(4) : null }))
  }, [drivers])

  if (err) return <div className="loading">⚠️ {err}</div>
  if (!drivers) return <div className="loading">Loading…</div>

  const rangeOf = k => byKey[k] ? `${byKey[k].from} → ${byKey[k].to}` : ''
  const lmMw = splitHistFcst(merged, ['lme', 'midwest_premium'])
  const total = splitHistFcst(merged, ['total'])
  const cng = splitHistFcst(merged, ['gas'])
  const ppi = splitHistFcst(merged, ['labour'])

  return (
    <div className="page">
      <div className="grid-2">
        <ChartCard title="LME & Midwest Premium ($/lb)" range={rangeOf('lme')}
          data={lmMw} unit="$/lb"
          lines={[
            { key: 'lme', name: 'LME', color: COLORS.coral },
            { key: 'midwest_premium', name: 'Midwest Premium', color: '#B45309' },
          ]} />
        <ChartCard title="Total: LME + Midwest ($/lb)" range={rangeOf('lme')}
          data={total} unit="$/lb"
          lines={[{ key: 'total', name: 'LME + Midwest', color: COLORS.navy }]} />
        <ChartCard title="CNG Cost ($/lb)" range={rangeOf('gas')}
          data={cng} unit="$/lb"
          lines={[{ key: 'gas', name: 'CNG', color: '#0E7490' }]} />
        <ChartCard title="PPI Index" range={rangeOf('labour')}
          data={ppi} unit=""
          lines={[{ key: 'labour', name: 'PPI', color: '#5B21B6' }]} />
      </div>
    </div>
  )
}
