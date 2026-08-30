import { useEffect, useState } from 'react'
import {
  ResponsiveContainer, ComposedChart, Line, Bar, XAxis, YAxis,
  CartesianGrid, Tooltip, Legend,
} from 'recharts'
import { COLORS, Loading, tooltipStyle } from '../components/ui.jsx'

const VARIANCE_FILL = 'rgba(180, 83, 9, 0.85)'

export default function Validation() {
  const [cmpRange, setCmpRange] = useState(60)
  const [cmp, setCmp] = useState(null)

  useEffect(() => {
    const q = cmpRange ? `?months=${cmpRange}` : ''
    fetch(`/api/comparison${q}`).then(r => r.json()).then(setCmp).catch(() => {})
  }, [cmpRange])

  const cmpData = cmp
    ? cmp.labels.map((m, i) => ({
        month: m,
        'FRED (PALUMUSDM)': cmp.fred[i],
        'LME 3-month': cmp.lme_3m[i],
        'Variance %': cmp.variance_pct[i],
      }))
    : []

  return (
    <div className="page">
      <div className="card">
        <div className="card-head">
          <h3>Source Cross-Check — FRED vs LME 3-month</h3>
          <div className="seg" role="tablist" aria-label="Comparison range">
            {[[24, '24M'], [60, '5Y'], [null, 'All']].map(([m, label]) => (
              <button key={label} className={cmpRange === m ? 'active' : ''}
                onClick={() => setCmpRange(m)}>{label}</button>
            ))}
          </div>
        </div>
        <div className="card-body">
          {!cmp ? <Loading label="Loading comparison…" /> : (
            <div style={{ height: 380 }}>
              <ResponsiveContainer>
                <ComposedChart data={cmpData}>
                  <CartesianGrid stroke={COLORS.grid} vertical={false} />
                  <XAxis dataKey="month" tick={{ fontSize: 10 }} minTickGap={34} />
                  <YAxis yAxisId="p" tick={{ fontSize: 10.5 }} width={58}
                    domain={['auto', 'auto']}
                    tickFormatter={v => '$' + v.toLocaleString('en-US')} />
                  <YAxis yAxisId="v" orientation="right" tick={{ fontSize: 10.5 }} width={44}
                    tickFormatter={v => v + '%'} />
                  <Tooltip {...tooltipStyle}
                    formatter={(v, name) => name === 'Variance %'
                      ? [v?.toFixed(2) + '%', name]
                      : ['$' + v?.toLocaleString('en-US'), name]} />
                  <Legend wrapperStyle={{ fontSize: 11.5 }} />
                  <Bar yAxisId="v" dataKey="Variance %" fill={VARIANCE_FILL} barSize={7} />
                  <Line yAxisId="p" dataKey="FRED (PALUMUSDM)" stroke={COLORS.coral}
                    strokeWidth={2.1} dot={false} />
                  <Line yAxisId="p" dataKey="LME 3-month" stroke={COLORS.navy}
                    strokeWidth={2.1} dot={false} />
                </ComposedChart>
              </ResponsiveContainer>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
