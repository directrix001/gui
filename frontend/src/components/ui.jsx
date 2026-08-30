export const COLORS = {
  coral: '#ff555f',
  coralDark: '#e13d47',
  navy: '#041c2c',
  slate: '#5b6b76',
  teal: '#0fb5a6',
  blue: '#2f6fed',
  amber: '#e8a13c',
  band: 'rgba(255, 85, 95, 0.12)',
  grid: '#e5e9ec',
}

export function Loading({ label = 'Loading data…' }) {
  return (
    <div className="loading">
      <div className="spinner" />
      {label}
    </div>
  )
}

export function ErrorBox({ error }) {
  return (
    <div className="card" style={{ padding: 24 }}>
      <strong style={{ color: COLORS.coralDark }}>Couldn't reach the API.</strong>
      <div style={{ color: COLORS.slate, marginTop: 6, fontSize: 13 }}>
        {String(error?.message || error)} — make sure the FastAPI backend is running on port 8000
        (<span className="mono">uvicorn app.main:app --reload</span>).
      </div>
    </div>
  )
}

export function Delta({ value }) {
  const cls = value > 0.05 ? 'up' : value < -0.05 ? 'down' : 'flat'
  const arrow = value > 0.05 ? '▲' : value < -0.05 ? '▼' : '—'
  return <div className={`delta ${cls}`}>{arrow} {Math.abs(value).toFixed(2)}% MoM</div>
}

export const usd = (v) =>
  v == null ? '—' : `$${Number(v).toLocaleString('en-US', { maximumFractionDigits: 0 })}`

export const tooltipStyle = {
  contentStyle: {
    borderRadius: 10, border: `1px solid ${COLORS.grid}`,
    fontSize: 12.5, fontFamily: 'Inter, sans-serif',
    boxShadow: '0 8px 24px rgba(4,28,44,0.12)',
  },
}
