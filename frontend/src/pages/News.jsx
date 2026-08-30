import { useEffect, useRef, useState } from 'react'
import {
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
} from 'recharts'
import { COLORS, tooltipStyle } from '../components/ui.jsx'

const FACTOR_LABELS = {
  lme: 'LME', midwest_premium: 'Premium', gas: 'Gas / Energy',
  labour: 'Labour', macro: 'Supply−Demand', external: 'Tariffs / External',
}
const FACTOR_BADGE = {
  lme: 'info', midwest_premium: 'warn', gas: 'err',
  labour: 'ok', macro: 'info', external: 'warn',
}

const NEWS_SUGGESTIONS = [
  'Summarize aluminum news for 2026-Q2',
  'What are the future trends for aluminum prices?',
  'News about tariffs in April 2026',
  'Yearly summary of 2026 so far',
]

export default function News() {
  const [data, setData] = useState(null)
  const [stats, setStats] = useState(null)
  const [factor, setFactor] = useState(null)
  const [refreshing, setRefreshing] = useState(false)
  const [refreshMsg, setRefreshMsg] = useState(null)
  const [bfStart, setBfStart] = useState('2022-01')
  const [bfEnd, setBfEnd] = useState('2024-12')
  const [backfilling, setBackfilling] = useState(false)

  // news AI panel
  const [thread, setThread] = useState([])
  const [q, setQ] = useState('')
  const [busy, setBusy] = useState(false)
  const threadRef = useRef(null)

  function load(f = factor) {
    const params = f ? `?factor=${f}&limit=200` : '?limit=200'
    fetch(`/api/news${params}`).then(r => r.json()).then(setData).catch(() => {})
  }

  useEffect(() => { load() }, [factor])
  useEffect(() => {
    fetch('/api/news/stats').then(r => r.json()).then(setStats).catch(() => {})
  }, [refreshMsg])
  useEffect(() => {
    if (threadRef.current) threadRef.current.scrollTop = threadRef.current.scrollHeight
  }, [thread, busy])

  async function refresh() {
    setRefreshing(true); setRefreshMsg(null)
    try {
      const r = await fetch('/api/news/refresh', { method: 'POST' })
      const j = await r.json()
      setRefreshMsg(j.message)
      load()
    } catch { setRefreshMsg('Refresh failed — kept existing articles.') }
    finally { setRefreshing(false) }
  }

  async function backfill() {
    setBackfilling(true); setRefreshMsg(null)
    try {
      const r = await fetch('/api/news/backfill', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ start: bfStart, end: bfEnd }),
      })
      const j = await r.json()
      setRefreshMsg(j.message || j.detail || 'Backfill finished.')
      load()
    } catch { setRefreshMsg('Backfill failed — existing articles kept.') }
    finally { setBackfilling(false) }
  }

  async function ask(text) {
    const question = (text ?? q).trim()
    if (!question || busy) return
    setThread(t => [...t, { role: 'user', content: question }])
    setQ(''); setBusy(true)
    try {
      const r = await fetch('/api/news/chat', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ messages: [{ role: 'user', content: question }] }),
      })
      const j = await r.json()
      setThread(t => [...t, { role: 'assistant', content: j.reply,
        articles: j.articles, mode: j.mode }])
    } catch (e) {
      setThread(t => [...t, { role: 'assistant', content: `⚠️ ${e.message}`, error: true }])
    } finally { setBusy(false) }
  }

  const chart = stats?.monthly_counts?.map(r => ({ month: r.month, Articles: r.count })) || []
  const isLive = data?.articles?.some(a => a.origin?.startsWith('live'))

  return (
    <div className="page">
      {/* bulletin */}
      {data?.bulletin && (
        <a className="bulletin" href={data.bulletin.url} target="_blank" rel="noreferrer">
          <span className="bulletin-tag">Latest</span>
          <span className="bulletin-text">
            {data.bulletin.title} — {data.bulletin.source} · {data.bulletin.published}
          </span>
          <span className="bulletin-arrow">→</span>
        </a>
      )}

      <div className="grid-2">
        {/* chart */}
        <div className="card">
          <div className="card-head">
            <h3>Coverage that could move the price</h3>
            <span className="hint">Articles per month, all factors</span>
          </div>
          <div className="card-body" style={{ height: 260 }}>
            <ResponsiveContainer>
              <BarChart data={chart}>
                <CartesianGrid stroke={COLORS.grid} vertical={false} />
                <XAxis dataKey="month" tick={{ fontSize: 10.5 }} minTickGap={20} />
                <YAxis allowDecimals={false} tick={{ fontSize: 11 }} width={30} />
                <Tooltip {...tooltipStyle} />
                <Bar dataKey="Articles" fill={COLORS.coral} radius={[5, 5, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* news AI */}
        <div className="card news-ai">
          <div className="card-head">
            <h3>News Intelligence AI</h3>
            <span className="hint">gpt-4.1-mini · news &amp; trends only</span>
          </div>
          <div className="card-body" style={{ display: 'flex', flexDirection: 'column', minHeight: 260 }}>
            <div className="news-thread" ref={threadRef}>
              {thread.length === 0 && (
                <div className="chat-suggestions" style={{ marginTop: 0 }}>
                  {NEWS_SUGGESTIONS.map(s => (
                    <button key={s} onClick={() => ask(s)}>{s}</button>
                  ))}
                </div>
              )}
              {thread.map((m, i) => (
                <div key={i} style={{ display: 'flex', flexDirection: 'column',
                  alignItems: m.role === 'user' ? 'flex-end' : 'flex-start' }}>
                  <div className={`chat-msg ${m.role}${m.error ? ' err' : ''}`}>{m.content}</div>
                  {m.articles?.rows?.length > 0 && (
                    <div className="sql-table-wrap news-ai-table">
                      <table className="sql-table">
                        <thead><tr>
                          <th>#</th><th>date</th><th>title</th><th>source</th><th>link</th>
                        </tr></thead>
                        <tbody>
                          {m.articles.rows.slice(0, 12).map((r, ri) => (
                            <tr key={ri}>
                              <td>[{ri + 1}]</td><td>{r[0]}</td>
                              <td style={{ whiteSpace: 'normal', minWidth: 180 }}>{r[1]}</td>
                              <td>{r[2]}</td>
                              <td><a className="news-link" href={r[4]} target="_blank"
                                rel="noreferrer">Open ↗</a></td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              ))}
              {busy && <div className="chat-msg assistant typing"><span /><span /><span /></div>}
            </div>
            <div className="chat-input" style={{ padding: '10px 0 0', borderTop: '1px solid var(--gp-line)' }}>
              <textarea rows={1} placeholder="Ask about any month, quarter, year or trends…"
                value={q} onChange={e => setQ(e.target.value)}
                onKeyDown={e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); ask() } }} />
              <button className="btn" onClick={() => ask()} disabled={busy || !q.trim()}>Ask</button>
            </div>
          </div>
        </div>
      </div>

      {/* article table */}
      <div className="card">
        <div className="card-head">
          <h3>Factor News Feed</h3>
          <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
            <span className={`badge ${isLive ? 'ok' : 'warn'}`}>
              {isLive ? 'live: Google News RSS' : 'curated sample — refresh for live'}
            </span>
            <button className="btn ghost" style={{ padding: '6px 12px', fontSize: 12 }}
              onClick={refresh} disabled={refreshing}>
              {refreshing ? 'Fetching…' : 'Refresh live news'}
            </button>
          </div>
        </div>
        <div className="card-body">
          <div className="news-filters">
            <button className={`chip${!factor ? ' active' : ''}`}
              onClick={() => setFactor(null)}>All factors</button>
            {Object.entries(FACTOR_LABELS).map(([k, label]) => (
              <button key={k} className={`chip${factor === k ? ' active' : ''}`}
                onClick={() => setFactor(k)}>{label}</button>
            ))}
          </div>
          <div className="backfill-bar">
            <span className="backfill-label">Historical backfill (GDELT · free · 2017 onward):</span>
            <input type="month" min="2017-01" max="2026-12" value={bfStart}
              onChange={e => setBfStart(e.target.value)} />
            <span style={{ color: 'var(--gp-slate)' }}>to</span>
            <input type="month" min="2017-01" max="2026-12" value={bfEnd}
              onChange={e => setBfEnd(e.target.value)} />
            <button className="btn ghost" style={{ padding: '6px 12px', fontSize: 12 }}
              onClick={backfill} disabled={backfilling}>
              {backfilling ? 'Backfilling… (can take a minute)' : 'Backfill history'}
            </button>
          </div>
          {refreshMsg && <div className="calc-alert warn" style={{ marginBottom: 10 }}>{refreshMsg}</div>}
          <div style={{ overflowX: 'auto' }}>
            <table className="news-table">
              <thead>
                <tr><th>Date</th><th>Article &amp; summary</th><th>Factors</th>
                  <th>Source</th><th>Link</th></tr>
              </thead>
              <tbody>
                {(data?.articles || []).map(a => (
                  <tr key={a.id}>
                    <td className="mono" style={{ whiteSpace: 'nowrap' }}>{a.published}</td>
                    <td>
                      <div className="news-title">{a.title}</div>
                      <div className="news-summary">{a.summary}</div>
                    </td>
                    <td>
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                        {a.factors.split(',').map(f => (
                          <span key={f} className={`badge ${FACTOR_BADGE[f] || 'info'}`}
                            style={{ fontSize: 10.5 }}>{FACTOR_LABELS[f] || f}</span>
                        ))}
                      </div>
                    </td>
                    <td style={{ whiteSpace: 'nowrap' }}>{a.source}</td>
                    <td><a className="news-link" href={a.url} target="_blank"
                      rel="noreferrer">Read ↗</a></td>
                  </tr>
                ))}
              </tbody>
            </table>
            {(!data || data.count === 0) && (
              <div className="loading" style={{ padding: 30 }}>No articles for this filter.</div>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}
