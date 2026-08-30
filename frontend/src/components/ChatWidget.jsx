import { useEffect, useRef, useState } from 'react'

const HELLO = {
  role: 'assistant',
  content: "Hi! I'm the multi-agent assistant. I can query the factor databases, explain why the price moved on a given date, search the web for geopolitical context, and give the model's outlook. Try a suggestion below.",
}

const SUGGESTIONS = [
  'Why did LME vary in 2026-08?',
  'Compare all factors for 2026-03',
  'Top 5 months by Midwest premium',
  'PPI change wrt previous quarter',
  'Average LME price per year',
]

const AGENT_ICONS = {
  router: '🧭', sql: '🗄️', events: '📌', web: '🌐',
  forecaster: '🔮', synthesizer: '✍️',
}

export default function ChatWidget() {
  const [open, setOpen] = useState(false)
  const [minimized, setMinimized] = useState(false)
  const [full, setFull] = useState(false)
  const [status, setStatus] = useState(null)
  const [messages, setMessages] = useState([HELLO])
  const [input, setInput] = useState('')
  const [busy, setBusy] = useState(false)
  const bodyRef = useRef(null)

  useEffect(() => {
    fetch('/api/agent/status').then(r => r.json())
      .then(setStatus)
      .catch(() => setStatus({ llm_mode: 'none', web_search_installed: false }))
  }, [])

  useEffect(() => {
    if (bodyRef.current) bodyRef.current.scrollTop = bodyRef.current.scrollHeight
  }, [messages, open, busy, minimized, full])

  async function send(textOverride) {
    const text = (textOverride ?? input).trim()
    if (!text || busy) return
    const next = [...messages, { role: 'user', content: text }]
    setMessages([...next, { role: 'assistant', content: '', streaming: true }])
    setInput('')
    setBusy(true)
    try {
      const res = await fetch('/api/agent/chat/stream', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          messages: next.filter(m => m !== HELLO)
            .map(({ role, content }) => ({ role, content })),
        }),
      })
      if (!res.ok || !res.body) {
        const data = await res.json().catch(() => ({}))
        throw new Error(data.detail || `Request failed (${res.status})`)
      }
      const reader = res.body.getReader()
      const decoder = new TextDecoder()
      let buffer = ''
      const patch = (fn) => setMessages(m => {
        const copy = [...m]
        copy[copy.length - 1] = fn(copy[copy.length - 1])
        return copy
      })
      while (true) {
        const { done, value } = await reader.read()
        if (done) break
        buffer += decoder.decode(value, { stream: true })
        let idx
        while ((idx = buffer.indexOf('\n\n')) !== -1) {
          const chunk = buffer.slice(0, idx).trim()
          buffer = buffer.slice(idx + 2)
          if (!chunk.startsWith('data: ')) continue
          let ev
          try { ev = JSON.parse(chunk.slice(6)) } catch { continue }
          if (ev.type === 'meta') {
            patch(msg => ({ ...msg, trace: ev.trace, citations: ev.citations,
              mode: ev.mode, sqlTable: ev.sql_table }))
          } else if (ev.type === 'delta') {
            patch(msg => ({ ...msg, content: msg.content + ev.text }))
          } else if (ev.type === 'done') {
            patch(msg => ({ ...msg, streaming: false }))
          }
        }
      }
      patch(msg => ({ ...msg, streaming: false }))
    } catch (e) {
      setMessages(m => {
        const copy = [...m]
        const last = copy[copy.length - 1]
        copy[copy.length - 1] = last?.streaming
          ? { role: 'assistant', content: `⚠️ ${e.message}`, error: true }
          : last
        if (!last?.streaming) copy.push({ role: 'assistant', content: `⚠️ ${e.message}`, error: true })
        return copy
      })
    } finally {
      setBusy(false)
    }
  }

  async function sendFeedback(i, rating) {
    const msg = messages[i]
    const q = [...messages].slice(0, i).reverse().find(m => m.role === 'user')?.content || ''
    setMessages(ms => ms.map((m, j) => (j === i ? { ...m, feedback: rating } : m)))
    try {
      await fetch('/api/agent/feedback', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rating, question: q.slice(0, 1900),
          answer: (msg.content || '').slice(0, 3900), mode: msg.mode || '' }),
      })
    } catch { /* stored locally in UI regardless */ }
  }

  function exportChat() {
    const lines = messages.map(m =>
      `${m.role === 'user' ? 'You' : 'Assistant'}:\n${m.content}\n`)
    const stamp = new Date().toISOString().slice(0, 16).replace('T', ' ')
    const text = `Aluminum Price Intelligence — chat export (${stamp})\n\n` + lines.join('\n')
    const url = URL.createObjectURL(new Blob([text], { type: 'text/plain' }))
    const a = document.createElement('a')
    a.href = url; a.download = 'assistant-chat.txt'; a.click()
    URL.revokeObjectURL(url)
  }

  const llmLabel = !status ? 'Checking…'
    : status.llm_mode === 'openai' ? 'GPT-4o-mini · OpenAI'
    : status.llm_mode === 'azure' ? 'GPT-4o-mini · Azure'
    : 'Basic mode · DB + rules (no key)'

  const panelCls = `chat-panel${full ? ' fullscreen' : ''}${minimized ? ' minimized' : ''}`

  return (
    <>
      {!open && (
        <button className="chat-fab" aria-label="Open assistant" onClick={() => setOpen(true)}>
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.9"
            strokeLinecap="round" strokeLinejoin="round">
            <path d="M21 12a8 8 0 01-8 8H5l-2 2V12a8 8 0 018-8h2a8 8 0 018 8z" />
          </svg>
        </button>
      )}

      {open && (
        <div className={panelCls} role="dialog" aria-label="Multi-agent assistant">
          <div className="chat-head">
            <div style={{ minWidth: 0 }}>
              <div className="chat-title">Multi-Agent Assistant</div>
              {!minimized && <div className="chat-sub">{llmLabel}</div>}
            </div>
            <div className="chat-controls">
              <span className={`chat-dot ${status?.llm_mode && status.llm_mode !== 'none' ? 'on' : 'off'}`} />
              <button className="chat-ctl" title="Export chat" aria-label="Export chat"
                onClick={exportChat}>
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"
                  strokeLinecap="round" strokeLinejoin="round">
                  <path d="M12 3v12M7 10l5 5 5-5M4 21h16" /></svg>
              </button>
              <button className="chat-ctl" title={minimized ? 'Restore' : 'Minimize'}
                aria-label={minimized ? 'Restore' : 'Minimize'}
                onClick={() => { setMinimized(m => !m); if (full) setFull(false) }}>
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2"
                  strokeLinecap="round"><path d={minimized ? 'M7 14l5-5 5 5' : 'M5 12h14'} /></svg>
              </button>
              <button className="chat-ctl" title={full ? 'Exit full screen' : 'Full screen'}
                aria-label={full ? 'Exit full screen' : 'Full screen'}
                onClick={() => { setFull(f => !f); setMinimized(false) }}>
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"
                  strokeLinecap="round" strokeLinejoin="round">
                  {full
                    ? <path d="M9 3v6H3M15 3v6h6M9 21v-6H3M15 21v-6h6" />
                    : <path d="M3 9V3h6M21 9V3h-6M3 15v6h6M21 15v6h-6" />}
                </svg>
              </button>
              <button className="chat-ctl" title="Close" aria-label="Close"
                onClick={() => { setOpen(false); setFull(false); setMinimized(false) }}>
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2"
                  strokeLinecap="round"><path d="M6 6l12 12M18 6L6 18" /></svg>
              </button>
            </div>
          </div>

          {!minimized && (
            <>
              {status && status.llm_mode === 'none' && (
                <div className="chat-notice">
                  Agents are running on databases + rules. Add <span className="mono">OPENAI_API_KEY</span>{' '}
                  (or Azure vars) to <span className="mono">backend/.env</span> for GPT-4o-mini answers.
                </div>
              )}

              <div className="chat-body" ref={bodyRef}>
                {messages.map((m, i) => (
                  <div key={i} style={{ display: 'flex', flexDirection: 'column',
                    alignItems: m.role === 'user' ? 'flex-end' : 'flex-start' }}>
                    {m.trace && (
                      <div className="trace-row">
                        {m.trace.map((t, j) => (
                          <span key={j} className="trace-chip" title={t.detail}>
                            {AGENT_ICONS[t.agent] || '•'} {t.agent}
                          </span>
                        ))}
                      </div>
                    )}
                    <div className={`chat-msg ${m.role}${m.error ? ' err' : ''}`}>
                      {m.content}
                      {m.streaming && <span className="stream-cursor">▍</span>}
                    </div>
                    {m.sqlTable?.sql && !m.streaming && (
                      <div className="sql-block">
                        <div className="sql-head">
                          <button className="sql-copy"
                            onClick={() => setSqlOpen(o => ({ ...o, [i]: !o[i] }))}>
                            {sqlOpen[i] ? 'Hide SQL' : 'View SQL'}
                          </button>
                          <span className="hint" style={{ color: '#8fa3af', fontSize: 10 }}>
                            {m.sqlTable.generator}
                          </span>
                          {sqlOpen[i] && (
                            <button className="sql-copy"
                              onClick={() => navigator.clipboard?.writeText(m.sqlTable.sql)}>
                              Copy
                            </button>
                          )}
                        </div>
                        {sqlOpen[i] && <pre>{m.sqlTable.sql}</pre>}
                        {m.sqlTable.error ? (
                          <div className="sql-error">{m.sqlTable.error}</div>
                        ) : m.sqlTable.columns?.length > 0 && (
                          <div className="sql-table-wrap">
                            <table className="sql-table">
                              <thead><tr>
                                {m.sqlTable.columns.map(c => <th key={c}>{c}</th>)}
                              </tr></thead>
                              <tbody>
                                {m.sqlTable.rows.slice(0, 25).map((row, ri) => (
                                  <tr key={ri}>
                                    {row.map((v, ci) => (
                                      <td key={ci}>
                                        {typeof v === 'number'
                                          ? v.toLocaleString('en-US', { maximumFractionDigits: 2 })
                                          : String(v ?? '—')}
                                      </td>
                                    ))}
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                            {m.sqlTable.rows.length > 25 && (
                              <div className="sql-more">
                                +{m.sqlTable.rows.length - 25} more rows (export chat for full text)
                              </div>
                            )}
                          </div>
                        )}
                      </div>
                    )}
                    {m.citations?.length > 0 && !m.streaming && (
                      <div className="cite-row">
                        {m.citations.map((c, k) => (
                          <a key={k} href={c.url} target="_blank" rel="noreferrer"
                            className="cite-chip" title={c.url}>
                            🔗 {c.title || new URL(c.url).hostname}
                          </a>
                        ))}
                      </div>
                    )}
                    {m.role === 'assistant' && !m.streaming && !m.error && i > 0 && (
                      <div className="fb-row">
                        {m.feedback
                          ? <span className="fb-done">
                              {m.feedback === 'up' ? '👍' : '👎'} Thanks for the feedback
                            </span>
                          : <>
                              <button className="fb-btn" title="Good answer"
                                onClick={() => sendFeedback(i, 'up')}>👍</button>
                              <button className="fb-btn" title="Bad answer"
                                onClick={() => sendFeedback(i, 'down')}>👎</button>
                            </>}
                      </div>
                    )}
                  </div>
                ))}
                {busy && (
                  <div className="chat-msg assistant typing"><span /><span /><span /></div>
                )}
                {messages.length === 1 && !busy && (
                  <div className="chat-suggestions">
                    {SUGGESTIONS.map(s => (
                      <button key={s} onClick={() => send(s)}>{s}</button>
                    ))}
                  </div>
                )}
              </div>

              <div className="chat-input">
                <textarea
                  rows={1}
                  placeholder="Ask about price movements, outlook, history…"
                  value={input}
                  onChange={e => setInput(e.target.value)}
                  onKeyDown={e => {
                    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send() }
                  }}
                />
                <button className="btn" onClick={() => send()} disabled={busy || !input.trim()}>
                  Send
                </button>
              </div>
            </>
          )}
        </div>
      )}
    </>
  )
}
