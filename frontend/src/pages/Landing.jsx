import { useEffect, useState } from 'react'
import { Link } from 'react-router-dom'
import { api } from '../api/client.js'
import { usd } from '../components/ui.jsx'

const FEATURES = [
  {
    to: '/dashboard', title: 'Live Dashboard',
    text: 'The six-component price equation with current values, KPI cards and 36 months of component history.',
    icon: 'M3 13h4v8H3zM10 8h4v13h-4zM17 3h4v18h-4z',
  },
  // {
  //   to: '/forecast', title: '12-Month Forecast',
  //   text: 'Monthly predictions with honest 80% confidence bands and a full month-by-month table.',
  //   icon: 'M3 17l6-6 4 4 8-8M15 7h6v6',
  // },
 
  {
    to: '/validation', title: 'Data Validation',
    text: 'Source-link health for LME, Platts and Westmetall plus automated rule checks on every series.',
    icon: 'M9 12l2 2 4-4M12 21a9 9 0 100-18 9 9 0 000 18z',
  },
  {
    to: '/calculator', title: 'Price Calculator',
    text: 'Compute P_New from PPI, metal-cost and CNG movements — one part manually, or a whole Excel at once.',
    icon: 'M4 19h16M6 16V9M10 16V5M14 16v-6M18 16V8',
  },
  
  {
    to: '/data', title: 'Data Viewer',
    text: 'View historical data',
    icon: 'M4 6c0-1.7 3.6-3 8-3s8 1.3 8 3-3.6 3-8 3-8-1.3-8-3zM4 6v12c0 1.7 3.6 3 8 3s8-1.3 8-3V6',
  },
]

const STEPS = [
  ['01', 'Validate the data', 'Source links are health-checked and every series passes rule checks — the pricing identity, no missing months, outlier flags.'],
  ['02', 'Forecast the drivers', 'Each input has history only, so the pipeline projects gas, labour, macro balance and external factors 12 months ahead first.'],
  ['03', 'Predict the price', 'A GBM ensemble turns those driver paths into a monthly all-in price forecast with 80% confidence bands.'],
]

function Icon({ d, size = 22 }) {
  return (
    <svg viewBox="0 0 24 24" width={size} height={size} fill="none" stroke="currentColor"
      strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d={d} /></svg>
  )
}

export default function Landing() {
  const [s, setS] = useState(null)

  useEffect(() => { api.summary().then(setS).catch(() => {}) }, [])

  return (
    <div className="landing">
      {/* hero */}
      <section className="hero">
        <div className="hero-inner">
          <span className="hero-kicker">Metals Analytics · Aluminum</span>
          <h1>
            Know your aluminum price,<br />
            <em>twelve months before you pay it.</em>
          </h1>
          <p className="hero-sub">
            The all-in US aluminum price is more than the exchange quote — it's LME plus the
            Midwest Premium, gas, labour, macro balance and external factors. This platform
            forecasts that full cost, month by month, with honest uncertainty bands.
          </p>
          <div className="hero-actions">
            <Link to="/dashboard" className="btn">Open Dashboard</Link>

          </div>

    
        </div>
      </section>

      {/* formula band */}
      <section className="formula-band">
        <span className="fb-item">Six connected views, one consistent price identity</span><span className="fb-op"></span>

      </section>

      {/* features */}
      <section className="section">
        <h2>Everything the engagement covers, in one place</h2>
       
        <div className="feature-grid">
          {FEATURES.map(f => (
            <Link to={f.to} className="feature-card" key={f.to}>
              <span className="feature-icon"><Icon d={f.icon} /></span>
              <h3>{f.title}</h3>
              <p>{f.text}</p>
              <span className="feature-more">Open →</span>
            </Link>
          ))}
        </div>
      </section>

      {/* how it works */}
      {/* <section className="section steps-section">
        <h2>How the forecast is built</h2>
        <p className="section-sub">A two-step approach, because the inputs only have history.</p>
        <div className="steps">
          {STEPS.map(([n, t, d]) => (
            <div className="step" key={n}>
              <span className="step-n">{n}</span>
              <h3>{t}</h3>
              <p>{d}</p>
            </div>
          ))}
        </div>
      </section> */}

      {/* CTA */}
      {/* <section className="cta">
        <h2>Have questions? Just ask.</h2>
        <p>
          The assistant in the corner answers questions about the formula, forecasts and bands —
          it works out of the box, and connects to Azure OpenAI for full AI answers when a key
          is provided.
        </p>
        <Link to="/dashboard" className="btn cta-btn">Explore the Platform</Link>
      </section> */}
    </div>
  )
}
