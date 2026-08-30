import { NavLink, Link, Routes, Route, useLocation, Navigate } from 'react-router-dom'
import Landing from './pages/Landing.jsx'
import Dashboard from './pages/Dashboard.jsx'
import Forecast from './pages/Forecast.jsx'
import Validation from './pages/Validation.jsx'
import Calculator from './pages/Calculator.jsx'
import DataManager from './pages/DataManager.jsx'
import News from './pages/News.jsx'
import ChatWidget from './components/ChatWidget.jsx'

const NAV = [
  { to: '/', label: 'Home', end: true },
  { to: '/dashboard', label: 'Dashboard' },
  { to: '/calculator', label: 'Cost Estimation' },
  { to: '/validation', label: 'Validator' },
  // { to: '/forecast', label: 'Forecast' },
  { to: '/data', label: 'Data' },
]

const TITLES = {
  '/dashboard': ['Dashboard', 'Real driver data — history and forward months'],
  '/forecast': ['12-Month Forecast', 'Quarterly formula engine — per Part Number + Tier 1'],
  '/validation': ['Validator', 'Source cross-check and rule checks'],
  '/calculator': ['Cost Estimation', 'Compute P_New manually or from an Excel/CSV batch'],
  '/data': ['Data', 'All four driver series — searchable by month'],
  '/news': ['Factor News & Intelligence', 'Open-source coverage of every aluminum price driver — with its own News AI'],
}

function GenpactLogo({ light, height = 22 }) {
  const fill = light ? '#FFFFFF' : '#041C2C'
  return (
    <svg height={height} viewBox="0 0 132 30" role="img" aria-label="Genpact"
      xmlns="http://www.w3.org/2000/svg">
      <text x="0" y="22" fontFamily="'Space Grotesk', Inter, sans-serif"
        fontWeight="700" fontSize="24" letterSpacing="-0.5" fill={fill}>genpact</text>
      <circle cx="124" cy="21" r="3.4" fill="#FF555F" />
    </svg>
  )
}

function Brand({ light }) {
  return (
    <Link to="/" className="brand-inline" aria-label="Home">
      <GenpactLogo light={light} />
      <span className="brand-tag">Aluminum Price Intelligence</span>
    </Link>
  )
}

export default function App() {
  const { pathname } = useLocation()
  const meta = TITLES[pathname]

  return (
    <div className="site">
      <header className="site-header">
        <Brand light />
        <nav className="site-nav" aria-label="Main">
          {NAV.map(n => (
            <NavLink key={n.to} to={n.to} end={n.end}
              className={({ isActive }) => `site-link${isActive ? ' active' : ''}`}>
              {n.label}
            </NavLink>
          ))}
        </nav>
      </header>

      {meta && (
        <div className="page-head">
          <div className="page-head-inner">
            <h1>{meta[0]}</h1>
            <span>{meta[1]}</span>
          </div>
        </div>
      )}

      <main className="site-main">
        <Routes>
          <Route path="/" element={<Landing />} />
          <Route path="/dashboard" element={<Dashboard />} />
          <Route path="/forecast" element={<Forecast />} />
          <Route path="/drivers" element={<Navigate to="/dashboard" replace />} />
          <Route path="/validation" element={<Validation />} />
          <Route path="/calculator" element={<Calculator />} />
          <Route path="/model" element={<Calculator />} />
          <Route path="/data" element={<DataManager />} />
          <Route path="/news" element={<News />} />
          <Route path="*" element={<Landing />} />
        </Routes>
      </main>

      <footer className="site-footer">
        <div className="foot-inner">
          <div className="foot-col foot-brand">
            <div style={{ marginBottom: 4 }}><GenpactLogo light height={20} /></div>
            <div className="brand-mark" style={{ color: '#fff', display: 'none' }}>
              genpact<span className="dot">•</span>
            </div>
            <p>
              Aluminum Price Intelligence — a Metals Analytics platform forecasting the
              all-in US aluminum purchase price 12 months ahead, monthly.
            </p>
            <p className="foot-tag">Transformation happens here.</p>
          </div>
          <div className="foot-col">
            <h4>Platform</h4>
            <Link to="/dashboard">Dashboard</Link>
            <Link to="/forecast">12-Month Forecast</Link>
            <Link to="/calculator">Price Calculator</Link>
          </div>
          <div className="foot-col">
            <h4>Data</h4>
            <Link to="/validation">Data Validation</Link>
            <Link to="/data">Data Manager</Link>
            <a href="/docs" target="_blank" rel="noreferrer">API Docs</a>
          </div>
          <div className="foot-col">
            <h4>Pricing Formula</h4>
            <p className="foot-formula">
              All-in = LME + Midwest Premium + Gas Coefficient + Labour Index
              − Macro (Supply−Demand) + External Factor
            </p>
          </div>
        </div>
        <div className="foot-bar">
          <span>© {new Date().getFullYear()} Genpact · Aluminum Price Intelligence</span>
          <span>Assistant available bottom-right — works with or without an Azure OpenAI key</span>
        </div>
      </footer>

      <ChatWidget />
    </div>
  )
}
