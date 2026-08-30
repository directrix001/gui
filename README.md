# Aluminum Price Intelligence · Genpact Metals Analytics

Full-stack platform built from the client kickoff discussion:
all-in aluminum price = **LME + Midwest Premium + Gas Coefficient + Labour Index − Macro (Supply−Demand) + External Factor**, with a
**12-month monthly ML forecast**, **four input drivers** (historical-only inputs
that are forecast first, then used to predict the target column), source-link
**data validation**, and **model performance** tracking.

Currently runs on deterministic **mock data** — drop in your real Excel/CSV via
the Data Manager page (or wire `backend/app/main.py` to your data layer) when ready.

## Stack
- **Backend** — FastAPI + NumPy/Pandas (`backend/`)
- **Frontend** — React 18 + Vite + Recharts + React Router (`frontend/`)
- Genpact-aligned theme: coral `#FF555F`, deep navy `#041C2C`, Space Grotesk / Inter

## Deploy it
See **DEPLOY.md** for the full GitHub → Render walkthrough (single-server mode, free tier).

## Run it

### 1. Backend (port 8000)
```bash
cd backend
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000
```
Interactive API docs: http://localhost:8000/docs

### 2. Frontend (port 5173)
```bash
cd frontend
npm install
npm run dev
```
Open http://localhost:5173 — Vite proxies `/api/*` to the backend.

## Pages
| Page | What it covers |
|---|---|
| Dashboard | Price equation hero (LME + MWP = All-in), KPIs, 12/24/36-month history, feature importance |
| 12-Month Forecast | Monthly forecast with 80% confidence bands, switchable per component, full forecast table |
| Input Drivers | The 4 inputs (alumina, energy index, DXY, China output): history + their own 12-month forecasts |
| Data Validation | Source-link health (LME, Platts, Westmetall, client Excel) + rule checks incl. the 4–5 quarter history caveat |
| Model Performance | MAPE / RMSE / MAE / R² / band coverage, 12-month backtest, champion–challenger leaderboard |
| Data Manager | Drag-and-drop CSV/XLSX upload → parsed by the backend, echoes rows/columns for mapping |

## API endpoints
`GET /api/summary` · `GET /api/history?months=` · `GET /api/forecast?component=` ·
`GET /api/drivers` · `GET /api/drivers/importance` · `GET /api/model/metrics` ·
`GET /api/validation` · `POST /api/upload` · `GET /api/health` ·
`GET /api/chat/status` · `POST /api/chat` (Azure OpenAI proxy)

## Enabling the chat assistant (Azure OpenAI)
The floating chat button (bottom-right) talks to `POST /api/chat`, which proxies
Azure OpenAI **server-side** so your key never reaches the browser.

Windows PowerShell (before starting the backend):
```powershell
$env:AZURE_OPENAI_ENDPOINT="https://your-resource.openai.azure.com"
$env:AZURE_OPENAI_KEY="your-key-here"
$env:AZURE_OPENAI_DEPLOYMENT="gpt-4o-mini"   # your deployment name
python -m uvicorn app.main:app --reload --port 8000
```
Or copy `backend/.env.example` to `.env` and load it with your preferred method.
Until configured, the chat panel shows a friendly "not configured" notice.

## Swapping in real data
1. Upload your file on the **Data Manager** page to confirm columns parse.
2. Replace the `_series` mock generators in `backend/app/main.py` with reads from
   your validated dataset (same shapes: monthly labels + float arrays).
3. Everything downstream — forecasts, bands, tables, charts — updates automatically.
