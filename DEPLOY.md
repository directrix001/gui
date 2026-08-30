# Deploying — from this folder to a public link

The app runs as **one service**: FastAPI serves the API *and* the built React site.
Locally you can still develop with two terminals (`uvicorn` + `npm run dev`);
deployment uses the single-server mode automatically.

---

## Part 1 — Push to GitHub (one time)

From the project root (the folder containing `backend/`, `frontend/`, `.gitignore`):

```powershell
git init
git add .
git commit -m "Aluminum price intelligence platform"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/YOUR-REPO.git
git push -u origin main
```

Already added a remote before? Skip `git remote add` (or fix it with
`git remote set-url origin <url>`). If commit complains about identity:

```powershell
git config --global user.email "you@example.com"
git config --global user.name "Your Name"
```

> `.gitignore` is already set up so `node_modules/`, `dist/`, `.env` (your Azure
> key) and the generated SQLite DB are **never** uploaded. Keep it that way —
> a key pushed to a public repo must be treated as leaked and regenerated.

---

## Part 2 — Deploy on Render (free)

1. Go to **https://render.com** → sign up **with your GitHub account**.
2. Click **New +** → **Blueprint**.
3. Select your repo. Render reads **`render.yaml`** from the repo root and
   configures everything: build command (installs Python deps, builds React)
   and start command (uvicorn). Click **Apply / Deploy**.
4. Wait ~5 minutes for the first build. You get a public URL like
   `https://aluminum-price-intelligence.onrender.com`.

That's it — open the URL: landing page, dashboard, calculator, validation,
chatbot (basic mode) all work with **no keys required**.

### Manual alternative (if you skip the Blueprint)
New + → **Web Service** → pick the repo → set:
- **Runtime**: Python
- **Build Command**:
  `pip install -r backend/requirements.txt && cd frontend && npm install && npm run build`
- **Start Command**:
  `cd backend && uvicorn app.main:app --host 0.0.0.0 --port $PORT`

---

## Part 3 — Optional: enable Azure OpenAI chat

In the Render dashboard → your service → **Environment** → add:

| Key | Value |
|---|---|
| `AZURE_OPENAI_ENDPOINT` | `https://your-resource.openai.azure.com` |
| `AZURE_OPENAI_KEY` | your key |
| `AZURE_OPENAI_DEPLOYMENT` | your deployment name |

Save → Render restarts → the chat header flips to "Azure OpenAI · connected".
Without these, the chatbot keeps working in basic mode. Set a spending cap in
the Azure portal so costs can never surprise you.

The live-feed switch works the same way when you get a price API:
`PRICE_SOURCE_MODE=api`, `PRICE_API_URL=...`, optional `PRICE_API_KEY`.

---

## Part 4 — Updating the live site

```powershell
git add .
git commit -m "describe the change"
git push
```
Render auto-redeploys on every push. Nothing else to do.

---

## Costs & gotchas

- **Free tier**: ₹0, but the service sleeps after ~15 idle minutes — first
  visit takes ~30–50 s to wake. Fine while building.
- **Starter (~$7/mo)**: always-on, instant loads. Switch `plan: free` →
  `plan: starter` in `render.yaml` (or in the dashboard) when the link goes on
  your resume.
- The SQLite DB is rebuilt from the bundled CSV on every deploy — by design.
  On the free tier the filesystem is ephemeral; that's fine because ingestion
  runs at startup.
- Before sharing the link publicly, complete the rebrand (name, colors,
  no client/engagement references).

## Local development (unchanged)

Two terminals as before — Vite proxies `/api` to 8000:
```powershell
cd backend  → python -m uvicorn app.main:app --reload --port 8000
cd frontend → npm run dev
```
To test the production single-server mode locally:
```powershell
cd frontend → npm run build
cd backend  → python -m uvicorn app.main:app --port 8000
# open http://localhost:8000  (site + API on one port)
```
