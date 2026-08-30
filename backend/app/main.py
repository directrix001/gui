"""
Aluminum Price Intelligence — FastAPI backend
Genpact | Metals Analytics

Serves mock (deterministic) data for:
  • LME Aluminum base price + US Midwest Premium  →  All-in price = LME + MWP
  • 12-month monthly ML forecast with confidence bands
  • 4 input drivers (historical + forecast)
  • Feature importance, model backtest metrics
  • Data-source validation (link health checks)
  • CSV/XLSX upload endpoint (replace mock data with real data later)
"""
from datetime import date
from io import BytesIO
from typing import Optional

import numpy as np
from dateutil.relativedelta import relativedelta
from fastapi import FastAPI, File, UploadFile, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(
    title="Aluminum Price Intelligence API",
    version="1.0.0",
    description="Aluminum price intelligence backend.",
)

TODAY_MONTH = date.today().strftime("%Y-%m")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/api/health")
def health():
    return {"status": "ok"}

# --------------------------------------------- two-source price comparison
# CSV → SQLite → this endpoint → dashboard. Set PRICE_SOURCE_MODE=api (plus
# PRICE_API_URL / PRICE_API_KEY) to read from a live feed instead — see sources.py.
from .sources import ingest_csv_to_sql, get_comparison

@app.on_event("startup")
def _ingest_market_data():
    try:
        ingest_csv_to_sql()
    except Exception as e:      # never block startup on ingestion
        print(f"[warn] market data ingestion failed: {e}")

@app.get("/api/comparison")
def comparison(months: Optional[int] = Query(None, ge=6, le=600)):
    rows, source_used = get_comparison(months)
    labels, fred, lme3m, variance, variance_pct = [], [], [], [], []
    for r in rows:
        a, b = r["palumusdm"], r["lme_3m"]
        labels.append(r["observation_date"])
        fred.append(round(a, 2) if a is not None else None)
        lme3m.append(round(b, 2) if b is not None else None)
        if a is not None and b is not None and b != 0:
            v = a - b
            variance.append(round(v, 2))
            variance_pct.append(round(v / b * 100, 3))
        else:
            variance.append(None)
            variance_pct.append(None)

    vp = [v for v in variance_pct if v is not None]
    latest = next((i for i in range(len(rows) - 1, -1, -1)
                   if variance_pct[i] is not None), None)
    return {
        "source_mode": source_used,
        "sources": {"fred": "FRED — Global Price of Aluminum (PALUMUSDM)",
                    "lme_3m": "LME Aluminium 3-month"},
        "labels": labels, "fred": fred, "lme_3m": lme3m,
        "variance": variance, "variance_pct": variance_pct,
        "stats": {
            "rows": len(rows),
            "latest_month": labels[latest] if latest is not None else None,
            "latest_fred": fred[latest] if latest is not None else None,
            "latest_lme_3m": lme3m[latest] if latest is not None else None,
            "latest_variance_pct": variance_pct[latest] if latest is not None else None,
            "avg_abs_variance_pct": round(sum(abs(v) for v in vp) / len(vp), 3) if vp else None,
            "max_abs_variance_pct": round(max(abs(v) for v in vp), 3) if vp else None,
        },
    }

# ------------------------------------------------------- price calculator
from pydantic import BaseModel as _BM, Field
from .calculator import PriceInputs, calculate_new_price

class CalcRequest(_BM):
    weight: float = Field(..., description="Part weight (PWt)")
    current_price: float = Field(..., description="Current price (P_Current)")
    ppi_q: float = Field(..., description="PPI current quarter")
    ppi_q1: float = Field(..., description="PPI previous quarter (PPI_Q-1)")
    drauss_factor: float = Field(1.44, description="Drauss / conversion factor (DF_c)")
    mc_q: float = Field(..., description="Metal cost current quarter (MC_Q)")
    mc_q_1: float = Field(..., description="Metal cost previous quarter (MC_Q-1)")
    cng_q: float = Field(..., description="CNG cost current quarter")
    cng_q_1: float = Field(..., description="CNG cost previous quarter (CNG_Q-1)")

@app.post("/api/calculator/single")
def calculator_single(req: CalcRequest):
    try:
        r = calculate_new_price(PriceInputs(**req.model_dump()))
    except ValueError as e:
        raise HTTPException(422, str(e))
    return {"new_price": r.new_price, "ams_q": r.ams_q,
            "ams_q_1": r.ams_q_1, "ppi_factor": r.ppi_factor}

# flexible column-name matching for uploaded files
_COL_ALIASES = {
    "weight":        ["weight", "pwt", "part_weight", "wt"],
    "current_price": ["current_price", "p_current", "pcurrent", "price", "current"],
    "ppi_q":         ["ppi_q", "ppi", "ppi_current", "ppi_curr"],
    "ppi_q1":        ["ppi_q_1", "ppi_q1", "ppi_prev", "ppi_previous", "ppi_last"],
    "drauss_factor": ["drauss_factor", "df_c", "dfc", "conversion_factor", "drauss"],
    "mc_q":          ["mc_q", "metal_cost", "mc_current", "mc"],
    "mc_q_1":        ["mc_q_1", "mc_q1", "mc_prev", "mc_previous", "mc_last"],
    "cng_q":         ["cng_q", "cng", "cng_current", "cng_curr"],
    "cng_q_1":       ["cng_q_1", "cng_q1", "cng_prev", "cng_previous", "cng_last"],
}

def _normalize(col: str) -> str:
    return str(col).strip().lower().replace("-", "_").replace(" ", "_").replace("__", "_")

@app.post("/api/calculator/batch")
async def calculator_batch(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith((".csv", ".xlsx", ".xls")):
        raise HTTPException(400, "Upload a .csv or .xlsx file.")
    content = await file.read()
    import pandas as pd
    try:
        df = (pd.read_csv(BytesIO(content)) if file.filename.lower().endswith(".csv")
              else pd.read_excel(BytesIO(content)))
    except Exception as e:
        raise HTTPException(422, f"Could not parse file: {e}")
    if df.empty:
        raise HTTPException(422, "The file has no rows.")
    if len(df) > 5000:
        raise HTTPException(422, "Please keep files under 5,000 rows.")

    norm_map = {_normalize(c): c for c in df.columns}
    mapping, missing = {}, []
    for field, aliases in _COL_ALIASES.items():
        found = next((norm_map[a] for a in aliases if a in norm_map), None)
        if found is not None:
            mapping[field] = found
        elif field != "drauss_factor":   # DF_c is optional → defaults to 1.44
            missing.append(field)
    if missing:
        raise HTTPException(422,
            f"Missing required columns: {', '.join(missing)}. "
            f"Found columns: {', '.join(map(str, df.columns))}. "
            "Accepted names include e.g. weight/pwt, current_price/p_current, ppi_q, "
            "ppi_q_1, mc_q, mc_q_1, cng_q, cng_q_1, drauss_factor (optional).")

    results, errors = [], []
    for i, row in df.iterrows():
        try:
            vals = {f: float(row[c]) for f, c in mapping.items()}
            vals.setdefault("drauss_factor", 1.44)
            if pd.isna(list(vals.values())).any():
                raise ValueError("blank / non-numeric value")
            r = calculate_new_price(PriceInputs(**vals))
            out = {k: (None if pd.isna(v) else v if not hasattr(v, "item") else v.item())
                   for k, v in row.items()}
            out.update({"ams_q": r.ams_q, "ams_q_1": r.ams_q_1,
                        "ppi_factor": r.ppi_factor, "new_price": r.new_price})
            results.append(out)
        except Exception as e:
            errors.append({"row": int(i) + 2, "error": str(e)})  # +2 = header + 1-index

    return {"filename": file.filename, "count": len(results),
            "column_mapping": mapping, "errors": errors, "results": results}

# ----------------------------------------------------------- assistant chat
# Azure OpenAI proxy — the key stays server-side. Configure via environment
# variables (or a .env file loaded before starting uvicorn):
#   AZURE_OPENAI_ENDPOINT    e.g. https://<your-resource>.openai.azure.com
#   AZURE_OPENAI_KEY         your Azure OpenAI API key
#   AZURE_OPENAI_DEPLOYMENT  your deployment name, e.g. gpt-4o-mini
#   AZURE_OPENAI_API_VERSION optional, defaults to 2024-06-01
from pydantic import BaseModel

class ChatMessage(BaseModel):
    role: str      # "user" | "assistant"
    content: str

class ChatRequest(BaseModel):
    messages: list[ChatMessage]

SYSTEM_PROMPT = (
    "You are the in-app assistant for the Aluminum Price Intelligence platform "
    "(Genpact Metals Analytics). The platform forecasts the all-in US aluminum "
    "price for 12 months ahead at monthly granularity, where "
    "All-in = LME base + US Midwest Premium + Gas Coefficient + Labour Index "
    "− Macro factor (supply − demand) + External factor. Input drivers have "
    "historical data only and are forecast first, then the target price is "
    "predicted (two-step approach). Answer general questions helpfully and "
    "concisely; explain forecasting, pricing and platform concepts in simple language."
)

@app.get("/api/chat/status")
def chat_status():
    configured = bool(os.getenv("AZURE_OPENAI_ENDPOINT") and os.getenv("AZURE_OPENAI_KEY")
                      and os.getenv("AZURE_OPENAI_DEPLOYMENT"))
    return {"configured": configured,
            "mode": "azure" if configured else "basic",
            "deployment": os.getenv("AZURE_OPENAI_DEPLOYMENT") if configured else None}

# ---- basic fallback assistant (no key needed) ------------------------------
def _fmt_usd(v): return f"${v:,.0f}"

def _basic_answer(text: str) -> str:
    q = text.lower()
    allin = float(ALLIN_HIST[-1]); lme = float(LME_HIST[-1]); mwp = float(MWP_HIST[-1])
    gas = float(GAS_HIST[-1]); lab = float(LABOUR_HIST[-1])
    mac = float(MACRO_HIST[-1]); ext = float(EXT_HIST[-1])

    if any(k in q for k in ["formula", "equation", "how is", "calculated", "components", "made of"]):
        return (
            "The all-in aluminum price is built from six components:\n\n"
            f"LME base ({_fmt_usd(lme)}) + Midwest Premium ({_fmt_usd(mwp)}) "
            f"+ Gas Coefficient ({_fmt_usd(gas)}) + Labour Index ({_fmt_usd(lab)}) "
            f"− Macro Factor, supply−demand ({_fmt_usd(mac)}) + External Factor ({_fmt_usd(ext)}) "
            f"= All-in {_fmt_usd(allin)}/t.\n\n"
            "Only these components move the price on this platform."
        )
    if any(k in q for k in ["price", "all-in", "all in", "current", "today", "latest", "cost"]):
        return (
            f"The latest all-in aluminum price is {_fmt_usd(allin)}/tonne "
            f"(LME {_fmt_usd(lme)} + Midwest Premium {_fmt_usd(mwp)} + Gas {_fmt_usd(gas)} "
            f"+ Labour {_fmt_usd(lab)} − Macro {_fmt_usd(mac)} + External {_fmt_usd(ext)}), "
            f"as of {TODAY_MONTH}."
        )
    if any(k in q for k in ["band", "confidence", "80", "lower", "upper", "uncertain", "range"]):
        return (
            "The lower/upper 80% values form a confidence band: the model is 80% sure the actual "
            "price will land inside that range (about a 10% chance below, 10% above). The band "
            "widens further into the future because uncertainty grows — predicting next month is "
            "easier than predicting month 12. Use the mean as the best guess and the upper bound "
            "for worst-case budgeting."
        )
    if any(k in q for k in ["forecast", "12 month", "next month", "predict", "future", "outlook"]):
        peak_i = int(np.argmax(ALLIN_F))
        return (
            f"The 12-month monthly forecast averages {_fmt_usd(float(np.mean(ALLIN_F)))}/t for the "
            f"all-in price. Next month is expected around {_fmt_usd(float(ALLIN_F[0]))} "
            f"({_fmt_usd(float(ALLIN_LO[0]))}–{_fmt_usd(float(ALLIN_HI[0]))} at 80% confidence), "
            f"with a peak near {_fmt_usd(float(ALLIN_F[peak_i]))} in {FCST_LABELS[peak_i]}. "
            "See the 12-Month Forecast page for the full table."
        )
    if any(k in q for k in ["driver", "input", "gas", "labour", "labor", "macro", "external",
                            "supply", "demand", "variable", "factor"]):
        return (
            "Four inputs drive the price beyond LME and the Midwest Premium:\n\n"
            f"• Gas Coefficient ({_fmt_usd(gas)}) — adds to price, energy pass-through\n"
            f"• Labour Index ({_fmt_usd(lab)}) — adds to price\n"
            f"• Macro Factor, supply−demand ({_fmt_usd(mac)}) — subtracts; surplus pulls price down\n"
            f"• External Factor ({_fmt_usd(ext)}) — adds; tariffs, freight, geopolitics\n\n"
            "Each has historical data only, so the pipeline forecasts every input 12 months ahead "
            "first, then predicts the target price (the two-step approach)."
        )
    if any(k in q for k in ["model", "accuracy", "mape", "rmse", "how good", "reliable", "trust"]):
        return (
            "The champion model is a GBM ensemble with ~3.7% MAPE (average error) on backtests, "
            "R² 0.91, and 84% of actuals falling inside the 80% band — close to the promised 80%, "
            "which means the bands are honest. Challengers (SARIMAX, LSTM, Prophet) are re-scored "
            "monthly. Note: with only 4–5 quarters of driver history, expect wider bands early on."
        )
    if any(k in q for k in ["validat", "source", "link", "lme site", "platts", "data quality"]):
        return (
            "Data validation covers two things: source-link health (LME, Platts Midwest Premium, "
            "Westmetall cross-check, client Excel) and rule checks — the pricing identity must hold, "
            "no missing months, outliers beyond 3σ flagged, and a warning that driver history is "
            "only 4–5 quarters. Note that LME and Platts feeds are licensed/commercial, so live "
            "linking is a business decision as much as a technical one."
        )
    if any(k in q for k in ["upload", "excel", "csv", "real data", "my data", "file"]):
        return (
            "Use the Data Manager page: drag and drop your Excel/CSV (columns like month, "
            "lme_usd_t, midwest_premium_usd_t, gas_coeff, labour_index, macro_supply_demand, "
            "external_factor). The backend parses it and echoes the rows/columns it found so the "
            "mapping can be confirmed before switching from mock to real data."
        )
    if any(k in q for k in ["hi", "hello", "hey", "help", "what can you"]):
        return (
            "Hello! I can answer questions about this platform — try asking about the pricing "
            "formula, the current price, the 12-month forecast, the 80% confidence bands, the "
            "input drivers, model accuracy, data validation, or how to upload real data.\n\n"
            "(I'm currently in basic mode. Add an Azure OpenAI key to the backend to unlock "
            "full AI answers to any general question.)"
        )
    return (
        "I'm in basic mode (no Azure OpenAI key configured), so I can answer platform questions "
        "about: the pricing formula, current price, 12-month forecast, confidence bands, input "
        "drivers, model accuracy, data validation, and data upload. Try one of those — or add "
        "your Azure key to backend/.env for full AI answers to anything."
    )

@app.post("/api/chat")
async def chat(req: ChatRequest):
    endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")
    key = os.getenv("AZURE_OPENAI_KEY")
    deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT")
    api_version = os.getenv("AZURE_OPENAI_API_VERSION", "2024-06-01")

    last_user = next((m.content for m in reversed(req.messages) if m.role == "user"), "")

    # ---- no key → basic mode, never fail --------------------------------
    if not (endpoint and key and deployment):
        return {"reply": _basic_answer(last_user), "mode": "basic"}

    url = (f"{endpoint.rstrip('/')}/openai/deployments/{deployment}"
           f"/chat/completions?api-version={api_version}")
    payload = {
        "messages": [{"role": "system", "content": SYSTEM_PROMPT}]
                    + [m.model_dump() for m in req.messages][-20:],
        "max_tokens": 700,
        "temperature": 0.4,
    }

    import httpx
    try:
        async with httpx.AsyncClient(timeout=60) as client:
            r = await client.post(url, json=payload, headers={"api-key": key})
        if r.status_code == 200:
            data = r.json()
            reply = data["choices"][0]["message"]["content"]
            return {"reply": reply, "mode": "azure", "usage": data.get("usage", {})}
        # Azure returned an error → degrade gracefully to basic mode
        note = f"(Azure OpenAI returned {r.status_code} — answered in basic mode instead.)\n\n"
        return {"reply": note + _basic_answer(last_user), "mode": "basic"}
    except Exception:
        note = "(Couldn't reach Azure OpenAI — answered in basic mode instead.)\n\n"
        return {"reply": note + _basic_answer(last_user), "mode": "basic"}

# ---------------------------------------------------------- multi-agent chat
from . import agents as _agents
from . import factor_dbs as _fdb

@app.on_event("startup")
def _seed_factor_databases():
    try:
        counts = _fdb.build_master()
        _agents.configure({"latest_month": TODAY_MONTH})
        print(f"[ok] master.db built: {counts}")
    except Exception as e:
        print(f"[warn] master.db build failed: {e}")

@app.get("/api/agent/status")
def agent_status():
    try:
        inv = _fdb.db_inventory()
    except Exception:
        inv = []
    web_ok = True
    try:
        import ddgs  # noqa
    except ImportError:
        try:
            import duckduckgo_search  # noqa
        except ImportError:
            web_ok = False
    return {"llm_mode": _agents.llm_mode(), "web_search_installed": web_ok,
            "databases": inv}

@app.post("/api/agent/chat")
async def agent_chat(req: ChatRequest):
    # refresh comparison stats for variance questions
    try:
        rows, _src = get_comparison(None)
        last = next((r for r in reversed(rows)
                     if r["palumusdm"] is not None and r["lme_3m"] is not None), None)
        if last:
            _agents.configure({"comparison_stats": {
                "latest_month": last["observation_date"],
                "latest_variance_pct": round(
                    (last["palumusdm"] - last["lme_3m"]) / last["lme_3m"] * 100, 3),
                "rows": len(rows),
                "avg_abs_variance_pct": round(sum(
                    abs((r["palumusdm"] - r["lme_3m"]) / r["lme_3m"] * 100)
                    for r in rows if r["palumusdm"] is not None and r["lme_3m"] is not None
                ) / max(1, sum(1 for r in rows if r["palumusdm"] is not None
                               and r["lme_3m"] is not None)), 3),
            }})
    except Exception:
        pass
    return _agents.answer([m.model_dump() for m in req.messages])

# ------------------------------------------------- streaming chat + feedback
from fastapi.responses import StreamingResponse
import sqlite3 as _sqlite3
import time as _time
import json as _json

def _refresh_comparison_stats():
    try:
        rows, _src = get_comparison(None)
        valid = [r for r in rows
                 if r["palumusdm"] is not None and r["lme_3m"] is not None]
        if valid:
            last = valid[-1]
            _agents.configure({"comparison_stats": {
                "latest_month": last["observation_date"],
                "latest_variance_pct": round(
                    (last["palumusdm"] - last["lme_3m"]) / last["lme_3m"] * 100, 3),
                "rows": len(rows),
                "avg_abs_variance_pct": round(sum(
                    abs((r["palumusdm"] - r["lme_3m"]) / r["lme_3m"] * 100)
                    for r in valid) / len(valid), 3),
            }})
    except Exception:
        pass

@app.post("/api/agent/chat/stream")
async def agent_chat_stream(req: ChatRequest):
    _refresh_comparison_stats()
    msgs = [m.model_dump() for m in req.messages]

    def gen():
        try:
            for ev in _agents.stream(msgs):
                yield "data: " + _json.dumps(ev) + "\n\n"
        except Exception as e:
            yield "data: " + _json.dumps(
                {"type": "delta", "text": f"⚠️ {e}"}) + "\n\n"
            yield "data: " + _json.dumps({"type": "done"}) + "\n\n"

    return StreamingResponse(gen(), media_type="text/event-stream",
                             headers={"Cache-Control": "no-cache",
                                      "X-Accel-Buffering": "no"})

# ---- feedback storage (SQLite) ----
from pathlib import Path as _Path
_FEEDBACK_DB = _Path(__file__).parent / "data" / "feedback.db"

class FeedbackIn(_BM):
    rating: str = Field(..., pattern="^(up|down)$")
    question: str = Field("", max_length=2000)
    answer: str = Field("", max_length=4000)
    mode: str = Field("", max_length=40)

@app.post("/api/agent/feedback")
def agent_feedback(fb: FeedbackIn):
    _FEEDBACK_DB.parent.mkdir(exist_ok=True)
    conn = _sqlite3.connect(_FEEDBACK_DB)
    try:
        conn.execute("""CREATE TABLE IF NOT EXISTS feedback (
            id INTEGER PRIMARY KEY AUTOINCREMENT, ts INTEGER,
            rating TEXT, question TEXT, answer TEXT, mode TEXT)""")
        conn.execute("INSERT INTO feedback (ts, rating, question, answer, mode) "
                     "VALUES (?,?,?,?,?)",
                     (int(_time.time()), fb.rating, fb.question, fb.answer, fb.mode))
        conn.commit()
        n_up = conn.execute("SELECT COUNT(*) FROM feedback WHERE rating='up'").fetchone()[0]
        n_dn = conn.execute("SELECT COUNT(*) FROM feedback WHERE rating='down'").fetchone()[0]
    finally:
        conn.close()
    return {"ok": True, "totals": {"up": n_up, "down": n_dn}}

# ----------------------------------------------------------------- news module
from . import news as _news

@app.on_event("startup")
def _init_news():
    try:
        r = _news.ensure_news()
        print(f"[ok] news ready: {r}")
    except Exception as e:
        print(f"[warn] news init failed: {e}")

@app.get("/api/news")
def news_list(month: Optional[str] = None, factor: Optional[str] = None,
              limit: int = Query(100, ge=1, le=300)):
    arts = _news.get_news(month=month, factor=factor, limit=limit)
    latest = arts[0] if arts else None
    return {"articles": arts, "count": len(arts),
            "bulletin": {"title": latest["title"], "url": latest["url"],
                         "published": latest["published"],
                         "source": latest["source"]} if latest else None}

@app.get("/api/news/stats")
def news_stats():
    return {"monthly_counts": _news.monthly_counts()}

@app.post("/api/news/refresh")
def news_refresh():
    try:
        inserted = _news.fetch_live()
        return {"ok": True, "live_inserted": inserted,
                "message": f"Fetched live feeds — {inserted} new article(s)."
                if inserted else "Feeds reachable but no new relevant articles."}
    except Exception as e:
        return {"ok": False, "live_inserted": 0,
                "message": f"Live fetch unavailable ({e}); existing articles kept."}

@app.post("/api/news/chat")
async def news_chat(req: ChatRequest):
    return _news.news_answer([m.model_dump() for m in req.messages])

class BackfillIn(_BM):
    start: str = Field(..., pattern=r"^20\d{2}-(0[1-9]|1[0-2])$")
    end: str = Field(..., pattern=r"^20\d{2}-(0[1-9]|1[0-2])$")

@app.post("/api/news/backfill")
def news_backfill(req: BackfillIn):
    if req.start > req.end:
        raise HTTPException(422, "start must be <= end")
    # cap one request to 36 months to keep it responsive
    sy, sm = int(req.start[:4]), int(req.start[5:7])
    ey, em = int(req.end[:4]), int(req.end[5:7])
    if (ey - sy) * 12 + (em - sm) > 36:
        raise HTTPException(422, "Backfill max 36 months per request — run multiple ranges.")
    try:
        r = _news.backfill_gdelt(req.start, req.end)
        note = ("GDELT indexes news from 2017-01 onward; earlier months are skipped. "
                if req.start <= "2017-01" else "")
        return {"ok": r["failed_calls"] < r["api_calls"], **r,
                "message": f"{note}Inserted {r['inserted']} historical article(s) "
                           f"({r['api_calls']} API calls, {r['failed_calls']} failed)."}
    except Exception as e:
        return {"ok": False, "inserted": 0, "message": f"Backfill unavailable: {e}"}


class ManualRow(_BM):
    month: str = Field(..., pattern=r"^20\d{2}-(0[1-9]|1[0-2])$")
    value: float

@app.get("/api/real/drivers")
def real_drivers():
    out = []
    for key, meta in _fdb.FACTORS.items():
        series = _fdb.get_factor_series(key)
        latest = series[-1] if series else None
        prev = series[-2] if len(series) > 1 else None
        mom = None
        if latest and prev and prev["value"]:
            mom = round((latest["value"] - prev["value"]) / prev["value"] * 100, 2)
        out.append({"key": key, "name": meta["name"], "unit": meta["unit"],
                    "db": meta["db"], "rows": len(series),
                    "from": series[0]["month"] if series else None,
                    "to": series[-1]["month"] if series else None,
                    "latest": latest, "mom_pct": mom,
                    "series": series})
    return {"drivers": out}

@app.post("/api/real/manual/{factor}")
def real_manual(factor: str, row: ManualRow):
    if factor not in _fdb.FACTORS:
        raise HTTPException(404, f"Unknown factor '{factor}'")
    _fdb.upsert_rows(factor, [(row.month, row.value)], source="real:manual")
    return {"ok": True, "factor": factor, "month": row.month, "value": row.value,
            "message": f"Saved {row.month} = {row.value} into {_fdb.FACTORS[factor]['db']}"}

@app.post("/api/real/upload/{factor}")
async def real_upload(factor: str, file: UploadFile = File(...)):
    if factor not in _fdb.FACTORS:
        raise HTTPException(404, f"Unknown factor '{factor}'")
    import io, re as _re
    import pandas as pd
    raw = await file.read()
    name = (file.filename or "").lower()
    try:
        if name.endswith((".xlsx", ".xls")):
            df = pd.read_excel(io.BytesIO(raw))
        else:
            text = raw.decode("utf-8", errors="replace")
            sep = "\t" if "\t" in text.splitlines()[0] else ","
            df = pd.read_csv(io.StringIO(text), sep=sep)
    except Exception as e:
        raise HTTPException(422, f"Could not parse file: {e}")
    df.columns = [str(c).strip().lower() for c in df.columns]
    month_col = next((c for c in df.columns if "month" in c or "date" in c), df.columns[0])
    value_col = next((c for c in df.columns if c != month_col and
                      ("value" in c or "price" in c or "cost" in c or "premium" in c
                       or "index" in c or "$" in c)), None)
    if value_col is None:
        others = [c for c in df.columns if c != month_col]
        if not others:
            raise HTTPException(422, "Need two columns: month and value")
        value_col = others[0]
    rows, errors = [], []
    for i, r in df.iterrows():
        mraw = str(r[month_col]).strip()
        mm = _re.match(r"^(20\d{2})[-/](\d{1,2})", mraw)
        if not mm:
            try:
                dt = pd.to_datetime(mraw)
                mm_month = f"{dt.year:04d}-{dt.month:02d}"
            except Exception:
                errors.append(f"row {i+1}: bad month '{mraw}'")
                continue
        else:
            mm_month = f"{mm.group(1)}-{int(mm.group(2)):02d}"
        vraw = str(r[value_col]).replace("$", "").replace(",", "").strip()
        try:
            rows.append((mm_month, float(vraw)))
        except ValueError:
            errors.append(f"row {i+1}: bad value '{r[value_col]}'")
    if not rows:
        raise HTTPException(422, "No valid rows found. " + "; ".join(errors[:3]))
    n = _fdb.upsert_rows(factor, rows, source="real:upload")
    return {"ok": True, "factor": factor, "inserted": n,
            "errors": errors[:5],
            "message": f"{n} row(s) pushed into {_fdb.FACTORS[factor]['db']}"
                       + (f" ({len(errors)} row(s) skipped)" if errors else "")}


from .api.v1.router import api_v1_router
app.include_router(api_v1_router)

@app.get("/health")
def health_check():
    return {"status": "ok", "engine": "quarterly-formula", "parts_file": "aluminium_data.xlsx"}

# ---------------------------------------------------------------- single-server
# Production mode: serve the built React app (frontend/dist) from FastAPI so the
# whole platform — site + API + chatbot — runs as ONE service on ONE port.
# Locally you can still use `npm run dev` (Vite on 5173 proxying to 8000);
# this block only activates when frontend/dist exists (after `npm run build`).
from pathlib import Path as _Path
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse

_FRONTEND_DIST = _Path(__file__).resolve().parents[2] / "frontend" / "dist"

if _FRONTEND_DIST.exists():
    app.mount("/assets", StaticFiles(directory=_FRONTEND_DIST / "assets"), name="assets")

    @app.get("/{full_path:path}", include_in_schema=False)
    async def serve_spa(full_path: str):
        # /api/* routes are registered above and take precedence; everything
        # else falls through here: real files are served, unknown SPA paths
        # get index.html so React Router handles /dashboard, /calculator, etc.
        candidate = _FRONTEND_DIST / full_path
        if full_path and candidate.is_file():
            return FileResponse(candidate)
        return FileResponse(_FRONTEND_DIST / "index.html")
