"""
Factor News module.

Sources (all free/open): Google News RSS feeds — one query per price factor.
Articles are relevance-filtered (must be aluminum-related), classified by the
factor(s) they touch, deduplicated, and stored in the `news` table of
master.db. If live fetching is unavailable (no network / RSS blocked), a
curated sample dataset is seeded so the section always works; every sample
links to a live Google News search for that headline.

News AI (separate from the platform assistant):
  • answers ONLY from the news table — month/quarter/year summaries and
    forward-looking trend readings grounded in recent articles
  • refuses database/value questions (those belong to the main assistant)
  • model: gpt-4.1-mini by default (OPENAI_NEWS_MODEL to override)
"""
import os
import re
import sqlite3
import time
import urllib.parse
import xml.etree.ElementTree as ET
from datetime import datetime
from email.utils import parsedate_to_datetime

from .factor_dbs import MASTER_DB

FEEDS = {
    "lme":             "LME aluminium price",
    "midwest_premium": "US midwest aluminum premium tariff",
    "gas":             "aluminum smelter energy gas prices",
    "labour":          "aluminum smelter workers labor",
    "macro":           "aluminum supply demand China output",
    "external":        "aluminum tariffs sanctions exports",
}

RELEVANCE = ["aluminum", "aluminium", "smelter", "lme", "alumina", "bauxite"]

FACTOR_TAGS = {
    "lme": ["lme", "price", "exchange", "futures", "metal price"],
    "midwest_premium": ["premium", "midwest", "duty"],
    "gas": ["energy", "gas", "power", "electricity"],
    "labour": ["labor", "labour", "worker", "strike", "wage", "union"],
    "macro": ["supply", "demand", "output", "production", "deficit", "surplus",
              "china", "inventory", "stocks"],
    "external": ["tariff", "sanction", "export", "import", "trade", "war",
                 "geopolit", "freight"],
}

SAMPLE_NEWS = [
    ("2026-08-14", "Aluminium holds near multi-year highs as supply deficit persists", "Market wire", ["lme", "macro"],
     "Analysts point to a widening primary aluminium deficit as smelter restarts lag demand recovery."),
    ("2026-08-08", "US Midwest premium steady as tariff regime keeps imports costly", "Trade daily", ["midwest_premium", "external"],
     "The delivered Midwest premium held near record levels with duty-paid imports still uneconomic."),
    ("2026-07-30", "European smelters warn on power costs into winter", "Energy desk", ["gas", "lme"],
     "Producers flagged that forward power prices threaten restart economics for curtailed capacity."),
    ("2026-07-18", "China smelter output edges lower on hydropower shortfall", "Metals brief", ["macro"],
     "Yunnan hydro constraints trimmed monthly output, tightening the global balance."),
    ("2026-07-05", "Wage settlements lift smelter labor costs across North America", "Industry note", ["labour"],
     "New agreements added mid-single-digit increases to cash costs at several plants."),
    ("2026-06-20", "Aluminium slips as profit-taking follows strong first half", "Market wire", ["lme"],
     "Prices eased from May peaks though fundamentals stayed constructive, traders said."),
    ("2026-06-11", "Freight rates add to landed aluminum costs on Atlantic routes", "Logistics watch", ["external", "midwest_premium"],
     "Container and breakbulk rates pushed delivered premiums higher in the US market."),
    ("2026-05-28", "LME aluminium touches cycle high on restocking wave", "Market wire", ["lme", "macro"],
     "Downstream restocking met thin exchange inventories, extending the rally."),
    ("2026-05-15", "Gas benchmark spike feeds through to smelting cost curves", "Energy desk", ["gas"],
     "Higher spot gas lifted the marginal cost of production in Europe and Asia."),
    ("2026-04-22", "Tariff review keeps US aluminum import duties elevated", "Trade daily", ["external", "midwest_premium"],
     "Officials signalled continuity on aluminum tariffs, supporting domestic premiums."),
    ("2026-04-09", "Energy pass-through accelerates for aluminium producers", "Industry note", ["gas", "lme"],
     "Producers cited faster pass-through of power costs into contract pricing."),
    ("2026-03-25", "Supply deficit narrative builds as inventories fall", "Metals brief", ["macro", "lme"],
     "Visible stocks fell for a sixth month, reinforcing deficit expectations."),
    ("2026-03-12", "US buyers face higher all-in aluminum costs", "Trade daily", ["midwest_premium", "lme"],
     "The combination of firm LME prices and record premiums lifted all-in costs."),
    ("2026-02-18", "China exports of semis slow amid domestic demand pickup", "Metals brief", ["macro", "external"],
     "Slower semi-fabricated exports tightened availability outside China."),
    ("2026-01-28", "Aluminium rally extends into the new year", "Market wire", ["lme"],
     "Momentum funds added length as prices cleared multi-year resistance."),
    ("2026-01-10", "Restocking and deficit talk lift base metals", "Market wire", ["macro", "lme"],
     "Aluminium led base-metal gains on tightening balance projections."),
    ("2025-12-15", "Smelter curtailment risk returns with winter power pricing", "Energy desk", ["gas"],
     "Cold-season power auctions revived curtailment concerns in Europe."),
    ("2025-11-20", "Premiums firm as tariff-era trade flows settle", "Trade daily", ["midwest_premium", "external"],
     "US premiums consolidated at elevated levels as flows adjusted to duties."),
    ("2025-10-16", "Labor negotiations open at major aluminum producers", "Industry note", ["labour"],
     "Contract talks began with unions seeking cost-of-living adjustments."),
    ("2025-09-24", "Aluminium steadies after tariff-driven volatility", "Market wire", ["lme", "external"],
     "Prices found a range as markets digested the year's tariff escalation."),
]


def _conn():
    c = sqlite3.connect(MASTER_DB)
    c.row_factory = sqlite3.Row
    return c


def _ensure_table():
    conn = _conn()
    try:
        conn.execute("""CREATE TABLE IF NOT EXISTS news (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            published TEXT, month TEXT, title TEXT UNIQUE, summary TEXT,
            url TEXT, source TEXT, factors TEXT, origin TEXT, fetched_at INTEGER)""")
        conn.commit()
    finally:
        conn.close()


def _classify(text: str) -> list[str]:
    t = text.lower()
    tags = [k for k, words in FACTOR_TAGS.items() if any(w in t for w in words)]
    return tags or ["lme"]


def _relevant(text: str) -> bool:
    t = text.lower()
    return any(w in t for w in RELEVANCE)


def _strip_html(s: str) -> str:
    return re.sub(r"<[^>]+>", "", s or "").strip()


def fetch_live() -> int:
    """Pull Google News RSS for every factor query. Returns rows inserted."""
    import httpx
    _ensure_table()
    inserted = 0
    conn = _conn()
    try:
        for fk, q in FEEDS.items():
            url = ("https://news.google.com/rss/search?q="
                   + urllib.parse.quote(q) + "&hl=en-US&gl=US&ceid=US:en")
            try:
                r = httpx.get(url, timeout=15, follow_redirects=True)
                r.raise_for_status()
                root = ET.fromstring(r.text)
            except Exception:
                continue
            for item in root.iter("item"):
                title = _strip_html(item.findtext("title") or "")
                link = (item.findtext("link") or "").strip()
                desc = _strip_html(item.findtext("description") or "")[:280]
                src = item.findtext("{*}source") or item.findtext("source") or "Google News"
                pub = item.findtext("pubDate") or ""
                if not title or not link or not _relevant(title + " " + desc):
                    continue
                try:
                    dt = parsedate_to_datetime(pub)
                    published = dt.strftime("%Y-%m-%d")
                    month = dt.strftime("%Y-%m")
                except Exception:
                    published = datetime.utcnow().strftime("%Y-%m-%d")
                    month = published[:7]
                factors = ",".join(sorted(set(_classify(title + " " + desc) + [fk])))
                try:
                    conn.execute(
                        "INSERT OR IGNORE INTO news (published, month, title, summary, "
                        "url, source, factors, origin, fetched_at) VALUES (?,?,?,?,?,?,?,?,?)",
                        (published, month, title, desc or "(no summary provided)",
                         link, str(src).strip(), factors, "live:google_news_rss",
                         int(time.time())))
                    inserted += conn.total_changes and 1
                except Exception:
                    pass
        conn.commit()
    finally:
        conn.close()
    return inserted


def seed_samples() -> int:
    _ensure_table()
    conn = _conn()
    try:
        n = 0
        for published, title, source, factors, summary in SAMPLE_NEWS:
            url = ("https://news.google.com/search?q="
                   + urllib.parse.quote(title))
            cur = conn.execute(
                "INSERT OR IGNORE INTO news (published, month, title, summary, url, "
                "source, factors, origin, fetched_at) VALUES (?,?,?,?,?,?,?,?,?)",
                (published, published[:7], title, summary, url, source,
                 ",".join(factors), "sample:curated", int(time.time())))
            n += cur.rowcount
        conn.commit()
        return n
    finally:
        conn.close()


def ensure_news() -> dict:
    """Startup: try live fetch; fall back to samples if the table is empty."""
    _ensure_table()
    live = 0
    try:
        live = fetch_live()
    except Exception:
        pass
    conn = _conn()
    try:
        total = conn.execute("SELECT COUNT(*) FROM news").fetchone()[0]
    finally:
        conn.close()
    seeded = 0
    if total == 0:
        seeded = seed_samples()
    return {"live_inserted": live, "sample_seeded": seeded}


def get_news(month: str | None = None, factor: str | None = None,
             limit: int = 100) -> list[dict]:
    _ensure_table()
    conn = _conn()
    try:
        q, args = "SELECT * FROM news", []
        conds = []
        if month:
            conds.append("month = ?"); args.append(month)
        if factor:
            conds.append("factors LIKE ?"); args.append(f"%{factor}%")
        if conds:
            q += " WHERE " + " AND ".join(conds)
        q += " ORDER BY published DESC LIMIT ?"; args.append(limit)
        return [dict(r) for r in conn.execute(q, args).fetchall()]
    finally:
        conn.close()


def monthly_counts() -> list[dict]:
    _ensure_table()
    conn = _conn()
    try:
        return [dict(r) for r in conn.execute(
            "SELECT month, COUNT(*) AS count FROM news "
            "GROUP BY month ORDER BY month").fetchall()]
    finally:
        conn.close()


# --------------------------------------------------------------- news AI
DB_QUESTION_MARKERS = [
    "what was the price", "value of", "how much was", "average price",
    "sql", "query", "database", "table", "monthly value", "calculate",
    "p_new", "calculator", "variance between fred",
]

NEWS_SCOPE_REFUSAL = (
    "This assistant covers news and trends only. For prices, values, changes "
    "or calculations from the database, use the main platform assistant on "
    "any other page."
)


def _period_months(question: str) -> list[str] | None:
    from .agents import _extract_month
    q = question.lower()
    m = re.search(r"\b(20\d{2})\s*-?\s*q([1-4])\b", q) or \
        re.search(r"\bq([1-4])\s*,?\s*(20\d{2})\b", q)
    if m:
        g = m.groups()
        year, qtr = (g[0], int(g[1])) if len(g[0]) == 4 else (g[1], int(g[0]))
        start = (qtr - 1) * 3 + 1
        return [f"{year}-{mm:02d}" for mm in range(start, start + 3)]
    ym = re.search(r"\b(20\d{2})[-/](\d{1,2})\b", q)
    if ym:
        return [f"{ym.group(1)}-{int(ym.group(2)):02d}"]
    mon = _extract_month(q)
    if mon and re.search(r"\b(20\d{2})\b", q) and not re.search(r"[a-z]{3,9}\s+20\d{2}|20\d{2}\s+[a-z]{3,9}", q):
        y = re.search(r"\b(20\d{2})\b", q).group(1)
        return [f"{y}-{mm:02d}" for mm in range(1, 13)]
    return [mon] if mon else None


def news_answer(messages: list[dict]) -> dict:
    from .agents import _call_llm, llm_mode, _is_injection
    question = next((m["content"] for m in reversed(messages)
                     if m.get("role") == "user"), "")[:1500]
    ql = question.lower()

    if _is_injection(question):
        return {"reply": "I can't change my role or instructions. " + NEWS_SCOPE_REFUSAL,
                "articles": {"columns": [], "rows": []}, "mode": "guardrail"}
    if any(m in ql for m in DB_QUESTION_MARKERS):
        return {"reply": NEWS_SCOPE_REFUSAL, "articles": {"columns": [], "rows": []}, "mode": "guardrail"}

    months = _period_months(question)
    factor = None
    for fk, words in FACTOR_TAGS.items():
        if any(w in ql for w in words[:3]):
            factor = fk
            break

    if months:
        arts = []
        for m in months:
            arts += get_news(month=m, factor=factor, limit=40)
    else:
        arts = get_news(factor=factor, limit=25)
    arts = sorted({a["title"]: a for a in arts}.values(),
                  key=lambda a: a["published"], reverse=True)[:20]

    table = {"columns": ["date", "title", "source", "factors", "link"],
             "rows": [[a["published"], a["title"], a["source"],
                       a["factors"], a["url"]] for a in arts]}

    period_label = (", ".join(months[:3]) + ("…" if months and len(months) > 3 else "")
                    if months else "recent coverage")

    if not arts:
        return {"reply": f"No stored articles match {period_label}"
                         + (f" for that factor" if factor else "")
                         + ". Try the Refresh button to pull the latest live "
                           "feeds, or ask about a different period.",
                "articles": table, "mode": "basic"}

    trend_q = any(w in ql for w in ["trend", "future", "outlook", "expect",
                                    "going forward", "next"])
    mode = llm_mode()
    if mode != "none":
        model = os.getenv("OPENAI_NEWS_MODEL", "gpt-4.1-mini")
        sys_p = (
            "You are the News Intelligence assistant of an aluminum price "
            "platform. Answer ONLY from the provided articles: summarize what "
            "the coverage says about aluminum prices and their drivers (LME, "
            "Midwest premium, gas/energy, labour, macro supply-demand, "
            "external/tariffs) for the requested period, referencing articles "
            "as [1], [2] in listed order. For trend/outlook questions, infer "
            "cautiously from the articles with hedged language — never give "
            "price numbers from memory, never answer database/value questions "
            "(redirect those to the main assistant), no financial advice, no "
            "markdown headers, be concise.")
        arts_txt = "\n".join(
            f"[{i+1}] {a['published']} | {a['title']} | {a['source']} | "
            f"factors: {a['factors']} | {a['summary']}"
            for i, a in enumerate(arts))
        out = _call_llm(sys_p, f"Question: {question}\n\nArticles:\n{arts_txt}",
                        max_tokens=700, model=model)
        if out:
            return {"reply": out, "articles": table, "mode": f"{mode}:{model}"}

    # template fallback
    by_factor = {}
    for a in arts:
        for f in a["factors"].split(","):
            by_factor.setdefault(f, 0)
        for f in a["factors"].split(","):
            by_factor[f] += 1
    top = sorted(by_factor.items(), key=lambda x: -x[1])[:3]
    lines = [f"News summary for {period_label} — {len(arts)} article(s), "
             f"most covered drivers: "
             + ", ".join(f"{k} ({v})" for k, v in top) + "."]
    for i, a in enumerate(arts[:5], 1):
        lines.append(f"[{i}] {a['published']} — {a['title']} ({a['source']}): "
                     f"{a['summary'][:120]}")
    if trend_q:
        lines.append("\nTrend read (from coverage only): themes above suggest "
                     "the direction of travel, but this is a qualitative "
                     "reading of headlines — not a forecast or advice.")
    lines.append("\nAll articles with links are in the table below.")
    return {"reply": "\n".join(lines), "articles": table, "mode": "basic"}

# ------------------------------------------------- GDELT historical backfill
# GDELT DOC 2.0 API — free, no key, indexes worldwide news back to Jan 2017.
# https://api.gdeltproject.org/api/v2/doc/doc
GDELT_QUERIES = {
    "lme":             '"aluminum price" OR "aluminium price" OR "LME aluminium"',
    "midwest_premium": '"midwest premium" aluminum',
    "gas":             'aluminum smelter energy prices',
    "labour":          'aluminum smelter workers strike wages',
    "macro":           'aluminum supply demand China production',
    "external":        'aluminum tariffs sanctions exports',
}

GDELT_MIN = "2017-01"


def _month_range(start: str, end: str) -> list[tuple[str, str]]:
    """Quarter slices [(YYYYMMDDHHMMSS start, end), ...] between two months."""
    ys, ms = int(start[:4]), int(start[5:7])
    ye, me = int(end[:4]), int(end[5:7])
    slices, y, m = [], ys, ((ms - 1) // 3) * 3 + 1
    while (y, m) <= (ye, me):
        m2, y2 = m + 2, y
        last_day = "31" if m2 in (1, 3, 5, 7, 8, 10, 12) else \
                   ("30" if m2 != 2 else "28")
        slices.append((f"{y:04d}{m:02d}01000000",
                       f"{y2:04d}{m2:02d}{last_day}235959",
                       f"{y:04d}-{m:02d}"))
        m += 3
        if m > 12:
            m, y = 1, y + 1
    return slices


def backfill_gdelt(start: str, end: str, per_slice: int = 25) -> dict:
    """Fetch historical articles from GDELT for [start, end] months (inclusive),
    one call per factor per quarter. Stores into the same news table."""
    import httpx
    _ensure_table()
    start = max(start, GDELT_MIN)
    inserted, calls, errors = 0, 0, 0
    conn = _conn()
    try:
        for s_dt, e_dt, label in _month_range(start, end):
            for fk, q in GDELT_QUERIES.items():
                calls += 1
                url = ("https://api.gdeltproject.org/api/v2/doc/doc?"
                       + urllib.parse.urlencode({
                           "query": f"({q}) sourcelang:english",
                           "mode": "artlist", "format": "json",
                           "maxrecords": per_slice,
                           "startdatetime": s_dt, "enddatetime": e_dt,
                           "sort": "hybridrel"}))
                try:
                    r = httpx.get(url, timeout=20)
                    r.raise_for_status()
                    arts = r.json().get("articles", [])
                except Exception:
                    errors += 1
                    continue
                for a in arts:
                    title = _strip_html(a.get("title", ""))
                    link = a.get("url", "")
                    if not title or not link or not _relevant(title):
                        continue
                    seen = a.get("seendate", "")          # 20220315T120000Z
                    try:
                        published = f"{seen[:4]}-{seen[4:6]}-{seen[6:8]}"
                        month = published[:7]
                    except Exception:
                        continue
                    factors = ",".join(sorted(set(_classify(title) + [fk])))
                    cur = conn.execute(
                        "INSERT OR IGNORE INTO news (published, month, title, "
                        "summary, url, source, factors, origin, fetched_at) "
                        "VALUES (?,?,?,?,?,?,?,?,?)",
                        (published, month, title,
                         "(headline via GDELT historical index)",
                         link, a.get("domain", "GDELT"), factors,
                         "live:gdelt_backfill", int(time.time())))
                    inserted += cur.rowcount
                time.sleep(0.25)                          # be polite to the API
        conn.commit()
    finally:
        conn.close()
    return {"inserted": inserted, "api_calls": calls, "failed_calls": errors,
            "range": [start, end]}
