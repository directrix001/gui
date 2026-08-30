"""
Price-comparison data layer.

Pipeline: CSV → SQLite → API → dashboard.

On startup the bundled CSV (two sources: FRED PALUMUSDM global aluminum price
and LME Aluminium 3-month) is ingested into SQLite. All reads then come from
SQL — not the CSV — so swapping in a live feed later touches only this file.

Future API integration (no code changes needed elsewhere):
    set PRICE_SOURCE_MODE=api
    set PRICE_API_URL=https://your-feed.example.com/prices
    set PRICE_API_KEY=...            (optional, sent as X-API-Key)
The API must return JSON: [{"date": "YYYY-MM", "palumusdm": <num>, "lme_3m": <num>}, ...]
If the API call fails, reads fall back to SQL automatically.
"""
import os
import sqlite3
from pathlib import Path

import pandas as pd

DATA_DIR = Path(__file__).parent / "data"
CSV_PATH = DATA_DIR / "lme_comparison.csv"
DB_PATH = DATA_DIR / "market.db"

TABLE = "price_comparison"


def _connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def ingest_csv_to_sql(force: bool = False) -> int:
    """Load the CSV into SQLite. Returns number of rows in the table."""
    DATA_DIR.mkdir(exist_ok=True)
    conn = _connect()
    try:
        cur = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?", (TABLE,)
        )
        exists = cur.fetchone() is not None
        if exists and not force:
            n = conn.execute(f"SELECT COUNT(*) FROM {TABLE}").fetchone()[0]
            if n > 0:
                return n

        df = pd.read_csv(CSV_PATH)
        df["observation_date"] = df["observation_date"].astype(str)
        conn.execute(f"DROP TABLE IF EXISTS {TABLE}")
        conn.execute(
            f"""CREATE TABLE {TABLE} (
                    observation_date TEXT PRIMARY KEY,
                    palumusdm REAL,
                    lme_3m REAL
                )"""
        )
        conn.executemany(
            f"INSERT INTO {TABLE} (observation_date, palumusdm, lme_3m) VALUES (?, ?, ?)",
            [
                (
                    r.observation_date,
                    None if pd.isna(r.palumusdm) else float(r.palumusdm),
                    None if pd.isna(r.lme_3m) else float(r.lme_3m),
                )
                for r in df.itertuples()
            ],
        )
        conn.commit()
        return conn.execute(f"SELECT COUNT(*) FROM {TABLE}").fetchone()[0]
    finally:
        conn.close()


def _read_from_sql(months: int | None = None) -> list[dict]:
    conn = _connect()
    try:
        q = f"SELECT observation_date, palumusdm, lme_3m FROM {TABLE} ORDER BY observation_date"
        rows = [dict(r) for r in conn.execute(q).fetchall()]
        return rows[-months:] if months else rows
    finally:
        conn.close()


def _read_from_api(months: int | None = None) -> list[dict]:
    """Live-feed reader — activated with PRICE_SOURCE_MODE=api."""
    import httpx

    url = os.getenv("PRICE_API_URL")
    if not url:
        raise RuntimeError("PRICE_API_URL is not set.")
    headers = {}
    if os.getenv("PRICE_API_KEY"):
        headers["X-API-Key"] = os.environ["PRICE_API_KEY"]
    r = httpx.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    data = r.json()
    rows = [
        {
            "observation_date": str(d["date"]),
            "palumusdm": d.get("palumusdm"),
            "lme_3m": d.get("lme_3m"),
        }
        for d in data
    ]
    rows.sort(key=lambda x: x["observation_date"])
    return rows[-months:] if months else rows


def get_comparison(months: int | None = None) -> tuple[list[dict], str]:
    """Returns (rows, source_used). Prefers API when configured, falls back to SQL."""
    if os.getenv("PRICE_SOURCE_MODE", "sql").lower() == "api":
        try:
            return _read_from_api(months), "api"
        except Exception:
            return _read_from_sql(months), "sql (api fallback)"
    return _read_from_sql(months), "sql"
