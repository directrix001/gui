"""
Real-data factor database layer.

Four SEPARATE SQLite databases, one per driver, each a single table with month
as PRIMARY KEY (data/factors/<key>.db):

    monthly_values(month TEXT PRIMARY KEY, value REAL, source TEXT)

    lme.db              LME Price ($/lb)          real: lme sheet (yahoo finance)
    midwest_premium.db  Midwest Premium ($/lb)    real: yahoo finance + lme midwest site
    gas.db              CNG Cost ($/lb)           real: client CNG sheet
    labour.db           PPI Index (dimensionless) real: client PPI sheet

master.db keeps shared tables: factors catalog, events + event_factors,
price_comparison (FRED vs LME USD/t history), news.

The SQL agent still sees ONE logical schema: run_sql() ATTACHes the four
factor DBs read-only and creates a temp view `monthly_values` that UNIONs
them with a factor_key column, so pivots and joins keep working unchanged.
"""
import csv
import sqlite3
from pathlib import Path

DATA_DIR = Path(__file__).parent / "data"
FACTOR_DIR = DATA_DIR / "factors"
REAL_DIR = DATA_DIR / "real"
MASTER_DB = DATA_DIR / "master.db"
LME_CSV = DATA_DIR / "lme_comparison.csv"

FACTORS = {
    "lme":             {"name": "LME Price",       "unit": "$/lb",  "db": "lme.db"},
    "midwest_premium": {"name": "Midwest Premium", "unit": "$/lb",  "db": "midwest_premium.db"},
    "gas":             {"name": "CNG Cost",        "unit": "$/lb",  "db": "gas.db"},
    "labour":          {"name": "PPI Index",       "unit": "index", "db": "labour.db"},
}

EVENTS = [
    ("2025-10", "Premium climbs on tariff-era import costs", "tariff", "up",
     "Duty-paid Midwest premium stepped up as tariff costs passed through.", ["midwest_premium"]),
    ("2026-01", "PPI jump signals broad input-cost inflation", "macro_shock", "up",
     "Producer price index moved sharply higher into the new year.", ["labour"]),
    ("2026-01", "Aluminium rally extends into the new year", "macro_shock", "up",
     "LME pricing cleared multi-month resistance on restocking.", ["lme"]),
    ("2026-03", "LME spikes toward cycle highs", "geopolitical", "up",
     "Supply-fear rally lifted exchange prices sharply.", ["lme", "midwest_premium"]),
    ("2026-04", "Energy cost pass-through firms", "energy", "up",
     "Energy benchmarks fed into smelting cost expectations.", ["gas", "lme"]),
    ("2026-06", "Prices ease on profit-taking", "macro_shock", "down",
     "LME retraced from May peak as length was trimmed.", ["lme"]),
    ("2026-08", "Sharp LME pullback", "macro_shock", "down",
     "Exchange price corrected on demand worries.", ["lme"]),
]

SCHEMA_DOC = """Four separate SQLite factor databases exposed through one
logical view. Months are TEXT 'YYYY-MM'; month is the PRIMARY KEY in every
factor database.

VIEW monthly_values(month TEXT, factor_key TEXT, value REAL, source TEXT)
  factor_key values: lme ($/lb), midwest_premium ($/lb), gas (CNG $/lb),
  labour (PPI index, dimensionless). Real data, roughly 2025-07 .. 2027-12.
TABLE factors(key TEXT PK, name TEXT, unit TEXT)
TABLE events(id INTEGER PK, month TEXT, title TEXT, category TEXT,
  impact TEXT('up'/'down'), note TEXT)
TABLE event_factors(event_id INTEGER FK->events.id, factor_key TEXT FK->factors.key)
TABLE price_comparison(month TEXT PK, fred_palumusdm REAL, lme_3m REAL)  -- USD/t history 2008..2026

Pivot example (factors side by side per month):
  SELECT month,
    MAX(CASE WHEN factor_key='lme' THEN value END) AS lme,
    MAX(CASE WHEN factor_key='midwest_premium' THEN value END) AS premium
  FROM monthly_values GROUP BY month;
Events for a factor:
  SELECT e.* FROM events e JOIN event_factors ef ON ef.event_id=e.id
  WHERE ef.factor_key='lme';"""


def _ro_connect(path) -> sqlite3.Connection:
    """Read-only connection that works on Windows and Linux.
    Windows paths break the naive f"file:{path}" URI (backslashes, drive
    colon), so build a proper file URI via pathlib; if the URI form still
    fails on an exotic setup, fall back to a normal connection — queries are
    already SELECT-only-validated upstream, so safety is preserved."""
    try:
        uri = Path(path).resolve().as_uri() + "?mode=ro"
        return sqlite3.connect(uri, uri=True)
    except Exception:
        return sqlite3.connect(str(path))


def _factor_conn(key: str, readonly: bool = False) -> sqlite3.Connection:
    path = FACTOR_DIR / FACTORS[key]["db"]
    c = _ro_connect(path) if readonly else sqlite3.connect(str(path))
    c.row_factory = sqlite3.Row
    return c


def _master_conn(readonly: bool = False) -> sqlite3.Connection:
    c = _ro_connect(MASTER_DB) if readonly else sqlite3.connect(str(MASTER_DB))
    c.row_factory = sqlite3.Row
    return c


def build_master() -> dict:
    FACTOR_DIR.mkdir(parents=True, exist_ok=True)
    counts = {}
    for key, meta in FACTORS.items():
        conn = _factor_conn(key)
        try:
            conn.execute("""CREATE TABLE IF NOT EXISTS monthly_values (
                month TEXT PRIMARY KEY, value REAL NOT NULL,
                source TEXT NOT NULL DEFAULT 'real:client_sheet')""")
            csv_path = REAL_DIR / f"{key}.csv"
            if csv_path.exists():
                with open(csv_path) as f:
                    rows = [(r["month"], float(r["value"]), "real:client_sheet")
                            for r in csv.DictReader(f)]
                conn.executemany(
                    "INSERT OR REPLACE INTO monthly_values VALUES (?,?,?)", rows)
            conn.commit()
            counts[key] = conn.execute(
                "SELECT COUNT(*) FROM monthly_values").fetchone()[0]
        finally:
            conn.close()

    reg = _master_conn()
    try:
        reg.executescript("""
            DROP TABLE IF EXISTS factors;
            DROP TABLE IF EXISTS events;
            DROP TABLE IF EXISTS event_factors;
            DROP TABLE IF EXISTS price_comparison;
            DROP TABLE IF EXISTS monthly_values;
            CREATE TABLE factors(key TEXT PRIMARY KEY, name TEXT, unit TEXT);
            CREATE TABLE events(id INTEGER PRIMARY KEY, month TEXT, title TEXT,
                category TEXT, impact TEXT, note TEXT);
            CREATE TABLE event_factors(
                event_id INTEGER REFERENCES events(id),
                factor_key TEXT REFERENCES factors(key),
                PRIMARY KEY(event_id, factor_key));
            CREATE TABLE price_comparison(
                month TEXT PRIMARY KEY, fred_palumusdm REAL, lme_3m REAL);
        """)
        reg.executemany("INSERT INTO factors VALUES (?,?,?)",
            [(k, m["name"], m["unit"]) for k, m in FACTORS.items()])
        for i, (month, title, cat, impact, note, fks) in enumerate(EVENTS, start=1):
            reg.execute("INSERT INTO events VALUES (?,?,?,?,?,?)",
                        (i, month, title, cat, impact, note))
            reg.executemany("INSERT INTO event_factors VALUES (?,?)",
                            [(i, fk) for fk in fks])
        if LME_CSV.exists():
            with open(LME_CSV) as f:
                cmp_rows = [(r["observation_date"],
                             float(r["palumusdm"]) if r["palumusdm"] else None,
                             float(r["lme_3m"]) if r["lme_3m"] else None)
                            for r in csv.DictReader(f)]
            reg.executemany(
                "INSERT OR REPLACE INTO price_comparison VALUES (?,?,?)", cmp_rows)
        reg.commit()
        counts["events"] = len(EVENTS)
    finally:
        reg.close()
    return counts


def _union_view_sql() -> str:
    parts = [
        f"SELECT month, '{k}' AS factor_key, value, source "
        f"FROM {k}_db.monthly_values"
        for k in FACTORS
    ]
    return "CREATE TEMP VIEW monthly_values AS " + " UNION ALL ".join(parts)


def run_sql(sql: str, limit: int = 200) -> dict:
    conn = _master_conn(readonly=True)
    try:
        for k, meta in FACTORS.items():
            path = (FACTOR_DIR / meta["db"]).resolve()
            try:
                conn.execute(f"ATTACH DATABASE ? AS {k}_db",
                             (path.as_uri() + "?mode=ro",))
            except Exception:
                conn.execute(f"ATTACH DATABASE ? AS {k}_db", (str(path),))
        conn.execute(_union_view_sql())
        cur = conn.execute(sql)
        cols = [d[0] for d in cur.description] if cur.description else []
        rows = [list(r) for r in cur.fetchmany(limit)]
        return {"columns": cols, "rows": rows, "row_count": len(rows)}
    finally:
        conn.close()


def upsert_rows(key: str, rows: list[tuple[str, float]],
                source: str = "real:manual") -> int:
    if key not in FACTORS:
        raise ValueError(f"Unknown factor '{key}'")
    conn = _factor_conn(key)
    try:
        conn.execute("""CREATE TABLE IF NOT EXISTS monthly_values (
            month TEXT PRIMARY KEY, value REAL NOT NULL,
            source TEXT NOT NULL DEFAULT 'real:manual')""")
        conn.executemany(
            "INSERT OR REPLACE INTO monthly_values VALUES (?,?,?)",
            [(m, v, source) for m, v in rows])
        conn.commit()
        return len(rows)
    finally:
        conn.close()


def get_factor_series(key: str) -> list[dict]:
    conn = _factor_conn(key, readonly=True)
    try:
        return [dict(r) for r in conn.execute(
            "SELECT month, value, source FROM monthly_values ORDER BY month")]
    finally:
        conn.close()


def get_factor_value(key: str, month: str):
    conn = _factor_conn(key, readonly=True)
    try:
        r = conn.execute("SELECT value FROM monthly_values WHERE month=?",
                         (month,)).fetchone()
        return r["value"] if r else None
    finally:
        conn.close()


def month_snapshot(month: str) -> dict:
    out = {}
    for key, meta in FACTORS.items():
        conn = _factor_conn(key, readonly=True)
        try:
            rows = conn.execute(
                "SELECT month, value FROM monthly_values WHERE month <= ? "
                "ORDER BY month DESC LIMIT 2", (month,)).fetchall()
        finally:
            conn.close()
        cur = rows[0] if rows and rows[0]["month"] == month else None
        prev = rows[1] if len(rows) > 1 and cur else (rows[0] if rows and not cur else None)
        cv = cur["value"] if cur else None
        pv = prev["value"] if prev else None
        out[key] = {
            "name": meta["name"], "unit": meta["unit"], "value": cv, "prev": pv,
            "change": round(cv - pv, 4) if cv is not None and pv is not None else None,
            "change_pct": round((cv - pv) / pv * 100, 2)
                          if cv is not None and pv not in (None, 0) else None,
        }
    return out


def events_near(month: str, window: int = 1) -> list[dict]:
    y, m = int(month[:4]), int(month[5:7])
    keep = set()
    for d in range(-window, window + 1):
        mm = m + d
        yy = y + (mm - 1) // 12
        mm = (mm - 1) % 12 + 1
        keep.add(f"{yy:04d}-{mm:02d}")
    conn = _master_conn(readonly=True)
    try:
        rows = conn.execute(
            f"""SELECT e.*, GROUP_CONCAT(ef.factor_key) AS factor_keys
                FROM events e LEFT JOIN event_factors ef ON ef.event_id = e.id
                WHERE e.month IN ({','.join('?' * len(keep))})
                GROUP BY e.id ORDER BY e.month""", list(keep)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def db_inventory() -> list[dict]:
    inv = []
    for key, meta in FACTORS.items():
        try:
            conn = _factor_conn(key, readonly=True)
            n, lo, hi = conn.execute(
                "SELECT COUNT(*), MIN(month), MAX(month) FROM monthly_values").fetchone()
            src = conn.execute("SELECT source FROM monthly_values LIMIT 1").fetchone()
            conn.close()
        except Exception:
            n, lo, hi, src = 0, None, None, None
        inv.append({"key": key, "db": meta["db"], "name": meta["name"],
                    "unit": meta["unit"], "rows": n, "from": lo, "to": hi,
                    "primary_source": src["source"] if src else None})
    return inv
