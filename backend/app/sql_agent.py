"""
SQL agent — converts natural language to a validated, read-only SQLite query
against master.db, executes it, and returns the SQL + results table.

Safety model (the query is ALWAYS checked, whoever wrote it):
  • single statement, must start with SELECT or WITH
  • forbidden keywords blocked (INSERT/UPDATE/DELETE/DROP/ATTACH/PRAGMA/...)
  • LIMIT forced (appended if missing, capped at 200)
  • executed on a read-only SQLite connection (mode=ro) — writes are impossible
    even if something slipped through

Generation:
  • with an OpenAI/Azure key → gpt-4o-mini writes the SQL from the schema doc
  • without a key → rule-based templates cover the common question shapes
"""
import os
import re

from . import factor_dbs as fdb

FORBIDDEN = re.compile(
    r"\b(insert|update|delete|drop|alter|create|replace|attach|detach|pragma|"
    r"vacuum|reindex|trigger|grant|revoke)\b", re.I)

FACTOR_WORDS = {
    "lme": "lme", "base price": "lme", "price": "lme",
    "premium": "midwest_premium", "midwest": "midwest_premium",
    "gas": "gas", "cng": "gas", "energy": "gas",
    "labour": "labour", "labor": "labour", "ppi": "labour", "index": "labour",
}


def validate(sql: str) -> str:
    """Return a safe single SELECT statement or raise ValueError."""
    s = sql.strip().rstrip(";").strip()
    s = re.sub(r"^```(sql)?|```$", "", s, flags=re.I | re.M).strip()
    if ";" in s:
        raise ValueError("Multiple statements are not allowed.")
    if not re.match(r"^(select|with)\b", s, re.I):
        raise ValueError("Only SELECT queries are allowed.")
    if FORBIDDEN.search(s):
        raise ValueError("Query contains a forbidden keyword.")
    m = re.search(r"\blimit\s+(\d+)", s, re.I)
    if m:
        if int(m.group(1)) > 200:
            s = re.sub(r"\blimit\s+\d+", "LIMIT 200", s, flags=re.I)
    else:
        s += " LIMIT 200"
    return s


def _detect_factor(q: str) -> str:
    for w, key in FACTOR_WORDS.items():
        if w in q:
            return key
    return "lme"


def _detect_year(q: str):
    m = re.search(r"\b(20\d{2})\b", q)
    return m.group(1) if m else None


def template_sql(question: str, month: str | None) -> str:
    """Rule-based NL→SQL for the common shapes (no LLM needed)."""
    q = question.lower()
    fk = _detect_factor(q)
    year = _detect_year(q)
    n = 5
    mn = re.search(r"\btop\s+(\d+)", q)
    if mn:
        n = min(int(mn.group(1)), 50)

    if "average" in q or "avg" in q or "mean" in q:
        if "per year" in q or "by year" in q or "yearly" in q or "each year" in q:
            return (f"SELECT substr(month,1,4) AS year, ROUND(AVG(value),2) AS avg_value "
                    f"FROM monthly_values WHERE factor_key='{fk}'"
                    + (f" AND month >= '{year}-01'" if year and "since" in q else "")
                    + " GROUP BY year ORDER BY year")
        scope = f" AND substr(month,1,4)='{year}'" if year else ""
        return (f"SELECT ROUND(AVG(value),2) AS avg_value, MIN(value) AS min_value, "
                f"MAX(value) AS max_value FROM monthly_values "
                f"WHERE factor_key='{fk}'{scope}")

    if "top" in q or "highest" in q or "peak" in q or "maximum" in q:
        return (f"SELECT month, value FROM monthly_values WHERE factor_key='{fk}' "
                f"ORDER BY value DESC LIMIT {n}")
    if "lowest" in q or "minimum" in q or "cheapest" in q:
        return (f"SELECT month, value FROM monthly_values WHERE factor_key='{fk}' "
                f"ORDER BY value ASC LIMIT {n}")

    if "event" in q:
        return ("SELECT e.month, e.title, e.category, e.impact, "
                "GROUP_CONCAT(ef.factor_key) AS factors "
                "FROM events e JOIN event_factors ef ON ef.event_id=e.id "
                + (f"WHERE ef.factor_key='{fk}' " if any(
                    w in q for w in FACTOR_WORDS) else "")
                + "GROUP BY e.id ORDER BY e.month")

    if ("compare" in q or "side by side" in q or "all factor" in q) or (
            month and "everything" in q):
        cases = ", ".join(
            f"MAX(CASE WHEN factor_key='{k}' THEN value END) AS {k}"
            for k in fdb.FACTORS)
        scope = f" WHERE month='{month}'" if month else \
                (f" WHERE substr(month,1,4)='{year}'" if year else "")
        return f"SELECT month, {cases} FROM monthly_values{scope} GROUP BY month ORDER BY month"

    if month:
        return (f"SELECT mv.month, f.name AS factor, mv.value, mv.source "
                f"FROM monthly_values mv JOIN factors f ON f.key=mv.factor_key "
                f"WHERE mv.month='{month}' ORDER BY f.key")
    if year:
        return (f"SELECT month, value FROM monthly_values "
                f"WHERE factor_key='{fk}' AND substr(month,1,4)='{year}' ORDER BY month")
    return (f"SELECT month, value FROM monthly_values WHERE factor_key='{fk}' "
            f"ORDER BY month DESC LIMIT 12")


FEWSHOT = """Q: average lme price per year since 2020
SQL: SELECT substr(month,1,4) AS year, ROUND(AVG(value),2) AS avg_lme FROM monthly_values WHERE factor_key='lme' AND month>='2020-01' GROUP BY year ORDER BY year
Q: show all factors for March 2026
SQL: SELECT month, MAX(CASE WHEN factor_key='lme' THEN value END) AS lme, MAX(CASE WHEN factor_key='midwest_premium' THEN value END) AS midwest_premium, MAX(CASE WHEN factor_key='gas' THEN value END) AS gas, MAX(CASE WHEN factor_key='labour' THEN value END) AS labour FROM monthly_values WHERE month='2026-03' GROUP BY month
Q: which events pushed the premium up?
SQL: SELECT e.month, e.title, e.note FROM events e JOIN event_factors ef ON ef.event_id=e.id WHERE ef.factor_key='midwest_premium' AND e.impact='up' ORDER BY e.month"""


def llm_sql(question: str) -> str | None:
    from .agents import _call_llm  # reuse the same LLM helper
    sys_p = (
        "You translate questions into a SINGLE SQLite SELECT query for this "
        "schema. Rules: SELECT/WITH only; no comments; no semicolons; always "
        "include LIMIT <=200 for row-listing queries; months are 'YYYY-MM' "
        "strings; use the pivot pattern for side-by-side factors. Reply with "
        "the SQL only, nothing else.\n\n" + fdb.SCHEMA_DOC + "\n\n" + FEWSHOT)
    out = _call_llm(sys_p, question, max_tokens=300)
    return out.strip() if out else None


def query(question: str, month: str | None = None) -> dict:
    """Full pipeline: generate → validate → execute. Never raises."""
    generator = "template"
    sql = None
    from .agents import llm_mode
    if llm_mode() != "none":
        candidate = llm_sql(question)
        if candidate:
            try:
                sql = validate(candidate)
                generator = "gpt-4o-mini"
            except ValueError:
                sql = None
    if sql is None:
        try:
            sql = validate(template_sql(question, month))
        except ValueError as e:
            return {"ok": False, "error": str(e), "generator": generator}
    try:
        result = fdb.run_sql(sql)
        return {"ok": True, "sql": sql, "generator": generator, **result}
    except Exception as e:
        return {"ok": False, "sql": sql, "generator": generator,
                "error": f"SQL execution failed: {e}"}
