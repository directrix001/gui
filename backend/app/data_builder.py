"""
Builds data/master_data.csv — the single long-format dataset behind master.db.

Coverage: 2008-01 .. 2026-06 (222 months) for ALL six factors, so every
cross-factor join is complete.

Sources, tagged per row in the `source` column:
  • lme               → REAL data (user-provided FRED PALUMUSDM / LME 3-month csv)
  • all other factors → 'approx:public_anchor' — monthly interpolation between
    documented public anchor points (e.g. 2018 Section-232 premium spike, the
    2021-22 energy crisis, 2025 tariff era), with small deterministic noise.
    Replace with real series by editing this file's ANCHORS or by dropping a
    new master_data.csv in place — the DB rebuilds from CSV at startup.

Run manually to regenerate:  python -m app.data_builder
"""
import csv
from pathlib import Path

import numpy as np

DATA_DIR = Path(__file__).parent / "data"
OUT_CSV = DATA_DIR / "master_data.csv"
LME_CSV = DATA_DIR / "lme_comparison.csv"

START, END = (2008, 1), (2026, 6)


def _months():
    y, m = START
    out = []
    while (y, m) <= END:
        out.append(f"{y:04d}-{m:02d}")
        m += 1
        if m > 12:
            m, y = 1, y + 1
    return out

MONTHS = _months()
IDX = {m: i for i, m in enumerate(MONTHS)}

# anchor points: (month, value USD/t). Linear interpolation between anchors.
ANCHORS = {
    "midwest_premium": [
        ("2008-01", 90), ("2009-06", 70), ("2011-06", 130), ("2013-06", 260),
        ("2014-11", 440), ("2015-10", 180), ("2016-12", 165), ("2017-12", 190),
        ("2018-06", 440), ("2019-06", 400), ("2020-06", 240), ("2021-06", 500),
        ("2021-12", 650), ("2022-05", 880), ("2022-12", 480), ("2023-09", 420),
        ("2024-09", 460), ("2025-03", 700), ("2025-08", 880), ("2026-06", 900),
    ],
    "gas": [
        ("2008-07", 130), ("2009-06", 60), ("2012-04", 50), ("2014-02", 78),
        ("2016-03", 42), ("2018-11", 65), ("2020-06", 38), ("2021-10", 150),
        ("2022-08", 205), ("2023-04", 92), ("2024-02", 68), ("2025-01", 82),
        ("2026-04", 112), ("2026-06", 108),
    ],
    "labour": [
        ("2008-01", 95), ("2012-01", 105), ("2016-01", 118), ("2019-01", 130),
        ("2020-06", 128), ("2021-06", 138), ("2022-06", 152), ("2023-06", 163),
        ("2024-06", 172), ("2025-06", 181), ("2026-06", 190),
    ],
    # supply − demand balance: positive = surplus (pushes price DOWN),
    # negative = deficit (pushes price UP)
    "macro": [
        ("2008-01", 40), ("2009-03", 130), ("2011-01", 60), ("2013-01", 95),
        ("2015-09", 115), ("2017-06", 55), ("2018-06", 45), ("2020-05", 135),
        ("2021-09", 25), ("2022-03", -25), ("2022-10", 40), ("2023-08", 85),
        ("2024-06", 55), ("2025-06", 20), ("2026-01", -15), ("2026-06", -5),
    ],
    "external": [
        ("2008-01", 30), ("2010-01", 25), ("2014-06", 45), ("2018-04", 95),
        ("2019-06", 75), ("2020-04", 40), ("2022-03", 125), ("2023-01", 80),
        ("2024-04", 95), ("2025-04", 145), ("2026-06", 120),
    ],
}


def _interpolate(anchors: list[tuple[str, float]], noise_seed: int) -> list[float]:
    xs = [IDX[m] for m, _ in anchors]
    ys = [v for _, v in anchors]
    series = np.interp(range(len(MONTHS)), xs, ys)
    rng = np.random.default_rng(noise_seed)
    noise = rng.normal(0, 0.02, len(series)) * np.abs(series).clip(min=10)
    return [round(float(v), 2) for v in series + noise]


def _real_lme() -> dict[str, float]:
    """LME 3-month where available, FRED PALUMUSDM as fill — user's real data."""
    out = {}
    with open(LME_CSV) as f:
        for row in csv.DictReader(f):
            v = row["lme_3m"] or row["palumusdm"]
            if v:
                out[row["observation_date"]] = round(float(v), 2)
    return out


def build_csv() -> int:
    DATA_DIR.mkdir(exist_ok=True)
    rows = []
    lme = _real_lme()
    for m in MONTHS:
        if m in lme:
            rows.append((m, "lme", lme[m], "real:fred_lme_csv"))
    for i, (key, anchors) in enumerate(ANCHORS.items()):
        vals = _interpolate(anchors, noise_seed=100 + i)
        for m, v in zip(MONTHS, vals):
            rows.append((m, key, v, "approx:public_anchor"))
    with open(OUT_CSV, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["month", "factor_key", "value", "source"])
        w.writerows(rows)
    return len(rows)


if __name__ == "__main__":
    print(f"wrote {build_csv()} rows → {OUT_CSV}")
