"""
seed_cache.py — One-time: load the latest output Excel into cache/dhis2_cache.db
so future incremental runs skip re-fetching past period data.

Run: .venv\Scripts\python seed_cache.py
"""

import sqlite3
import pandas as pd
from pathlib import Path
from datetime import datetime

CACHE_DB   = Path(__file__).parent / "cache" / "dhis2_cache.db"
OUTPUT_DIR = Path(__file__).parent / "output"

# Pick latest xlsx (prefer _patched, then regular)
xlsx_files = sorted(OUTPUT_DIR.glob("dhis2_acebay_74facilities_FY26_*.xlsx"), reverse=True)
if not xlsx_files:
    raise FileNotFoundError(f"No output Excel found in {OUTPUT_DIR}")
src = xlsx_files[0]
print(f"Seeding from: {src.name}")

raw = pd.read_excel(src, sheet_name="Raw")
print(f"  {len(raw)} rows in Raw sheet")

CACHE_DB.parent.mkdir(exist_ok=True)
conn = sqlite3.connect(CACHE_DB)
conn.executescript("""
    CREATE TABLE IF NOT EXISTS cache (
        ou_uid    TEXT,
        period    TEXT,
        col_name  TEXT,
        value     REAL,
        view      TEXT,
        pulled_at TEXT,
        PRIMARY KEY (ou_uid, period, col_name)
    );
    CREATE TABLE IF NOT EXISTS run_log (
        id           INTEGER PRIMARY KEY AUTOINCREMENT,
        mode         TEXT,
        ran_at       TEXT,
        open_periods TEXT,
        rows_api     INTEGER,
        rows_cache   INTEGER
    );
""")

now = datetime.now().isoformat()
inserted = 0
skipped  = 0

for _, row in raw.iterrows():
    ou_uid   = row.get("ou_uid")
    col_name = row.get("column")
    value    = row.get("value")
    view     = str(row.get("view", ""))

    # period key: Month label for monthly rows, Quarter label for quarterly
    month_val   = str(row.get("Month", "")).strip()
    quarter_val = str(row.get("Quarter", "")).strip()
    period = month_val if view == "monthly" and month_val else quarter_val

    if not ou_uid or pd.isna(ou_uid) or not col_name or not period:
        skipped += 1
        continue

    val = float(value) if pd.notna(value) else None
    conn.execute(
        "INSERT OR REPLACE INTO cache (ou_uid, period, col_name, value, view, pulled_at) "
        "VALUES (?, ?, ?, ?, ?, ?)",
        (str(ou_uid), period, str(col_name), val, view, now),
    )
    inserted += 1

conn.commit()

# Summary
total = conn.execute("SELECT COUNT(*) FROM cache").fetchone()[0]
facilities = conn.execute("SELECT COUNT(DISTINCT ou_uid) FROM cache").fetchone()[0]
periods    = conn.execute("SELECT COUNT(DISTINCT period) FROM cache").fetchone()[0]
conn.close()

print(f"\n  Inserted : {inserted:,} rows")
print(f"  Skipped  : {skipped} rows")
print(f"  Cache now: {total:,} total rows | {facilities} facilities | {periods} distinct periods")
print(f"  Location : {CACHE_DB}")
print("\nDone — future incremental runs will read past periods from this cache.")
