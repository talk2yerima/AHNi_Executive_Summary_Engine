"""
service_runner.py  —  AHNi Executive Service Engine (AESE)
Wraps dhis2_pull.py as a long-running process managed by NSSM.

Schedule:
  - 00:00 every day  → full pull  (re-pulls all periods)
  - Every 3 hours    → incremental pull (API for open period only)
"""

import logging
import signal
import sqlite3
import subprocess
import sys
import time
from datetime import date, datetime, timedelta
from logging.handlers import RotatingFileHandler
from pathlib import Path

INTERVAL_HOURS = 3

SCRIPT_DIR  = Path(__file__).parent
PYTHON      = SCRIPT_DIR / ".venv" / "Scripts" / "python.exe"
PULL_SCRIPT = SCRIPT_DIR / "dhis2_pull.py"
CACHE_DB    = SCRIPT_DIR / "cache" / "dhis2_cache.db"
LOG_DIR     = SCRIPT_DIR / "logs"

# ── Logging  (short name: AESE) ───────────────────────────────────────────────
LOG_DIR.mkdir(exist_ok=True)
_fmt = logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
_fh  = RotatingFileHandler(LOG_DIR / "AESE.log", maxBytes=5_000_000, backupCount=5, encoding="utf-8")
_fh.setFormatter(_fmt)
_sh  = logging.StreamHandler(sys.stdout)
_sh.setFormatter(_fmt)
log  = logging.getLogger("AESE")
log.setLevel(logging.INFO)
log.addHandler(_fh)
log.addHandler(_sh)

# ── Graceful stop ─────────────────────────────────────────────────────────────
_stop = False

def _handle_stop(signum, frame):
    global _stop
    log.info("AESE — stop signal received. Will exit after current sleep.")
    _stop = True

signal.signal(signal.SIGTERM, _handle_stop)
signal.signal(signal.SIGINT,  _handle_stop)


# ── Helpers ───────────────────────────────────────────────────────────────────
def _cache_has_data():
    """Return True if the cache DB has at least one prior run."""
    if not CACHE_DB.exists():
        return False
    try:
        conn = sqlite3.connect(CACHE_DB)
        count = conn.execute("SELECT COUNT(*) FROM run_log").fetchone()[0]
        conn.close()
        return count > 0
    except Exception:
        return False


def run_pull(mode: str):
    log.info("=" * 65)
    log.info("AESE — starting pull  mode=%s", mode)
    try:
        result = subprocess.run(
            [str(PYTHON), str(PULL_SCRIPT), "--mode", mode],
            cwd=str(SCRIPT_DIR),
            timeout=7_200,      # 2-hour hard limit per run
        )
        if result.returncode == 0:
            log.info("AESE — pull finished OK  mode=%s", mode)
        else:
            log.error("AESE — pull exited with code %d  mode=%s", result.returncode, mode)
    except subprocess.TimeoutExpired:
        log.error("AESE — pull timed out after 2 hours  mode=%s", mode)
    except Exception as exc:
        log.error("AESE — pull error: %s", exc)


def _next_wake(last_run: datetime) -> datetime:
    """
    Next wake time = earliest of:
      - last_run + 3 hours
      - tomorrow midnight (triggers daily full pull)
    """
    next_interval = last_run + timedelta(hours=INTERVAL_HOURS)
    tomorrow_midnight = (datetime.now() + timedelta(days=1)).replace(
        hour=0, minute=0, second=0, microsecond=0
    )
    return min(next_interval, tomorrow_midnight)


def sleep_until(wake_at: datetime):
    """Sleep in 1-second ticks until wake_at, respecting stop signal."""
    while not _stop:
        remaining = (wake_at - datetime.now()).total_seconds()
        if remaining <= 0:
            break
        if int(remaining) % 600 == 0 and remaining > 1:
            log.info("AESE — next pull in %d min.", int(remaining // 60))
        time.sleep(min(1, remaining))


# ── Main loop ─────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    log.info("AESE — AHNi Executive Service Engine starting.")
    log.info("AESE — schedule: full pull at midnight | incremental every %dh.", INTERVAL_HOURS)

    if not PYTHON.exists():
        log.error("AESE — Python venv not found at %s. Exiting.", PYTHON)
        sys.exit(1)

    last_full_date: date | None = None

    while not _stop:
        today = date.today()

        # Full pull: first start ever, fresh VM (no cache), or new calendar day
        if last_full_date != today or not _cache_has_data():
            run_pull("full")
            last_full_date = today
        else:
            run_pull("incremental")

        if _stop:
            break

        wake_at = _next_wake(datetime.now())
        log.info("AESE — sleeping until %s.", wake_at.strftime("%Y-%m-%d %H:%M:%S"))
        sleep_until(wake_at)

    log.info("AESE — service stopped.")
