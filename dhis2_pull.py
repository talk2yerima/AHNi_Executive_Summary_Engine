"""
dhis2_pull.py  —  Pull all ACEBAY indicators for 74 facilities, THIS_FY
Run: .venv\Scripts\python dhis2_pull.py

Output: output/dhis2_acebay_74facilities_FY26_<timestamp>.xlsx
  Sheet "Quarterly" — STATE, FACILITY, Qtr1..Qtr4 + FY26 total
  Sheet "Monthly"   — STATE, FACILITY, Oct-25..Sep-26
"""

import argparse
import calendar as _cal
import json
import os
import sqlite3
import sys
import time
import requests
import pandas as pd
from datetime import date, datetime as _dt
from pathlib import Path
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / ".env")

# ── Logging setup ─────────────────────────────────────────────────────────────
import logging
from logging.handlers import RotatingFileHandler as _RFH

_LOG_DIR = Path(__file__).parent / "logs"
_LOG_DIR.mkdir(exist_ok=True)
_fmt = logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
_fh  = _RFH(_LOG_DIR / "dhis2_pull.log", maxBytes=5_000_000, backupCount=5, encoding="utf-8")
_fh.setFormatter(_fmt)
_sh  = logging.StreamHandler()
_sh.setFormatter(_fmt)
log  = logging.getLogger("dhis2")
log.setLevel(logging.INFO)
log.addHandler(_fh)
log.addHandler(_sh)

# ── Startup validation ────────────────────────────────────────────────────────
_REQUIRED_VARS = ["DHIS2_URL", "DHIS2_USER", "DHIS2_PASS"]
_missing = [v for v in _REQUIRED_VARS if not os.getenv(v)]
if _missing:
    log.error("Missing required environment variable(s): %s", ", ".join(_missing))
    log.error("Set them in your .env file and try again.")
    sys.exit(1)

# ── Azure Blob helpers ────────────────────────────────────────────────────────
def _azure_configured():
    return bool(os.getenv("AZURE_CONNECTION_STRING") and os.getenv("AZURE_CONTAINER_NAME"))

def _azure_download_excel():
    from azure.storage.blob import BlobServiceClient
    client = BlobServiceClient.from_connection_string(os.getenv("AZURE_CONNECTION_STRING"))
    blob   = client.get_blob_client(
        container=os.getenv("AZURE_CONTAINER_NAME"),
        blob=os.getenv("AZURE_EXCEL_BLOB_NAME", "targetA.xlsx"),
    )
    import io
    data = blob.download_blob().readall()
    return pd.read_excel(io.BytesIO(data), engine="openpyxl")

def _azure_upload(local_path, blob_name):
    from azure.storage.blob import BlobServiceClient
    client = BlobServiceClient.from_connection_string(os.getenv("AZURE_CONNECTION_STRING"))
    blob   = client.get_blob_client(
        container=os.getenv("AZURE_CONTAINER_NAME"),
        blob=blob_name,
    )
    with open(local_path, "rb") as f:
        blob.upload_blob(f, overwrite=True)
    account = BlobServiceClient.from_connection_string(
        os.getenv("AZURE_CONNECTION_STRING")
    ).account_name
    return f"https://{account}.blob.core.windows.net/{os.getenv('AZURE_CONTAINER_NAME')}/{blob_name}"


# ── RADET helpers ─────────────────────────────────────────────────────────────
import io as _io

_RADET_MEM_CACHE   = {}   # date_str → {"radet_df": df, "hts_df": df}
_DATIMID_MAP_FILE  = Path(__file__).parent / "config" / "datimid_map.json"
_datimid_map: dict = {}   # populated lazily on first RADET download

_STATE_CODES = {"ad", "bo", "ta", "yo"}

def _strip_state_prefix(name: str) -> str:
    parts = name.strip().split(" ", 1)
    if len(parts) == 2 and parts[0].lower() in _STATE_CODES:
        return parts[1].strip()
    return name.strip()


def _build_datimid_map(radet_df) -> dict:
    """Fuzzy-match RADET DatimId → ACEBAY OU_UID via facility names. Saved to config/datimid_map.json."""
    from rapidfuzz import process, fuzz
    fac_names = {_strip_state_prefix(f["name"]): f["id"] for f in FACILITIES}
    unique = radet_df[["DatimId", "Facility Name"]].drop_duplicates("DatimId")
    mapping = {}
    for _, row in unique.iterrows():
        clean = _strip_state_prefix(str(row["Facility Name"]))
        if clean in fac_names:
            mapping[row["DatimId"]] = fac_names[clean]
        else:
            hit = process.extractOne(clean, list(fac_names.keys()), scorer=fuzz.WRatio, score_cutoff=65)
            if hit:
                mapping[row["DatimId"]] = fac_names[hit[0]]
            else:
                log.warning("   DatimId map: no match for '%s' (%s)", clean, row["DatimId"])
    _DATIMID_MAP_FILE.parent.mkdir(exist_ok=True)
    import json as _json
    _json.dump(mapping, open(_DATIMID_MAP_FILE, "w"), indent=2)
    log.info("   DatimId map: built %d/%d entries → config/datimid_map.json", len(mapping), len(unique))
    return mapping


def _load_datimid_map() -> dict:
    import json as _json
    if _DATIMID_MAP_FILE.exists():
        m = _json.load(open(_DATIMID_MAP_FILE))
        log.info("   DatimId map: loaded %d entries from config/datimid_map.json", len(m))
        return m
    return {}


def _quarter_start(q_raw: str) -> str:
    """'2026Q1' → '2026-01-01'"""
    yr = int(q_raw[:4])
    qn = int(q_raw[5])
    month = (qn - 1) * 3 + 1
    return f"{yr}-{month:02d}-01"


def _month_start(m_raw: str) -> str:
    """'202601' → '2026-01-01'"""
    return f"{m_raw[:4]}-{m_raw[4:]}-01"


def _download_radet_stream(snapshot_date_str: str):
    """Download RADET blob → BytesIO. Returns None on failure."""
    blob_name = RADET_BLOB_PATTERN.format(date=snapshot_date_str)
    log.info("      Downloading RADET: %s ...", blob_name)
    try:
        from azure.storage.blob import BlobServiceClient
        client = BlobServiceClient.from_connection_string(os.getenv("AZURE_CONNECTION_STRING"))
        blob   = client.get_container_client(RADET_CONTAINER).get_blob_client(blob_name)
        stream = _io.BytesIO()
        blob.download_blob().readinto(stream)
        stream.seek(0)
        log.info("      Download complete: %s", blob_name)
        return stream
    except Exception as e:
        log.error("      RADET download failed (%s): %s", blob_name, e)
        return None


def pull_radet_for_date(snapshot_date_str: str, period_start_str: str, period_end_str: str) -> dict:
    """
    Return {col_name: {ou_uid: float}} for the given period.

    snapshot_date_str : 'YYYY-MM-DD' — which RADET file to download
    period_start_str  : 'YYYY-MM-DD' — start of HTS flow period
    period_end_str    : 'YYYY-MM-DD' — end of HTS flow period (= snapshot date)

    Indicators computed:
      TX_CURR     — Active/Active-Restart patients (snapshot)
      TX_PVLS_D   — eligible for VL: Active, on ART ≥180 d, VL in last 365 d
      TX_PVLS_N   — TX_PVLS_D with Current Viral Load < 1000
      HTS_TST     — all HIV tests in [period_start, period_end]
      HTS_TST_POS — positive tests in [period_start, period_end]
    """
    global _datimid_map

    # ── Use in-memory cache to avoid re-downloading the same 119 MB file ──────
    if snapshot_date_str not in _RADET_MEM_CACHE:
        stream = _download_radet_stream(snapshot_date_str)
        if stream is None:
            return {}

        # CombinedRADET sheet ─────────────────────────────────────────────────
        RADET_COLS = [
            "DatimId", "Facility Name", "State",
            "Current ART Status",
            "ART Start Date (yyyy-mm-dd)",
            "Date of Current ViralLoad Result Sample (yyyy-mm-dd)",
            "Current Viral Load (c/ml)",
        ]
        try:
            radet_df = pd.read_excel(stream, sheet_name="CombinedRADET",
                                     usecols=RADET_COLS, engine="openpyxl")
        except Exception as e:
            log.error("      CombinedRADET sheet read failed: %s", e)
            radet_df = pd.DataFrame()

        # Build/load DatimId → OU_UID map on first successful load ─────────────
        if not radet_df.empty and not _datimid_map:
            _datimid_map = _load_datimid_map()
            if not _datimid_map:
                _datimid_map = _build_datimid_map(radet_df)

        # CombinedHTS sheet ───────────────────────────────────────────────────
        HTS_COLS = ["datimCode", "dateOfHIVTesting", "finalHIVTestResult"]
        stream.seek(0)
        try:
            hts_df = pd.read_excel(stream, sheet_name="CombinedHTS",
                                   usecols=HTS_COLS, engine="openpyxl")
        except Exception as e:
            log.error("      CombinedHTS sheet read failed: %s", e)
            hts_df = pd.DataFrame()

        _RADET_MEM_CACHE[snapshot_date_str] = {"radet_df": radet_df, "hts_df": hts_df}
        log.info("      RADET loaded: %d ART rows, %d HTS rows",
                 len(radet_df), len(hts_df))
    else:
        log.info("      RADET %s: using in-memory cache", snapshot_date_str)

    radet_df = _RADET_MEM_CACHE[snapshot_date_str]["radet_df"]
    hts_df   = _RADET_MEM_CACHE[snapshot_date_str]["hts_df"]
    dm       = _datimid_map
    fac_uid_set = {f["id"] for f in FACILITIES}

    snap_dt  = _dt.strptime(snapshot_date_str, "%Y-%m-%d").date()
    start_dt = _dt.strptime(period_start_str,  "%Y-%m-%d").date()
    end_dt   = _dt.strptime(period_end_str,    "%Y-%m-%d").date()

    result: dict = {}

    # ── TX_CURR ───────────────────────────────────────────────────────────────
    if not radet_df.empty:
        active = radet_df[
            radet_df["Current ART Status"].str.contains("active", case=False, na=False)
        ]
        tx_curr = active.groupby("DatimId").size()
        result["TX_CURR"] = {
            dm[d]: float(v) for d, v in tx_curr.items()
            if d in dm and dm[d] in fac_uid_set
        }

        # ── TX_PVLS base filter ────────────────────────────────────────────────
        pvls = radet_df[
            radet_df["Current ART Status"].isin(["Active", "Active Restart"])
        ].copy()
        pvls["vl_date"]  = pd.to_datetime(
            pvls["Date of Current ViralLoad Result Sample (yyyy-mm-dd)"], errors="coerce"
        ).dt.date
        pvls["art_date"] = pd.to_datetime(
            pvls["ART Start Date (yyyy-mm-dd)"], errors="coerce"
        ).dt.date

        one_year_ago = snap_dt - pd.Timedelta(days=365)
        pvls = pvls[
            pvls["vl_date"].notna() &
            pvls["art_date"].notna() &
            (pvls["vl_date"] <= snap_dt) &
            (pvls["vl_date"] >= one_year_ago) &
            ((pvls["vl_date"] - pvls["art_date"]).apply(
                lambda d: d.days if hasattr(d, "days") else -1
            ) >= 180)
        ]

        tx_pvls_d = pvls.groupby("DatimId").size()
        result["TX_PVLS_D"] = {
            dm[d]: float(v) for d, v in tx_pvls_d.items()
            if d in dm and dm[d] in fac_uid_set
        }

        # ── TX_PVLS_N (suppressed: VL < 1000) ─────────────────────────────────
        pvls["vl_num"] = pd.to_numeric(pvls["Current Viral Load (c/ml)"], errors="coerce")
        supp = pvls[pvls["vl_num"] < 1000]
        tx_pvls_n = supp.groupby("DatimId").size()
        result["TX_PVLS_N"] = {
            dm[d]: float(v) for d, v in tx_pvls_n.items()
            if d in dm and dm[d] in fac_uid_set
        }

    # ── HTS_TST / HTS_TST_POS (flow — filter by date range) ──────────────────
    if not hts_df.empty:
        hts_df["hts_date"] = pd.to_datetime(hts_df["dateOfHIVTesting"], errors="coerce").dt.date
        period_hts = hts_df[
            hts_df["hts_date"].notna() &
            (hts_df["hts_date"] >= start_dt) &
            (hts_df["hts_date"] <= end_dt)
        ]
        hts_tst = period_hts.groupby("datimCode").size()
        result["HTS_TST"] = {
            dm.get(d, d): float(v) for d, v in hts_tst.items()
            if dm.get(d, d) in fac_uid_set
        }
        pos_hts = period_hts[
            period_hts["finalHIVTestResult"].str.strip().str.lower() == "positive"
        ]
        hts_pos = pos_hts.groupby("datimCode").size()
        result["HTS_TST_POS"] = {
            dm.get(d, d): float(v) for d, v in hts_pos.items()
            if dm.get(d, d) in fac_uid_set
        }

    counts = {k: len(v) for k, v in result.items()}
    log.info("      RADET results: %s", counts)
    return result


# ── Config ────────────────────────────────────────────────────────────────────
BASE_URL = os.getenv("DHIS2_URL", "").rstrip("/")
USERNAME = os.getenv("DHIS2_USER")
PASSWORD = os.getenv("DHIS2_PASS")

# ── Dynamic FY config (auto-detects current PEPFAR fiscal year) ───────────────
def _build_fy():
    """
    Derive all period maps from today's date.
    PEPFAR FY = Oct–Sep, so FY26 = Oct-2025 to Sep-2026.
    Automatically advances to FY27 on Oct 1, 2026, etc.
    """
    today       = date.today()
    fy          = today.year + 1 if today.month >= 10 else today.year
    start_yr    = fy - 1   # calendar year FY starts (Oct)
    end_yr      = fy       # calendar year FY ends   (Sep)
    _MN = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

    quarters = [
        f"{start_yr}Q4", f"{end_yr}Q1", f"{end_yr}Q2", f"{end_yr}Q3",
    ]
    months = (
        [f"{start_yr}{m:02d}" for m in (10, 11, 12)] +
        [f"{end_yr}{m:02d}"   for m in range(1, 10)]
    )
    quarter_last_day = {
        f"{start_yr}Q4": f"{start_yr}1231",
        f"{end_yr}Q1"  : f"{end_yr}0331",
        f"{end_yr}Q2"  : f"{end_yr}0630",
        f"{end_yr}Q3"  : f"{end_yr}0930",
    }
    month_last_day = {
        m: f"{m[:4]}{m[4:]}{_cal.monthrange(int(m[:4]), int(m[4:]))[1]:02d}"
        for m in months
    }
    m2q = {}
    for m in months[0:3]:  m2q[m] = f"{start_yr}Q4"
    for m in months[3:6]:  m2q[m] = f"{end_yr}Q1"
    for m in months[6:9]:  m2q[m] = f"{end_yr}Q2"
    for m in months[9:12]: m2q[m] = f"{end_yr}Q3"

    q_label = {q: f"Qtr{i+1}" for i, q in enumerate(quarters)}
    m_label = {m: f"{_MN[int(m[4:])-1]}-{m[2:4]}" for m in months}

    return fy, quarters, months, quarter_last_day, month_last_day, m2q, q_label, m_label


(THIS_FY, THIS_FY_QUARTERS, THIS_FY_MONTHS,
 QUARTER_LAST_DAY, MONTH_LAST_DAY, MONTH_TO_QUARTER_RAW,
 QUARTER_LABEL, MONTH_LABEL) = _build_fy()

# Reverse maps for cache label → DHIS2 code lookups
LABEL_TO_QUARTER = {v: k for k, v in QUARTER_LABEL.items()}   # 'Qtr1' → '2025Q4'
LABEL_TO_MONTH   = {v: k for k, v in MONTH_LABEL.items()}     # 'Oct-25' → '202510'

SNAPSHOT_COLS = {"TX_CURR", "TX_PVLS_D", "TX_PVLS_N"}

# ── Indicators ────────────────────────────────────────────────────────────────
INDICATOR_MAP = {
    "ACEBAY_PMTCT_ART-Already on ART" : "PMTCT_ART.Already.T",
    "ACEBAY PMTCT_ART-New on ART"     : "PMTCT_ART.New.T",
    "ACEBAY PMTCT_Known Positive"     : "PMTCT_STAT.N.Known.Pos.T",
    "ACEBAY PMTCT_New Negative"       : "PMTCT_STAT.N.New.Neg.T",
    "ACEBAY PMTCT_New Positive"       : "PMTCT_STAT.N.New.Pos.T",
    "ACEBAY TB_ART Already on ART"    : "TB_ART.Already.T",
    "ACEBAY TB_ART New on ART"        : "TB_ART.New.T",
}
SKIP_INDICATORS = {
    "ACEBAY HTS_TST",
    "ACEBAY HTS_TST_POS",
    "ACEBAY TB_ART-Already on ART",
    "ACEBAY TB_ART-New on ART",
    "ACEBAY TB_PREV_D",
    "ACEBAY TB_PREV_N",
    "ACEBAY TX_TB_D",
    "ACEBAY TX_TB_N",
    "ACEBAY TX_TB New on ART",
}

# Flow data elements (SUM)
DATA_ELEMENT_MAP = {
    "Jt8x99v4EA0" : "TX_NEW",
    "y1C9SlSiZhQ" : "POST_RESP.PE.T",   # DR POST-RESP Physical/Emotional Violence_v22
    "uXsX7BTOuPa" : "POST_RESP.S.T",    # DR POST-RESP Sexual Violence_v22
}

# TX_CURR and TX_PVLS are now sourced from the Combined RADET (see section 5c/6c)
SNAPSHOT_DE_MAP  = {}   # was TX_CURR — now from RADET
SNAPSHOT_BAY_MAP = {}   # was TX_PVLS_D/N — now from RADET

# RADET blob config
RADET_CONTAINER    = os.getenv("RADET_CONTAINER",    "combine-radets-xls")
RADET_BLOB_PATTERN = os.getenv("RADET_BLOB_PATTERN", "{date}/ACE-1_Combined_RADET-{date}.xlsx")

OUTPUT_DIR    = Path(__file__).parent / "output"
FACILITY_FILE = Path(__file__).parent / "config" / "facility_list.json"
CACHE_DB      = Path(__file__).parent / "cache" / "dhis2_cache.db"

# ── Session ───────────────────────────────────────────────────────────────────
def _new_session():
    s = requests.Session()
    s.auth = HTTPBasicAuth(USERNAME, PASSWORD)
    s.headers.update({"Accept": "application/json"})
    return s


session = _new_session()


def get(path, params=None, _retry=2):
    global session
    url = f"{BASE_URL.rstrip('/')}/api/{path}"
    for attempt in range(_retry + 1):
        try:
            r = session.get(url, params=params, timeout=300)
            r.raise_for_status()
            return r.json()
        except Exception:
            if attempt < _retry:
                session = _new_session()
                time.sleep(3 * (attempt + 1))
            else:
                raise


def sep(char="=", n=65):
    log.info(char * n)


# ── Blob / target helpers ─────────────────────────────────────────────────────
_STATE_PREFIXES = {"ad", "bo", "ta", "yo"}


def _clean_facility_name(name):
    """Strip state prefix ('ad ', 'bo ', 'ta ', 'yo ') from target-file facility names."""
    if not isinstance(name, str):
        return str(name)
    parts = name.strip().split(" ", 1)
    if len(parts) == 2 and parts[0].lower() in _STATE_PREFIXES:
        return parts[1].strip()
    return name.strip()


def _load_targets():
    """
    Download target.xlsx from Azure Blob and return targets.

    Supports two formats:
    - Old: one FY row per facility → returns DataFrame indexed (STATE, FACILITY).
    - New running-target: one row per facility × weekday with DATE column →
      returns dict {"fy": fy_df, "quarterly": q_df, "monthly": m_df} where
      q_df/m_df contain per-period incremental targets.
    Returns None on failure.
    """
    try:
        try:
            from dotenv import load_dotenv as _ldenv
            _ldenv(Path(__file__).parent / ".env")
        except ImportError:
            pass
        if not _azure_configured():
            log.warning("Azure credentials not configured in .env — targets skipped")
            return None
        df = _azure_download_excel()
        df["FACILITY"] = df["FACILITY"].apply(_clean_facility_name)
        if "STATE" in df.columns:
            df["STATE"] = df["STATE"].str.strip().str.title()
        FACILITY_NAME_FIXES = {
            "Geidam General Hospital"        : "Geidem General Hospital",
            "FHI360 Clinic Banki IDP Camp"   : "FHI360 Clinic, Banki IDP Camp",
            "Rapha Hospital"                 : "Rapah Hospital",
        }
        if "FACILITY" in df.columns:
            df["FACILITY"] = df["FACILITY"].replace(FACILITY_NAME_FIXES)
        RENAME = {
            "TX_NEW.T"      : "TX_NEW",
            "TX_CURR.T"     : "TX_CURR",
            "TX_PVLS.D.T"   : "TX_PVLS_D",
            "TX_PVLS.N.T"   : "TX_PVLS_N",
            "HTS_TST.Pos.T" : "HTS_TST_POS",
            "HTS_TST.Neg.T" : "HTS_TST_Neg",
        }
        df = df.rename(columns=RENAME)
        if "HTS_TST_Neg" in df.columns and "HTS_TST_POS" in df.columns:
            neg = pd.to_numeric(df["HTS_TST_Neg"], errors="coerce").fillna(0)
            pos = pd.to_numeric(df["HTS_TST_POS"], errors="coerce").fillna(0)
            df["HTS_TST"] = neg + pos
        df = df.drop(columns=["PrEP_NEW.T"], errors="ignore")

        # ── NEW FORMAT: daily running targets ────────────────────────────────
        if "DATE" in df.columns:
            return _parse_running_targets(df)

        # ── OLD FORMAT: single FY row per facility ───────────────────────────
        if "Qtr" in df.columns:
            fy_mask = df["Qtr"].astype(str).str.contains("FY", case=False, na=False)
            df = df[fy_mask].copy()
        LAST_SUMMARY = 24
        DROP_META = {"LGA", "Qtr"}
        META_COLS = {"STATE", "FACILITY"}
        COMPUTED_COLS = {"HTS_TST"}  # computed from Neg + Pos, added past positional cutoff
        keep = [c for i, c in enumerate(df.columns)
                if c in META_COLS or c in COMPUTED_COLS or (i < LAST_SUMMARY and c not in DROP_META)]
        df = df[keep]
        idx_cols = [c for c in ["STATE", "FACILITY"] if c in df.columns]
        df = df.set_index(idx_cols)
        df = df.apply(pd.to_numeric, errors="coerce")
        log.info("Targets loaded: %d facilities, %d indicators", len(df), len(df.columns))
        return df
    except Exception as e:
        log.warning("Could not load targets from blob: %s", e)
        return None


def _parse_running_targets(df):
    """
    Parse the new DATE-indexed running target sheet.

    Each row = one facility × one weekday; values are CUMULATIVE from Oct 1.
    The last date in the sheet (e.g. Jun 30) equals the FY annual target.

    Returns dict:
        "fy"        – DataFrame indexed (STATE, FACILITY)            → FY annual targets
        "quarterly" – DataFrame indexed (STATE, FACILITY, Quarter)   → period increments
        "monthly"   – DataFrame indexed (STATE, FACILITY, Month)     → period increments
    """
    import calendar as _cal_mod
    df = df.copy()
    df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce")
    df = df.dropna(subset=["DATE"])

    META = {"STATE", "LGA", "FACILITY", "Qtr", "DATE"}
    for c in df.columns:
        if c not in META:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    ind_cols = [c for c in df.columns
                if c not in META and pd.api.types.is_numeric_dtype(df[c])]

    idx_cols = [c for c in ["STATE", "FACILITY"] if c in df.columns]
    sheet_max = df["DATE"].max()

    start_yr, end_yr = THIS_FY - 1, THIS_FY
    q_ends = {
        "Qtr1": pd.Timestamp(f"{start_yr}-12-31"),
        "Qtr2": pd.Timestamp(f"{end_yr}-03-31"),
        "Qtr3": pd.Timestamp(f"{end_yr}-06-30"),
        "Qtr4": pd.Timestamp(f"{end_yr}-09-30"),
    }
    m_ends = {
        m_label: pd.Timestamp(int(m_code[:4]), int(m_code[4:]),
                               _cal_mod.monthrange(int(m_code[:4]), int(m_code[4:]))[1])
        for m_code, m_label in MONTH_LABEL.items()
    }

    fy_rows, q_rows, m_rows = [], [], []

    for keys, grp in df.groupby(idx_cols, sort=False):
        state, facility = keys if len(idx_cols) == 2 else ("", keys)
        grp = grp.sort_values("DATE")

        def cum_at(target_date):
            cap  = min(target_date, sheet_max)
            mask = grp["DATE"] <= cap
            if not mask.any():
                return pd.Series(0.0, index=ind_cols)
            return grp.loc[mask, ind_cols].iloc[-1].fillna(0)

        fy_val = cum_at(sheet_max)
        fy_rows.append({"STATE": state, "FACILITY": facility, **fy_val.to_dict()})

        # Quarterly: raw cumulative value at each quarter-end (YTD target)
        for qtr in ["Qtr1", "Qtr2", "Qtr3", "Qtr4"]:
            if q_ends[qtr] > sheet_max:
                q_rows.append({"STATE": state, "FACILITY": facility, "Quarter": qtr,
                               **{c: float("nan") for c in ind_cols}})
            else:
                cum = cum_at(q_ends[qtr])
                q_rows.append({"STATE": state, "FACILITY": facility, "Quarter": qtr,
                               **cum.to_dict()})

        # Monthly: raw cumulative value at each month-end (YTD target)
        for m_label, mend in m_ends.items():
            if mend > sheet_max:
                m_rows.append({"STATE": state, "FACILITY": facility, "Month": m_label,
                               **{c: float("nan") for c in ind_cols}})
            else:
                cum = cum_at(mend)
                m_rows.append({"STATE": state, "FACILITY": facility, "Month": m_label,
                               **cum.to_dict()})

    fy_df = pd.DataFrame(fy_rows).set_index(idx_cols).apply(pd.to_numeric, errors="coerce")
    q_df  = (pd.DataFrame(q_rows)
               .set_index(idx_cols + ["Quarter"])
               .apply(pd.to_numeric, errors="coerce"))
    m_df  = (pd.DataFrame(m_rows)
               .set_index(idx_cols + ["Month"])
               .apply(pd.to_numeric, errors="coerce"))

    log.info("Running targets loaded: %d facilities, %d indicators, %d dates",
             len(fy_df), len(ind_cols), df["DATE"].nunique())
    return {"fy": fy_df, "quarterly": q_df, "monthly": m_df}


def _apply_targets_quarterly(df, targets):
    """
    Interleave {col}_Target and {col}_Ach% after every numeric indicator column.

    Accepts two target formats:
    - New (dict from _parse_running_targets): cumulative YTD target at each quarter-end
      for ALL indicators (flow and snapshot). Q4/future quarters with no sheet data = NaN.
    - Old (DataFrame indexed (STATE, FACILITY)): FY-interpolation logic.
      Snapshot cols: FY_Target every quarter. Flow cols: FY − prior quarters.

    FY row always receives FY_Target regardless of format.
    """
    if targets is None or df.empty:
        return df

    if isinstance(targets, dict):
        fy_tgt = targets.get("fy")
        q_tgt  = targets.get("quarterly")
    else:
        fy_tgt = targets
        q_tgt  = None

    if fy_tgt is None or fy_tgt.empty:
        return df

    df = df.copy()
    fy_label = f"FY{THIS_FY % 100}"
    q_order  = [QUARTER_LABEL[q] for q in THIS_FY_QUARTERS]

    skip = {"STATE", "FACILITY", "Quarter", "VLC", "VLS", "Linkage Rate"}
    indicator_cols = [c for c in df.select_dtypes(include="number").columns if c not in skip]

    for col in indicator_cols:
        if col not in fy_tgt.columns:
            continue
        t_col   = f"{col}_Target"
        ach_col = f"{col}_Ach%"
        df[t_col]   = pd.Series(dtype=float)
        df[ach_col] = pd.Series(dtype=float)

        fy_col_dict = fy_tgt[col].dropna().to_dict()
        q_col_dict  = (q_tgt[col].dropna().to_dict()
                       if q_tgt is not None and col in q_tgt.columns else {})

        for (state, facility), grp in df.groupby(["STATE", "FACILITY"]):
            key = (state, facility) if fy_tgt.index.nlevels == 2 else facility
            raw = fy_col_dict.get(key)
            if raw is None:
                continue
            fy_target = float(raw)

            if q_col_dict and col not in SNAPSHOT_COLS:
                # New format, flow indicator: cumulative YTD target at each quarter-end
                for qtr in q_order:
                    idx = grp.index[grp["Quarter"] == qtr]
                    if not len(idx):
                        continue
                    q_key = (state, facility, qtr) if fy_tgt.index.nlevels == 2 else (facility, qtr)
                    q_val = q_col_dict.get(q_key)
                    if q_val is not None and not pd.isna(q_val):
                        df.at[idx[0], t_col] = max(0.0, float(q_val))
            else:
                # Snapshot cols (TX_CURR, TX_PVLS_D/N) or old format:
                # FY_Target for all periods — stock count vs annual target
                for qtr in q_order:
                    idx = grp.index[grp["Quarter"] == qtr]
                    if len(idx):
                        df.at[idx[0], t_col] = fy_target

            fy_idx = grp.index[grp["Quarter"] == fy_label]
            if len(fy_idx):
                df.at[fy_idx[0], t_col] = fy_target

        t_s = pd.to_numeric(df[t_col], errors="coerce")
        df[ach_col] = pd.to_numeric(df[col], errors="coerce") / t_s.replace(0, float("nan"))

    # Reorder: meta | (col, col_Target, col_Ach%)... | VLC, VLS, Linkage Rate
    meta_cols = ["STATE", "FACILITY", "Quarter"]
    ordered   = []
    for col in indicator_cols:
        ordered.append(col)
        if f"{col}_Target" in df.columns: ordered.append(f"{col}_Target")
        if f"{col}_Ach%"   in df.columns: ordered.append(f"{col}_Ach%")
    derived   = [c for c in ["VLC", "VLS", "Linkage Rate"] if c in df.columns]
    rest      = [c for c in df.columns if c not in set(meta_cols + ordered + derived)]
    return df[[c for c in meta_cols + ordered + derived + rest if c in df.columns]]


def _apply_targets_monthly(df, targets):
    """
    Interleave {col}_Target and {col}_Ach% after every numeric indicator column
    in the Monthly sheet.

    Accepts two target formats:
    - New (dict from _parse_running_targets): cumulative YTD target at each month-end
      for ALL indicators. Months beyond the sheet (Jul–Sep) receive NaN targets.
    - Old (DataFrame): FY_Target for every month row — YTD actuals vs annual target.
    """
    if targets is None or df.empty:
        return df

    if isinstance(targets, dict):
        fy_tgt = targets.get("fy")
        m_tgt  = targets.get("monthly")
    else:
        fy_tgt = targets
        m_tgt  = None

    if fy_tgt is None or fy_tgt.empty:
        return df

    df = df.copy()
    m_order = list(MONTH_LABEL.values())
    _, open_month = get_open_periods()
    open_m_idx = m_order.index(open_month) if open_month in m_order else len(m_order)

    skip = {"STATE", "FACILITY", "Quarter", "Month", "VLC", "VLS", "Linkage Rate"}
    indicator_cols = [c for c in df.select_dtypes(include="number").columns if c not in skip]

    for col in indicator_cols:
        if col not in fy_tgt.columns:
            continue
        t_col   = f"{col}_Target"
        ach_col = f"{col}_Ach%"
        df[t_col]   = pd.Series(dtype=float)
        df[ach_col] = pd.Series(dtype=float)

        fy_col_dict = fy_tgt[col].dropna().to_dict()
        m_col_dict  = (m_tgt[col].dropna().to_dict()
                       if m_tgt is not None and col in m_tgt.columns else {})

        for (state, facility), grp in df.groupby(["STATE", "FACILITY"]):
            key = (state, facility) if fy_tgt.index.nlevels == 2 else facility
            raw = fy_col_dict.get(key)
            if raw is None:
                continue
            fy_target = float(raw)

            if m_col_dict and col not in SNAPSHOT_COLS:
                # New format, flow indicator: cumulative YTD target at each month-end
                for m in m_order:
                    idx = grp.index[grp["Month"] == m]
                    if not len(idx):
                        continue
                    m_key = (state, facility, m) if fy_tgt.index.nlevels == 2 else (facility, m)
                    m_val = m_col_dict.get(m_key)
                    if m_val is not None and not pd.isna(m_val):
                        df.at[idx[0], t_col] = max(0.0, float(m_val))
            else:
                # Snapshot cols (TX_CURR, TX_PVLS_D/N) or old format:
                # FY_Target for all periods — stock count vs annual target
                for m in m_order:
                    idx = grp.index[grp["Month"] == m]
                    if len(idx):
                        df.at[idx[0], t_col] = fy_target

        t_s = pd.to_numeric(df[t_col], errors="coerce")
        df[ach_col] = pd.to_numeric(df[col], errors="coerce") / t_s.replace(0, float("nan"))

    # Reorder: meta | (col, col_Target, col_Ach%)... | VLC, VLS, Linkage Rate
    meta_cols = ["STATE", "FACILITY", "Quarter", "Month"]
    ordered   = []
    for col in indicator_cols:
        ordered.append(col)
        if f"{col}_Target" in df.columns: ordered.append(f"{col}_Target")
        if f"{col}_Ach%"   in df.columns: ordered.append(f"{col}_Ach%")
    derived   = [c for c in ["VLC", "VLS", "Linkage Rate"] if c in df.columns]
    rest      = [c for c in df.columns if c not in set(meta_cols + ordered + derived)]
    return df[[c for c in meta_cols + ordered + derived + rest if c in df.columns]]


# ── CLI args ──────────────────────────────────────────────────────────────────
_parser = argparse.ArgumentParser()
_parser.add_argument(
    "--mode",
    choices=["full", "incremental"],
    default="full",
    help="full = pull all periods; incremental = API only for current open period, rest from cache",
)
_parser.add_argument(
    "--skip-upload",
    action="store_true",
    default=False,
    help="skip the Azure Blob Storage upload step",
)
ARGS = _parser.parse_args()


# ── Cache helpers ─────────────────────────────────────────────────────────────
def init_cache():
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
    conn.commit()
    return conn


def get_open_periods():
    """Return (open_quarter_label, open_month_label) for today using the current FY config."""
    today  = date.today()
    m_code = f"{today.year}{today.month:02d}"
    q_raw  = MONTH_TO_QUARTER_RAW.get(m_code)
    if q_raw is None:
        # Today is outside the current FY — shouldn't normally happen
        return list(QUARTER_LABEL.values())[-1], list(MONTH_LABEL.values())[-1]
    return QUARTER_LABEL[q_raw], MONTH_LABEL[m_code]


def cache_get_recs(conn, period_labels, view):
    """Load records from cache for given period labels; returns list of record dicts."""
    if not period_labels or conn is None:
        return []
    placeholders = ",".join("?" * len(period_labels))
    rows = conn.execute(
        f"SELECT ou_uid, period, col_name, value FROM cache "
        f"WHERE period IN ({placeholders}) AND view = ?",
        (*period_labels, view),
    ).fetchall()
    recs = []
    for ou_uid, period, col_name, value in rows:
        if view == "monthly":
            month_code  = LABEL_TO_MONTH.get(period, "")
            quarter_raw = MONTH_TO_QUARTER_RAW.get(month_code, "")
            recs.append(make_rec(ou_uid, quarter_raw, month_code, col_name, value, "monthly"))
        else:
            quarter_raw = LABEL_TO_QUARTER.get(period, "")
            recs.append(make_rec(ou_uid, quarter_raw, None, col_name, value, "quarterly"))
    return recs


def cache_put_recs(conn, recs, view):
    """Upsert record dicts into cache."""
    if conn is None or not recs:
        return
    now = _dt.now().isoformat()
    for rec in recs:
        period = rec.get("Month") if view == "monthly" and rec.get("Month") else rec.get("Quarter")
        if not period or not rec.get("ou_uid") or not rec.get("column"):
            continue
        conn.execute(
            "INSERT OR REPLACE INTO cache (ou_uid, period, col_name, value, view, pulled_at) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            (rec["ou_uid"], period, rec["column"], rec.get("value"), view, now),
        )
    conn.commit()


def safe_div(n, d, pct=False):
    if isinstance(n, pd.Series):
        result = n / d.replace(0, float("nan"))
        return (result * 100).round(2) if pct else result.round(4)
    if not d or d == 0:
        return None
    result = n / d
    return round(result * 100, 2) if pct else round(result, 4)


def add_derived(df):
    hts_neg = (
        pd.to_numeric(df.get("HTS_TST",     pd.Series(dtype=float)), errors="coerce") -
        pd.to_numeric(df.get("HTS_TST_POS", pd.Series(dtype=float)), errors="coerce")
    )
    if "HTS_TST_Neg" not in df.columns:
        if "HTS_TST_POS" in df.columns:
            df.insert(df.columns.get_loc("HTS_TST_POS") + 1, "HTS_TST_Neg", hts_neg)
        else:
            df["HTS_TST_Neg"] = hts_neg
    else:
        df["HTS_TST_Neg"] = hts_neg
    df["Linkage Rate"] = safe_div(
        df.get("TX_NEW",      pd.Series(dtype=float)),
        df.get("HTS_TST_POS", pd.Series(dtype=float)).replace(0, float("nan")),
    )
    df["VLC"] = safe_div(
        df.get("TX_PVLS_D", pd.Series(dtype=float)),
        df.get("TX_CURR",   pd.Series(dtype=float)).replace(0, float("nan")),
    )
    df["VLS"] = safe_div(
        df.get("TX_PVLS_N", pd.Series(dtype=float)),
        df.get("TX_PVLS_D", pd.Series(dtype=float)).replace(0, float("nan")),
    )
    return df


def pull_analytics(dx_uid_list, periods, ou_uid_list, col_map, ou_chunk=20):
    """
    Pull analytics chunked by OU (20 facilities per request) to avoid
    DHIS2 server computation errors (HTTP 500) with large OU sets.
    Returns (results, failed_ou_uids) where results is a list of
    (period, ou_uid, col_name, value) and failed_ou_uids is the set of
    OU UIDs whose chunk raised an exception.
    """
    if not dx_uid_list or not ou_uid_list:
        return [], set()
    dx_param    = ";".join(dx_uid_list)
    results     = []
    failed_uids = set()
    for i in range(0, len(ou_uid_list), ou_chunk):
        ou_chunk_list = ou_uid_list[i:i + ou_chunk]
        ou_param = ";".join(ou_chunk_list)
        try:
            data = get("analytics", {
                "dimension"       : [
                    f"dx:{dx_param}",
                    f"pe:{';'.join(periods)}",
                    f"ou:{ou_param}",
                ],
                "ouMode"          : "SELECTED",
                "skipMeta"        : "false",
                "displayProperty" : "NAME",
            })
            for row in data.get("rows", []):
                dx_uid, period, ou_uid, value = row
                col_name = col_map.get(dx_uid, dx_uid)
                val = float(value) if value not in (None, "") else None
                results.append((period, ou_uid, col_name, val))
        except Exception as e:
            log.error("OU chunk %d analytics pull failed: %s", i // ou_chunk + 1, e)
            failed_uids.update(ou_chunk_list)
    return results, failed_uids


def pull_bay_sum(uid_list, periods, ou_uid_list, col_name, ou_chunk=20):
    """
    Sum all BAY disaggregated elements per (period, ou_uid), chunked by OU.
    Returns ({(period, ou_uid): total}, failed_ou_uids).
    """
    if not uid_list or not ou_uid_list:
        return {}, set()
    dx_param    = ";".join(uid_list)
    totals      = {}
    failed_uids = set()
    for i in range(0, len(ou_uid_list), ou_chunk):
        ou_chunk_list = ou_uid_list[i:i + ou_chunk]
        ou_param = ";".join(ou_chunk_list)
        try:
            data = get("analytics", {
                "dimension"       : [
                    f"dx:{dx_param}",
                    f"pe:{';'.join(periods)}",
                    f"ou:{ou_param}",
                ],
                "ouMode"          : "SELECTED",
                "skipMeta"        : "false",
                "displayProperty" : "NAME",
                "aggregationType" : "SUM",
            })
            for row in data.get("rows", []):
                _, period, ou_uid, value = row
                val = float(value) if value not in (None, "") else 0.0
                key = (period, ou_uid)
                totals[key] = totals.get(key, 0.0) + val
        except Exception as e:
            log.error("OU chunk %d pulling %s: %s", i // ou_chunk + 1, col_name, e)
            failed_uids.update(ou_chunk_list)
    return totals, failed_uids


# ─────────────────────────────────────────────────────────────────────────────
# 1. Connection
# ─────────────────────────────────────────────────────────────────────────────
sep()
log.info("1. Connecting to DHIS2 ...")
try:
    me = get("me", {"fields": "id,name,username"})
    log.info("   Logged in as: %s (%s)", me.get("name"), me.get("username"))
except Exception as e:
    log.error("   FAILED: %s", e)
    sys.exit(1)


# ─────────────────────────────────────────────────────────────────────────────
# 2. Load facility list
# ─────────────────────────────────────────────────────────────────────────────
log.info("2. Loading facility list from %s ...", FACILITY_FILE.name)
with open(FACILITY_FILE) as f:
    FACILITIES = json.load(f)

ou_uid_to_facility = {fc["id"]: fc for fc in FACILITIES}
OU_UID_LIST = [fc["id"] for fc in FACILITIES]
OU_PARAM    = ";".join(OU_UID_LIST)   # kept for reference

state_counts = {}
for fc in FACILITIES:
    state_counts[fc["state"]] = state_counts.get(fc["state"], 0) + 1
log.info("   %d facilities loaded:", len(FACILITIES))
for state, count in sorted(state_counts.items()):
    log.info("     %s: %d", state, count)


# ─────────────────────────────────────────────────────────────────────────────
# 2b. Init cache + determine what to pull vs. load from cache
# ─────────────────────────────────────────────────────────────────────────────
conn_cache = init_cache() if CACHE_DB.exists() or ARGS.mode == "incremental" else init_cache()

open_q_label, open_m_label = get_open_periods()
open_q_raw   = LABEL_TO_QUARTER.get(open_q_label, "2026Q2")
open_m_raw   = LABEL_TO_MONTH.get(open_m_label,   "202606")

if ARGS.mode == "incremental":
    # Only hit the API for the current open period
    PULL_QUARTERS = [open_q_raw]
    PULL_MONTHS   = [open_m_raw]
    CACHE_Q_LABELS = [QUARTER_LABEL[q] for q in THIS_FY_QUARTERS if q != open_q_raw]
    CACHE_M_LABELS = [MONTH_LABEL[m]   for m in THIS_FY_MONTHS   if m != open_m_raw]
    log.info("   Mode: INCREMENTAL — API for %s / %s; cache for %d quarters, %d months",
             open_q_label, open_m_label, len(CACHE_Q_LABELS), len(CACHE_M_LABELS))
else:
    PULL_QUARTERS  = THIS_FY_QUARTERS
    PULL_MONTHS    = THIS_FY_MONTHS
    CACHE_Q_LABELS = []
    CACHE_M_LABELS = []
    log.info("   Mode: FULL — pulling all periods from API")

_today = date.today().strftime("%Y%m%d")
# For open (unfinished) periods use today's date so we get actual data instead of blanks
PULL_LAST_DAYS_Q = {q: min(QUARTER_LAST_DAY[q], _today) for q in PULL_QUARTERS}
PULL_LAST_DAYS_M = {m: min(MONTH_LAST_DAY[m],   _today) for m in PULL_MONTHS}


# ─────────────────────────────────────────────────────────────────────────────
# 3. Discover ACEBAY indicators
# ─────────────────────────────────────────────────────────────────────────────
log.info("3. Fetching ACEBAY indicators ...")
try:
    result = get("indicators", {
        "filter" : "name:ilike:ACEBAY",
        "fields" : "id,name",
        "paging" : "false",
    })
    all_indicators = {item["name"]: item["id"] for item in result.get("indicators", [])}
    log.info("   Found %d ACEBAY indicator(s)", len(all_indicators))
except Exception as e:
    log.error("   FAILED: %s", e)
    sys.exit(1)

# Only pull indicators that map to a known output column — avoids server errors
# from unmapped indicators (CXCA, GEND_GBV, HTS_INDEX, etc.) that lack data
active_indicators = {
    n: uid for n, uid in all_indicators.items()
    if n in INDICATOR_MAP
}
acebay_col_map = {uid: INDICATOR_MAP[name] for name, uid in active_indicators.items()}
log.info("   Pulling %d mapped indicator(s) (skipping %d unmapped)",
         len(active_indicators), len(all_indicators) - len(active_indicators))


# ─────────────────────────────────────────────────────────────────────────────
# 4. TX_PVLS → now from RADET (SNAPSHOT_BAY_MAP is empty)
# ─────────────────────────────────────────────────────────────────────────────
log.info("4. TX_PVLS/TX_CURR now sourced from Combined RADET — skipping DHIS2 discovery.")
pvls_uids = {}   # kept for compatibility; will be empty

# 4d. Pre-load DatimId map if config file already exists (avoids rebuild on every run)
log.info("4d. Loading DatimId → OU_UID map ...")
_datimid_map = _load_datimid_map()
if _datimid_map:
    log.info("   Map ready (%d entries). Will rebuild if RADET is downloaded fresh.", len(_datimid_map))
else:
    log.info("   No map file yet — will build on first RADET download.")

log.info("4b. Discovering TX_TB data element UIDs ...")
# TX_TB.N.Already.T → DR ...Clinic_Encounter  |  TX_TB.N.New.T → DR ...Care_New
TXTB_SEARCH_MAP = {
    "TX_TB.N.Already.T": "DR Number of PLHIV on ART with active TB disease who initiated TB treatment (Clinic_Encounter)",
    "TX_TB.N.New.T"    : "DR Number of PLHIV on ART with active TB disease who initiated TB treatment (Care_New)",
}
txtb_uids = {}
for col_name, search_name in TXTB_SEARCH_MAP.items():
    res   = get("dataElements", {
        "filter" : f"name:ilike:{search_name}",
        "fields" : "id,name",
        "paging" : "false",
    })
    items = res.get("dataElements", [])
    txtb_uids[col_name] = [i["id"] for i in items]
    log.info("   %s: %d element(s)  %s", col_name, len(items),
             [i["name"] for i in items] if items else "NOT FOUND")

log.info("4c. Discovering TB_PREV data element UIDs ...")
# TB_PREV.N.Already.T → DR ...Clinic_Encounter (completed IPT)
# TB_PREV.N.New.T     → DR ...Care_New         (completed INH 6 months)
TBPREV_SEARCH_MAP = {
    "TB_PREV.N.Already.T": "DR - Number of ART patients who started a course of TB preventive therapy and completed - (Clinic_Encounter)",
    "TB_PREV.N.New.T"    : "DR - Number of ART patients who started a course of TB preventive therapy (INH) 6months prior and completed - (Care_New)",
}
tbprev_uids = {}
for col_name, search_name in TBPREV_SEARCH_MAP.items():
    res   = get("dataElements", {
        "filter" : f"name:ilike:{search_name}",
        "fields" : "id,name",
        "paging" : "false",
    })
    items = res.get("dataElements", [])
    tbprev_uids[col_name] = [i["id"] for i in items]
    log.info("   %s: %d element(s)  %s", col_name, len(items),
             [i["name"] for i in items] if items else "NOT FOUND")


# ─────────────────────────────────────────────────────────────────────────────
# Helper: build a record dict
# ─────────────────────────────────────────────────────────────────────────────
def make_rec(ou_uid, quarter_raw, month_period, col_name, value, view):
    fc = ou_uid_to_facility.get(ou_uid, {"state": "Unknown", "name": ou_uid})
    q_raw = MONTH_TO_QUARTER_RAW.get(month_period, quarter_raw) if month_period else quarter_raw
    return {
        "STATE"    : fc["state"],
        "FACILITY" : fc["name"],
        "ou_uid"   : ou_uid,
        "Quarter"  : QUARTER_LABEL.get(q_raw, q_raw),
        "Month"    : MONTH_LABEL.get(month_period, "") if month_period else "",
        "column"   : col_name,
        "value"    : value,
        "view"     : view,
    }


# ─────────────────────────────────────────────────────────────────────────────
# 5. Pull QUARTERLY data (all 74 facilities at once)
# ─────────────────────────────────────────────────────────────────────────────
log.info("5. Pulling QUARTERLY data — %d facilities ...", len(FACILITIES))
q_records = []

# Load past quarters from cache (incremental mode only)
if CACHE_Q_LABELS:
    cached = cache_get_recs(conn_cache, CACHE_Q_LABELS, "quarterly")
    q_records.extend(cached)
    log.info("   Cache: %d rows for %s", len(cached), CACHE_Q_LABELS)

# helper: report facilities with missing data
def _report_missing(uid_set, indicator):
    names = ", ".join(ou_uid_to_facility.get(u, {}).get("name", u) for u in sorted(uid_set))
    log.warning("   MISSING %s — no data for %d facility(ies): %s", indicator, len(uid_set), names)

# 5a. ACEBAY indicators (flow) — one facility at a time to avoid server timeout/500
log.info("   5a. ACEBAY indicators (flow, 1 facility at a time) ...")
_ace_dx = ";".join(active_indicators.values())

total_rows = 0
new_q_ace = []
_failed_q_ace = set()
for idx, fc in enumerate(FACILITIES):
    if idx % 10 == 0:
        log.info("      [%d/%d] %s ...", idx + 1, len(FACILITIES), fc["name"])
    try:
        data = get("analytics", {
            "dimension"       : [f"dx:{_ace_dx}",
                                 f"pe:{';'.join(PULL_QUARTERS)}",
                                 f"ou:{fc['id']}"],
            "ouMode"          : "SELECTED",
            "skipMeta"        : "false",
            "displayProperty" : "NAME",
        })
        for row in data.get("rows", []):
            dx_uid, period, ou_uid, value = row
            col_name = acebay_col_map.get(dx_uid, dx_uid)
            val = float(value) if value not in (None, "") else None
            rec = make_rec(ou_uid, period, None, col_name, val, "quarterly")
            new_q_ace.append(rec)
            total_rows += 1
    except Exception as e:
        log.error("      ERROR [%s]: %s", fc["name"], e)
        _failed_q_ace.add(fc["id"])
if _failed_q_ace:
    _report_missing(_failed_q_ace, "ACEBAY indicators (quarterly)")
q_records.extend(new_q_ace)
log.info("      Total: %d row(s)", total_rows)

# 5b. TX_NEW + POST_RESP (flow)
log.info("   5b. TX_NEW, POST_RESP (flow) ...")
rows, _failed_5b = pull_analytics(list(DATA_ELEMENT_MAP.keys()), PULL_QUARTERS, OU_UID_LIST, DATA_ELEMENT_MAP)
log.info("      %d row(s)", len(rows))
new_q_flow = []
for period, ou_uid, col_name, value in rows:
    rec = make_rec(ou_uid, period, None, col_name, value, "quarterly")
    new_q_flow.append(rec)
if _failed_5b:
    _report_missing(_failed_5b, "TX_NEW/POST_RESP (quarterly)")
q_records.extend(new_q_flow)

# 5c. TX_CURR (snapshot — last day of each quarter)
log.info("   5c. RADET pull: TX_CURR, TX_PVLS_D/N, HTS_TST/POS (quarterly) ...")
new_q_radet = []
for q_raw in PULL_QUARTERS:
    snap_date   = PULL_LAST_DAYS_Q[q_raw]          # 'YYYYMMDD'
    snap_iso    = f"{snap_date[:4]}-{snap_date[4:6]}-{snap_date[6:]}"
    q_start_iso = _quarter_start(q_raw)
    q_label     = QUARTER_LABEL[q_raw]
    if q_start_iso > snap_iso:
        log.info("      %s: quarter hasn't started yet (%s > today) — skipping", q_label, q_start_iso)
        continue
    log.info("      %s: snapshot=%s  HTS range=%s → %s", q_label, snap_iso, q_start_iso, snap_iso)
    radet_data = pull_radet_for_date(snap_iso, q_start_iso, snap_iso)
    for col_name, by_ou in radet_data.items():
        for ou_uid, value in by_ou.items():
            rec = make_rec(ou_uid, q_raw, None, col_name, value, "quarterly")
            new_q_radet.append(rec)
    if not radet_data:
        log.warning("      %s: no RADET data — these columns will be blank", q_label)
q_records.extend(new_q_radet)
new_q_snap = new_q_radet   # kept for new_q_all reference below
new_q_pvls = []             # was separate; now folded into new_q_radet

# 5e. TX_TB.N.Already/New (flow — sum over quarter period)
log.info("   5e. TX_TB.N.Already/New (flow — sum over quarter) ...")
new_q_txtb = []
for col_name, uid_list in txtb_uids.items():
    if not uid_list:
        log.warning("      %s: no elements found — skipping", col_name)
        continue
    totals, _failed_5e = pull_bay_sum(uid_list, PULL_QUARTERS, OU_UID_LIST, col_name)
    log.info("      %s: %d (period, facility) combinations", col_name, len(totals))
    for (period, ou_uid), total in totals.items():
        rec = make_rec(ou_uid, period, None, col_name, total, "quarterly")
        new_q_txtb.append(rec)
    if _failed_5e:
        _report_missing(_failed_5e, f"{col_name} (quarterly)")
q_records.extend(new_q_txtb)

# 5f. TB_PREV.N.Already/New (flow — sum over quarter period)
log.info("   5f. TB_PREV.N.Already/New (flow — sum over quarter) ...")
new_q_tbprev = []
for col_name, uid_list in tbprev_uids.items():
    if not uid_list:
        log.warning("      %s: no elements found — skipping", col_name)
        continue
    totals, _failed_5f = pull_bay_sum(uid_list, PULL_QUARTERS, OU_UID_LIST, col_name)
    log.info("      %s: %d (period, facility) combinations", col_name, len(totals))
    for (period, ou_uid), total in totals.items():
        rec = make_rec(ou_uid, period, None, col_name, total, "quarterly")
        new_q_tbprev.append(rec)
    if _failed_5f:
        _report_missing(_failed_5f, f"{col_name} (quarterly)")
q_records.extend(new_q_tbprev)

# Save new quarterly records to cache
new_q_all = new_q_ace + new_q_flow + new_q_radet + new_q_txtb + new_q_tbprev
if new_q_all:
    cache_put_recs(conn_cache, new_q_all, "quarterly")
if not new_q_all:
    log.warning("API returned 0 quarterly rows for %s — output will be empty for this period", open_q_label)

log.info("   Total quarterly records: %d", len(q_records))


# ─────────────────────────────────────────────────────────────────────────────
# 6. Pull MONTHLY data (all 74 facilities at once)
# ─────────────────────────────────────────────────────────────────────────────
log.info("6. Pulling MONTHLY data — %d facilities ...", len(FACILITIES))
m_records = []

# Load past months from cache (incremental mode only)
if CACHE_M_LABELS:
    cached = cache_get_recs(conn_cache, CACHE_M_LABELS, "monthly")
    m_records.extend(cached)
    log.info("   Cache: %d rows for %d past months", len(cached), len(CACHE_M_LABELS))

# 6a. ACEBAY indicators (flow) — one facility at a time to avoid server timeout/500
log.info("   6a. ACEBAY indicators (flow, 1 facility at a time) ...")

total_rows = 0
new_m_ace = []
_failed_m_ace = set()
for idx, fc in enumerate(FACILITIES):
    if idx % 10 == 0:
        log.info("      [%d/%d] %s ...", idx + 1, len(FACILITIES), fc["name"])
    try:
        data = get("analytics", {
            "dimension"       : [f"dx:{_ace_dx}",
                                 f"pe:{';'.join(PULL_MONTHS)}",
                                 f"ou:{fc['id']}"],
            "ouMode"          : "SELECTED",
            "skipMeta"        : "false",
            "displayProperty" : "NAME",
        })
        for row in data.get("rows", []):
            dx_uid, period, ou_uid, value = row
            col_name = acebay_col_map.get(dx_uid, dx_uid)
            val = float(value) if value not in (None, "") else None
            rec = make_rec(ou_uid, None, period, col_name, val, "monthly")
            new_m_ace.append(rec)
            total_rows += 1
    except Exception as e:
        log.error("      ERROR [%s]: %s", fc["name"], e)
        _failed_m_ace.add(fc["id"])
if _failed_m_ace:
    _report_missing(_failed_m_ace, "ACEBAY indicators (monthly)")
m_records.extend(new_m_ace)
log.info("      Total: %d row(s)", total_rows)

# 6b. TX_NEW + POST_RESP (flow)
log.info("   6b. TX_NEW, POST_RESP (flow) ...")
rows, _failed_6b = pull_analytics(list(DATA_ELEMENT_MAP.keys()), PULL_MONTHS, OU_UID_LIST, DATA_ELEMENT_MAP)
log.info("      %d row(s)", len(rows))
new_m_flow = []
for period, ou_uid, col_name, value in rows:
    rec = make_rec(ou_uid, None, period, col_name, value, "monthly")
    new_m_flow.append(rec)
if _failed_6b:
    _report_missing(_failed_6b, "TX_NEW/POST_RESP (monthly)")
m_records.extend(new_m_flow)

# 6c. RADET pull: TX_CURR, TX_PVLS_D/N, HTS_TST/POS (monthly)
log.info("   6c. RADET pull: TX_CURR, TX_PVLS_D/N, HTS_TST/POS (monthly) ...")
new_m_radet = []
for m_raw in PULL_MONTHS:
    snap_date   = PULL_LAST_DAYS_M[m_raw]          # 'YYYYMMDD'
    snap_iso    = f"{snap_date[:4]}-{snap_date[4:6]}-{snap_date[6:]}"
    m_start_iso = _month_start(m_raw)
    m_label     = MONTH_LABEL[m_raw]
    if m_start_iso > snap_iso:
        log.info("      %s: month hasn't started yet (%s > today) — skipping", m_label, m_start_iso)
        continue
    log.info("      %s: snapshot=%s  HTS range=%s → %s", m_label, snap_iso, m_start_iso, snap_iso)
    radet_data = pull_radet_for_date(snap_iso, m_start_iso, snap_iso)
    for col_name, by_ou in radet_data.items():
        for ou_uid, value in by_ou.items():
            rec = make_rec(ou_uid, None, m_raw, col_name, value, "monthly")
            new_m_radet.append(rec)
    if not radet_data:
        log.warning("      %s: no RADET data — these columns will be blank", m_label)
m_records.extend(new_m_radet)
new_m_snap = new_m_radet   # kept for new_m_all reference below
new_m_pvls = []             # was separate; now folded into new_m_radet

# 6e. TX_TB.N.Already/New (flow — sum over month period)
log.info("   6e. TX_TB.N.Already/New (flow — sum over month) ...")
new_m_txtb = []
for col_name, uid_list in txtb_uids.items():
    if not uid_list:
        continue
    totals, _failed_6e = pull_bay_sum(uid_list, PULL_MONTHS, OU_UID_LIST, col_name)
    log.info("      %s: %d (period, facility) combinations", col_name, len(totals))
    for (period, ou_uid), total in totals.items():
        rec = make_rec(ou_uid, None, period, col_name, total, "monthly")
        new_m_txtb.append(rec)
    if _failed_6e:
        _report_missing(_failed_6e, f"{col_name} (monthly)")
m_records.extend(new_m_txtb)

# 6f. TB_PREV.N.Already/New (flow — sum over month period)
log.info("   6f. TB_PREV.N.Already/New (flow — sum over month) ...")
new_m_tbprev = []
for col_name, uid_list in tbprev_uids.items():
    if not uid_list:
        continue
    totals, _failed_6f = pull_bay_sum(uid_list, PULL_MONTHS, OU_UID_LIST, col_name)
    log.info("      %s: %d (period, facility) combinations", col_name, len(totals))
    for (period, ou_uid), total in totals.items():
        rec = make_rec(ou_uid, None, period, col_name, total, "monthly")
        new_m_tbprev.append(rec)
    if _failed_6f:
        _report_missing(_failed_6f, f"{col_name} (monthly)")
m_records.extend(new_m_tbprev)

# Save new monthly records to cache
new_m_all = new_m_ace + new_m_flow + new_m_radet + new_m_txtb + new_m_tbprev
if new_m_all:
    cache_put_recs(conn_cache, new_m_all, "monthly")
if not new_m_all:
    log.warning("API returned 0 monthly rows for %s — output will be empty for this period", open_m_label)

log.info("   Total monthly records: %d", len(m_records))


# ─────────────────────────────────────────────────────────────────────────────
# 7. Build quarterly summary (STATE, FACILITY, Quarter as index)
# ─────────────────────────────────────────────────────────────────────────────
log.info("7. Building quarterly summary ...")
IDX_Q = ["STATE", "FACILITY", "Quarter"]

df_q = pd.DataFrame(q_records)
if df_q.empty:
    log.warning("   No quarterly data.")
    quarterly = pd.DataFrame()
else:
    quarterly = (
        df_q.pivot_table(index=IDX_Q, columns="column", values="value", aggfunc="sum")
        .reset_index()
    )
    quarterly.columns.name = None
    quarterly = add_derived(quarterly)

    # Sort by state, facility, quarter order
    q_order = {q: i for i, q in enumerate(list(QUARTER_LABEL.values()) + ["FY26"])}
    quarterly["_sort"] = quarterly["Quarter"].map(q_order)
    quarterly = (
        quarterly.sort_values(["STATE", "FACILITY", "_sort"])
        .drop(columns="_sort")
        .reset_index(drop=True)
    )

    # FY26 row per facility
    numeric_cols = quarterly.select_dtypes(include="number").columns.tolist()
    fy_rows = []
    for (state, facility), grp in quarterly.groupby(["STATE", "FACILITY"], sort=False):
        fy = {"STATE": state, "FACILITY": facility, "Quarter": "FY26"}
        for col in numeric_cols:
            if col in ("Linkage Rate", "VLC", "VLS"):
                fy[col] = None
            elif col in SNAPSHOT_COLS:
                # Use the open quarter's value only; if missing, FY row stays blank
                open_idx = grp.index[grp["Quarter"] == open_q_label]
                if len(open_idx) and pd.notna(grp.at[open_idx[0], col]):
                    fy[col] = float(grp.at[open_idx[0], col])
                else:
                    fy[col] = None
            else:
                fy[col] = float(grp[col].sum(skipna=True))
        fy["Linkage Rate"] = safe_div(fy.get("TX_NEW"),    fy.get("HTS_TST_POS"))
        fy["VLC"]          = safe_div(fy.get("TX_PVLS_D"), fy.get("TX_CURR"))
        fy["VLS"]          = safe_div(fy.get("TX_PVLS_N"), fy.get("TX_PVLS_D"))
        fy_rows.append(fy)

    quarterly = pd.concat([quarterly, pd.DataFrame(fy_rows)], ignore_index=True)

    # ── YTD cumulative actuals for flow indicators (Q1..Q4 rows only) ──
    _snap_derived = SNAPSHOT_COLS | {"VLC", "VLS", "Linkage Rate"}
    _flow_cols_q  = [c for c in quarterly.select_dtypes(include="number").columns
                     if c not in _snap_derived]
    _q_order_ytd  = [QUARTER_LABEL[q] for q in THIS_FY_QUARTERS]
    _fy_lbl       = f"FY{THIS_FY % 100}"

    for _, _grp in quarterly.groupby(["STATE", "FACILITY"], sort=False):
        _running: dict = {}
        for _qtr in _q_order_ytd:
            _idx = _grp.index[_grp["Quarter"] == _qtr]
            if not len(_idx):
                continue
            for _c in _flow_cols_q:
                _pv = quarterly.at[_idx[0], _c]
                _running[_c] = _running.get(_c, 0.0) + (float(_pv) if pd.notna(_pv) else 0.0)
                quarterly.at[_idx[0], _c] = _running[_c]

    # Re-derive Linkage Rate from cumulative TX_NEW / HTS_TST_POS for Q1-Q4 rows
    if "TX_NEW" in quarterly.columns and "HTS_TST_POS" in quarterly.columns:
        _q_mask = quarterly["Quarter"] != _fy_lbl
        quarterly.loc[_q_mask, "Linkage Rate"] = safe_div(
            quarterly.loc[_q_mask, "TX_NEW"],
            quarterly.loc[_q_mask, "HTS_TST_POS"].replace(0, float("nan"))
        )

    # Re-sort with FY row last per facility
    _q_sort_map = {q: i for i, q in enumerate(list(QUARTER_LABEL.values()) + [_fy_lbl])}
    quarterly["_sort"] = quarterly["Quarter"].map(_q_sort_map)
    quarterly = quarterly.sort_values(["STATE", "FACILITY", "_sort"]).drop(columns="_sort").reset_index(drop=True)

    log.info("   %d rows  (%d facilities)", len(quarterly), quarterly["FACILITY"].nunique())


# ─────────────────────────────────────────────────────────────────────────────
# 8. Build monthly summary (STATE, FACILITY, Quarter, Month as index)
# ─────────────────────────────────────────────────────────────────────────────
log.info("8. Building monthly summary ...")
IDX_M = ["STATE", "FACILITY", "Quarter", "Month"]

df_m = pd.DataFrame(m_records)
if df_m.empty:
    log.warning("   No monthly data.")
    monthly = pd.DataFrame()
else:
    monthly = (
        df_m.pivot_table(index=IDX_M, columns="column", values="value", aggfunc="sum")
        .reset_index()
    )
    monthly.columns.name = None
    monthly = add_derived(monthly)

    month_order = {m: i for i, m in enumerate(MONTH_LABEL.values())}
    monthly["_sort"] = monthly["Month"].map(month_order)
    monthly = (
        monthly.sort_values(["STATE", "FACILITY", "_sort"])
        .drop(columns="_sort")
        .reset_index(drop=True)
    )

    # ── YTD cumulative actuals for flow indicators ──
    _snap_derived_m = SNAPSHOT_COLS | {"VLC", "VLS", "Linkage Rate"}
    _flow_cols_m    = [c for c in monthly.select_dtypes(include="number").columns
                       if c not in _snap_derived_m]
    _m_order_ytd    = list(MONTH_LABEL.values())

    for _, _grp in monthly.groupby(["STATE", "FACILITY"], sort=False):
        _running_m: dict = {}
        for _m in _m_order_ytd:
            _idx = _grp.index[_grp["Month"] == _m]
            if not len(_idx):
                continue
            for _c in _flow_cols_m:
                _pv = monthly.at[_idx[0], _c]
                _running_m[_c] = _running_m.get(_c, 0.0) + (float(_pv) if pd.notna(_pv) else 0.0)
                monthly.at[_idx[0], _c] = _running_m[_c]

    # Re-derive Linkage Rate from cumulative values
    if "TX_NEW" in monthly.columns and "HTS_TST_POS" in monthly.columns:
        monthly["Linkage Rate"] = safe_div(
            monthly["TX_NEW"],
            monthly["HTS_TST_POS"].replace(0, float("nan"))
        )

    log.info("   %d rows  (%d facilities)", len(monthly), monthly["FACILITY"].nunique())


# ─────────────────────────────────────────────────────────────────────────────
# 9. Print preview
# ─────────────────────────────────────────────────────────────────────────────
if not quarterly.empty:
    sep()
    preview_cols = ["STATE", "FACILITY", "Quarter", "HTS_TST", "HTS_TST_POS",
                    "TX_NEW", "TX_CURR", "TX_PVLS_D", "TX_PVLS_N", "VLC", "VLS"]
    preview_cols = [c for c in preview_cols if c in quarterly.columns]
    log.info("QUARTERLY PREVIEW (first 10 rows):")
    log.info("\n%s", quarterly[preview_cols].head(10).to_string(index=False))
    sep()


# ─────────────────────────────────────────────────────────────────────────────
# 10. Load targets from Azure Blob + merge into quarterly / monthly sheets
# ─────────────────────────────────────────────────────────────────────────────
log.info("10. Loading FY targets from Azure Blob ...")
_targets = _load_targets()
if _targets is not None:
    quarterly = _apply_targets_quarterly(quarterly, _targets)
    monthly   = _apply_targets_monthly(monthly, _targets)
    _tgt_cols = sum(1 for c in quarterly.columns if c.endswith("_Target"))
    log.info("    %d target column(s) added to Quarterly; %d to Monthly.",
             _tgt_cols, sum(1 for c in monthly.columns if c.endswith("_Target")))
else:
    log.warning("    Continuing without targets.")


# ─────────────────────────────────────────────────────────────────────────────
# 11. Save to Excel
# ─────────────────────────────────────────────────────────────────────────────
from datetime import datetime as _dt

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
out_file = OUTPUT_DIR / f"ACEBAY_Dhis2_Output_FY{THIS_FY}.xlsx"

PCT_FMT  = "0.00%"
SKIP_RND = {"VLC", "VLS", "Linkage Rate"}

def _round_counts(df):
    """Round all numeric columns to 0 dp, except ratios/percentages."""
    df = df.copy()
    for col in df.select_dtypes(include="number").columns:
        if col not in SKIP_RND and not col.endswith("_Ach%"):
            df[col] = df[col].round(0)
    return df

with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
    for df, sheet in [(quarterly, "Quarterly"), (monthly, "Monthly")]:
        if df is None or df.empty:
            continue
        _round_counts(df).to_excel(writer, sheet_name=sheet, index=False)
        ws = writer.sheets[sheet]
        for col_idx, col_name in enumerate(df.columns, start=1):
            if col_name in {"VLC", "VLS", "Linkage Rate"} or col_name.endswith("_Ach%"):
                for row_cells in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                    for c in row_cells:
                        c.number_format = PCT_FMT

log.info("Saved: %s", out_file)
log.info("  Sheet 'Quarterly' — STATE, FACILITY, Qtr1..Qtr4 + FY26 per facility")
log.info("  Sheet 'Monthly'   — STATE, FACILITY, Oct-25..Sep-26 per facility")

# ─────────────────────────────────────────────────────────────────────────────
# 12. Upload to Azure Blob Storage
# ─────────────────────────────────────────────────────────────────────────────
dashboard_blob = f"FY{THIS_FY}_dashboard.xlsx"
log.info("12. Uploading to Azure Blob as '%s' ...", dashboard_blob)
if ARGS.skip_upload:
    log.info("    Skipped — --skip-upload flag set.")
else:
    try:
        if _azure_configured():
            blob_url = _azure_upload(out_file, dashboard_blob)
            log.info("    Uploaded: %s", blob_url)
        else:
            log.warning("    Skipped — Azure credentials not configured.")
    except Exception as _upload_err:
        log.warning("    Upload failed: %s", _upload_err)

# Log this run and close cache
_rows_api   = len(new_q_all) + len(new_m_all)
_rows_cache = len(q_records) + len(m_records) - _rows_api
conn_cache.execute(
    "INSERT INTO run_log (mode, ran_at, open_periods, rows_api, rows_cache) VALUES (?,?,?,?,?)",
    (
        ARGS.mode,
        _dt.now().isoformat(),
        f"{open_q_label} / {open_m_label}",
        _rows_api,
        _rows_cache,
    ),
)
conn_cache.commit()

# ── Item 6: Run log summary (last 5 runs) ────────────────────────────────────
sep("-")
log.info("Recent run history:")
for _row in conn_cache.execute(
    "SELECT ran_at, mode, open_periods, rows_api, rows_cache FROM run_log ORDER BY id DESC LIMIT 5"
):
    log.info("  %s | mode=%-11s | period=%-12s | api=%5d | cache=%5d",
             _row[0][:19], _row[1], _row[2], _row[3], _row[4])

conn_cache.close()

log.info("Done.")
sep()
