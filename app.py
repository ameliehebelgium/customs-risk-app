import streamlit as st
import pandas as pd
from datetime import date, datetime
from pathlib import Path
import hashlib
import json
import io
import re

# ─── reportlab imports for PDF export ────────────────────────────────────────
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)

# ─── Google Sheets ────────────────────────────────────────────────────────────
import gspread
from google.oauth2.service_account import Credentials

SHEET_ID   = "1X6EGJDD3gmzAtR_VJF076Wlg-FuspiLsVyekGQ6fJG4"
SHEET_NAME = "Feuille 1"   # expected tab name — used as a hint only, not load-bearing
SHEET_GID  = None          # optional: stable numeric tab id (see .../edit#gid=NNNN in the
                           # sheet's browser URL). If set, this is tried first and survives
                           # the tab being renamed or reordered. Safe to leave as None.


class WorksheetLocateError(RuntimeError):
    """Raised when the reference-table worksheet can't be found by gid, name, or header."""


def _find_worksheet_by_header(spreadsheet):
    """Self-healing fallback: scan every tab in the spreadsheet and return the one
    whose header row matches our expected columns. This is what makes the app
    survive a tab being renamed — instead of silently reading an empty/wrong tab,
    it finds the real data wherever it currently lives."""
    candidates = [
        ws for ws in spreadsheet.worksheets()
        if ws.row_values(1)[:1] == REF_COLUMNS[:1]
    ]
    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0]
    # More than one tab has a matching header (shouldn't normally happen) —
    # prefer the one that actually holds the most rows.
    return max(candidates, key=lambda w: w.row_count)


@st.cache_resource(ttl=30)
def _get_gclient():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    if "gcp_service_account" in st.secrets:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], scopes=scopes
        )
    else:
        creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
    return gspread.authorize(creds)


@st.cache_resource(ttl=30)
def _get_gsheet():
    gc = _get_gclient()
    spreadsheet = gc.open_by_key(SHEET_ID)

    # 1) Preferred: stable numeric gid — immune to renames/reordering.
    if SHEET_GID is not None:
        try:
            return spreadsheet.get_worksheet_by_id(SHEET_GID)
        except Exception:
            pass

    # 2) By configured tab name.
    try:
        return spreadsheet.worksheet(SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        pass

    # 3) Self-healing: locate the tab by its header row, whatever it's named.
    ws = _find_worksheet_by_header(spreadsheet)
    if ws is not None:
        return ws

    raise WorksheetLocateError(
        "Could not find the reference-table worksheet by gid, name, or header row. "
        f"Tabs currently in the sheet: {[w.title for w in spreadsheet.worksheets()]}. "
        "The tab may have been deleted — check the Google Sheet directly."
    )


# ─── Automatic backups ────────────────────────────────────────────────────────
# A full copy of the live spreadsheet, at most once per day, named to match the
# manual backup convention that already saved this database once ("Customs Risk
# Database BACKUP YYYY-MM-DD"). Purely additive — never touches or deletes
# existing data or older backups.
BACKUP_TITLE_PREFIX = "Customs Risk Database BACKUP "


def _today_backup_title() -> str:
    return f"{BACKUP_TITLE_PREFIX}{date.today().isoformat()}"


def backup_spreadsheet_if_due(force: bool = False) -> bool:
    """Create a dated full copy of the live spreadsheet (see app.py for full
    docstring — kept identical between the two files on purpose). Never raises."""
    try:
        gc = _get_gclient()
        if force:
            title = f"{_today_backup_title()} {datetime.now().strftime('%H-%M')}"
        else:
            title = _today_backup_title()
            if gc.list_spreadsheet_files(title=title):
                return False
        gc.copy(SHEET_ID, title=title, copy_permissions=True, copy_comments=False)
        return True
    except Exception as e:
        st.session_state["_last_backup_error"] = str(e)
        return False


def list_backups() -> list:
    try:
        gc = _get_gclient()
        files = gc.list_spreadsheet_files()
        backups = [f for f in files if f.get("name", "").startswith(BACKUP_TITLE_PREFIX)]
        backups.sort(key=lambda f: f.get("createdTime", ""), reverse=True)
        return backups
    except Exception:
        return []


# ─── File paths ───────────────────────────────────────────────────────────────
USERS_FILE = Path("users.json")
LOGO_FILE  = Path("vevor_logo.png")

# ─── Column definitions ───────────────────────────────────────────────────────
# Reference table: SKU | Corrected DESCRIPTION | Customs Commodity Code | Duty Rate
REF_COLUMNS = ["SKU", "Corrected DESCRIPTION", "Customs Commodity Code", "Duty Rate"]
HS_EXPECTED_LEN = 10   # EU CN/TARIC import codes are 10 digits; anything else is flagged

# Packing list parsed columns (internal use)
DOC_COLUMNS = ["Source File", "Current Container", "Job Number", "SKU Number",
               "Line No", "Product Description", "HS Code", "Qty"]


# ════════════════════════════════════════════════════════════════════════════════
# AUTH HELPERS  (unchanged)
# ════════════════════════════════════════════════════════════════════════════════

def _hash(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()


def load_users() -> dict:
    if USERS_FILE.exists():
        with open(USERS_FILE) as f:
            return json.load(f)
    default = {
        "admin":   {"password": _hash("admin123"),      "role": "admin",   "display": "Admin"},
        "broker1": {"password": _hash("broker2026"),    "role": "broker",  "display": "Broker User"},
        "visitor": {"password": _hash("visit2026"),     "role": "visitor", "display": "Visitor"},
        "geodis":  {"password": _hash("geodis123456!"), "role": "broker",  "display": "Geodis"},
        "rhenus":  {"password": _hash("rhenus123456!"), "role": "broker",  "display": "Rhenus"},
    }
    save_users(default)
    return default


def save_users(users: dict):
    with open(USERS_FILE, "w") as f:
        json.dump(users, f, indent=2)


def check_login(username: str, password: str):
    users = load_users()
    if username in users and users[username]["password"] == _hash(password):
        return users[username]
    return None


def login_screen():
    st.set_page_config(page_title="Customs Risk Database", layout="centered")
    st.markdown("""
        <div style='text-align:center; padding: 2rem 0 1rem 0;'>
            <h2 style='color:#1a3c6e;'>🛃 Customs Risk Database</h2>
            <p style='color:#666;'>Vevor EU — Internal Use Only</p>
        </div>
    """, unsafe_allow_html=True)
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Log In", use_container_width=True)
    if submitted:
        user = check_login(username.strip(), password)
        if user:
            st.session_state["logged_in"] = True
            st.session_state["username"]  = username.strip()
            st.session_state["role"]      = user["role"]
            st.session_state["display"]   = user["display"]
            st.rerun()
        else:
            st.error("Incorrect username or password.")


# ════════════════════════════════════════════════════════════════════════════════
# DATA HELPERS
# ════════════════════════════════════════════════════════════════════════════════

def clean_text(value) -> str:
    if pd.isna(value) or value is None:
        return ""
    text = str(value).strip()
    if text.lower() in ["nan", "none", "nat"]:
        return ""
    return text


def clean_hs(value) -> str:
    value = clean_text(value)
    if value.endswith(".0"):
        value = value[:-2]
    return value.replace(" ", "")


def format_duty_rate(value) -> str:
    """Normalise a duty rate to e.g. '2,7%'. Accepts 0.027 (fraction), '2.7%' or '2,7%'."""
    text = clean_text(value)
    if text == "":
        return ""
    has_pct = "%" in text
    text = text.replace("%", "").replace(",", ".").strip()
    try:
        number = float(text)
    except Exception:
        return clean_text(value)
    if not has_pct and 0 <= number < 1:
        number = number * 100
    number = round(number, 2)
    out = str(int(number)) if number == int(number) else f"{number:.2f}".rstrip("0").rstrip(".")
    return out.replace(".", ",") + "%"


def hs_len_abnormal(value) -> bool:
    """True if an HS code is present but is not a 10-digit number (e.g. '0', or a
    leading zero lost on a 9-digit code)."""
    hs = clean_hs(value)
    return bool(hs) and (len(hs) != HS_EXPECTED_LEN or not hs.isdigit())


def status_series(df: pd.DataFrame) -> pd.Series:
    """Vectorised version of sku_status() — needed for a ~45k-row reference table."""
    has_desc = df["Corrected DESCRIPTION"].astype(str).str.strip() != ""
    has_hs   = df["Customs Commodity Code"].astype(str).str.strip() != ""
    out = pd.Series("empty", index=df.index)
    out[has_hs & ~has_desc] = "hs_only"
    out[has_desc & ~has_hs] = "desc_only"
    out[has_desc & has_hs]  = "confirmed"
    return out


@st.cache_data(ttl=120, show_spinner=False)
def load_reference_db() -> pd.DataFrame:
    """Load SKU reference table from Google Sheets."""
    try:
        ws = _get_gsheet()
    except WorksheetLocateError as e:
        st.error(
            f"⚠️ Could not locate the reference table sheet ({e}). "
            "This is different from having zero entries — please check the Google "
            "Sheet before assuming data is missing."
        )
        return pd.DataFrame(columns=REF_COLUMNS)
    try:
        data = ws.get_all_records(default_blank="")
        if not data:
            return pd.DataFrame(columns=REF_COLUMNS)
        df = pd.DataFrame(data, dtype=str)
        for col in REF_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        df = df[REF_COLUMNS].fillna("")
        df["SKU"]                    = df["SKU"].str.strip()
        df["Corrected DESCRIPTION"]  = df["Corrected DESCRIPTION"].str.strip()
        df["Customs Commodity Code"] = df["Customs Commodity Code"].str.strip().str.replace(" ", "", regex=False)
        df["Duty Rate"]              = df["Duty Rate"].apply(format_duty_rate)
        # Drop fully empty rows
        df = df[df["SKU"] != ""]
        return df
    except Exception as e:
        st.error(f"Error loading reference database: {e}")
        return pd.DataFrame(columns=REF_COLUMNS)


def save_reference_db(df: pd.DataFrame):
    """Overwrite the Google Sheet with the reference table."""
    # Best-effort daily safety snapshot, taken before we overwrite anything.
    backup_spreadsheet_if_due()
    try:
        df = df.copy().fillna("")
        for _c in REF_COLUMNS:
            if _c not in df.columns:
                df[_c] = ""
        df["Duty Rate"]              = df["Duty Rate"].apply(format_duty_rate)
        df["SKU"]                    = df["SKU"].apply(clean_text)
        df["Corrected DESCRIPTION"]  = df["Corrected DESCRIPTION"].apply(clean_text)
        df["Customs Commodity Code"] = df["Customs Commodity Code"].apply(clean_hs)
        ws = _get_gsheet()

        # Safety guard: never let an accidentally-empty save silently wipe real data.
        existing_rows = max(0, len(ws.get_all_values()) - 1)
        if len(df) == 0 and existing_rows > 0:
            st.error(
                "⚠️ Refused to save: the new data is empty but the sheet currently "
                f"has {existing_rows} row(s). This looks like a lookup/display bug, "
                "not an intentional 'delete everything' — check the Google Sheet "
                "before retrying."
            )
            return

        ws.clear()
        # RAW: keep HS codes as text (USER_ENTERED would drop leading zeros)
        ws.update([REF_COLUMNS] + df[REF_COLUMNS].values.tolist(), value_input_option="RAW")
        load_reference_db.clear()
    except Exception as e:
        st.error(f"Error saving reference database: {e}")


def sku_status(row) -> str:
    """
    Return status of a reference row:
      'confirmed'  — has both description and HS code
      'hs_only'    — has HS but no description (inspected, awaiting confirmation)
      'desc_only'  — has description but no HS
      'empty'      — neither description nor HS
    """
    has_desc = bool(clean_text(row.get("Corrected DESCRIPTION", "")))
    has_hs   = bool(clean_hs(row.get("Customs Commodity Code", "")))
    if has_desc and has_hs:
        return "confirmed"
    if has_hs and not has_desc:
        return "hs_only"
    if has_desc and not has_hs:
        return "desc_only"
    return "empty"


def build_sku_lookup(ref_df: pd.DataFrame) -> dict:
    """
    Build a dict: SKU (upper) -> last row (as dict) + status.
    For conflicting SKUs the last occurrence wins (most recent data).
    """
    lookup = {}
    statuses = status_series(ref_df).tolist()
    for d, stt in zip(ref_df.to_dict("records"), statuses):
        sku = clean_text(d["SKU"]).upper()
        if sku:
            d["_status"] = stt
            lookup[sku] = d
    return lookup


# ════════════════════════════════════════════════════════════════════════════════
# PACKING LIST PARSING  (preserved from original, unchanged)
# ════════════════════════════════════════════════════════════════════════════════

def _looks_like_container(value: str) -> bool:
    return bool(re.match(r'^[A-Z]{3,4}\d{6,8}$', value.strip().upper()))


def extract_containers_from_header(raw_df) -> list:
    containers = []
    for i in range(min(len(raw_df), 20)):
        row_values = [clean_text(x) for x in raw_df.iloc[i].tolist()]
        row_text   = " ".join(row_values).upper()
        if "CONTAINER" not in row_text:
            continue
        full_text = " ".join(row_values)
        m = re.search(r'Container\s*NO\.?\s*:?\s*([A-Z0-9+\s]+)', full_text, re.IGNORECASE)
        if m:
            raw_ids = re.split(r'[+\s]+', m.group(1).strip())
            found = [x.strip().upper() for x in raw_ids if _looks_like_container(x)]
            if found:
                containers.extend(found)
                break
        for j, value in enumerate(row_values):
            if "CONTAINER" in value.upper():
                for k in range(j + 1, min(j + 4, len(row_values))):
                    candidate = clean_text(row_values[k])
                    parts = re.split(r'[+\s]+', candidate)
                    found = [p.strip().upper() for p in parts if _looks_like_container(p)]
                    if found:
                        containers.extend(found)
                        break
                if containers:
                    break
        if containers:
            break
    return containers


def extract_po_number(raw_df) -> str:
    for i in range(min(len(raw_df), 20)):
        row_values = [str(x).strip() for x in raw_df.iloc[i].tolist()]
        for j, val in enumerate(row_values):
            if not re.search(r'PO\s*NO', val, re.IGNORECASE):
                continue
            m = re.search(r'PO\s*NO\.?\s*:?\s*([A-Z]{1,4}-?\d{4,8})', val, re.IGNORECASE)
            if m:
                return m.group(1).strip()
            for k in range(j + 1, min(j + 4, len(row_values))):
                candidate = row_values[k].strip()
                if candidate and candidate.lower() not in ("", "none", "nan"):
                    return candidate
    return ""


def pick_packing_sheet(xls) -> str:
    """Which sheet holds the packing list. Single-sheet files: that sheet.
    Three-in-one files (发票 / 箱单 / 合同): the 箱单 sheet."""
    names = list(xls.sheet_names)
    if len(names) == 1:
        return names[0]
    for n in names:                      # 1) by sheet name
        nl = str(n).lower()
        if "箱单" in nl or "packing" in nl:
            return n
    for n in names:                      # 2) by content: has a 'Packing List' title
        try:
            head = xls.parse(n, header=None, dtype=str, nrows=6)
            if head.astype(str).apply(lambda c: c.str.contains("packing list", case=False, na=False)).any().any():
                return n
        except Exception:
            pass
    return names[0]                      # 3) fallback: first sheet (old behaviour)


def _normalize_document_file(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheet = pick_packing_sheet(xls)
    raw = xls.parse(sheet, header=None, dtype=str)

    first_row = [str(x).strip().lower() for x in raw.iloc[0].tolist()]
    is_vevor_format1 = (
        any("sku" in c for c in first_row) and
        any("hs" in c for c in first_row) and
        any("job" in c for c in first_row)
    )
    is_vevor_format2 = (
        any("排柜单号" in c for c in first_row) and
        any("sku" in c for c in first_row) and
        any("英文申报名称" in c for c in first_row)
    )

    if is_vevor_format1:
        df = xls.parse(sheet, header=0, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        job_col = sku_col = desc_col = hs_col = container_col = None
        for col in df.columns:
            col_l = col.lower()
            if job_col is None and "job" in col_l: job_col = col
            if sku_col is None and "sku" in col_l: sku_col = col
            if desc_col is None and "description" in col_l: desc_col = col
            if hs_col is None and "hs" in col_l: hs_col = col
            if container_col is None and "container" in col_l: container_col = col
        if not desc_col or not hs_col:
            return pd.DataFrame(columns=DOC_COLUMNS)
        def _extract_job(val):
            val = clean_text(val)
            m = re.search(r"(\d{6})$", val)
            return m.group(1) if m else val
        result = pd.DataFrame()
        result["Source File"]         = uploaded_file.name
        result["Line No"]             = range(1, len(df) + 1)
        result["Product Description"] = df[desc_col].apply(clean_text)
        result["HS Code"]             = df[hs_col].apply(clean_hs)
        result["Qty"]                 = ""
        result["SKU Number"]          = df[sku_col].apply(clean_text) if sku_col else ""
        result["Current Container"]   = df[container_col].apply(clean_text) if container_col else ""
        result["Job Number"]          = df[job_col].apply(_extract_job) if job_col else ""
        result = result[(result["Product Description"] != "") & (result["HS Code"] != "")]
        return result[DOC_COLUMNS]

    if is_vevor_format2:
        df = xls.parse(sheet, header=0, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        job_col  = next((c for c in df.columns if "排柜单号" in c), None)
        sku_col  = next((c for c in df.columns if c.upper() == "SKU"), None)
        desc_col = next((c for c in df.columns if "英文申报名称" in c), None)
        hs_col   = next((c for c in df.columns if "中国出口HS" in c or "出口HS" in c), None)
        con_col  = next((c for c in df.columns if "箱号" in c), None)
        if not desc_col or not hs_col:
            return pd.DataFrame(columns=DOC_COLUMNS)
        def _extract_job2(val):
            val = clean_text(val)
            m = re.search(r"(\d{6})$", val)
            return m.group(1) if m else val
        result = pd.DataFrame()
        result["Source File"]         = uploaded_file.name
        result["Line No"]             = range(1, len(df) + 1)
        result["Product Description"] = df[desc_col].apply(clean_text)
        result["HS Code"]             = df[hs_col].apply(clean_hs)
        result["Qty"]                 = ""
        result["SKU Number"]          = df[sku_col].apply(clean_text) if sku_col else ""
        result["Current Container"]   = df[con_col].apply(clean_text) if con_col else ""
        result["Job Number"]          = df[job_col].apply(_extract_job2) if job_col else ""
        result = result[(result["Product Description"] != "") & (result["HS Code"] != "") & (result["HS Code"].str.len() >= 6)]
        return result[DOC_COLUMNS]

    # ── Standard packing list / invoice ──
    header_containers = extract_containers_from_header(raw)
    fallback_container = "+".join(header_containers) if header_containers else ""

    header_row = 0
    for i in range(len(raw)):
        row_text = " ".join([clean_text(x).upper() for x in raw.iloc[i].tolist()])
        if ("DESCRIPTION" in row_text or "ITEM" in row_text) and ("HS" in row_text or "HTS" in row_text):
            header_row = i
            break

    df = raw.iloc[header_row + 1:].copy()
    raw_headers = [str(c).replace("\n", " ").strip() for c in raw.iloc[header_row].tolist()]
    named_headers = []
    for idx, h in enumerate(raw_headers):
        if h.lower() in ("nan", "none", ""):
            named_headers.append(f"Unnamed: {idx}")
        else:
            named_headers.append(h)
    df.columns = named_headers
    df = df.reset_index(drop=True)
    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]

    product_col = hs_col = qty_col = container_col = sku_col_pl = None
    for col in df.columns:
        col_upper = col.upper()
        if product_col is None and ("DESCRIPTION" in col_upper or "ITEM" in col_upper or "PRODUCT" in col_upper):
            product_col = col
        if hs_col is None and ("HS" in col_upper or "HTS" in col_upper):
            hs_col = col
        if qty_col is None and ("QTY" in col_upper or "QUANTITY" in col_upper):
            qty_col = col
        if container_col is None and ("柜号" in col or "CONTAINER" in col_upper):
            container_col = col
        if sku_col_pl is None and "SKU" in col_upper:
            sku_col_pl = col

    if sku_col_pl is None:
        def _looks_like_sku(val):
            s = str(val).strip()
            return (len(s) >= 8 and bool(re.search(r'[A-Za-z]', s))
                    and bool(re.search(r'\d', s)))
        for col in df.columns:
            if "Unnamed" in str(col) or str(col).strip() in ("", "None", "nan"):
                sample = df[col].dropna().head(20).tolist()
                sku_like = sum(1 for v in sample if _looks_like_sku(v))
                if sku_like >= max(1, len(sample) * 0.5):
                    sku_col_pl = col
                    break

    if product_col is None or hs_col is None:
        return pd.DataFrame(columns=DOC_COLUMNS)

    po_number = extract_po_number(raw)
    if not po_number:
        m = re.search(r'(\d{6})', uploaded_file.name)
        po_number = m.group(1) if m else ""
    source_label = po_number if po_number else uploaded_file.name

    # Duty rate printed in the packing list: unnamed column right of the SKU/Qty columns,
    # numeric values between 0 and 1 (e.g. 0.060 = 6%). Detected by content, not position.
    duty_col = None
    if sku_col_pl is not None:
        _cols = list(df.columns)
        _pos = _cols.index(sku_col_pl)
        for _c in _cols[_pos + 1:_pos + 4]:
            _nums = pd.to_numeric(df[_c], errors="coerce").dropna()
            _filled = df[_c].replace("", pd.NA).dropna()
            if len(_nums) > 0 and len(_nums) >= 0.5 * len(_filled) and (_nums >= 0).all() and (_nums < 1).all():
                duty_col = _c
                break

    result = pd.DataFrame({
        "Line No":             range(1, len(df) + 1),
        "Duty Rate":           df[duty_col].apply(format_duty_rate).values if duty_col else "",
        "_HS_Row":             [header_row + 2 + i for i in range(len(df))],   # Excel row number
        "_HS_Col":             openpyxl_col_letter(list(df.columns).index(hs_col) + 1),
        "_Sheet":              sheet,
        "Product Description": df[product_col].apply(clean_text).values,
        "HS Code":             df[hs_col].apply(clean_hs).values,
        "Qty":                 df[qty_col].apply(clean_text).values if qty_col else "",
        "SKU Number":          df[sku_col_pl].apply(clean_text).values if sku_col_pl else "",
        "Job Number":          "",
    })
    result["Source File"] = source_label
    result["Product Description"] = result["Product Description"].replace("", pd.NA).ffill().fillna("")
    result["HS Code"] = result["HS Code"].replace("", pd.NA).ffill().fillna("")

    if container_col:
        result["Current Container"] = df[container_col].apply(
            lambda x: clean_text(x) if clean_text(x) else fallback_container
        ).values
    else:
        result["Current Container"] = fallback_container

    result = result[(result["Product Description"] != "") & (result["HS Code"] != "")]
    return result[DOC_COLUMNS + ["Duty Rate", "_Sheet", "_HS_Row", "_HS_Col"]]


def openpyxl_col_letter(idx: int) -> str:
    from openpyxl.utils import get_column_letter
    return get_column_letter(idx)


def normalize_document_file(uploaded_file):
    d = _normalize_document_file(uploaded_file)
    for _c in ("Duty Rate", "_Sheet", "_HS_Row", "_HS_Col"):
        if _c not in d.columns:
            d[_c] = ""
    return d


# ════════════════════════════════════════════════════════════════════════════════
# RISK CHECK LOGIC  (new SKU-based)
# ════════════════════════════════════════════════════════════════════════════════

def check_packing_list_against_reference(doc_df: pd.DataFrame, ref_df: pd.DataFrame) -> pd.DataFrame:
    """
    For each line in doc_df, look up the SKU in the reference table.
    Returns a DataFrame of flagged lines with severity and reason.

    Severity levels:
      RED    — SKU confirmed (has desc + HS), but PL description or HS doesn't match
      YELLOW — SKU found but pending confirmation (HS only, no desc), any mismatch present
      ORANGE — SKU found but both description and HS are missing in reference
      INFO   — SKU not found in reference at all (unknown SKU — log for awareness)
    """
    lookup = build_sku_lookup(ref_df)
    results = []

    for _, doc_row in doc_df.iterrows():
        sku_raw  = clean_text(doc_row.get("SKU Number", ""))
        sku_key  = sku_raw.upper()
        pl_desc  = clean_text(doc_row.get("Product Description", "")).upper()
        pl_hs    = clean_hs(doc_row.get("HS Code", ""))

        if not sku_key:
            continue  # No SKU on this line — skip

        if sku_key not in lookup:
            continue  # SKU not in reference — no alert

        ref = lookup[sku_key]
        status   = ref["_status"]
        ref_desc = clean_text(ref.get("Corrected DESCRIPTION", "")).upper()
        ref_hs   = clean_hs(ref.get("Customs Commodity Code", ""))
        ref_duty = clean_text(ref.get("Duty Rate", ""))
        ref_hs_bad = hs_len_abnormal(ref_hs)

        # ── Determine mismatches ──
        pl_duty = format_duty_rate(doc_row.get("Duty Rate", ""))
        _norm = lambda s: " ".join(str(s).split()).upper()
        def _state(a, b):
            if not a or not b:
                return "na"
            return "ok" if a == b else "diff"
        cmp_hs   = _state(pl_hs, ref_hs)
        cmp_desc = _state(_norm(pl_desc), _norm(ref_desc))
        cmp_duty = _state(pl_duty, ref_duty)
        hs_mismatch   = cmp_hs == "diff"
        desc_mismatch = cmp_desc == "diff"

        if ref_hs_bad:
            # The reference HS itself is not a 10-digit code, so it cannot be
            # compared reliably — flag the reference data instead of the PL.
            severity = "HSLEN"
            reasons = [f"Reference HS '{ref_hs}' has {len(ref_hs)} digit(s), expected {HS_EXPECTED_LEN}"]
            if pl_hs:
                reasons.append(f"PL declares HS: {pl_hs}")
            if desc_mismatch:
                reasons.append(f"Description mismatch: PL has '{doc_row.get('Product Description','')}', reference is '{ref.get('Corrected DESCRIPTION','')}'")
            reason_str = " | ".join(reasons)
            action = "CHECK REFERENCE"
            message = "HS length abnormal - reference HS is not a 10-digit code, cannot be compared reliably."

        elif status == "confirmed":
            if not hs_mismatch and not desc_mismatch:
                continue  # All good
            severity = "RED"
            reasons = []
            if hs_mismatch:
                reasons.append(f"HS mismatch: PL has {pl_hs}, reference is {ref_hs}")
            if desc_mismatch:
                reasons.append(f"Description mismatch: PL has '{doc_row.get('Product Description','')}', reference is '{ref.get('Corrected DESCRIPTION','')}'")
            reason_str = " | ".join(reasons)
            action = "STOP & REVIEW"
            message = "SKU confirmed in reference — HS or description differs from validated data."

        elif status == "hs_only":
            # Has HS in reference but no confirmed description yet
            severity = "YELLOW"
            reasons = []
            if hs_mismatch:
                reasons.append(f"HS mismatch: PL has {pl_hs}, reference has {ref_hs}")
            if not ref_desc:
                reasons.append("Description not yet confirmed in reference")
            reason_str = " | ".join(reasons) if reasons else "SKU pending confirmation — description missing in reference"
            action = "PENDING REVIEW"
            message = "SKU previously inspected but description not yet confirmed by customs."

        elif status == "desc_only":
            # Has description but no HS confirmed yet
            severity = "YELLOW"
            reasons = [f"HS code not yet confirmed in reference"]
            if pl_hs:
                reasons.append(f"PL declares HS: {pl_hs}")
            reason_str = " | ".join(reasons)
            action = "PENDING REVIEW"
            message = "SKU previously inspected but HS code not yet confirmed by customs."

        else:  # empty
            severity = "ORANGE"
            reason_str = "Both description and HS code missing in reference — SKU registered but not yet processed"
            action = "LOG & MONITOR"
            message = "SKU exists in reference database but has no validated description or HS code yet."

        results.append({
            "Severity":           severity,
            "Action":             action,
            "Message":            message,
            "Reason":             reason_str,
            "Source File":        doc_row.get("Source File", ""),
            "Current Container":  doc_row.get("Current Container", ""),
            "Job Number":         doc_row.get("Job Number", ""),
            "Line No":            doc_row.get("Line No", ""),
            "SKU":                sku_raw,
            "PL Description":     doc_row.get("Product Description", ""),
            "PL HS Code":         pl_hs,
            "Qty":                doc_row.get("Qty", ""),
            "Ref Description":    ref.get("Corrected DESCRIPTION", ""),
            "Ref HS Code":        ref_hs,
            "Ref Duty Rate":      ref_duty,
            "PL Duty Rate":       pl_duty,
            "File Name":          doc_row.get("_File", ""),
            "Sheet":              doc_row.get("_Sheet", ""),
            "HS Row":             doc_row.get("_HS_Row", ""),
            "HS Col":             doc_row.get("_HS_Col", ""),
            "Cmp HS":             cmp_hs,
            "Cmp Desc":           cmp_desc,
            "Cmp Duty":           cmp_duty,
            "Ref Status":         status,
        })

    return pd.DataFrame(results)


# ════════════════════════════════════════════════════════════════════════════════
# PDF EXPORT
# ════════════════════════════════════════════════════════════════════════════════

def apply_hs_corrections(file_bytes: bytes, sheet_name: str, edits: list):
    """Write corrected HS codes into the ORIGINAL workbook and paint those cells yellow.

    edits: list of (excel_row, column_letter, new_hs) — one per packing-list line.
    Sub-SKU lines share the HS cell of their product row (merged or blank cells below it);
    the edit goes to that product-row cell. If SKUs under one product need different HS
    codes the cell is left untouched and reported as a conflict.

    The workbook XML is patched in place (one cell value + one new style), so formulas,
    other sheets, widths, merges and formatting are untouched.
    Returns (new_bytes, applied, skipped); applied = [(cell, old, new, n_lines)],
    skipped = [(cell, reason)].
    """
    import io, re, zipfile
    import openpyxl
    from xml.sax.saxutils import escape as _esc

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb[sheet_name]

    # 1) resolve each edit to the cell that really holds the HS value
    targets = {}
    for row, col, new_hs in edits:
        coord = f"{col}{int(row)}"
        for mr in ws.merged_cells.ranges:
            if coord in mr:
                coord = mr.start_cell.coordinate
                break
        else:
            r = int(row)
            while r > 1 and ws[f"{col}{r}"].value in (None, ""):
                r -= 1
            coord = f"{col}{r}"
        targets.setdefault(coord, []).append(clean_hs(new_hs))

    applied, skipped, todo = [], [], {}
    for coord, vals in targets.items():
        old = clean_hs(ws[coord].value)
        if len(set(vals)) > 1:
            skipped.append((coord, "SKUs under this product need different HS codes: " + ", ".join(sorted(set(vals)))))
        elif vals[0] == old:
            skipped.append((coord, "already has this HS"))
        elif not vals[0]:
            skipped.append((coord, "no new HS"))
        else:
            todo[coord] = vals[0]
            applied.append((coord, old, vals[0], len(vals)))

    if not todo:
        return file_bytes, applied, skipped

    # 2) patch the XML
    zin = zipfile.ZipFile(io.BytesIO(file_bytes))
    names = zin.namelist()
    wbxml = zin.read("xl/workbook.xml").decode("utf-8")
    rels  = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    m = None
    for sm in re.finditer(r"<sheet\b[^>]*>", wbxml):
        tag = sm.group(0)
        nm = re.search(r'\bname="([^"]*)"', tag)
        if nm and nm.group(1).replace("&amp;", "&") == sheet_name:
            m = tag
            break
    if m is None:
        raise RuntimeError(f"sheet '{sheet_name}' not found in workbook.xml")
    rid = re.search(r'r:id="([^"]+)"', m).group(1)
    target = None
    for rm in re.finditer(r"<Relationship\b[^>]*>", rels):
        if f'Id="{rid}"' in rm.group(0):
            target = re.search(r'Target="([^"]+)"', rm.group(0)).group(1)
    sheet_path = target.lstrip("/") if target.startswith("/") else "xl/" + target
    sheet_xml = zin.read(sheet_path).decode("utf-8")
    styles_xml = zin.read("xl/styles.xml").decode("utf-8")

    # yellow fill + a clone of each affected cell style with that fill
    fm = re.search(r'<fills count="(\d+)">(.*?)</fills>', styles_xml, re.S)
    fill_id = int(fm.group(1))
    new_fill = '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>'
    styles_xml = styles_xml.replace(fm.group(0), f'<fills count="{fill_id + 1}">{fm.group(2)}{new_fill}</fills>')
    xm = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', styles_xml, re.S)
    xfs = re.findall(r"<xf\b[^>]*?(?:/>|>.*?</xf>)", xm.group(2), re.S)
    new_xfs, style_map = [], {}

    def yellow_style(s_idx: int) -> int:
        if s_idx not in style_map:
            xf = xfs[s_idx]
            head_end = xf.index(">")
            head, rest = xf[:head_end], xf[head_end:]
            head = re.sub(r'\sfillId="\d+"', "", head)
            head = re.sub(r'\sapplyFill="\d+"', "", head)
            head = head.replace("<xf", f'<xf fillId="{fill_id}" applyFill="1"', 1)
            if rest.startswith("/>") is False and head.endswith("/"):
                head = head[:-1]; rest = "/" + rest
            new_xfs.append(head + rest)
            style_map[s_idx] = len(xfs) + len(new_xfs) - 1
        return style_map[s_idx]

    for coord, new_hs in todo.items():
        cm = re.search(r'<c r="' + coord + r'"([^>]*?)(/>|>.*?</c>)', sheet_xml, re.S)
        if cm is None:
            raise RuntimeError(f"cell {coord} not found in sheet XML")
        attrs = cm.group(1)
        sm_ = re.search(r'\bs="(\d+)"', attrs)
        s_new = yellow_style(int(sm_.group(1)) if sm_ else 0)
        cell = f'<c r="{coord}" s="{s_new}" t="inlineStr"><is><t>{_esc(new_hs)}</t></is></c>'
        sheet_xml = sheet_xml[:cm.start()] + cell + sheet_xml[cm.end():]

    styles_xml = styles_xml.replace(
        xm.group(0),
        f'<cellXfs count="{len(xfs) + len(new_xfs)}">{xm.group(2)}{"".join(new_xfs)}</cellXfs>')

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == sheet_path:
                data = sheet_xml.encode("utf-8")
            elif item.filename == "xl/styles.xml":
                data = styles_xml.encode("utf-8")
            zout.writestr(item, data)
    return out.getvalue(), applied, skipped


def build_card_html(row, border: str, bg: str, icon: str) -> str:
    """Side-by-side comparison card: DATABASE (left) vs PACKING LIST (right), one aligned
    row per field. Differences are highlighted in red, matches in green."""
    import html as _h

    def _v(val):
        s = str(val).strip() if val is not None else ""
        return _h.escape(s) if s and s.lower() not in ("nan", "none", "") else "—"

    def _cells(ref_val, pl_val, state):
        if state == "diff":
            cell = "background:#fdecea; color:#c0392b; font-weight:700;"
            res = '<span style="color:#c0392b; font-weight:700;">≠ DIFF</span>'
        elif state == "ok":
            cell = "color:#1a6e3c; font-weight:600;"
            res = '<span style="color:#1a6e3c; font-weight:700;">✓</span>'
        else:
            cell = "color:#333; font-weight:600;"
            res = '<span style="color:#999;">—</span>'
        pad = "padding:7px 10px; border-bottom:1px solid #eee; word-break:break-word;"
        return (f'<div style="{pad}{cell}">{_v(ref_val)}</div>'
                f'<div style="{pad}{cell}">{_v(pl_val)}</div>'
                f'<div style="{pad}text-align:center;">{res}</div>')

    def _label(text):
        return ('<div style="padding:7px 10px; border-bottom:1px solid #eee; color:#666;'
                f' font-size:0.85rem;">{text}</div>')

    hd = "padding:8px 10px; font-weight:700; font-size:0.9rem; border-bottom:2px solid #ddd;"
    grid = (
        '<div style="display:grid; grid-template-columns:100px 1fr 1fr 70px; background:white;'
        ' border:1px solid #ddd; border-radius:8px; overflow:hidden; font-size:0.9rem;">'
        f'<div style="{hd}"></div>'
        f'<div style="{hd} color:#1a6e3c;">✅ DATABASE (reference)</div>'
        f'<div style="{hd} color:#1a3c6e;">📦 PACKING LIST</div>'
        f'<div style="{hd} text-align:center;"></div>'
        + _label("SKU")         + _cells(row.get("SKU"),             row.get("SKU"),           "ok")
        + _label("Description") + _cells(row.get("Ref Description"), row.get("PL Description"), row.get("Cmp Desc", "na"))
        + _label("HS Code")     + _cells(row.get("Ref HS Code"),     row.get("PL HS Code"),     row.get("Cmp HS", "na"))
        + _label("Duty Rate")   + _cells(row.get("Ref Duty Rate"),   row.get("PL Duty Rate"),   row.get("Cmp Duty", "na"))
        + '</div>'
    )

    ref_status = {
        "confirmed": "✅ Confirmed", "hs_only": "🟡 HS only — pending desc",
        "desc_only": "🟡 Desc only — pending HS", "empty": "⬜ Not yet processed",
    }.get(str(row.get("Ref Status", "")), str(row.get("Ref Status", "")))

    def _chip(label, val):
        return (f'<span style="display:inline-block; margin:2px 14px 2px 0; font-size:0.82rem;">'
                f'<span style="color:#777;">{label}:</span> <b>{_v(val)}</b></span>')

    meta = (_chip("PO / Source", row.get("Source File")) + _chip("Container", row.get("Current Container"))
            + _chip("Job", row.get("Job Number")) + _chip("Line", row.get("Line No"))
            + _chip("Qty", row.get("Qty")) + _chip("Reference status", ref_status))
    reason = (f'<div style="margin-top:6px; font-size:0.84rem; color:#555; font-style:italic;">'
              f'{_v(row.get("Reason"))}</div>')

    return (
        f'<div style="border:2px solid {border}; border-radius:10px; background:{bg};'
        f' padding:16px; margin-bottom:16px;">'
        f'<div style="font-weight:700; font-size:1.05rem; color:{border}; margin-bottom:12px;">'
        f'{icon} {_v(row.get("Action"))} &nbsp;|&nbsp; {_v(row.get("Message"))}</div>'
        + grid
        + f'<div style="margin-top:10px;">{meta}</div>' + reason
        + '</div>'
    )


def build_pdf_report(check_df: pd.DataFrame, doc_files_info: str) -> bytes:
    buf    = io.BytesIO()
    doc    = SimpleDocTemplate(buf, pagesize=landscape(A4),
                               leftMargin=1.5*cm, rightMargin=1.5*cm,
                               topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    story  = []

    title_style = ParagraphStyle("title", parent=styles["Title"],
                                 fontSize=16, textColor=colors.HexColor("#1a3c6e"), spaceAfter=6)
    sub_style   = ParagraphStyle("sub",   parent=styles["Normal"],
                                 fontSize=9,  textColor=colors.grey)
    h2_style    = ParagraphStyle("h2",    parent=styles["Heading2"],
                                 fontSize=12, textColor=colors.HexColor("#1a3c6e"), spaceBefore=12)
    cell_style  = ParagraphStyle("cell",  parent=styles["Normal"], fontSize=7.5, leading=10)
    label_style = ParagraphStyle("label", parent=styles["Normal"],
                                 fontSize=7.5, textColor=colors.HexColor("#666666"))
    value_style = ParagraphStyle("value", parent=styles["Normal"],
                                 fontSize=7.5, fontName="Helvetica-Bold")
    red_val     = ParagraphStyle("redval", parent=styles["Normal"],
                                 fontSize=7.5, fontName="Helvetica-Bold",
                                 textColor=colors.HexColor("#c0392b"))
    yellow_val  = ParagraphStyle("yelval", parent=styles["Normal"],
                                 fontSize=7.5, fontName="Helvetica-Bold",
                                 textColor=colors.HexColor("#b7770d"))
    green_val   = ParagraphStyle("greenval", parent=styles["Normal"],
                                 fontSize=7.5, fontName="Helvetica-Bold",
                                 textColor=colors.HexColor("#1a6e3c"))
    section_style = ParagraphStyle("sect", parent=styles["Normal"],
                                   fontSize=8, fontName="Helvetica-Bold",
                                   textColor=colors.HexColor("#1a3c6e"))

    red_df    = check_df[check_df["Severity"] == "RED"]
    yellow_df = check_df[check_df["Severity"] == "YELLOW"]
    orange_df = check_df[check_df["Severity"] == "ORANGE"]
    hslen_df  = check_df[check_df["Severity"] == "HSLEN"]

    story.append(Paragraph("Customs Risk Check Report", title_style))
    story.append(Paragraph(
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')} &nbsp;|&nbsp; "
        f"Files checked: {doc_files_info} &nbsp;|&nbsp; "
        f"Total alerts: {len(check_df)}",
        sub_style
    ))
    story.append(Spacer(1, 0.4*cm))

    summary_data = [
        ["RED - Action Required", "YELLOW - Pending Review", "ORANGE - Log & Monitor", "HS LENGTH ABNORMAL", "Total Flagged"],
        [str(len(red_df)), str(len(yellow_df)), str(len(orange_df)), str(len(hslen_df)), str(len(check_df))],
    ]
    summary_table = Table(summary_data, colWidths=[5*cm, 5*cm, 5*cm, 5*cm, 5*cm])
    summary_table.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (-1,0), colors.HexColor("#1a3c6e")),
        ("TEXTCOLOR",    (0,0), (-1,0), colors.white),
        ("BACKGROUND",   (0,1), (0,1),  colors.HexColor("#fde8e8")),
        ("BACKGROUND",   (1,1), (1,1),  colors.HexColor("#fffbe6")),
        ("BACKGROUND",   (2,1), (2,1),  colors.HexColor("#fff3e0")),
        ("BACKGROUND",   (3,1), (3,1),  colors.HexColor("#f4ecf7")),
        ("BACKGROUND",   (4,1), (4,1),  colors.HexColor("#e8f0fe")),
        ("FONTSIZE",     (0,1), (-1,1), 18),
        ("ALIGN",        (0,0), (-1,-1), "CENTER"),
        ("VALIGN",       (0,0), (-1,-1), "MIDDLE"),
        ("BOX",          (0,0), (-1,-1), 0.5, colors.HexColor("#cccccc")),
        ("INNERGRID",    (0,0), (-1,-1), 0.5, colors.HexColor("#cccccc")),
        ("TOPPADDING",   (0,0), (-1,-1), 8),
        ("BOTTOMPADDING",(0,0), (-1,-1), 8),
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 0.5*cm))

    def v(row, key, style=None):
        val = str(row.get(key, "") or "—")
        if not val or val in ("", "None", "nan"):
            val = "—"
        return Paragraph(val, style or value_style)

    def make_card_section(section_df, heading, border_color, bg_color):
        if len(section_df) == 0:
            return
        story.append(Paragraph(heading, h2_style))
        story.append(Spacer(1, 0.2*cm))
        for _, row in section_df.iterrows():
            alert_msg = f"<b>{row.get('Action','')} | {row.get('Message','')}</b>"
            banner = Table([[Paragraph(alert_msg, ParagraphStyle(
                "banner", parent=styles["Normal"], fontSize=8,
                textColor=border_color))]],
                colWidths=[25*cm])
            banner.setStyle(TableStyle([
                ("BACKGROUND",    (0,0), (-1,-1), bg_color),
                ("TOPPADDING",    (0,0), (-1,-1), 6),
                ("BOTTOMPADDING", (0,0), (-1,-1), 6),
                ("LEFTPADDING",   (0,0), (-1,-1), 8),
            ]))
            story.append(banner)

            from xml.sax.saxutils import escape as _esc

            def _cell(val, state):
                s = str(val if val is not None else "").strip()
                if not s or s.lower() in ("nan", "none"):
                    s = "—"
                st_ = {"diff": red_val, "ok": green_val}.get(state, value_style)
                return Paragraph(_esc(s), st_)

            def _res(state):
                txt = {"diff": "DIFF", "ok": "OK"}.get(state, "—")
                return Paragraph(txt, {"diff": red_val, "ok": green_val}.get(state, value_style))

            fields = [
                ("SKU",         row.get("SKU"),             row.get("SKU"),            "ok"),
                ("Description", row.get("Ref Description"), row.get("PL Description"), row.get("Cmp Desc", "na")),
                ("HS Code",     row.get("Ref HS Code"),     row.get("PL HS Code"),     row.get("Cmp HS", "na")),
                ("Duty Rate",   row.get("Ref Duty Rate"),   row.get("PL Duty Rate"),   row.get("Cmp Duty", "na")),
            ]
            data = [[Paragraph("", label_style),
                     Paragraph("DATABASE (reference)", section_style),
                     Paragraph("PACKING LIST", section_style),
                     Paragraph("", label_style)]]
            for lab, ref_v, pl_v, state in fields:
                data.append([Paragraph(lab, label_style), _cell(ref_v, state), _cell(pl_v, state), _res(state)])
            card = Table(data, colWidths=[3*cm, 10.2*cm, 10.2*cm, 1.6*cm])
            style_cmds = [
                ("BACKGROUND",    (0,0), (-1,0),  colors.HexColor("#f0f4fa")),
                ("TOPPADDING",    (0,0), (-1,-1), 4),
                ("BOTTOMPADDING", (0,0), (-1,-1), 4),
                ("LEFTPADDING",   (0,0), (-1,-1), 6),
                ("INNERGRID",     (0,0), (-1,-1), 0.3, colors.HexColor("#dddddd")),
                ("BOX",           (0,0), (-1,-1), 1.5, border_color),
                ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
            ]
            for i, f in enumerate(fields, start=1):
                if f[3] == "diff":
                    style_cmds.append(("BACKGROUND", (1,i), (2,i), colors.HexColor("#fdecea")))
            card.setStyle(TableStyle(style_cmds))
            story.append(card)
            meta_txt = (
                f"<b>PO / Source:</b> {_esc(str(row.get('Source File','') or '—'))} &nbsp; "
                f"<b>Container:</b> {_esc(str(row.get('Current Container','') or '—'))} &nbsp; "
                f"<b>Job:</b> {_esc(str(row.get('Job Number','') or '—'))} &nbsp; "
                f"<b>Line:</b> {_esc(str(row.get('Line No','') or '—'))} &nbsp; "
                f"<b>Qty:</b> {_esc(str(row.get('Qty','') or '—'))} &nbsp; "
                f"<b>Reference status:</b> {_esc(str(row.get('Ref Status','') or '—'))}"
            )
            story.append(Spacer(1, 0.1*cm))
            story.append(Paragraph(meta_txt, cell_style))
            story.append(Paragraph(f"<b>Reason:</b> {_esc(str(row.get('Reason','') or '—'))}", cell_style))
            story.append(Spacer(1, 0.4*cm))

    make_card_section(red_df,    "🔴 RED — Confirmed SKU with mismatched HS or Description",
                      colors.HexColor("#c0392b"), colors.HexColor("#fdf0f0"))
    make_card_section(yellow_df, "🟡 YELLOW — SKU pending customs confirmation",
                      colors.HexColor("#b7770d"), colors.HexColor("#fffdf0"))
    make_card_section(orange_df, "🟠 ORANGE — SKU registered but not yet processed",
                      colors.HexColor("#d35400"), colors.HexColor("#fdf6ec"))
    make_card_section(hslen_df,  "HS LENGTH ABNORMAL - reference HS is not a 10-digit code",
                      colors.HexColor("#6c3483"), colors.HexColor("#f9f0ff"))

    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph(
        "Customs Risk Database — Vevor EU | Confidential — Internal Use Only",
        ParagraphStyle("footer", parent=styles["Normal"], fontSize=7,
                       textColor=colors.grey, alignment=1)
    ))
    doc.build(story)
    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════════════════
# DASHBOARD TAB
# ════════════════════════════════════════════════════════════════════════════════

def render_dashboard(df: pd.DataFrame):
    st.subheader("📊 Reference Database Dashboard")
    if len(df) == 0:
        st.info("No SKU data in the reference database yet.")
        return

    _st        = status_series(df)
    confirmed  = df[_st == "confirmed"]
    hs_only    = df[_st == "hs_only"]
    desc_only  = df[_st == "desc_only"]
    empty      = df[_st == "empty"]
    hs_abnormal = df[df["Customs Commodity Code"].apply(hs_len_abnormal)]

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c6.metric("⚠️ HS 位数异常", len(hs_abnormal))
    c1.metric("Total SKUs",                len(df))
    c2.metric("✅ Confirmed",              len(confirmed))
    c3.metric("🟡 HS only (pending desc)", len(hs_only))
    c4.metric("🟡 Desc only (pending HS)", len(desc_only))
    c5.metric("⬜ Not yet processed",      len(empty))

    st.markdown("---")

    col_left, col_right = st.columns(2)
    with col_left:
        st.markdown("**Top 15 HS Codes in Reference**")
        hs_counts = (
            df[df["Customs Commodity Code"].str.len() > 0]["Customs Commodity Code"]
            .value_counts().head(15).reset_index()
        )
        hs_counts.columns = ["HS Code", "Count"]
        st.dataframe(hs_counts, use_container_width=True, hide_index=True)

    with col_right:
        st.markdown("**SKU Status Breakdown**")
        status_df = pd.DataFrame([
            {"Status": "✅ Confirmed (desc + HS)",    "Count": len(confirmed)},
            {"Status": "🟡 HS only — pending desc",   "Count": len(hs_only)},
            {"Status": "🟡 Desc only — pending HS",   "Count": len(desc_only)},
            {"Status": "⬜ No data yet",              "Count": len(empty)},
        ])
        st.dataframe(status_df, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("**SKUs pending confirmation (HS only — no description)**")
    if len(hs_only) > 0:
        st.dataframe(hs_only[REF_COLUMNS], use_container_width=True, hide_index=True)
    else:
        st.success("No SKUs pending description confirmation.")

    st.markdown("**SKUs not yet processed (both fields empty)**")
    if len(empty) > 0:
        st.dataframe(empty[REF_COLUMNS].head(50), use_container_width=True, hide_index=True)
        if len(empty) > 50:
            st.caption(f"Showing 50 of {len(empty)} unprocessed SKUs.")
    else:
        st.success("All SKUs have at least partial data.")

    st.markdown(f"**⚠️ HS 位数异常 (HS is not {HS_EXPECTED_LEN} digits)**")
    if len(hs_abnormal) > 0:
        st.dataframe(hs_abnormal[REF_COLUMNS], use_container_width=True, hide_index=True)
        st.caption("Packing-list lines with these SKUs are reported as 'HS LENGTH ABNORMAL' until the reference HS is corrected.")
    else:
        st.success("All HS codes in the reference have 10 digits.")


# ════════════════════════════════════════════════════════════════════════════════
# ADMIN TAB  (unchanged)
# ════════════════════════════════════════════════════════════════════════════════

def render_admin():
    st.subheader("⚙️ Admin Panel")
    st.markdown("Manage user accounts. Only admins can access this panel.")
    users = load_users()

    st.markdown("**Current Users**")
    user_table = [{"Username": u, "Role": v["role"], "Display Name": v["display"]}
                  for u, v in users.items()]
    st.dataframe(pd.DataFrame(user_table), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("**Add / Update User**")
    with st.form("add_user_form"):
        new_username = st.text_input("Username")
        new_display  = st.text_input("Display Name")
        new_password = st.text_input("Password", type="password")
        new_role     = st.selectbox("Role", ["broker", "visitor", "admin"])
        save_btn     = st.form_submit_button("Save User")
        if save_btn:
            if not new_username or not new_password:
                st.error("Username and password are required.")
            else:
                users[new_username] = {
                    "password": _hash(new_password),
                    "role":     new_role,
                    "display":  new_display or new_username,
                }
                save_users(users)
                st.success(f"User '{new_username}' saved successfully.")
                st.rerun()

    st.markdown("---")
    st.markdown("**Delete User**")
    deletable = [u for u in users if u != st.session_state["username"]]
    if deletable:
        del_user = st.selectbox("Select user to delete", deletable)
        if st.button("Delete User", type="secondary"):
            del users[del_user]
            save_users(users)
            st.success(f"User '{del_user}' deleted.")
            st.rerun()
    else:
        st.info("No other users to delete.")

    st.markdown("---")
    st.markdown("**Change My Password**")
    with st.form("change_pw_form"):
        old_pw  = st.text_input("Current Password", type="password")
        new_pw  = st.text_input("New Password",     type="password")
        new_pw2 = st.text_input("Confirm New Password", type="password")
        change_btn = st.form_submit_button("Change Password")
        if change_btn:
            me = st.session_state["username"]
            if users[me]["password"] != _hash(old_pw):
                st.error("Current password is incorrect.")
            elif new_pw != new_pw2:
                st.error("New passwords do not match.")
            elif len(new_pw) < 6:
                st.error("Password must be at least 6 characters.")
            else:
                users[me]["password"] = _hash(new_pw)
                save_users(users)
                st.success("Password changed successfully.")

    st.markdown("---")
    st.markdown("**Database Backups**")
    st.caption(
        "A full copy of the database is saved automatically, at most once per day, "
        "right before any edit is written — named 'Customs Risk Database BACKUP YYYY-MM-DD'."
    )
    if st.button("Back up now"):
        with st.spinner("Creating backup..."):
            ok = backup_spreadsheet_if_due(force=True)
        if ok:
            st.success("Backup created.")
        else:
            st.error(
                "Backup failed: "
                f"{st.session_state.get('_last_backup_error', 'unknown error')}"
            )
        st.rerun()

    backups = list_backups()
    if backups:
        backup_table = [
            {"Name": b.get("name", ""), "Created": b.get("createdTime", "")[:19].replace("T", " ")}
            for b in backups[:15]
        ]
        st.dataframe(pd.DataFrame(backup_table), use_container_width=True, hide_index=True)
        if len(backups) > 15:
            st.caption(f"...and {len(backups) - 15} older backup(s), all still kept in Drive.")
    else:
        st.info(
            "No backups found yet — one is created automatically on the next save, "
            "or click 'Back up now' above."
        )


# ════════════════════════════════════════════════════════════════════════════════
# MAIN APP
# ════════════════════════════════════════════════════════════════════════════════

def main():
    st.set_page_config(
        page_title="Customs Risk Database — Vevor EU",
        layout="wide",
        page_icon="🛃"
    )

    # ── Sidebar ──────────────────────────────────────────────────────────────
    with st.sidebar:
        if LOGO_FILE.exists():
            st.image(str(LOGO_FILE), use_container_width=True)
        else:
            st.markdown("""
                <div style='text-align:center; padding: 0.5rem 0;'>
                    <span style='font-size:1.5rem; font-weight:700; color:#1a3c6e;'>VEVOR</span>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("---")
        role        = st.session_state.get("role", "")
        display     = st.session_state.get("display", "")
        role_colors = {"admin": "#c0392b", "broker": "#1a6e3c", "visitor": "#555555"}
        role_labels = {"admin": "Admin", "broker": "Broker", "visitor": "Visitor"}
        badge_color = role_colors.get(role, "#555555")
        badge_label = role_labels.get(role, role)
        st.markdown(f"""
            <div style='margin-bottom:0.3rem;'>
                <span style='font-size:1rem; font-weight:600;'>👤 {display}</span>
            </div>
            <div>
                <span style='background:{badge_color}; color:white; padding:2px 10px;
                             border-radius:12px; font-size:0.75rem; font-weight:600;'>
                    {badge_label}
                </span>
            </div>
        """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚪 Log Out", use_container_width=True):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()
        st.markdown("---")
        st.markdown("""
            <div style='font-size:0.72rem; color:#999; line-height:1.6;'>
                Customs Risk Database<br>
                Vevor EU — Internal Use Only<br><br>
                <span style='color:#bbb;'>Designed and created by<br>
                <b style='color:#888;'>Amelie — Vevor EU</b></span>
            </div>
        """, unsafe_allow_html=True)

    st.title("🛃 Customs Risk Database — Vevor EU")

    role      = st.session_state.get("role", "visitor")
    is_admin  = (role == "admin")
    is_broker = (role in ("admin", "broker"))

    tabs = ["📋 Risk Check", "📥 Manage Reference Data", "🗄️ Database", "📊 Dashboard"]
    if is_admin:
        tabs.append("⚙️ Admin")
    tab_objects = st.tabs(tabs)

    # ────────────────────────────────────────────────────────────────────────
    # TAB 1 — Risk Check
    # ────────────────────────────────────────────────────────────────────────
    with tab_objects[0]:
        st.subheader("1. Upload Packing List / Invoice for Risk Check")
        st.markdown("""
        Upload one or more Excel packing lists. Each line with a **SKU** will be checked
        against the reference database:

        | Alert | Meaning |
        |-------|---------|
        | 🔴 **RED** | SKU confirmed in reference — HS code or description differs |
        | 🟡 **YELLOW** | SKU inspected but awaiting customs confirmation — mismatch detected |
        | 🟠 **ORANGE** | SKU registered in reference but no data yet — flag for monitoring |
        | ⚠️ **HS 位数异常** | Reference HS is not a 10-digit code — reference data must be corrected first |
        """)

        doc_files = st.file_uploader(
            "Upload Packing List Excel files",
            type=["xlsx"],
            accept_multiple_files=True,
            key="document_upload"
        )

        if doc_files:
            all_doc_rows = []
            file_bytes = {}
            for file in doc_files:
                file_bytes[file.name] = file.getvalue()
                doc_df = normalize_document_file(file)
                doc_df["_File"] = file.name
                try:
                    _xls = pd.ExcelFile(file)
                    if len(_xls.sheet_names) > 1:
                        st.caption(f"📄 {file.name}: multi-sheet workbook — using sheet '{pick_packing_sheet(_xls)}'")
                except Exception:
                    pass
                if len(doc_df) > 0:
                    all_doc_rows.append(doc_df)
                else:
                    st.warning(f"No valid product / HS lines detected in {file.name}.")

            if all_doc_rows:
                document_df = pd.concat(all_doc_rows, ignore_index=True)
                with st.expander("📄 Preview parsed packing list lines"):
                    st.dataframe(document_df, use_container_width=True, hide_index=True)

                ref_df = load_reference_db()
                if len(ref_df) == 0:
                    st.warning("Reference database is empty. Please import data first (Tab 2).")
                else:
                    check_df = check_packing_list_against_reference(document_df, ref_df)

                    if len(check_df) == 0:
                        st.success("✅ No issues detected — all SKUs match the reference data.")
                    else:
                        red_df    = check_df[check_df["Severity"] == "RED"]
                        yellow_df = check_df[check_df["Severity"] == "YELLOW"]
                        orange_df = check_df[check_df["Severity"] == "ORANGE"]
                        hslen_df  = check_df[check_df["Severity"] == "HSLEN"]

                        col_r, col_y, col_o, col_h = st.columns(4)
                        col_r.error(f"🔴 RED: {len(red_df)} line(s)")
                        col_y.warning(f"🟡 YELLOW: {len(yellow_df)} line(s)")
                        col_o.info(f"🟠 ORANGE: {len(orange_df)} line(s)")
                        col_h.warning(f"⚠️ HS 位数异常: {len(hslen_df)} line(s)")

                        # ── Card renderer ──
                        def render_cards(df, color):
                            if color == "red":
                                border, bg, icon = "#c0392b", "#fff5f5", "🔴"
                            elif color == "yellow":
                                border, bg, icon = "#b7770d", "#fffdf0", "🟡"
                            elif color == "purple":
                                border, bg, icon = "#6c3483", "#f9f0ff", "⚠️"
                            else:
                                border, bg, icon = "#d35400", "#fff8f0", "🟠"

                            for _, row in df.iterrows():
                                st.markdown(build_card_html(row, border, bg, icon), unsafe_allow_html=True)

                        if len(red_df) > 0:
                            st.markdown("#### 🔴 Action Required — Confirmed Mismatches")
                            render_cards(red_df, "red")
                        if len(yellow_df) > 0:
                            st.markdown("#### 🟡 Pending Review — Awaiting Customs Confirmation")
                            render_cards(yellow_df, "yellow")
                        if len(orange_df) > 0:
                            st.markdown("#### 🟠 Log & Monitor — Unprocessed SKUs")
                            render_cards(orange_df, "orange")
                        if len(hslen_df) > 0:
                            st.markdown("#### ⚠️ HS 位数异常 — Reference HS is not a 10-digit code")
                            render_cards(hslen_df, "purple")

                        # PDF export
                        st.markdown("---")
                        st.markdown("#### 📄 Export Risk Report")
                        files_info = ", ".join([f.name for f in doc_files])
                        pdf_bytes  = build_pdf_report(check_df, files_info)
                        report_name = f"RiskReport_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                        st.download_button(
                            label="⬇️ Download PDF Report",
                            data=pdf_bytes,
                            file_name=report_name,
                            mime="application/pdf",
                            type="primary",
                            use_container_width=False,
                        )

                        # ── Corrected packing list: HS taken from the reference, changed cells yellow ──
                        st.markdown("---")
                        st.markdown("#### ✏️ Corrected Packing List")
                        st.caption("Replaces the HS in the packing list with the reference HS and marks each changed cell "
                                   "yellow. Everything else in the file stays as it was. Untick lines you do not want to change.")
                        fix_df = check_df[(check_df["Cmp HS"] == "diff") & (check_df["Severity"] != "HSLEN")]
                        fix_df = fix_df[fix_df["HS Row"].astype(str) != ""]
                        if len(fix_df) == 0:
                            st.info("No HS differences that can be corrected automatically.")
                        for fname in fix_df["File Name"].unique() if len(fix_df) else []:
                            sub = fix_df[fix_df["File Name"] == fname].reset_index(drop=True)
                            st.markdown(f"**{fname}** — {len(sub)} line(s) with a different HS")
                            table = pd.DataFrame({
                                "Apply":                  True,
                                "Line":                   sub["Line No"].values,
                                "SKU":                    sub["SKU"].values,
                                "Description":            sub["PL Description"].values,
                                "HS in packing list":     sub["PL HS Code"].values,
                                "HS in reference (new)":  sub["Ref HS Code"].values,
                                "Duty (ref)":             sub["Ref Duty Rate"].values,
                            })
                            edited_fix = st.data_editor(
                                table, hide_index=True, use_container_width=True, key=f"fix_{fname}",
                                disabled=[c for c in table.columns if c != "Apply"],
                            )
                            chosen = sub[edited_fix["Apply"].values]
                            if len(chosen) == 0:
                                st.caption("No line selected.")
                                continue
                            try:
                                new_bytes, applied, skipped = apply_hs_corrections(
                                    file_bytes[fname], chosen.iloc[0]["Sheet"],
                                    [(int(r["HS Row"]), r["HS Col"], r["Ref HS Code"]) for _, r in chosen.iterrows()],
                                )
                            except Exception as e:
                                st.error(f"Could not edit {fname}: {e}")
                                continue
                            if applied:
                                st.caption("Changed cells: " + ", ".join(f"{c} ({o} → {n})" for c, o, n, _ in applied))
                            for cell, why in skipped:
                                st.warning(f"Cell {cell} left unchanged — {why}")
                            if len(pd.ExcelFile(io.BytesIO(file_bytes[fname])).sheet_names) > 1:
                                st.caption(f"Only the '{chosen.iloc[0]['Sheet']}' sheet is changed; other sheets in this workbook "
                                           "(e.g. invoice, contract) are not modified.")
                            if applied:
                                stem = fname.rsplit(".", 1)[0]
                                st.download_button(
                                    label=f"⬇️ Download corrected packing list — {fname}",
                                    data=new_bytes,
                                    file_name=f"{stem}_HS_corrected.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key=f"dl_fix_{fname}",
                                )

    # ────────────────────────────────────────────────────────────────────────
    # TAB 2 — Manage Reference Data
    # ────────────────────────────────────────────────────────────────────────
    with tab_objects[1]:
        if not is_broker:
            st.info("🔒 Visitors cannot modify the reference database. Please contact your broker or admin.")
        else:
            st.subheader("2a. Import Reference Data from Excel")
            st.markdown("""
            Upload an Excel file with columns: **SKU**, **Corrected DESCRIPTION** (or *Description*),
            **Customs Commodity Code** (or *Hscode*) and optionally **Duty Rate** (or *进口税率*).
            Existing SKUs will be updated; new SKUs will be added. HS codes are imported exactly as given —
            codes that are not 10 digits are flagged as **HS 位数异常**.
            """)
            uploaded_ref = st.file_uploader(
                "Upload Reference Excel (.xlsx)", type=["xlsx"], key="ref_import_upload"
            )
            if uploaded_ref is not None:
                try:
                    imp_df = pd.read_excel(uploaded_ref, dtype=str)
                    imp_df.columns = [str(c).strip() for c in imp_df.columns]

                    # Try to auto-map columns
                    col_map = {}
                    for col in imp_df.columns:
                        cu = col.upper()
                        if "SKU" in cu and "SKU" not in col_map:
                            col_map["SKU"] = col
                        elif "DESCRIPTION" in cu and "Corrected DESCRIPTION" not in col_map:
                            col_map["Corrected DESCRIPTION"] = col
                        elif ("税率" in col or "DUTY" in cu or "RATE" in cu) and "Duty Rate" not in col_map:
                            col_map["Duty Rate"] = col
                        elif ("HS" in cu or "COMMODITY" in cu or "CODE" in cu) and "Customs Commodity Code" not in col_map:
                            col_map["Customs Commodity Code"] = col

                    if not all(k in col_map for k in ("SKU", "Corrected DESCRIPTION", "Customs Commodity Code")):
                        # Show manual mapping
                        st.warning("Could not auto-detect all columns. Please map them manually:")
                        all_cols = ["(skip)"] + list(imp_df.columns)
                        sku_col  = st.selectbox("SKU column",                all_cols, key="map_sku")
                        desc_col = st.selectbox("Corrected DESCRIPTION col", all_cols, key="map_desc")
                        hs_col   = st.selectbox("Customs Commodity Code col",all_cols, key="map_hs")
                        duty_col = st.selectbox("Duty Rate col",             all_cols, key="map_duty")
                        col_map = {}
                        if sku_col != "(skip)":
                            col_map["SKU"] = sku_col
                            if desc_col != "(skip)": col_map["Corrected DESCRIPTION"]  = desc_col
                            if hs_col   != "(skip)": col_map["Customs Commodity Code"] = hs_col
                            if duty_col != "(skip)": col_map["Duty Rate"]              = duty_col

                    if "SKU" in col_map:
                        preview = pd.DataFrame()
                        preview["SKU"] = imp_df[col_map["SKU"]].apply(clean_text)
                        preview["Corrected DESCRIPTION"] = (
                            imp_df[col_map["Corrected DESCRIPTION"]].apply(clean_text)
                            if "Corrected DESCRIPTION" in col_map else ""
                        )
                        preview["Customs Commodity Code"] = (
                            imp_df[col_map["Customs Commodity Code"]].apply(clean_hs)
                            if "Customs Commodity Code" in col_map else ""
                        )
                        preview["Duty Rate"] = (
                            imp_df[col_map["Duty Rate"]].apply(format_duty_rate)
                            if "Duty Rate" in col_map else ""
                        )
                        preview = preview[preview["SKU"] != ""]
                        n_abn = int(preview["Customs Commodity Code"].apply(hs_len_abnormal).sum())
                        st.write(f"Preview — {len(preview)} valid rows detected:")
                        if n_abn:
                            st.warning(f"⚠️ HS 位数异常: {n_abn} row(s) have an HS that is not {HS_EXPECTED_LEN} digits. "
                                       "They will be imported as-is and flagged in the Dashboard and Risk Check.")
                        st.dataframe(preview.head(20), use_container_width=True, hide_index=True)

                        if st.button("📥 Import & Merge into Reference Database"):
                            mapped_cols = [c for c in REF_COLUMNS[1:] if c in col_map]
                            existing = load_reference_db().copy()
                            new = preview.drop_duplicates("SKU", keep="last").copy()
                            new["_k"] = new["SKU"].str.upper()
                            existing["_k"] = existing["SKU"].str.upper()
                            is_upd = existing["_k"].isin(set(new["_k"]))
                            new_by_key = new.set_index("_k")
                            for c in mapped_cols:
                                existing.loc[is_upd, c] = existing.loc[is_upd, "_k"].map(new_by_key[c])
                            to_add = new[~new["_k"].isin(set(existing["_k"]))]
                            updated = int(is_upd.sum())
                            added   = len(to_add)
                            existing = pd.concat(
                                [existing[REF_COLUMNS], to_add[REF_COLUMNS]], ignore_index=True
                            )
                            save_reference_db(existing)
                            st.success(f"Done — {added} new SKUs added, {updated} existing SKUs updated.")
                            st.rerun()
                except Exception as e:
                    st.error(f"Error reading file: {e}")
            st.markdown("---")

            st.subheader("2b. Add / Update a Single SKU")
            with st.form("add_sku_form"):
                new_sku  = st.text_input("SKU *")
                new_desc = st.text_input("Corrected DESCRIPTION")
                new_hs   = st.text_input("Customs Commodity Code (10-digit CN code)")
                new_duty = st.text_input("Duty Rate (e.g. 2,7%)")
                save_btn = st.form_submit_button("Save SKU")
                if save_btn:
                    if not new_sku.strip():
                        st.error("SKU is required.")
                    else:
                        existing = load_reference_db()
                        key = new_sku.strip().upper()
                        mask = existing["SKU"].str.upper() == key
                        if mask.any():
                            existing.loc[mask, "Corrected DESCRIPTION"]  = clean_text(new_desc)
                            existing.loc[mask, "Customs Commodity Code"] = clean_hs(new_hs)
                            if new_duty.strip():
                                existing.loc[mask, "Duty Rate"] = format_duty_rate(new_duty)
                            save_reference_db(existing)
                            st.success(f"SKU '{new_sku.strip()}' updated.")
                        else:
                            new_row = {
                                "SKU":                    new_sku.strip(),
                                "Corrected DESCRIPTION":  clean_text(new_desc),
                                "Customs Commodity Code": clean_hs(new_hs),
                                "Duty Rate":              format_duty_rate(new_duty),
                            }
                            existing = pd.concat(
                                [existing, pd.DataFrame([new_row])], ignore_index=True
                            )
                            save_reference_db(existing)
                            st.success(f"SKU '{new_sku.strip()}' added to reference database.")

    # ────────────────────────────────────────────────────────────────────────
    # TAB 3 — Database viewer / editor
    # ────────────────────────────────────────────────────────────────────────
    with tab_objects[2]:
        st.subheader("3. Reference Database")
        ref_df = load_reference_db()

        if len(ref_df) == 0:
            st.info("No data in reference database yet.")
        else:
            # Filters
            fc1, fc2, fc3 = st.columns(3)
            with fc1:
                filter_status = st.selectbox("Filter by Status", [
                    "All", "confirmed", "hs_only", "desc_only", "empty", "hs_len_abnormal"
                ])
            with fc2:
                filter_sku = st.text_input("Filter by SKU (contains)")
            with fc3:
                filter_hs = st.text_input("Filter by HS Code (contains)")

            filtered = ref_df.copy()
            filtered["_status"] = status_series(filtered)

            if filter_status == "hs_len_abnormal":
                filtered = filtered[filtered["Customs Commodity Code"].apply(hs_len_abnormal)]
            elif filter_status != "All":
                filtered = filtered[filtered["_status"] == filter_status]
            if filter_sku:
                filtered = filtered[filtered["SKU"].str.contains(filter_sku.strip(), case=False, na=False)]
            if filter_hs:
                filtered = filtered[filtered["Customs Commodity Code"].str.contains(filter_hs.strip(), case=False, na=False)]

            display_df = filtered[REF_COLUMNS + ["_status"]].copy()
            display_df = display_df.rename(columns={"_status": "Status"})
            display_df["HS Check"] = display_df["Customs Commodity Code"].apply(
                lambda h: "⚠️ HS 位数异常" if hs_len_abnormal(h) else "")
            MAX_ROWS = 1000   # the reference has ~45k rows — never render them all in an editor
            st.markdown(f"**{len(filtered)}** of **{len(ref_df)}** SKUs match the filters"
                        + (f" — showing the first {MAX_ROWS}. Use the filters to narrow down." if len(filtered) > MAX_ROWS else ""))
            display_df = display_df.head(MAX_ROWS)

            if is_broker:
                display_df.insert(0, "Delete", False)
                edited = st.data_editor(
                    display_df,
                    use_container_width=True,
                    num_rows="fixed",
                    hide_index=True,
                    column_config={
                        "Delete": st.column_config.CheckboxColumn("Delete"),
                        "Status": st.column_config.TextColumn("Status", disabled=True),
                        "HS Check": st.column_config.TextColumn("HS Check", disabled=True),
                        "SKU": st.column_config.TextColumn("SKU", disabled=True),
                    }
                )
                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    if st.button("💾 Save Changes", use_container_width=True):
                        edit_cols = ["Corrected DESCRIPTION", "Customs Commodity Code", "Duty Rate"]
                        kept = edited[edited["Delete"] == False]
                        deleted_idx = edited.index[edited["Delete"] == True]
                        upd = ref_df.copy()
                        upd.loc[kept.index, "Corrected DESCRIPTION"]  = kept["Corrected DESCRIPTION"].apply(clean_text).values
                        upd.loc[kept.index, "Customs Commodity Code"] = kept["Customs Commodity Code"].apply(clean_hs).values
                        upd.loc[kept.index, "Duty Rate"]              = kept["Duty Rate"].apply(format_duty_rate).values
                        final_df = upd.drop(index=deleted_idx)
                        save_reference_db(final_df)
                        st.success("Changes saved.")
                        st.rerun()
                with col_b2:
                    if st.button(f"🗑️ Delete ALL {len(filtered)} filtered rows",
                                 type="secondary", use_container_width=True):
                        filtered_skus = set(filtered["SKU"].str.upper().tolist())
                        final_df = ref_df[~ref_df["SKU"].str.upper().isin(filtered_skus)]
                        save_reference_db(final_df)
                        st.success(f"Deleted {len(filtered)} rows.")
                        st.rerun()
            else:
                st.info("👁️ Visitor mode — read only.")
                st.dataframe(display_df, use_container_width=True, hide_index=True)

    # ────────────────────────────────────────────────────────────────────────
    # TAB 4 — Dashboard
    # ────────────────────────────────────────────────────────────────────────
    with tab_objects[3]:
        render_dashboard(load_reference_db())

    # ────────────────────────────────────────────────────────────────────────
    # TAB 5 — Admin
    # ────────────────────────────────────────────────────────────────────────
    if is_admin:
        with tab_objects[4]:
            render_admin()


# ════════════════════════════════════════════════════════════════════════════════
# Entry point
# ════════════════════════════════════════════════════════════════════════════════
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    login_screen()
else:
    main()
