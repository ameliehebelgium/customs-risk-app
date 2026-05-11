import streamlit as st
import pandas as pd
from datetime import date, datetime
from pathlib import Path
import hashlib
import json
import io

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
SHEET_NAME = "Feuille 1"

@st.cache_resource(ttl=30)
def _get_gsheet():
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
    gc = gspread.authorize(creds)
    return gc.open_by_key(SHEET_ID).worksheet(SHEET_NAME)

# ─── File paths ───────────────────────────────────────────────────────────────
USERS_FILE = Path("users.json")
LOGO_FILE  = Path("vevor_logo.png")

# ─── Column definitions ───────────────────────────────────────────────────────
COLUMNS = [
    "Risk ID", "Input Date", "CC Date", "Inspection Date", "Container No", "MRN",
    "BL Number", "Job Number", "Inspector",
    "Product Name", "Product Alias", "Declaration Description",
    "Old HS", "Corrected HS", "Duty Before", "Duty After",
    "Findings Type", "Root Cause", "Risk Reason", "Customs Comment",
    "Status", "Notes",
]
DOC_COLUMNS = ["Source File", "Current Container", "Line No", "Product Description", "HS Code", "Qty"]


# ════════════════════════════════════════════════════════════════════════════════
# AUTH HELPERS
# ════════════════════════════════════════════════════════════════════════════════

def _hash(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()


def load_users() -> dict:
    if USERS_FILE.exists():
        with open(USERS_FILE) as f:
            return json.load(f)
    # Default users on first run
    default = {
        "admin":   {"password": _hash("admin123"),   "role": "admin",   "display": "Admin"},
        "broker1": {"password": _hash("broker2026"), "role": "broker",  "display": "Broker User"},
        "visitor": {"password": _hash("visit2026"),  "role": "visitor", "display": "Visitor"},
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
            st.session_state["logged_in"]   = True
            st.session_state["username"]    = username.strip()
            st.session_state["role"]        = user["role"]
            st.session_state["display"]     = user["display"]
            st.rerun()
        else:
            st.error("Incorrect username or password.")

    st.markdown("""
        <div style='text-align:center; margin-top:2rem; color:#aaa; font-size:0.8rem;'>
            
            
        </div>
    """, unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════════
# DATA HELPERS  (unchanged logic from original)
# ════════════════════════════════════════════════════════════════════════════════

def clean_text(value):
    if pd.isna(value) or value is None:
        return ""
    text = str(value).strip()
    if text.lower() in ["nan", "none", "nat"]:
        return ""
    return text


def clean_hs(value):
    value = clean_text(value)
    if value.endswith(".0"):
        value = value[:-2]
    return value.replace(" ", "")


def format_date(value):
    text = clean_text(value)
    if text == "":
        return ""
    try:
        return pd.to_datetime(text).strftime("%Y-%m-%d")
    except Exception:
        return text[:10]


def format_duty_rate(value):
    text = clean_text(value)
    if text == "":
        return ""
    text = text.replace("%", "").replace(",", ".").strip()
    try:
        number = float(text)
        if 0 < number < 1:
            number *= 100
        result = str(int(number)) if number == int(number) else f"{number:.2f}".rstrip("0").rstrip(".")
        return result.replace(".", ",") + "%"
    except Exception:
        return clean_text(value)


def load_database() -> pd.DataFrame:
    """Load risk database from Google Sheets."""
    try:
        ws = _get_gsheet()
        data = ws.get_all_records(default_blank="")
        if not data:
            return pd.DataFrame(columns=COLUMNS)
        df = pd.DataFrame(data, dtype=str)
        for col in COLUMNS:
            if col not in df.columns:
                df[col] = ""
        df = df[COLUMNS].fillna("")
        for col in ["Input Date", "CC Date", "Inspection Date"]:
            df[col] = df[col].apply(format_date)
        df["Old HS"]       = df["Old HS"].apply(clean_hs)
        df["Corrected HS"] = df["Corrected HS"].apply(clean_hs)
        df["Duty Before"]  = df["Duty Before"].apply(format_duty_rate)
        df["Duty After"]   = df["Duty After"].apply(format_duty_rate)
        return df
    except Exception as e:
        st.error(f"Error loading database: {e}")
        return pd.DataFrame(columns=COLUMNS)


def save_database(df: pd.DataFrame):
    """Save full risk database back to Google Sheets (overwrite)."""
    try:
        for col in ["Input Date", "CC Date", "Inspection Date"]:
            df[col] = df[col].apply(format_date)
        df = df.fillna("").astype(str)
        ws = _get_gsheet()
        # Clear and rewrite
        ws.clear()
        ws.update([COLUMNS] + df[COLUMNS].values.tolist())
    except Exception as e:
        st.error(f"Error saving database: {e}")


def generate_risk_id(df):
    year = date.today().year
    prefix = f"RSK-{year}-"
    existing = df["Risk ID"].dropna().astype(str).tolist() if "Risk ID" in df.columns else []
    numbers = []
    for rid in existing:
        if rid.startswith(prefix):
            try:
                numbers.append(int(rid.replace(prefix, "")))
            except Exception:
                pass
    return f"{prefix}{(max(numbers)+1 if numbers else 1):04d}"


def make_duplicate_key(row):
    return (
        clean_text(row.get("Container No", "")).upper(),
        clean_text(row.get("MRN", "")).upper(),
        clean_text(row.get("Product Name", "")).upper(),
        clean_hs(row.get("Old HS", "")),
        clean_hs(row.get("Corrected HS", "")),
    )


def find_header_row(uploaded_file):
    raw = pd.read_excel(uploaded_file, header=None, dtype=str)
    for i in range(len(raw)):
        row_text = " ".join([clean_text(x) for x in raw.iloc[i].tolist()])
        if "Product Name" in row_text and "OLD HS" in row_text:
            return i
    return 0


def normalize_import_file(uploaded_file):
    header_row = find_header_row(uploaded_file)
    df = pd.read_excel(uploaded_file, header=header_row, dtype=str)
    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
    mapping = {
        "Declaration Date": "CC Date",
        "Inspection Date": "Inspection Date",
        "Container No.": "Container No",
        "MRN (Declaration Ref)": "MRN",
        "BL number": "BL Number",
        "BL Number": "BL Number",
        "Job number": "Job Number",
        "Job Number": "Job Number",
        "Inspector (Customs Agent)": "Inspector",
        "Inspector": "Inspector",
        "Product Name (EN)": "Product Name",
        "Declaration Description (as filed)": "Declaration Description",
        "OLD HS Code (as declared)": "Old HS",
        "CORRECTED HS Code (by customs)": "Corrected HS",
        "Duty Rate BEFORE (%)": "Duty Before",
        "Duty Rate AFTER (%)": "Duty After",
        "Findings Type (see dropdown)": "Findings Type",
        "Root Cause": "Root Cause",
    }
    df = df.rename(columns=mapping)
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUMNS]
    for col in COLUMNS:
        df[col] = df[col].apply(clean_text)
    for col in ["Input Date", "CC Date", "Inspection Date"]:
        df[col] = df[col].apply(format_date)
    df["Old HS"]       = df["Old HS"].apply(clean_hs)
    df["Corrected HS"] = df["Corrected HS"].apply(clean_hs)
    df["Duty Before"]  = df["Duty Before"].apply(format_duty_rate)
    df["Duty After"]   = df["Duty After"].apply(format_duty_rate)
    df = df[(df["Product Name"] != "") & (df["Old HS"] != "") & (df["Corrected HS"] != "")]
    return df


def find_document_header_row(uploaded_file):
    raw = pd.read_excel(uploaded_file, header=None, dtype=str)
    for i in range(len(raw)):
        row_text = " ".join([clean_text(x).upper() for x in raw.iloc[i].tolist()])
        if ("DESCRIPTION" in row_text or "ITEM" in row_text) and ("HS" in row_text or "HTS" in row_text):
            return i
    return 0


def _looks_like_container(value: str) -> bool:
    """Container numbers are typically 4 letters + 7 digits, e.g. MSDU8622145."""
    import re
    return bool(re.match(r'^[A-Z]{3,4}\d{6,8}$', value.strip().upper()))


def extract_containers_from_header(raw_df) -> list:
    """
    Scan the first 20 rows for 'Container NO.' and extract ALL container numbers
    found in that row (handles multi-container strings like 'MSDU8622145+MEDU4532909').
    Returns a list of container number strings (may be empty).
    """
    import re
    containers = []
    for i in range(min(len(raw_df), 20)):
        row_values = [clean_text(x) for x in raw_df.iloc[i].tolist()]
        row_text   = " ".join(row_values).upper()
        if "CONTAINER" not in row_text:
            continue

        # Collect all cell values from this row
        full_text = " ".join(row_values)

        # Case 1: container number embedded in same cell as label
        # e.g. "Container NO.: MSDU8622145+MEDU4532909"
        m = re.search(r'Container\s*NO\.?\s*:?\s*([A-Z0-9+\s]+)', full_text, re.IGNORECASE)
        if m:
            raw_ids = re.split(r'[+\s]+', m.group(1).strip())
            found = [x.strip().upper() for x in raw_ids if _looks_like_container(x)]
            if found:
                containers.extend(found)
                break

        # Case 2: label in one cell, value(s) in adjacent cell
        for j, value in enumerate(row_values):
            if "CONTAINER" in value.upper():
                # Check next cells for container-like values
                for k in range(j + 1, min(j + 4, len(row_values))):
                    candidate = clean_text(row_values[k])
                    # May be "CSNU7617334+OOCU8995319"
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


def normalize_document_file(uploaded_file):
    raw = pd.read_excel(uploaded_file, header=None, dtype=str)

    # Extract container numbers from file header
    header_containers = extract_containers_from_header(raw)
    fallback_container = "+".join(header_containers) if header_containers else ""

    header_row = find_document_header_row(uploaded_file)
    df = pd.read_excel(uploaded_file, header=header_row, dtype=str)
    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]

    product_col = hs_col = qty_col = container_col = None
    for col in df.columns:
        col_upper = col.upper()
        if product_col is None and ("DESCRIPTION" in col_upper or "ITEM" in col_upper or "PRODUCT" in col_upper):
            product_col = col
        if hs_col is None and ("HS" in col_upper or "HTS" in col_upper):
            hs_col = col
        if qty_col is None and ("QTY" in col_upper or "QUANTITY" in col_upper):
            qty_col = col
        # Detect per-row container column (e.g. "对应柜号")
        if container_col is None and ("柜号" in col or "CONTAINER" in col_upper):
            container_col = col

    if product_col is None or hs_col is None:
        return pd.DataFrame(columns=DOC_COLUMNS)

    result = pd.DataFrame()
    result["Source File"]         = uploaded_file.name
    result["Line No"]             = range(1, len(df) + 1)
    result["Product Description"] = df[product_col].apply(clean_text)
    result["HS Code"]             = df[hs_col].apply(clean_hs)
    result["Qty"]                 = df[qty_col].apply(clean_text) if qty_col else ""

    # Per-row container column takes priority; fall back to file-level header value
    if container_col:
        result["Current Container"] = df[container_col].apply(
            lambda x: clean_text(x) if clean_text(x) else fallback_container
        )
    else:
        result["Current Container"] = fallback_container

    result = result[(result["Product Description"] != "") & (result["HS Code"] != "")]
    return result[DOC_COLUMNS]


def split_aliases(value):
    text = clean_text(value)
    if text == "":
        return []
    return [x.strip().upper() for x in text.replace(",", ";").split(";") if x.strip()]


def product_matches_risk(product_description, risk_row):
    product    = clean_text(product_description).upper()
    candidates = [
        clean_text(risk_row.get("Product Name", "")).upper(),
        clean_text(risk_row.get("Declaration Description", "")).upper(),
    ]
    candidates += split_aliases(risk_row.get("Product Alias", ""))
    candidates  = [x for x in candidates if x]
    for candidate in candidates:
        if candidate in product or product in candidate:
            return True
    return False


def check_documents_against_risks(doc_df, risk_df):
    results = []
    for _, doc_row in doc_df.iterrows():
        product = clean_text(doc_row["Product Description"])
        hs      = clean_hs(doc_row["HS Code"])
        for _, risk_row in risk_df.iterrows():
            old_hs        = clean_hs(risk_row["Old HS"])
            product_match = product_matches_risk(product, risk_row)
            hs_match      = (hs == old_hs and old_hs != "")
            if product_match and hs_match:
                severity = "RED"
                message  = "Both old HS and product name match a known risk. High probability of repeated classification error."
            elif hs_match:
                severity = "ORANGE"
                message  = "HS code matches a historical correction. Please verify the product manually."
            elif product_match:
                severity = "ORANGE"
                message  = "Product description matches a historical risk. Please confirm whether the HS has been updated."
            else:
                continue
            results.append({
                "Severity":                  severity,
                "Action Required":           "STOP & REVIEW" if severity == "RED" else "MANUAL REVIEW",
                "Current Container":         doc_row["Current Container"],
                "Current Product":           product,
                "Current HS":                hs,
                "Qty":                       doc_row["Qty"],
                "Source File":               doc_row["Source File"],
                "Line No":                   doc_row["Line No"],
                "Corrected HS":              risk_row["Corrected HS"],
                "Matched Risk ID":           risk_row["Risk ID"],
                "Previous Inspection Date":  risk_row["Inspection Date"],
                "Previous Container":        risk_row["Container No"],
                "Previous MRN":              risk_row["MRN"],
                "Historical Product":        risk_row["Product Name"],
                "Old HS Used Before":        old_hs,
                "Duty Before":               risk_row["Duty Before"],
                "Duty After":                risk_row["Duty After"],
                "Message":                   message,
                "Customs Comment":           risk_row["Customs Comment"],
                "Risk Reason":               risk_row["Risk Reason"],
            })
    return pd.DataFrame(results)


# ════════════════════════════════════════════════════════════════════════════════
# PDF EXPORT
# ════════════════════════════════════════════════════════════════════════════════

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

    # Header
    story.append(Paragraph("Customs Risk Check Report", title_style))
    story.append(Paragraph(
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')} &nbsp;|&nbsp; "
        f"Files checked: {doc_files_info} &nbsp;|&nbsp; "
        f"Total alerts: {len(check_df)}",
        sub_style
    ))
    story.append(Spacer(1, 0.4*cm))

    red_df    = check_df[check_df["Severity"] == "RED"]
    orange_df = check_df[check_df["Severity"] == "ORANGE"]

    # Summary box
    summary_data = [
        ["🚨 RED ALERTS", "⚠️ ORANGE WARNINGS", "Total Lines Flagged"],
        [str(len(red_df)), str(len(orange_df)), str(len(check_df))],
    ]
    summary_table = Table(summary_data, colWidths=[7*cm, 7*cm, 7*cm])
    summary_table.setStyle(TableStyle([
        ("BACKGROUND",  (0,0), (-1,0), colors.HexColor("#1a3c6e")),
        ("TEXTCOLOR",   (0,0), (-1,0), colors.white),
        ("BACKGROUND",  (0,1), (0,1),  colors.HexColor("#fde8e8")),
        ("BACKGROUND",  (1,1), (1,1),  colors.HexColor("#fff3e0")),
        ("BACKGROUND",  (2,1), (2,1),  colors.HexColor("#e8f0fe")),
        ("FONTSIZE",    (0,1), (-1,1), 18),
        ("ALIGN",       (0,0), (-1,-1), "CENTER"),
        ("VALIGN",      (0,0), (-1,-1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0,0), (-1,-1), [None, None]),
        ("BOX",         (0,0), (-1,-1), 0.5, colors.HexColor("#cccccc")),
        ("INNERGRID",   (0,0), (-1,-1), 0.5, colors.HexColor("#cccccc")),
        ("TOPPADDING",  (0,0), (-1,-1), 8),
        ("BOTTOMPADDING",(0,0),(-1,-1), 8),
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 0.5*cm))

    # Detail columns
    display_cols = [
        "Severity", "Action Required", "Current Container",
        "Current Product", "Current HS", "Corrected HS",
        "Matched Risk ID", "Previous Inspection Date", "Previous Container", "Message"
    ]
    col_widths = [1.8*cm, 2.8*cm, 3.2*cm, 4.5*cm, 2.2*cm, 2.2*cm, 2.8*cm, 2.6*cm, 2.8*cm, 4.5*cm]

    def make_section(section_df, heading, bg_header, bg_row):
        if len(section_df) == 0:
            return
        story.append(Paragraph(heading, h2_style))
        header_row = [Paragraph(f"<b>{c}</b>", cell_style) for c in display_cols]
        rows = [header_row]
        for _, row in section_df.iterrows():
            rows.append([Paragraph(str(row.get(c, "")), cell_style) for c in display_cols])
        t = Table(rows, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0), bg_header),
            ("TEXTCOLOR",     (0, 0), (-1, 0), colors.white),
            ("ROWBACKGROUNDS",(0, 1), (-1,-1), [colors.white, bg_row]),
            ("BOX",           (0, 0), (-1,-1), 0.4, colors.HexColor("#bbbbbb")),
            ("INNERGRID",     (0, 0), (-1,-1), 0.3, colors.HexColor("#dddddd")),
            ("VALIGN",        (0, 0), (-1,-1), "TOP"),
            ("TOPPADDING",    (0, 0), (-1,-1), 4),
            ("BOTTOMPADDING", (0, 0), (-1,-1), 4),
        ]))
        story.append(t)
        story.append(Spacer(1, 0.3*cm))

    make_section(red_df,    "🚨 HIGH RISK — Immediate Action Required",
                 colors.HexColor("#c0392b"), colors.HexColor("#fdf0f0"))
    make_section(orange_df, "⚠️  MANUAL REVIEW REQUIRED",
                 colors.HexColor("#d35400"), colors.HexColor("#fdf6ec"))

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
    st.subheader("📊 Risk Database Dashboard")

    if len(df) == 0:
        st.info("No risk cases in the database yet.")
        return

    active_df     = df[df["Status"] == "active"]
    monitoring_df = df[df["Status"] == "monitoring"]
    solved_df     = df[df["Status"] == "solved"]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Cases",      len(df))
    c2.metric("🔴 Active",        len(active_df))
    c3.metric("🟡 Monitoring",    len(monitoring_df))
    c4.metric("🟢 Solved",        len(solved_df))

    st.markdown("---")

    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown("**Top 10 Old HS Codes with Most Risk Cases**")
        hs_counts = (
            df[df["Old HS"] != ""]["Old HS"]
            .value_counts()
            .head(10)
            .reset_index()
        )
        hs_counts.columns = ["Old HS Code", "Count"]
        st.dataframe(hs_counts, use_container_width=True, hide_index=True)

    with col_right:
        st.markdown("**Cases by Findings Type**")
        ft_counts = (
            df[df["Findings Type"] != ""]["Findings Type"]
            .value_counts()
            .reset_index()
        )
        ft_counts.columns = ["Findings Type", "Count"]
        st.dataframe(ft_counts, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("**Duty Rate Impact — Before vs After (active cases)**")
    duty_df = active_df[
        (active_df["Duty Before"] != "") & (active_df["Duty After"] != "")
    ][["Risk ID", "Product Name", "Old HS", "Corrected HS", "Duty Before", "Duty After"]].copy()
    if len(duty_df) > 0:
        st.dataframe(duty_df, use_container_width=True, hide_index=True)
    else:
        st.info("No duty rate data available for active cases.")

    st.markdown("---")
    st.markdown("**Top Root Causes**")
    rc_counts = (
        df[df["Root Cause"] != ""]["Root Cause"]
        .value_counts()
        .head(8)
        .reset_index()
    )
    rc_counts.columns = ["Root Cause", "Count"]
    st.dataframe(rc_counts, use_container_width=True, hide_index=True)

    # Timeline
    st.markdown("---")
    st.markdown("**Cases by Input Month**")
    timeline_df = df[df["Input Date"] != ""].copy()
    if len(timeline_df) > 0:
        timeline_df["Month"] = pd.to_datetime(
            timeline_df["Input Date"], errors="coerce"
        ).dt.to_period("M").astype(str)
        monthly = (
            timeline_df.groupby("Month")
            .size()
            .reset_index(name="Cases")
            .sort_values("Month")
        )
        st.bar_chart(monthly.set_index("Month")["Cases"])


# ════════════════════════════════════════════════════════════════════════════════
# ADMIN TAB
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

    role     = st.session_state.get("role", "visitor")
    is_admin  = (role == "admin")
    is_broker = (role in ("admin", "broker"))   # broker + admin can upload & edit

    tabs = ["📋 Risk Check", "📥 Import / Add Cases", "🗄️ Database", "📊 Dashboard"]
    if is_admin:
        tabs.append("⚙️ Admin")

    tab_objects = st.tabs(tabs)

    # ────────────────────────────────────────────────────────────────────────
    # TAB 1 — Risk Check
    # ────────────────────────────────────────────────────────────────────────
    with tab_objects[0]:
        st.subheader("1. Upload Invoice / Packing List for Risk Check")

        doc_files = st.file_uploader(
            "Upload Invoice or Packing List Excel files",
            type=["xlsx"],
            accept_multiple_files=True,
            key="document_upload"
        )

        if doc_files:
            all_doc_rows = []
            for file in doc_files:
                doc_df = normalize_document_file(file)
                if len(doc_df) > 0:
                    all_doc_rows.append(doc_df)
                else:
                    st.warning(f"No valid product / HS lines detected in {file.name}.")

            if all_doc_rows:
                document_df = pd.concat(all_doc_rows, ignore_index=True)
                with st.expander("Preview uploaded document lines"):
                    st.dataframe(document_df, use_container_width=True)

                risk_df = load_database()
                if len(risk_df) == 0:
                    st.warning("Risk database is empty. Please import or add risk cases first.")
                else:
                    check_df = check_documents_against_risks(document_df, risk_df)

                    if len(check_df) == 0:
                        st.success("✅ No known risk detected.")
                    else:
                        red_df    = check_df[check_df["Severity"] == "RED"]
                        orange_df = check_df[check_df["Severity"] == "ORANGE"]

                        st.error(f"🚨 RED ALERT: {len(red_df)} high-risk line(s) detected.")
                        st.warning(f"⚠️ ORANGE WARNING: {len(orange_df)} line(s) require manual review.")

                        # Ensure Current Container is never blank
                        check_df["Current Container"] = check_df["Current Container"].apply(
                            lambda x: x if clean_text(x) not in ("", "None") else "⚠️ Not detected"
                        )
                        red_df    = check_df[check_df["Severity"] == "RED"]
                        orange_df = check_df[check_df["Severity"] == "ORANGE"]

                        def render_cards(df, color):
                            """Render risk cards with current shipment vs historical side by side."""
                            border = "#c0392b" if color == "red" else "#d35400"
                            bg     = "#fff5f5" if color == "red" else "#fff8f0"
                            icon   = "🚨" if color == "red" else "⚠️"
                            for i, (_, row) in enumerate(df.iterrows()):
                                st.markdown(f"""
                                <div style="border:2px solid {border}; border-radius:10px;
                                            background:{bg}; padding:16px; margin-bottom:16px;">
                                    <div style="font-weight:700; font-size:1.05rem;
                                                color:{border}; margin-bottom:12px;">
                                        {icon} {row.get("Action Required","")} &nbsp;|&nbsp;
                                        {row.get("Message","")}
                                    </div>
                                    <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px;">
                                        <div style="background:white; border-radius:8px;
                                                    padding:12px; border:1px solid #ddd;">
                                            <div style="font-weight:700; color:#1a3c6e;
                                                        margin-bottom:8px; font-size:0.9rem;">
                                                📦 CURRENT SHIPMENT
                                            </div>
                                            <table style="width:100%; font-size:0.85rem;
                                                          border-collapse:collapse;">
                                                <tr><td style="color:#666; padding:3px 8px 3px 0;
                                                    white-space:nowrap;">Container</td>
                                                    <td style="font-weight:600; padding:3px 0;">
                                                    {row.get("Current Container","—")}</td></tr>
                                                <tr><td style="color:#666; padding:3px 8px 3px 0;
                                                    white-space:nowrap;">BL Number</td>
                                                    <td style="font-weight:600; padding:3px 0;">
                                                    {row.get("BL Number","—") or "—"}</td></tr>
                                                <tr><td style="color:#666; padding:3px 8px 3px 0;
                                                    white-space:nowrap;">Job Number</td>
                                                    <td style="font-weight:600; padding:3px 0;">
                                                    {row.get("Job Number","—") or "—"}</td></tr>
                                                <tr><td style="color:#666; padding:3px 8px 3px 0;
                                                    white-space:nowrap;">SKU</td>
                                                    <td style="font-weight:600; padding:3px 0;">
                                                    {row.get("SKU Number","—") or "—"}</td></tr>
                                                <tr><td style="color:#666; padding:3px 8px 3px 0;
                                                    white-space:nowrap;">Product</td>
                                                    <td style="font-weight:600; padding:3px 0;">
                                                    {row.get("Current Product","—")}</td></tr>
                                                <tr><td style="color:#666; padding:3px 8px 3px 0;
                                                    white-space:nowrap;">Current HS</td>
                                                    <td style="font-weight:600; color:{border};
                                                    padding:3px 0;">
                                                    {row.get("Current HS","—")}</td></tr>
                                                <tr><td style="color:#666; padding:3px 8px 3px 0;
                                                    white-space:nowrap;">✅ Should be</td>
                                                    <td style="font-weight:700; color:#1a6e3c;
                                                    padding:3px 0;">
                                                    {row.get("Corrected HS","—")}</td></tr>
                                            </table>
                                        </div>
                                        <div style="background:white; border-radius:8px;
                                                    padding:12px; border:1px solid #ddd;">
                                            <div style="font-weight:700; color:#555;
                                                        margin-bottom:8px; font-size:0.9rem;">
                                                📋 HISTORICAL REFERENCE
                                            </div>
                                            <table style="width:100%; font-size:0.85rem;
                                                          border-collapse:collapse;">
                                                <tr><td style="color:#666; padding:3px 8px 3px 0;
                                                    white-space:nowrap;">Risk ID</td>
                                                    <td style="font-weight:600; padding:3px 0;">
                                                    {row.get("Matched Risk ID","—")}</td></tr>
                                                <tr><td style="color:#666; padding:3px 8px 3px 0;
                                                    white-space:nowrap;">Container</td>
                                                    <td style="font-weight:600; padding:3px 0;">
                                                    {row.get("Previous Container","—")}</td></tr>
                                                <tr><td style="color:#666; padding:3px 8px 3px 0;
                                                    white-space:nowrap;">BL Number</td>
                                                    <td style="font-weight:600; padding:3px 0;">
                                                    {row.get("BL Number","—") or "—"}</td></tr>
                                                <tr><td style="color:#666; padding:3px 8px 3px 0;
                                                    white-space:nowrap;">Job Number</td>
                                                    <td style="font-weight:600; padding:3px 0;">
                                                    {row.get("Job Number","—") or "—"}</td></tr>
                                                <tr><td style="color:#666; padding:3px 8px 3px 0;
                                                    white-space:nowrap;">SKU</td>
                                                    <td style="font-weight:600; padding:3px 0;">
                                                    {row.get("SKU Number","—") or "—"}</td></tr>
                                                <tr><td style="color:#666; padding:3px 8px 3px 0;
                                                    white-space:nowrap;">MRN</td>
                                                    <td style="font-weight:600; padding:3px 0;">
                                                    {row.get("Previous MRN","—")}</td></tr>
                                                <tr><td style="color:#666; padding:3px 8px 3px 0;
                                                    white-space:nowrap;">Inspection</td>
                                                    <td style="font-weight:600; padding:3px 0;">
                                                    {row.get("Previous Inspection Date","—")}</td></tr>
                                                <tr><td style="color:#666; padding:3px 8px 3px 0;
                                                    white-space:nowrap;">Product</td>
                                                    <td style="font-weight:600; padding:3px 0;">
                                                    {row.get("Historical Product","—")}</td></tr>
                                                <tr><td style="color:#666; padding:3px 8px 3px 0;
                                                    white-space:nowrap;">Old HS</td>
                                                    <td style="font-weight:600; color:{border};
                                                    padding:3px 0;">
                                                    {row.get("Old HS Used Before","—")}</td></tr>
                                                <tr><td style="color:#666; padding:3px 8px 3px 0;
                                                    white-space:nowrap;">✅ Corrected to</td>
                                                    <td style="font-weight:700; color:#1a6e3c;
                                                    padding:3px 0;">
                                                    {row.get("Corrected HS","—")}</td></tr>
                                            </table>
                                        </div>
                                    </div>
                                </div>
                                """, unsafe_allow_html=True)

                        if len(red_df) > 0:
                            st.markdown("#### 🚨 Action Required — High Risk Lines")
                            render_cards(red_df, "red")

                        if len(orange_df) > 0:
                            st.markdown("#### ⚠️ Manual Review Required")
                            render_cards(orange_df, "orange")

                        # PDF export
                        st.markdown("---")
                        st.markdown("#### 📄 Export Risk Report")
                        files_info = ", ".join([f.name for f in doc_files])
                        pdf_bytes = build_pdf_report(check_df, files_info)
                        report_name = f"RiskReport_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                        st.download_button(
                            label="⬇️ Download PDF Report",
                            data=pdf_bytes,
                            file_name=report_name,
                            mime="application/pdf",
                            type="primary",
                            use_container_width=False,
                        )

    # ────────────────────────────────────────────────────────────────────────
    # TAB 2 — Import / Add Cases  (admin + broker)
    # ────────────────────────────────────────────────────────────────────────
    with tab_objects[1]:
        if not is_broker:
            st.info("🔒 Visitors cannot add or import risk cases. Please contact your broker or admin.")
        else:
            # Batch import
            st.subheader("2a. Batch Import Risk Database from Excel")
            uploaded_file = st.file_uploader("Upload Risk Database Excel", type=["xlsx"], key="risk_database_upload")
            if uploaded_file is not None:
                imported_df = normalize_import_file(uploaded_file)
                st.write("Preview of valid risk cases:")
                st.dataframe(imported_df, use_container_width=True)
                if st.button("Import This Excel Into Database"):
                    existing_df   = load_database()
                    existing_keys = set(make_duplicate_key(row) for _, row in existing_df.iterrows())
                    new_rows, duplicate_count = [], 0
                    for _, row in imported_df.iterrows():
                        key = make_duplicate_key(row)
                        if key in existing_keys:
                            duplicate_count += 1
                            continue
                        row = row.copy()
                        temp_df = pd.concat([existing_df, pd.DataFrame(new_rows)], ignore_index=True)
                        row["Risk ID"]    = generate_risk_id(temp_df)
                        row["Input Date"] = str(date.today())
                        if not clean_text(row["Status"]):
                            row["Status"] = "active"
                        new_rows.append(row.to_dict())
                        existing_keys.add(key)
                    if new_rows:
                        combined_df = pd.concat([existing_df, pd.DataFrame(new_rows)], ignore_index=True)
                        save_database(combined_df)
                        st.success(f"Imported {len(new_rows)} new rows. Skipped {duplicate_count} duplicates.")
                    else:
                        st.warning(f"No new rows imported. {duplicate_count} duplicate rows detected.")

            st.markdown("---")

            # Manual entry
            st.subheader("2b. Add New Customs Risk Case Manually")
            with st.form("risk_case_form"):
                current_df = load_database()
                risk_id    = generate_risk_id(current_df)
                st.write(f"Risk ID: **{risk_id}**")

                col1, col2, col3 = st.columns(3)
                with col1: input_date      = st.date_input("Input Date", value=date.today())
                with col2: cc_date         = st.date_input("CC Date", value=None)
                with col3: inspection_date = st.date_input("Inspection Date", value=None)

                container_no = st.text_input("Container No")
                mrn          = st.text_input("MRN")

                col_bl = st.columns(3)
                with col_bl[0]: bl_number  = st.text_input("BL Number")
                with col_bl[1]: job_number = st.text_input("Job Number")
                with col_bl[2]: inspector  = st.text_input("Inspector")

                product_name = st.text_input("Product Name *")
                product_alias     = st.text_input("Product Alias / Possible Descriptions")
                declaration_desc  = st.text_input("Declaration Description")

                col4, col5 = st.columns(2)
                with col4:
                    old_hs      = st.text_input("Old HS *")
                    duty_before = st.text_input("Duty Before", placeholder="e.g. 2,7")
                with col5:
                    corrected_hs = st.text_input("Corrected HS *")
                    duty_after   = st.text_input("Duty After",  placeholder="e.g. 3,9 or 0")

                findings_type   = st.text_input("Findings Type")
                root_cause      = st.text_input("Root Cause")
                risk_reason     = st.text_area("Risk Reason")
                customs_comment = st.text_area("Customs / Broker Comment")
                status          = st.selectbox("Status", ["active", "monitoring", "solved"])
                notes           = st.text_area("Notes")

                submitted = st.form_submit_button("Save Risk Case")
                if submitted:
                    if not product_name or not old_hs or not corrected_hs:
                        st.error("Product Name, Old HS and Corrected HS are required.")
                    else:
                        current_df = load_database()
                        new_row = {
                            "Risk ID":                 generate_risk_id(current_df),
                            "Input Date":              str(input_date),
                            "CC Date":                 str(cc_date) if cc_date else "",
                            "Inspection Date":         str(inspection_date) if inspection_date else "",
                            "Container No":            container_no,
                            "MRN":                     mrn,
                            "BL Number":               bl_number,
                            "Job Number":              job_number,
                            "Inspector":               inspector,
                            "Product Name":            product_name,
                            "Product Alias":           product_alias,
                            "Declaration Description": declaration_desc,
                            "Old HS":                  clean_hs(old_hs),
                            "Corrected HS":            clean_hs(corrected_hs),
                            "Duty Before":             format_duty_rate(duty_before),
                            "Duty After":              format_duty_rate(duty_after),
                            "Findings Type":           findings_type,
                            "Root Cause":              root_cause,
                            "Risk Reason":             risk_reason,
                            "Customs Comment":         customs_comment,
                            "Status":                  status,
                            "Notes":                   notes,
                        }
                        new_key       = make_duplicate_key(new_row)
                        existing_keys = set(make_duplicate_key(r) for _, r in current_df.iterrows())
                        if new_key in existing_keys:
                            st.warning("This risk case already exists. It was not added again.")
                        else:
                            current_df = pd.concat([current_df, pd.DataFrame([new_row])], ignore_index=True)
                            save_database(current_df)
                            st.success("Risk case saved successfully.")

    # ────────────────────────────────────────────────────────────────────────
    # TAB 3 — Database viewer / editor
    # ────────────────────────────────────────────────────────────────────────
    with tab_objects[2]:
        st.subheader("3. Current Risk Database")
        current_df = load_database()

        if len(current_df) == 0:
            st.info("No risk cases available yet.")
        else:
            # Filter bar
            fc1, fc2, fc3 = st.columns(3)
            with fc1:
                filter_status = st.selectbox("Filter by Status", ["All", "active", "monitoring", "solved"])
            with fc2:
                filter_hs = st.text_input("Filter by Old HS (exact)")
            with fc3:
                filter_product = st.text_input("Filter by Product Name (contains)")

            filtered = current_df.copy()
            if filter_status != "All":
                filtered = filtered[filtered["Status"] == filter_status]
            if filter_hs:
                filtered = filtered[filtered["Old HS"] == filter_hs.strip()]
            if filter_product:
                filtered = filtered[filtered["Product Name"].str.contains(filter_product.strip(), case=False, na=False)]

            st.markdown(f"Showing **{len(filtered)}** of **{len(current_df)}** cases")

            if is_broker:
                editable = filtered.copy()
                editable.insert(0, "Delete", False)
                edited = st.data_editor(
                    editable,
                    use_container_width=True,
                    num_rows="fixed",
                    hide_index=True,
                    column_config={"Delete": st.column_config.CheckboxColumn("Delete")}
                )
                if st.button("💾 Save Changes"):
                    updated = edited[edited["Delete"] == False].copy().drop(columns=["Delete"])
                    # Merge back: rows not in filter remain unchanged
                    unchanged = current_df[~current_df["Risk ID"].isin(filtered["Risk ID"])]
                    merged    = pd.concat([unchanged, updated], ignore_index=True)
                    for col in COLUMNS:
                        if col not in merged.columns:
                            merged[col] = ""
                    merged = merged[COLUMNS]
                    for col in COLUMNS:
                        merged[col] = merged[col].apply(clean_text)
                    for col in ["Input Date", "CC Date", "Inspection Date"]:
                        merged[col] = merged[col].apply(format_date)
                    merged["Old HS"]       = merged["Old HS"].apply(clean_hs)
                    merged["Corrected HS"] = merged["Corrected HS"].apply(clean_hs)
                    merged["Duty Before"]  = merged["Duty Before"].apply(format_duty_rate)
                    merged["Duty After"]   = merged["Duty After"].apply(format_duty_rate)
                    save_database(merged)
                    st.success("Changes saved successfully.")
            else:
                st.info("👁️ Visitor mode — read only.")
                st.dataframe(filtered, use_container_width=True, hide_index=True)

    # ────────────────────────────────────────────────────────────────────────
    # TAB 4 — Dashboard
    # ────────────────────────────────────────────────────────────────────────
    with tab_objects[3]:
        render_dashboard(load_database())

    # ────────────────────────────────────────────────────────────────────────
    # TAB 5 — Admin (admins only)
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
