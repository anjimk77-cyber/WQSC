import pandas as pd
import streamlit as st
from datetime import date

import gspread
from google.oauth2.service_account import Credentials

# =========================================================================
# CONFIG
# =========================================================================
st.set_page_config(page_title="Water Quality Measures - Special Cases", layout="wide", page_icon="🧪")

CUSTOMER_FILE = "Customer List.xlsx"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# Name of the new column this app owns. It is appended AFTER whatever
# columns the main "Water Quality & Harvest Report" app already manages
# (its COLUMN_ORDER list), so it never collides with that app's own
# header self-heal logic (that app only ever rewrites the first
# len(COLUMN_ORDER) header cells — it never touches anything after them).
SPECIAL_COL_NAME = "WQ Special Cases"
# Separate column that records WHEN a special case was actually saved
# through this app (i.e. the submit date), as opposed to "Date" which is
# the water-quality visit/report date already owned by the main app.
SPECIAL_ENTERED_DATE_COL_NAME = "WQ Special Cases Entered Date"
SPECIAL_SEP = ", "

SPECIAL_CASE_OPTIONS = [
    "Low PH", "High PH",
    "Low Salinity", "High Salinity",
    "Low Ammonia", "High Ammonia",
    "Low Alkalinity", "High Alkalinity",
]

# =========================================================================
# STYLE (kept self-contained in this file — does not touch the main app)
# =========================================================================
st.markdown("""
<style>
ul[role="listbox"], div[role="listbox"] {
    width: max-content !important;
    min-width: 220px !important;
    max-width: 92vw !important;
}
[role="option"] {
    width: auto !important;
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
    word-break: break-word !important;
}
div[data-baseweb="tag"] { white-space: normal !important; max-width: 100% !important; }
span[data-baseweb="tag"] { white-space: normal !important; }

button[kind="primary"], button[kind="primaryFormSubmit"] {
    background-color: #e63946 !important;
    border-color: #e63946 !important;
    color: #ffffff !important;
}
button[kind="primary"]:hover, button[kind="primaryFormSubmit"]:hover {
    background-color: #c1121f !important;
    border-color: #c1121f !important;
    color: #ffffff !important;
}

.pond-badge {
    display: inline-block;
    border: 1px solid rgba(128,128,128,0.3);
    border-radius: 8px;
    padding: 0.5rem 0.8rem;
    margin: 0.25rem;
    background-color: rgba(128,128,128,0.04);
    min-width: 140px;
}
.pond-badge .pond-num { font-weight: 700; font-size: 1.05rem; }
.pond-badge .pond-meta { font-size: 0.8rem; color: gray; }
.pond-badge .pond-special { font-size: 0.8rem; color: #c1121f; font-weight: 600; }

@media (max-width: 700px) {
    div[data-testid="stHorizontalBlock"] {
        flex-direction: column !important;
    }
    div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
        width: 100% !important;
        min-width: 100% !important;
    }
}
</style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center;'>Water Quality Measures - Special Cases</h1>",
            unsafe_allow_html=True)
st.subheader("KMN Aqua Services")
st.markdown("---")

# =========================================================================
# GOOGLE SHEETS SETUP CHECK — reuses the exact same secrets structure as
# the main app, so both apps point at the same spreadsheet / worksheet.
# =========================================================================
def _gsheet_configured():
    return "gcp_service_account" in st.secrets and "gsheet" in st.secrets and "sheet_id" in st.secrets["gsheet"]

if not _gsheet_configured():
    st.error("❌ Google Sheets is not configured yet.")
    with st.expander("⚙️ How to connect this app to the Google Sheet", expanded=True):
        st.markdown(
            "This app must point at the **same** Google Sheet / secrets as the main "
            "Water Quality & Harvest Report app. Add the same `[gcp_service_account]` and "
            "`[gsheet]` blocks (sheet_id, worksheet_name) to this app's `.streamlit/secrets.toml`."
        )
    st.stop()

@st.cache_resource(show_spinner=False)
def get_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_resource(show_spinner=False)
def get_worksheet():
    """Opens the SAME worksheet the main app writes to (default
    'WaterQualityData'). Deliberately does NOT rewrite/reset the header
    row the way the main app does — this app only ever ADDS its own
    column if missing, and never touches any other existing header
    cell, so it can never clobber the main app's COLUMN_ORDER."""
    client = get_client()
    sheet_id = st.secrets["gsheet"]["sheet_id"]
    worksheet_name = st.secrets["gsheet"].get("worksheet_name", "WaterQualityData")
    sh = client.open_by_key(sheet_id)
    try:
        ws = sh.worksheet(worksheet_name)
    except gspread.WorksheetNotFound:
        st.error(
            f"❌ Worksheet '{worksheet_name}' was not found. Please run the main "
            "Water Quality & Harvest Report app first so it can create and set up the sheet."
        )
        st.stop()
    return ws

def ensure_special_column(ws):
    """Makes sure the 'WQ Special Cases' column AND the
    'WQ Special Cases Entered Date' column exist in the sheet. If either
    is missing, it's appended as a brand-new column right after every
    existing header cell — never overwrites or reorders any column the
    main app already owns. Returns (special_col_index, special_date_col_index),
    both 1-based."""
    header = ws.row_values(1)

    if SPECIAL_COL_NAME in header:
        special_col_index = header.index(SPECIAL_COL_NAME) + 1
    else:
        special_col_index = len(header) + 1
        ws.update_cell(1, special_col_index, SPECIAL_COL_NAME)
        header = ws.row_values(1)  # refresh so the next check sees it

    if SPECIAL_ENTERED_DATE_COL_NAME in header:
        special_date_col_index = header.index(SPECIAL_ENTERED_DATE_COL_NAME) + 1
    else:
        special_date_col_index = len(header) + 1
        ws.update_cell(1, special_date_col_index, SPECIAL_ENTERED_DATE_COL_NAME)

    return special_col_index, special_date_col_index

try:
    _ws = get_worksheet()
    SPECIAL_COL_INDEX, SPECIAL_DATE_COL_INDEX = ensure_special_column(_ws)
except Exception as e:
    st.error(f"❌ Could not connect to the Google Sheet. Check your secrets and sharing settings.\n\n{e}")
    st.stop()

# =========================================================================
# LOAD CUSTOMER LIST (same source file as the main app)
# =========================================================================
@st.cache_data
def load_customer_data():
    return pd.read_excel(CUSTOMER_FILE)

try:
    customer_df = load_customer_data()
except Exception as e:
    st.error(f"❌ Could not load '{CUSTOMER_FILE}'. Make sure it's in the app folder. ({e})")
    st.stop()

REQUIRED_COLS = ["Customer Name", "Farm Name with Code", "Zone", "Area"]
missing_cols = [c for c in REQUIRED_COLS if c not in customer_df.columns]
if missing_cols:
    st.error(f"❌ 'Customer List.xlsx' is missing required column(s): {', '.join(missing_cols)}")
    st.stop()

for _col in REQUIRED_COLS:
    customer_df[_col] = customer_df[_col].apply(
        lambda v: "" if pd.isna(v) else (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v))
    )

all_customers = sorted(customer_df["Customer Name"].replace("", pd.NA).dropna().unique().tolist())

# =========================================================================
# SHEET DATA HELPERS — read the raw sheet (values + real row numbers) so
# updates can target the exact saved row for a pond.
# =========================================================================
def bump_data_version():
    st.session_state["_data_version"] = st.session_state.get("_data_version", 0) + 1

def _load_sheet_df_cached(data_version):
    ws = get_worksheet()
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()
    header, rows = values[0], values[1:]
    df = pd.DataFrame(rows, columns=header)
    df["_row_number"] = range(2, 2 + len(df))
    return df

def load_sheet_df():
    """All rows in the sheet, with soft-deleted rows filtered out (same
    'Deleted' flag convention as the main app), as a DataFrame that also
    carries each row's real sheet row number in '_row_number'."""
    df = _load_sheet_df_cached(st.session_state.get("_data_version", 0))
    if len(df) == 0:
        return df
    if "Deleted" in df.columns:
        is_deleted = df["Deleted"].astype(str).str.strip().str.lower().isin(["yes", "true", "1"])
        df = df[~is_deleted].reset_index(drop=True)
    return df

def get_farm_ponds(df, customer, farm):
    """Every distinct Pond Number on record for this Customer + Farm,
    together with each pond's latest saved row (by parsed Date — same
    'take the most recent row' logic the main app uses for harvests),
    sorted for display."""
    required = {"Customer", "Farm Name with Code", "Pond Number"}
    if len(df) == 0 or not required.issubset(df.columns):
        return pd.DataFrame()
    sub = df[(df["Customer"] == customer) & (df["Farm Name with Code"] == farm)].copy()
    sub = sub[sub["Pond Number"].astype(str).str.strip() != ""]
    if len(sub) == 0:
        return pd.DataFrame()
    sub["_ParsedDate"] = pd.to_datetime(sub.get("Date"), errors="coerce")
    sub = sub.sort_values(by="_ParsedDate")
    latest_per_pond = sub.groupby("Pond Number", as_index=False).last()
    return latest_per_pond.sort_values(by="Pond Number").reset_index(drop=True)

def get_farm_last_visit_date(df, customer, farm):
    """Most recent Date across ALL rows on record for this Customer + Farm
    (not just the latest-per-pond rows used for the pond layout) — used
    to show a single 'Last Visit Date' for the farm as a whole."""
    required = {"Customer", "Farm Name with Code", "Date"}
    if len(df) == 0 or not required.issubset(df.columns):
        return None
    sub = df[(df["Customer"] == customer) & (df["Farm Name with Code"] == farm)]
    parsed = pd.to_datetime(sub.get("Date"), errors="coerce").dropna()
    if len(parsed) == 0:
        return None
    return parsed.max().date()

def _display_cycle(cycle_value):
    """Cycle Type value for badge display. 'Full Harvest' is shown
    abbreviated as 'Full H' to keep the pond badges compact; every other
    value (Partial Harvest, Culture, etc.) is shown as-is."""
    c = str(cycle_value or "").strip()
    if c.lower() == "full harvest":
        return "Full H"
    return c or "-"

def get_all_special_cases_entries(df):
    """Every row across the whole sheet that has a WQ Special Cases value
    recorded, reshaped for the summary table at the bottom of the page.
    'Entered Date' comes from the WQ Special Cases Entered Date column
    (i.e. when it was actually saved through this app), NOT the water
    quality visit 'Date'. This is independent of whichever Customer/Farm
    is currently selected above, and is sorted by Entered Date, most
    recent first."""
    empty_cols = ["Entered Date", "Customer Name", "Farm Name with Code", "Pond No", SPECIAL_COL_NAME]
    if len(df) == 0 or SPECIAL_COL_NAME not in df.columns:
        return pd.DataFrame(columns=empty_cols)
    sub = df[df[SPECIAL_COL_NAME].astype(str).str.strip() != ""].copy()
    if len(sub) == 0:
        return pd.DataFrame(columns=empty_cols)
    out = pd.DataFrame({
        "Entered Date": sub.get(SPECIAL_ENTERED_DATE_COL_NAME, ""),
        "Customer Name": sub.get("Customer", ""),
        "Farm Name with Code": sub.get("Farm Name with Code", ""),
        "Pond No": sub.get("Pond Number", ""),
        SPECIAL_COL_NAME: sub[SPECIAL_COL_NAME],
    })
    out["_ParsedDate"] = pd.to_datetime(out["Entered Date"], errors="coerce")
    out = out.sort_values(by="_ParsedDate", ascending=False).drop(columns=["_ParsedDate"])
    return out.reset_index(drop=True)

def update_special_case_for_pond(row_number, special_value):
    """Saves the special case value AND stamps today's date into the
    WQ Special Cases Entered Date column — this is the actual submit
    date shown as 'Entered Date' in the summary table below."""
    ws = get_worksheet()
    ws.update_cell(row_number, SPECIAL_COL_INDEX, special_value)
    ws.update_cell(row_number, SPECIAL_DATE_COL_INDEX, date.today().strftime("%Y-%m-%d"))
    bump_data_version()

# =========================================================================
# STEP 1: CUSTOMER / FARM
# =========================================================================
st.subheader("📋 Select Farm")

col1, col2 = st.columns(2)
with col1:
    customer = st.selectbox("Customer Name *", all_customers, key="customer_select")

farm_options = sorted(
    customer_df.loc[customer_df["Customer Name"] == customer, "Farm Name with Code"]
    .dropna().unique().tolist()
)
if not farm_options:
    farm_options = ["-- No farms found for this customer --"]

with col2:
    farm = st.selectbox("Farm Name with Code *", farm_options, key=f"farm_select_{customer}")

farm_row_match = customer_df[
    (customer_df["Customer Name"] == customer) & (customer_df["Farm Name with Code"] == farm)
]

# --- Zone is displayed automatically, read-only, from the Customer List ---
zone_val = farm_row_match.iloc[0]["Zone"] if len(farm_row_match) > 0 else ""
area_val = farm_row_match.iloc[0]["Area"] if len(farm_row_match) > 0 else ""

col3, col4 = st.columns(2)
with col3:
    st.text_input("Zone", value=zone_val, disabled=True, key=f"zone_display_{farm}")
with col4:
    st.text_input("Area", value=area_val, disabled=True, key=f"area_display_{farm}")

if len(farm_row_match) > 0 and "Marketing Manager" in customer_df.columns:
    mm = farm_row_match.iloc[0].get("Marketing Manager", "")
    if str(mm).strip():
        st.caption(f"Marketing Manager: {mm}")

# =========================================================================
# STEP 2: POND LAYOUT — every pond on record for this farm, shown as a
# card grid with its latest known status (Species/Cycle/DOC) and its
# current WQ Special Cases value (if any), for quick reference before
# picking one below.
# =========================================================================
st.markdown("---")
st.markdown("#### 🗺️ Pond Layout")

df_sheet = load_sheet_df()
farm_ponds_df = get_farm_ponds(df_sheet, customer, farm)

last_visit_date = get_farm_last_visit_date(df_sheet, customer, farm)
if last_visit_date is not None:
    st.caption(f"🗓️ Last Visit Date for {farm}: **{last_visit_date.strftime('%Y-%m-%d')}**")
else:
    st.caption(f"🗓️ Last Visit Date for {farm}: **-**")

if len(farm_ponds_df) == 0:
    st.info("No saved Pond Details records yet for this farm. Add pond records in the main "
            "Water Quality & Harvest Report app first, then come back here.")
else:
    badges_html = "<div>"
    for _, prow in farm_ponds_df.iterrows():
        pond = prow.get("Pond Number", "")
        species = prow.get("Species Culture", "") or "-"
        cycle = _display_cycle(prow.get("Cycle Type", ""))
        doc = prow.get("DOC", "") or "-"
        current_special = prow.get(SPECIAL_COL_NAME, "") if SPECIAL_COL_NAME in prow else ""
        special_html = (
            f"<div class='pond-special'>⚠️ {current_special}</div>" if str(current_special).strip() else ""
        )

        # Issues (from this pond's latest saved record) — shown at the
        # bottom inside the badge as "Disease: <issue>", with the issue
        # text in red. Stays empty when the latest record has no Issues.
        issues_val = str(prow.get("Issues", "") or "").strip()
        issues_html = (
            "<div style='font-size:0.8rem;font-weight:600;margin-top:4px;"
            "border-top:1px dashed rgba(128,128,128,0.4);padding-top:3px;'>"
            "<span></span>"
            "<span style='color:red;'>"
            f"{issues_val.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')}"
            "</span></div>"
            if issues_val and issues_val.lower() != "nan" else ""
        )

        badges_html += (
            "<div class='pond-badge'>"
            f"<div class='pond-num'>Pond {pond}</div>"
            f"<div class='pond-meta'>{species} · {cycle} · DOC {doc}</div>"
            f"{special_html}"
            f"{issues_html}"
            "</div>"
        )
    badges_html += "</div>"
    st.markdown(badges_html, unsafe_allow_html=True)
    st.caption(f"{len(farm_ponds_df)} pond(s) on record for {farm}.")

# =========================================================================
# STEP 3: WATER QUALITY MEASURES - SPECIAL CASES
# Select a pond, pick any number of special-case conditions, and save —
# this writes into the SAME Google Sheet the main app uses, into a
# 'WQ Special Cases' column, on that pond's most recent saved row.
# =========================================================================
st.markdown("---")
st.markdown("#### 🧪 Water Quality Measures - Special Cases")

# --- Show the "saved" confirmation left over from the previous run.
# The save button below triggers st.rerun() right after saving, which
# would otherwise wipe out an st.success() call before the user ever
# sees it. Stashing the message in session_state lets it survive the
# rerun; it's shown once here, then popped so it doesn't linger forever.
if st.session_state.get("_special_save_success"):
    st.success(st.session_state.pop("_special_save_success"))

if len(farm_ponds_df) == 0:
    st.info("No ponds available yet for this farm — nothing to record a special case against.")
else:
    pond_list = farm_ponds_df["Pond Number"].astype(str).tolist()
    special_scope = f"{customer}_{farm}"

    selected_pond = st.selectbox("Select Pond *", pond_list, key=f"special_pond_{special_scope}")

    pond_row = farm_ponds_df[farm_ponds_df["Pond Number"].astype(str) == selected_pond].iloc[0]
    existing_special_raw = str(pond_row.get(SPECIAL_COL_NAME, "") or "").strip()
    existing_special_list = [p.strip() for p in existing_special_raw.split(SPECIAL_SEP) if p.strip()]
    if existing_special_raw:
        st.caption(f"Current recorded special case(s) for Pond {selected_pond}: **{existing_special_raw}**")

    selected_cases = st.multiselect(
        "Special Case(s) *",
        SPECIAL_CASE_OPTIONS,
        default=[c for c in existing_special_list if c in SPECIAL_CASE_OPTIONS],
        key=f"special_cases_{special_scope}_{selected_pond}",
    )

    if st.button("✅ Save Special Case", type="primary", key=f"special_submit_{special_scope}"):
        if not selected_cases:
            st.error("❌ Please select at least one special case condition.")
        else:
            row_number = int(pond_row["_row_number"])
            special_value = SPECIAL_SEP.join(dict.fromkeys(selected_cases))
            update_special_case_for_pond(row_number, special_value)
            st.session_state["_special_save_success"] = (
                f"✅ Saved special case(s) for Pond {selected_pond}: {special_value}"
            )
            st.rerun()

# =========================================================================
# STEP 4: USER ENTERED SPECIAL CASES DATA — a log of every WQ Special
# Cases entry saved through this app, across all customers/farms, most
# recent first. "Entered Date" is the actual save/submit date, not the
# water quality visit date.
# =========================================================================
st.markdown("---")
st.markdown("#### 📝 User Entered Special Cases Data")

special_entries_df = get_all_special_cases_entries(df_sheet)
if len(special_entries_df) == 0:
    st.info("No special cases recorded yet.")
else:
    st.dataframe(special_entries_df, use_container_width=True, hide_index=True)

st.markdown("---")
st.markdown("<p style='text-align: center; color: gray;'>KMN Aqua Services - Water Quality Monitoring System (Special Cases)</p>",
            unsafe_allow_html=True)
