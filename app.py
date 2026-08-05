import io
import os
import sys
import gc
import asyncio
import datetime
import warnings
import pandas as pd
import streamlit as st

# -----------------------------------------------------------------------------
# MONKEY-PATCH STARLETTE GZIP RESPONDER (Fixes Streamlit Cloud Starlette >=0.40.0 TypeError)
# -----------------------------------------------------------------------------
try:
    import starlette.middleware.gzip
    _orig_gzip_init = starlette.middleware.gzip.GZipResponder.__init__
    
    def _patched_gzip_init(self, app, minimum_size=500, compresslevel=9, **kwargs):
        try:
            _orig_gzip_init(self, app, minimum_size=minimum_size, compresslevel=compresslevel, **kwargs)
        except TypeError:
            kwargs.setdefault('thread_minimum_size', minimum_size)
            _orig_gzip_init(self, app, minimum_size=minimum_size, compresslevel=compresslevel, **kwargs)

    starlette.middleware.gzip.GZipResponder.__init__ = _patched_gzip_init
except Exception:
    pass

# Ensure asyncio event loop is set for Python 3.14 main thread
try:
    asyncio.get_running_loop()
except RuntimeError:
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

# -----------------------------------------------------------------------------
# 1. PAGE CONFIG (MUST BE THE VERY FIRST STREAMLIT COMMAND IN APP.PY)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="General Ledger Aggregator & Vendor Analysis",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Filter out non-critical warnings
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=DeprecationWarning)

# Ensure project root is in sys.path
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from utils.data_loader import (
    load_ledger_data, 
    extract_unique_vendors, 
    get_default_data_dir
)


# Cached helper to build Excel files lazily without bloating memory on every rerun
@st.cache_data(show_spinner=False, max_entries=5)
def convert_df_to_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Filtered_Ledger')
    return output.getvalue()


@st.cache_data(show_spinner=False, max_entries=5)
def convert_df_to_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode('utf-8')


# -----------------------------------------------------------------------------
# 2. STYLING & AESTHETICS
# -----------------------------------------------------------------------------
st.markdown("""
<style>
    /* Global Container Padding */
    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        padding-left: 2rem;
        padding-right: 2rem;
    }
    
    /* Header Styling */
    .app-title {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E293B;
        margin-bottom: 0.2rem;
    }
    .app-subtitle {
        font-size: 1.0rem;
        color: #64748B;
        margin-bottom: 1.5rem;
    }
    
    /* Card Aesthetics */
    div[data-testid="stMetric"] {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 15px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    div[data-testid="stMetric"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    }
    div[data-testid="stMetricLabel"] {
        font-weight: 600;
        color: #475569;
        font-size: 0.9rem;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.6rem;
        font-weight: 700;
        color: #0F172A;
    }
</style>
""", unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# 3. SIDEBAR CONFIGURATION & DATA LOADING
# -----------------------------------------------------------------------------
st.sidebar.image("https://img.icons8.com/color/96/general-ledger.png", width=64)
st.sidebar.title("Ledger Settings & Filters")

# Mode selector: Local Repository vs Upload vs Remote API
data_source_mode = st.sidebar.radio(
    "Data Source Mode",
    ["Local / Repository Data (Recommended)", "Upload Excel Files", "GitHub API Stream (Remote)"],
    index=0,
    help="Select how to load the General Ledger Excel reports."
)

uploaded_files = None
folder_path = None

if data_source_mode == "Local / Repository Data (Recommended)":
    default_dir = get_default_data_dir()
    folder_path = st.sidebar.text_input(
        "Data Folder Path",
        value=default_dir,
        help="Directory containing *.xlsx General Ledger reports."
    )
    if st.sidebar.button("🔄 Refresh / Reload Folder"):
        st.cache_data.clear()
        st.rerun()
elif data_source_mode == "Upload Excel Files":
    uploaded_files = st.sidebar.file_uploader(
        "Upload General Ledger Excel Files",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

# Load raw ledger data safely
try:
    df_raw, load_stats = load_ledger_data(
        mode=data_source_mode, 
        folder_path=folder_path, 
        uploaded_files=uploaded_files
    )
except Exception as e:
    empty_cols = [
        'date', 'account_name', 'transaction_details', 'transaction_type', 
        'reference_number', 'entity_number', 'debit', 'credit', 'net_amount', 
        'contact_id', 'account_id', 'branch_name', '_source_file'
    ]
    df_raw = pd.DataFrame(columns=empty_cols)
    load_stats = {
        'files_count': 0,
        'total_rows': 0,
        'unique_rows': 0,
        'processed_files': [],
        'errors': [f"Initialization Warning: {str(e)}"]
    }

# Sidebar Folder Scanner Status Card
st.sidebar.markdown("---")
st.sidebar.subheader("📂 Ingestion Status")

if load_stats['files_count'] > 0:
    st.sidebar.success(f"Successfully loaded **{load_stats['files_count']}** Excel file(s).")
    st.sidebar.caption(f"• **Total Raw Records:** {load_stats['total_rows']:,}")
    st.sidebar.caption(f"• **Deduplicated Records:** {load_stats['unique_rows']:,}")
    
    with st.sidebar.expander("Loaded Files List", expanded=False):
        for fn in load_stats['processed_files']:
            st.write(f"📄 `{fn}`")
else:
    st.sidebar.warning("No valid General Ledger Excel files found.")
    if load_stats.get('errors'):
        with st.sidebar.expander("Ingestion Errors", expanded=True):
            for err in load_stats['errors']:
                st.error(err)


# -----------------------------------------------------------------------------
# 4. FILTERS & SEARCH CONTROLS
# -----------------------------------------------------------------------------
if not df_raw.empty:
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Filters & Search")

    # Date Range Filter
    valid_dates = df_raw['date'].dropna()
    if not valid_dates.empty:
        min_date = valid_dates.min().date()
        max_date = valid_dates.max().date()
        
        selected_date_range = st.sidebar.date_input(
            "Date Range",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date
        )
    else:
        selected_date_range = None

    # Branch Filter
    all_branches = sorted([b for b in df_raw['branch_name'].unique() if b])
    selected_branches = st.sidebar.multiselect(
        "Filter by Branch",
        options=all_branches,
        default=[],
        placeholder="All Branches"
    )

    # Transaction Type Filter
    all_types = sorted([t for t in df_raw['transaction_type'].unique() if t])
    selected_types = st.sidebar.multiselect(
        "Filter by Transaction Type",
        options=all_types,
        default=[],
        placeholder="All Types"
    )

    # Vendor Selector (transaction_details)
    unique_vendors = extract_unique_vendors(df_raw)
    
    st.sidebar.markdown("**Vendor Selection (`transaction_details`)**")
    vendor_mode = st.sidebar.radio(
        "Vendor Filter Method",
        ["All Vendors", "Select Vendor(s)", "Keyword Search"],
        index=0,
        horizontal=True
    )
    
    selected_vendors = []
    vendor_keyword = ""
    
    if vendor_mode == "Select Vendor(s)":
        selected_vendors = st.sidebar.multiselect(
            "Search / Select Vendor (`transaction_details`)",
            options=unique_vendors,
            placeholder="Type vendor or transaction detail..."
        )
    elif vendor_mode == "Keyword Search":
        vendor_keyword = st.sidebar.text_input(
            "Vendor / Account Keyword Search",
            placeholder="Enter detail keyword or account name..."
        )


# -----------------------------------------------------------------------------
# 5. FILTER EXECUTION LOGIC
# -----------------------------------------------------------------------------
df_filtered = df_raw

if not df_raw.empty:
    # Apply filters iteratively
    mask = pd.Series(True, index=df_raw.index)
    
    # Date filtering
    if selected_date_range and isinstance(selected_date_range, (tuple, list)):
        if len(selected_date_range) == 2:
            start_d, end_d = selected_date_range
            start_ts = pd.to_datetime(start_d)
            end_ts = pd.to_datetime(end_d) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            mask &= (df_raw['date'] >= start_ts) & (df_raw['date'] <= end_ts)
        elif len(selected_date_range) == 1:
            start_d = selected_date_range[0]
            start_ts = pd.to_datetime(start_d)
            mask &= (df_raw['date'] >= start_ts)

    # Branch filtering
    if selected_branches:
        mask &= df_raw['branch_name'].isin(selected_branches)

    # Transaction type filtering
    if selected_types:
        mask &= df_raw['transaction_type'].isin(selected_types)

    # Vendor filtering
    if vendor_mode == "Select Vendor(s)" and selected_vendors:
        mask &= (df_raw['transaction_details'].isin(selected_vendors) | df_raw['contact_id'].isin(selected_vendors))
    elif vendor_mode == "Keyword Search" and vendor_keyword.strip():
        kw = vendor_keyword.strip().lower()
        match_details = df_raw['transaction_details'].str.lower().str.contains(kw, na=False)
        match_contact = df_raw['contact_id'].str.lower().str.contains(kw, na=False)
        match_acc = df_raw['account_name'].str.lower().str.contains(kw, na=False)
        mask &= (match_details | match_contact | match_acc)

    df_filtered = df_raw[mask]
    
    # Force Garbage Collection to keep memory usage low
    gc.collect()


# -----------------------------------------------------------------------------
# 6. DASHBOARD MAIN CONTENT
# -----------------------------------------------------------------------------
st.markdown('<div class="app-title">VASANTA BHAVAN HOTELS INDIA (P) LTD</div>', unsafe_allow_html=True)
st.markdown('<div class="app-subtitle">Multi-File General Ledger Aggregator & Vendor Transaction Explorer</div>', unsafe_allow_html=True)

if df_raw.empty:
    st.info("👋 Welcome! Please select a valid Data Folder path in the sidebar or upload `.xlsx` files to get started.")
    st.stop()

# Key KPI Metrics
col1, col2, col3, col4 = st.columns(4)

total_tx = len(df_filtered)
total_debit = df_filtered['debit'].sum() if not df_filtered.empty else 0.0
total_credit = df_filtered['credit'].sum() if not df_filtered.empty else 0.0
net_amount = df_filtered['net_amount'].sum() if not df_filtered.empty else 0.0

col1.metric("Total Transactions", f"{total_tx:,}")
col2.metric("Total Debit (₹)", f"₹ {total_debit:,.2f}")
col3.metric("Total Credit (₹)", f"₹ {total_credit:,.2f}")
col4.metric("Net Payable / Balance (₹)", f"₹ {net_amount:,.2f}")

st.markdown("<br>", unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# 7. EXPORT BUTTONS & INTERACTIVE TABLE
# -----------------------------------------------------------------------------
tab_table, tab_summary = st.tabs(["📋 Filtered Ledger Transactions", "📊 Vendor / Branch Summary"])

display_columns = [
    'date', 'branch_name', 'contact_id', 'account_name', 
    'transaction_details', 'transaction_type', 'reference_number', 
    'debit', 'credit', 'net_amount'
]
present_cols = [c for c in display_columns if c in df_filtered.columns]

with tab_table:
    exp_col1, exp_col2, exp_col3 = st.columns([2, 1, 1])
    
    with exp_col1:
        if len(df_filtered) > 10000:
            st.markdown(f"Displaying top **10,000** of **{len(df_filtered):,}** matching transactions (Full dataset available in downloads)")
        else:
            st.markdown(f"Displaying **{len(df_filtered):,}** of **{len(df_raw):,}** transactions")
        
    with exp_col2:
        st.download_button(
            label="📥 Download CSV",
            data=convert_df_to_csv(df_filtered[present_cols]),
            file_name=f"General_Ledger_Filtered_{datetime.date.today()}.csv",
            mime="text/csv",
            use_container_width=True
        )
        
    with exp_col3:
        st.download_button(
            label="📊 Download Excel (.xlsx)",
            data=convert_df_to_excel(df_filtered[present_cols]),
            file_name=f"General_Ledger_Filtered_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # Slice top 10,000 rows for dataframe UI rendering to keep memory & DOM rendering ultra fast
    df_display = df_filtered[present_cols].head(10000)

    st.dataframe(
        df_display,
        column_config={
            "date": st.column_config.DateColumn("Date", format="DD/MM/YYYY"),
            "branch_name": st.column_config.TextColumn("Branch"),
            "contact_id": st.column_config.TextColumn("Vendor (`contact_id`)"),
            "account_name": st.column_config.TextColumn("Account Name"),
            "transaction_details": st.column_config.TextColumn("Details"),
            "transaction_type": st.column_config.TextColumn("Type"),
            "reference_number": st.column_config.TextColumn("Voucher / Ref No"),
            "debit": st.column_config.NumberColumn("Debit (₹)", format="₹ %.2f"),
            "credit": st.column_config.NumberColumn("Credit (₹)", format="₹ %.2f"),
            "net_amount": st.column_config.NumberColumn("Net Amount (₹)", format="₹ %.2f")
        },
        use_container_width=True,
        hide_index=True,
        height=550
    )


with tab_summary:
    st.subheader("Aggregated Summary by Vendor / Details & Branch")
    group_col = 'transaction_details' if 'transaction_details' in df_filtered.columns else 'contact_id'
    if not df_filtered.empty and group_col in df_filtered.columns:
        summary_df = df_filtered.groupby([group_col, 'branch_name'], as_index=False).agg(
            Transactions=('net_amount', 'count'),
            Total_Debit=('debit', 'sum'),
            Total_Credit=('credit', 'sum'),
            Net_Balance=('net_amount', 'sum')
        ).sort_values(by='Net_Balance', ascending=False)

        st.dataframe(
            summary_df,
            column_config={
                group_col: st.column_config.TextColumn(f"Vendor / Details (`{group_col}`)"),
                "branch_name": st.column_config.TextColumn("Branch"),
                "Transactions": st.column_config.NumberColumn("Transactions Count"),
                "Total_Debit": st.column_config.NumberColumn("Total Debit (₹)", format="₹ %.2f"),
                "Total_Credit": st.column_config.NumberColumn("Total Credit (₹)", format="₹ %.2f"),
                "Net_Balance": st.column_config.NumberColumn("Net Balance (₹)", format="₹ %.2f"),
            },
            use_container_width=True,
            hide_index=True,
            height=450
        )
    else:
        st.info("No records available to summarize.")
