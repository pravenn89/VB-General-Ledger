import io
import os
import datetime
import pandas as pd
import streamlit as st

from utils.data_loader import (
    load_ledger_data, 
    extract_unique_vendors, 
    get_default_data_dir
)

# -----------------------------------------------------------------------------
# 1. PAGE CONFIG & STYLING
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="General Ledger Aggregator & Vendor Analysis",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for UI aesthetics
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

    /* Custom badge status */
    .status-badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 12px;
        font-size: 0.85rem;
        font-weight: 600;
        background-color: #E0F2FE;
        color: #0369A1;
        margin-bottom: 10px;
    }
    .status-badge-warn {
        background-color: #FEF3C7;
        color: #92400E;
    }

    /* Filter Card */
    .filter-header {
        font-size: 1.1rem;
        font-weight: 600;
        color: #1E293B;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# 2. SIDEBAR CONFIGURATION & DATA LOADING
# -----------------------------------------------------------------------------
st.sidebar.image("https://img.icons8.com/color/96/general-ledger.png", width=64)
st.sidebar.title("Ledger Settings & Filters")

# Mode selector: Folder vs File Upload
data_source_mode = st.sidebar.radio(
    "Data Source Mode",
    ["Local Directory", "Upload Excel Files"],
    help="Select whether to scan a folder on disk or upload files manually."
)

uploaded_files = None
folder_path = None

if data_source_mode == "Local Directory":
    default_dir = get_default_data_dir()
    folder_path = st.sidebar.text_input(
        "Data Folder Path",
        value=default_dir,
        help="Directory containing *.xlsx General Ledger reports."
    )
    if st.sidebar.button("🔄 Refresh / Reload Folder"):
        st.cache_data.clear()
        st.rerun()
else:
    uploaded_files = st.sidebar.file_uploader(
        "Upload General Ledger Excel Files",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

# Load raw ledger data
df_raw, load_stats = load_ledger_data(folder_path=folder_path, uploaded_files=uploaded_files)

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
# 3. FILTERS & SEARCH CONTROLS
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

    # Vendor Selector (contact_id)
    unique_vendors = extract_unique_vendors(df_raw)
    
    st.sidebar.markdown("**Vendor Selection (`contact_id`)**")
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
            "Search / Select Vendor (`contact_id`)",
            options=unique_vendors,
            placeholder="Type vendor contact_id..."
        )
    elif vendor_mode == "Keyword Search":
        vendor_keyword = st.sidebar.text_input(
            "Vendor / Account Keyword Search",
            placeholder="Enter vendor ID or account name..."
        )


# -----------------------------------------------------------------------------
# 4. FILTER EXECUTION LOGIC
# -----------------------------------------------------------------------------
df_filtered = df_raw.copy()

if not df_raw.empty:
    # Date filtering
    if selected_date_range and isinstance(selected_date_range, (tuple, list)):
        if len(selected_date_range) == 2:
            start_d, end_d = selected_date_range
            start_ts = pd.to_datetime(start_d)
            end_ts = pd.to_datetime(end_d) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            df_filtered = df_filtered[
                (df_filtered['date'] >= start_ts) & (df_filtered['date'] <= end_ts)
            ]
        elif len(selected_date_range) == 1:
            start_d = selected_date_range[0]
            start_ts = pd.to_datetime(start_d)
            df_filtered = df_filtered[df_filtered['date'] >= start_ts]

    # Branch filtering
    if selected_branches:
        df_filtered = df_filtered[df_filtered['branch_name'].isin(selected_branches)]

    # Transaction type filtering
    if selected_types:
        df_filtered = df_filtered[df_filtered['transaction_type'].isin(selected_types)]

    # Vendor filtering
    if vendor_mode == "Select Vendor(s)" and selected_vendors:
        df_filtered = df_filtered[df_filtered['contact_id'].isin(selected_vendors)]
    elif vendor_mode == "Keyword Search" and vendor_keyword.strip():
        kw = vendor_keyword.strip().lower()
        match_contact = df_filtered['contact_id'].str.lower().str.contains(kw, na=False)
        match_acc = df_filtered['account_name'].str.lower().str.contains(kw, na=False)
        match_details = df_filtered['transaction_details'].str.lower().str.contains(kw, na=False)
        df_filtered = df_filtered[match_contact | match_acc | match_details]


# -----------------------------------------------------------------------------
# 5. DASHBOARD MAIN CONTENT
# -----------------------------------------------------------------------------
# Main Title & Subtitle
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
# 6. EXPORT BUTTONS & INTERACTIVE TABLE
# -----------------------------------------------------------------------------
# Create tabs for table view and aggregated view
tab_table, tab_summary = st.tabs(["📋 Filtered Ledger Transactions", "📊 Vendor / Branch Summary"])

# Display columns ordered logically
display_columns = [
    'date', 'branch_name', 'contact_id', 'account_name', 
    'transaction_details', 'transaction_type', 'reference_number', 
    'debit', 'credit', 'net_amount'
]
# Filter to existing columns
present_cols = [c for c in display_columns if c in df_filtered.columns]

with tab_table:
    # Action bar for export buttons
    exp_col1, exp_col2, exp_col3 = st.columns([2, 1, 1])
    
    with exp_col1:
        st.markdown(f"Displaying **{len(df_filtered):,}** of **{len(df_raw):,}** transactions")
        
    with exp_col2:
        # Download as CSV
        csv_data = df_filtered[present_cols].to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download CSV",
            data=csv_data,
            file_name=f"General_Ledger_Filtered_{datetime.date.today()}.csv",
            mime="text/csv",
            use_container_width=True
        )
        
    with exp_col3:
        # Download as Excel (.xlsx)
        if len(df_filtered) > 150000:
            st.caption("⚠️ Large dataset (>150k rows). Excel build may take a moment.")
            
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_filtered[present_cols].to_excel(writer, index=False, sheet_name='Filtered_Ledger')
        excel_data = output.getvalue()
        
        st.download_button(
            label="📊 Download Excel (.xlsx)",
            data=excel_data,
            file_name=f"General_Ledger_Filtered_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # Render interactive DataFrame
    st.dataframe(
        df_filtered[present_cols],
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
    st.subheader("Aggregated Summary by Vendor & Branch")
    if not df_filtered.empty and 'contact_id' in df_filtered.columns:
        summary_df = df_filtered.groupby(['contact_id', 'branch_name'], as_index=False).agg(
            Transactions=('net_amount', 'count'),
            Total_Debit=('debit', 'sum'),
            Total_Credit=('credit', 'sum'),
            Net_Balance=('net_amount', 'sum')
        ).sort_values(by='Net_Balance', ascending=False)

        st.dataframe(
            summary_df,
            column_config={
                "contact_id": st.column_config.TextColumn("Vendor (`contact_id`)"),
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
