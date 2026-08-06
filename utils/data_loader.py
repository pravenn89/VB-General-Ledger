import io
import os
import glob
import json
import time
import urllib.request
import urllib.error
import pandas as pd
import streamlit as st

# Base directory of the application (root directory containing app.py and utils/)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Absolute path to Data directory
DEFAULT_DATA_DIR = os.path.join(BASE_DIR, "Data")

# GitHub Repository details for remote auto-loading
GITHUB_REPO_OWNER = "pravenn89"
GITHUB_REPO_NAME = "VB-General-Ledger"
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/contents/Data"

FALLBACK_DATA_DIRS = [
    DEFAULT_DATA_DIR,
    os.path.join(BASE_DIR, "data"),
    r"C:\Users\rtrpr\.gemini\antigravity\scratch\VB General Ledger\Data",
    r"C:\Users\rtrpr\.gemini\antigravity\scratch\VB General Ledger\data"
]

def fetch_url_with_retry(url: str, retries: int = 2, backoff: float = 0.5) -> bytes:
    """Helper to fetch URL content with automatic retries on network hiccups."""
    last_exception = None
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 StreamlitApp/1.0'})
            with urllib.request.urlopen(req, timeout=15) as res:
                return res.read()
        except Exception as e:
            last_exception = e
            time.sleep(backoff * (attempt + 1))
    raise last_exception


def get_default_data_dir() -> str:
    """Return the absolute path to the default Data directory."""
    for p in FALLBACK_DATA_DIRS:
        if os.path.exists(p) and os.path.isdir(p):
            return p
    return DEFAULT_DATA_DIR


def find_excel_files(folder_path: str = None) -> list:
    """
    Smartly locate Excel files in folder_path, BASE_DIR/Data, or recursively.
    """
    search_paths = []
    
    if folder_path and folder_path.strip():
        tp = folder_path.strip()
        if not os.path.isabs(tp):
            tp = os.path.abspath(os.path.join(BASE_DIR, tp))
        search_paths.append(tp)
        
    search_paths.extend(FALLBACK_DATA_DIRS)
    
    for target_path in search_paths:
        if not os.path.exists(target_path):
            continue
            
        files = glob.glob(os.path.join(target_path, "*.xlsx")) + glob.glob(os.path.join(target_path, "*.xls"))
        files = [f for f in files if not os.path.basename(f).startswith("~$")]
        if files:
            return sorted(files)
            
        for sub in ["Data", "data", "DATA"]:
            sub_path = os.path.join(target_path, sub)
            if os.path.exists(sub_path):
                sub_files = glob.glob(os.path.join(sub_path, "*.xlsx")) + glob.glob(os.path.join(sub_path, "*.xls"))
                sub_files = [f for f in sub_files if not os.path.basename(f).startswith("~$")]
                if sub_files:
                    return sorted(sub_files)
                    
        rec_files = glob.glob(os.path.join(target_path, "**", "*.xlsx"), recursive=True) + \
                    glob.glob(os.path.join(target_path, "**", "*.xls"), recursive=True)
        rec_files = [f for f in rec_files if not os.path.basename(f).startswith("~$")]
        if rec_files:
            return sorted(rec_files)

    return []


def load_single_excel(file_source, source_name: str) -> pd.DataFrame:
    """
    Read a single Excel file skipping title row (header=1).
    Clean column names and format datatypes.
    """
    df = pd.read_excel(file_source, header=1, engine="openpyxl")
    
    if df.empty:
        return pd.DataFrame()
        
    df.columns = [str(c).strip().lower() for c in df.columns]
    
    numeric_cols = ['debit', 'credit', 'net_amount']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
        else:
            df[col] = 0.0
            
    if 'date' in df.columns:
        df['date'] = pd.to_datetime(df['date'], errors='coerce')
    else:
        df['date'] = pd.NaT

    text_cols = [
        'account_name', 'transaction_details', 'transaction_type', 
        'reference_number', 'entity_number', 'contact_id', 
        'account_id', 'branch_name'
    ]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().replace({'nan': '', 'None': '', 'NAT': ''})
        else:
            df[col] = ''
            
    df['_source_file'] = source_name
    return df


@st.cache_data(show_spinner=False)
def parse_single_uploaded_file_cached(file_name: str, file_bytes: bytes) -> pd.DataFrame:
    """
    Cache parsing of an individual uploaded Excel file by its name and raw bytes content.
    """
    buffer = io.BytesIO(file_bytes)
    return load_single_excel(buffer, file_name)


@st.cache_data(show_spinner="Loading General Ledger Dataset...")
def load_ledger_data_cached(folder_path: str = None):
    """
    Internal cached function to ingest, clean, aggregate, and deduplicate local folder ledger files.
    Prefers pre-compiled Parquet / compressed CSV files for instant crash-free loading on Streamlit Cloud.
    """
    check_dirs = [folder_path, DEFAULT_DATA_DIR, BASE_DIR]
    for d in check_dirs:
        if d and os.path.exists(d):
            pq_file = os.path.join(d, "ledger_data.parquet")
            if not os.path.exists(pq_file):
                pq_file = os.path.join(d, "Data", "ledger_data.parquet")
            if os.path.exists(pq_file):
                try:
                    df = pd.read_parquet(pq_file)
                    stats = {
                        'files_count': 25,
                        'total_rows': len(df),
                        'unique_rows': len(df),
                        'processed_files': ['ledger_data.parquet (Optimized Parquet)'],
                        'errors': []
                    }
                    return df, stats
                except Exception:
                    pass

            gz_file = os.path.join(d, "ledger_data.csv.gz")
            if not os.path.exists(gz_file):
                gz_file = os.path.join(d, "Data", "ledger_data.csv.gz")
            if os.path.exists(gz_file):
                try:
                    df = pd.read_csv(gz_file, compression='gzip')
                    if 'date' in df.columns:
                        df['date'] = pd.to_datetime(df['date'], errors='coerce')
                    stats = {
                        'files_count': 25,
                        'total_rows': len(df),
                        'unique_rows': len(df),
                        'processed_files': ['ledger_data.csv.gz (Optimized Compressed CSV)'],
                        'errors': []
                    }
                    return df, stats
                except Exception:
                    pass

    dfs = []
    processed_files = []
    errors = []
    
    excel_files = find_excel_files(folder_path)
    
    for filepath in excel_files:
        filename = os.path.basename(filepath)
        try:
            sub_df = load_single_excel(filepath, filename)
            if not sub_df.empty:
                dfs.append(sub_df)
                processed_files.append(filename)
        except Exception as e:
            errors.append(f"Error reading {filename}: {str(e)}")

    if not dfs:
        empty_cols = [
            'date', 'account_name', 'transaction_details', 'transaction_type', 
            'reference_number', 'entity_number', 'debit', 'credit', 'net_amount', 
            'contact_id', 'account_id', 'branch_name', '_source_file'
        ]
        empty_df = pd.DataFrame(columns=empty_cols)
        stats = {
            'files_count': 0,
            'total_rows': 0,
            'unique_rows': 0,
            'processed_files': [],
            'errors': errors
        }
        return empty_df, stats

    combined_df = pd.concat(dfs, ignore_index=True)
    total_rows = len(combined_df)
    
    dedup_cols = [c for c in combined_df.columns if c != '_source_file']
    combined_df = combined_df.drop_duplicates(subset=dedup_cols).reset_index(drop=True)
    unique_rows = len(combined_df)
    
    if 'date' in combined_df.columns:
        combined_df = combined_df.sort_values(by='date', ascending=True).reset_index(drop=True)
        
    stats = {
        'files_count': len(processed_files),
        'total_rows': total_rows,
        'unique_rows': unique_rows,
        'processed_files': processed_files,
        'errors': errors
    }
    
    return combined_df, stats


@st.cache_data(show_spinner="Fetching ledger reports from GitHub remote...")
def load_github_ledger_data_cached():
    """
    Fetch and parse all General Ledger Excel reports directly from GitHub raw API with safe error handling.
    """
    dfs = []
    processed_files = []
    errors = []
    
    try:
        api_bytes = fetch_url_with_retry(GITHUB_API_URL, retries=2)
        contents = json.loads(api_bytes.decode('utf-8'))
        excel_items = [item for item in contents if item['name'].endswith(('.xlsx', '.xls')) and not item['name'].startswith('~$')]
        
        for item in excel_items:
            download_url = item.get('download_url')
            file_name = item.get('name')
            if download_url:
                try:
                    file_bytes = fetch_url_with_retry(download_url, retries=2)
                    buffer = io.BytesIO(file_bytes)
                    sub_df = load_single_excel(buffer, file_name)
                    if not sub_df.empty:
                        dfs.append(sub_df)
                        processed_files.append(file_name)
                except Exception as ex:
                    errors.append(f"Error fetching {file_name} from GitHub: {str(ex)}")

    except urllib.error.HTTPError as he:
        if he.code == 403:
            errors.append("GitHub API rate limit exceeded for unauthenticated requests. Please use Local / Repository Data mode.")
        else:
            errors.append(f"GitHub API Error: {he.code} - {he.reason}")
    except Exception as e:
        errors.append(f"Failed to query GitHub repository: {str(e)}")

    if not dfs:
        empty_cols = [
            'date', 'account_name', 'transaction_details', 'transaction_type', 
            'reference_number', 'entity_number', 'debit', 'credit', 'net_amount', 
            'contact_id', 'account_id', 'branch_name', '_source_file'
        ]
        empty_df = pd.DataFrame(columns=empty_cols)
        return empty_df, {
            'files_count': 0,
            'total_rows': 0,
            'unique_rows': 0,
            'processed_files': [],
            'errors': errors
        }

    combined_df = pd.concat(dfs, ignore_index=True)
    total_rows = len(combined_df)
    dedup_cols = [c for c in combined_df.columns if c != '_source_file']
    combined_df = combined_df.drop_duplicates(subset=dedup_cols).reset_index(drop=True)
    if 'date' in combined_df.columns:
        combined_df = combined_df.sort_values(by='date', ascending=True).reset_index(drop=True)
        
    stats = {
        'files_count': len(processed_files),
        'total_rows': total_rows,
        'unique_rows': len(combined_df),
        'processed_files': processed_files,
        'errors': errors
    }
    return combined_df, stats


def load_ledger_data(mode: str = "Local / Repository Data", folder_path: str = None, uploaded_files=None):
    """
    Public data loading entry point handling Local/Repository files, GitHub API stream, or Uploaded files.
    """
    if "GitHub API" in mode:
        return load_github_ledger_data_cached()

    if mode == "Upload Excel Files" and uploaded_files:
        dfs = []
        processed_files = []
        errors = []
        
        progress_text = f"Parsing {len(uploaded_files)} uploaded file(s)..."
        progress_bar = st.sidebar.progress(0, text=progress_text)
        
        for idx, file_obj in enumerate(uploaded_files):
            filename = file_obj.name
            try:
                file_bytes = file_obj.getvalue()
                sub_df = parse_single_uploaded_file_cached(filename, file_bytes)
                if not sub_df.empty:
                    dfs.append(sub_df)
                    processed_files.append(filename)
            except Exception as e:
                errors.append(f"Error reading {filename}: {str(e)}")
            
            progress_bar.progress((idx + 1) / len(uploaded_files), text=f"Processed {idx + 1}/{len(uploaded_files)} files")

        progress_bar.empty()

        if not dfs:
            empty_cols = [
                'date', 'account_name', 'transaction_details', 'transaction_type', 
                'reference_number', 'entity_number', 'debit', 'credit', 'net_amount', 
                'contact_id', 'account_id', 'branch_name', '_source_file'
            ]
            return pd.DataFrame(columns=empty_cols), {
                'files_count': 0,
                'total_rows': 0,
                'unique_rows': 0,
                'processed_files': [],
                'errors': errors
            }

        combined_df = pd.concat(dfs, ignore_index=True)
        total_rows = len(combined_df)
        dedup_cols = [c for c in combined_df.columns if c != '_source_file']
        combined_df = combined_df.drop_duplicates(subset=dedup_cols).reset_index(drop=True)
        if 'date' in combined_df.columns:
            combined_df = combined_df.sort_values(by='date', ascending=True).reset_index(drop=True)
            
        return combined_df, {
            'files_count': len(processed_files),
            'total_rows': total_rows,
            'unique_rows': len(combined_df),
            'processed_files': processed_files,
            'errors': errors
        }

    # Default: Local / Repository Data mode
    target_dir = folder_path if folder_path else get_default_data_dir()
    return load_ledger_data_cached(folder_path=target_dir)


@st.cache_data(show_spinner=False)
def extract_unique_vendors(df: pd.DataFrame) -> list:
    """
    Extract unique, non-blank vendor names efficiently with caching to prevent RAM overhead.
    """
    if df is None or df.empty:
        return []
        
    col = 'transaction_details' if 'transaction_details' in df.columns else 'contact_id'
    if col in df.columns:
        unique_vals = df[col].dropna().unique()
        cleaned = set()
        for v in unique_vals:
            s = str(v).strip()
            if s.lower() not in ['', 'nan', 'none', 'nat', '-1', '0', 'null']:
                cleaned.add(s)
        return sorted(list(cleaned), key=lambda x: x.upper())
    return []
