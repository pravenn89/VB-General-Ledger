import os
import glob
import pandas as pd
import streamlit as st

# Default lookup directories for General Ledger Excel files
DEFAULT_DATA_DIRS = [
    "./Data",
    "./data",
    "Data",
    "data",
    r"C:\Users\rtrpr\.gemini\antigravity\scratch\VB General Ledger\Data"
]

def get_default_data_dir() -> str:
    """Find the first existing data directory from default options, returning clean relative path if possible."""
    for path in DEFAULT_DATA_DIRS:
        if os.path.exists(path) and os.path.isdir(path):
            return path
    return "./Data"


def find_excel_files(folder_path: str) -> list:
    """
    Smartly locate Excel files in folder_path, its subfolders (like Data/), or recursively.
    """
    if not folder_path or not os.path.exists(folder_path):
        return []
        
    # 1. Direct search in folder_path
    files = glob.glob(os.path.join(folder_path, "*.xlsx")) + glob.glob(os.path.join(folder_path, "*.xls"))
    # Filter out temporary Excel locking files (starting with ~$)
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    if files:
        return sorted(files)
        
    # 2. Check for Data/data subfolder inside folder_path
    for sub in ["Data", "data", "DATA"]:
        sub_path = os.path.join(folder_path, sub)
        if os.path.exists(sub_path):
            sub_files = glob.glob(os.path.join(sub_path, "*.xlsx")) + glob.glob(os.path.join(sub_path, "*.xls"))
            sub_files = [f for f in sub_files if not os.path.basename(f).startswith("~$")]
            if sub_files:
                return sorted(sub_files)
                
    # 3. Fallback: Recursive search in all subdirectories
    rec_files = glob.glob(os.path.join(folder_path, "**", "*.xlsx"), recursive=True) + \
                glob.glob(os.path.join(folder_path, "**", "*.xls"), recursive=True)
    rec_files = [f for f in rec_files if not os.path.basename(f).startswith("~$")]
    return sorted(rec_files)


def load_single_excel(file_source, source_name: str) -> pd.DataFrame:
    """
    Read a single Excel file skipping title row (header=1).
    Clean column names and format datatypes.
    """
    df = pd.read_excel(file_source, header=1)
    
    if df.empty:
        return pd.DataFrame()
        
    # Standardize column names: lowercase and stripped whitespace
    df.columns = [str(c).strip().lower() for c in df.columns]
    
    # Required core columns specification & fallbacks
    numeric_cols = ['debit', 'credit', 'net_amount']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
        else:
            df[col] = 0.0
            
    # Format date column
    if 'date' in df.columns:
        df['date'] = pd.to_datetime(df['date'], errors='coerce')
    else:
        df['date'] = pd.NaT

    # Ensure text string columns are clean
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


@st.cache_data(show_spinner="Scanning and loading General Ledger Excel files...")
def load_ledger_data_cached(folder_path: str = None):
    """
    Internal cached function to ingest, clean, aggregate, and deduplicate ledger files.
    """
    dfs = []
    processed_files = []
    errors = []
    
    excel_files = find_excel_files(folder_path) if folder_path else []
    
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
    
    # Deduplicate identical records
    dedup_cols = [c for c in combined_df.columns if c != '_source_file']
    combined_df = combined_df.drop_duplicates(subset=dedup_cols).reset_index(drop=True)
    unique_rows = len(combined_df)
    
    # Sort by date ascending by default
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


def load_ledger_data(folder_path: str = None, uploaded_files=None):
    """
    Public data loading entry point handling directory search or Streamlit uploaded files.
    """
    if uploaded_files:
        dfs = []
        processed_files = []
        errors = []
        for file_obj in uploaded_files:
            filename = file_obj.name
            try:
                sub_df = load_single_excel(file_obj, filename)
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

    # Folder mode
    target_dir = folder_path if folder_path else get_default_data_dir()
    return load_ledger_data_cached(folder_path=target_dir)


def extract_unique_vendors(df: pd.DataFrame) -> list:
    """
    Extract unique, non-blank vendor values from contact_id.
    """
    vendors = set()
    if 'contact_id' in df.columns:
        valid_contacts = df['contact_id'].dropna().astype(str).str.strip()
        valid = valid_contacts[~valid_contacts.str.lower().isin(['', 'nan', 'none', 'nat', '-1', '0'])]
        vendors.update(valid.tolist())
        
    sorted_vendors = sorted(list(vendors), key=lambda x: str(x).upper())
    return sorted_vendors
