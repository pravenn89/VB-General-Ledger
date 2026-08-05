# 📊 General Ledger Aggregator & Vendor Analysis App

A Streamlit web application that ingests multi-file General Ledger Excel reports, cleans and standardizes ledger data, calculates key financial metrics, and enables multi-level filtering by Vendor (`contact_id`), Branch, Transaction Type, and Date Range.

## 🚀 Features

- **Multi-File Excel Ingestion**: Automatically scans `./Data` for General Ledger `.xlsx` reports and handles metadata header rows (`header=1`).
- **Data Caching & Deduplication**: Fast data loading powered by `@st.cache_data` with automatic deduplication of identical records.
- **Vendor Search & Filtering**: Fast search and selection from unique vendor values (`contact_id`) plus keyword searching across account names and descriptions.
- **Key Financial KPI Metrics**: Interactive summary cards displaying Total Transactions, Total Debit (₹), Total Credit (₹), and Net Payable / Balance (₹).
- **Dual Format Exporting**: Filtered ledgers can be exported directly to `.xlsx` (Excel) or `.csv`.
- **Modular Code Base**: Clean separation between UI logic (`app.py`) and data loader pipelines (`utils/data_loader.py`).

## 📁 Repository Structure

```
├── app.py                     # Main Streamlit web application
├── utils/
│   ├── __init__.py
│   └── data_loader.py         # Data ingestion, caching, and cleaning module
├── requirements.txt           # Required Python packages
├── .gitignore                 # Git ignore rules
└── README.md                  # Project documentation
```

## 🛠️ Installation & Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/pravenn89/VB-General-Ledger.git
   cd VB-General-Ledger
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Data Directory Setup:**
   Ensure your General Ledger Excel files (`*.xlsx`) are placed inside a `Data` directory at the project root.

4. **Run the Streamlit application:**
   ```bash
   streamlit run app.py
   ```

5. Access the app in your browser at `http://localhost:8501`.
