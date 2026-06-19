import hashlib
import hmac
import io
import os
import sqlite3
from datetime import datetime

import pandas as pd
import streamlit as st

DATABASE_PATH = "accounting.db"
MIN_HEADER_MATCHES = 2
DEFAULT_PASSWORD_HASH = "008c70392e3abfbd0fa47bbc2ed96aa99bd49e159727fcba0f2e6abeb3a9d601"

MAIN_HEADERS = ["Posting Date", "Branch", "Description", "Debit", "Credit", "Running Balance", "Check Number"]

BANK_MAPPINGS = {
    "METROBANK": {
        "Posting Date": "Posting Date",
        "Branch/Channel": "Branch",
        "Transaction Description": "Description",
        "Debit Amount": "Debit",
        "Credit Amount": "Credit",
        "Balance": "Running Balance",
        "Check Number": "Check Number",
    },
    "EASTWEST": {
        "Value Date": "Posting Date",
        "Descript": "Branch",
        "Reference": "Description",
        "Debit": "Debit",
        "Credit": "Credit",
        "Closing Balance": "Running Balance",
        "Cheque Number": "Check Number",
    },
    "BDO": {
        "Posting Date": "Posting Date",
        "Branch": "Branch",
        "Description": "Description",
        "Debit": "Debit",
        "Credit": "Credit",
        "Running Balance": "Running Balance",
        "Check Number": "Check Number",
    },
}


def get_secret_or_env(key, default=None):
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return os.getenv(key, default)


def check_password(username, password):
    configured_user = get_secret_or_env("APP_USERNAME", "Accounting")
    configured_hash = get_secret_or_env("APP_PASSWORD_SHA256", DEFAULT_PASSWORD_HASH)
    password_hash = hashlib.sha256(password.encode("utf-8")).hexdigest()
    return hmac.compare_digest(username, configured_user) and hmac.compare_digest(password_hash, configured_hash)


def require_login():
    if st.session_state.get("authenticated"):
        with st.sidebar:
            st.success(f"Signed in as {st.session_state.get('username', 'Accounting')}")
            if st.button("Log out"):
                st.session_state.clear()
                st.rerun()
        return True
    st.title("Accounting System")
    st.write("Please sign in to continue.")
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Sign in")
    if submitted:
        if check_password(username, password):
            st.session_state["authenticated"] = True
            st.session_state["username"] = username
            st.rerun()
        else:
            st.error("Invalid username or password.")
    return False


def get_db():
    conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS uploads (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT NOT NULL,
            bank TEXT NOT NULL,
            account_name TEXT,
            account_code TEXT,
            uploaded_by TEXT,
            uploaded_at TEXT NOT NULL,
            row_count INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            upload_id INTEGER NOT NULL,
            bank TEXT NOT NULL,
            account_name TEXT,
            account_code TEXT,
            source_file TEXT NOT NULL,
            posting_date TEXT,
            month TEXT,
            branch TEXT,
            description TEXT,
            debit REAL DEFAULT 0,
            credit REAL DEFAULT 0,
            running_balance REAL DEFAULT 0,
            check_number TEXT,
            FOREIGN KEY(upload_id) REFERENCES uploads(id)
        );
        """
    )
    conn.commit()
    conn.close()


def query_df(sql, params=()):
    conn = get_db()
    df = pd.read_sql_query(sql, conn, params=params)
    conn.close()
    return df


def canonical_header(value):
    return " ".join(str(value).strip().upper().replace("_", " ").replace("-", " ").split())


def read_uploaded_file(uploaded_file):
    uploaded_file.seek(0)
    if uploaded_file.name.lower().endswith(".csv"):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file)


def normalize_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def detect_bank_from_headers(df):
    uploaded_headers = {canonical_header(column) for column in df.columns}
    detection_results = []
    for bank, mapping in BANK_MAPPINGS.items():
        expected_headers = {canonical_header(source_header) for source_header in mapping.keys()}
        matched_headers = uploaded_headers.intersection(expected_headers)
        detection_results.append({"bank": bank, "score": len(matched_headers), "matched_headers": sorted(matched_headers)})
    detection_results.sort(key=lambda item: item["score"], reverse=True)
    best = detection_results[0]
    tied_best = [item for item in detection_results if item["score"] == best["score"]]
    if best["score"] < MIN_HEADER_MATCHES:
        return None, detection_results, "Not enough matching headers to detect the bank."
    if len(tied_best) > 1:
        tied_names = ", ".join(item["bank"] for item in tied_best)
        return None, detection_results, f"Bank detection is ambiguous between: {tied_names}."
    return best["bank"], detection_results, None


def detect_account_from_filename(filename, bank):
    upper_name = filename.upper()
    if "709" in upper_name or "BDO MAIN" in upper_name:
        return "BDO MAIN", "709"
    if "253" in upper_name or "METROBANK" in upper_name:
        return "METROBANK", "253"
    if bank == "BDO":
        return "BDO MAIN", "709"
    if bank == "METROBANK":
        return "METROBANK", "253"
    return bank, ""


def build_mapping(df, bank):
    uploaded_lookup = {canonical_header(column): column for column in df.columns}
    mapping = {}
    missing = []
    for source_header, target_header in BANK_MAPPINGS[bank].items():
        actual_column = uploaded_lookup.get(canonical_header(source_header))
        if actual_column:
            mapping[actual_column] = target_header
        else:
            missing.append(source_header)
    return mapping, missing


def prepare_accounting_fields(df):
    df = df.copy()
    df["Posting Date"] = pd.to_datetime(df["Posting Date"], errors="coerce")
    for col in ["Debit", "Credit", "Running Balance"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df["Month"] = df["Posting Date"].dt.to_period("M").astype(str).replace("NaT", "")
    df["Check Number"] = df["Check Number"].fillna("").astype(str).str.strip()
    df["Description"] = df["Description"].fillna("").astype(str)
    df["Branch"] = df["Branch"].fillna("").astype(str)
    return df


def consolidate_file(uploaded_file):
    raw_df = normalize_columns(read_uploaded_file(uploaded_file))
    bank, detection_results, detection_error = detect_bank_from_headers(raw_df)
    if detection_error:
        raise ValueError(f"{detection_error} File: {uploaded_file.name}")
    mapping, missing = build_mapping(raw_df, bank)
    renamed = raw_df.rename(columns=mapping)
    output = pd.DataFrame()
    for header in MAIN_HEADERS:
        output[header] = renamed[header] if header in renamed.columns else pd.NA
    account_name, account_code = detect_account_from_filename(uploaded_file.name, bank)
    output.insert(0, "Bank", bank)
    output.insert(1, "Account Name", account_name)
    output.insert(2, "Account Code", account_code)
    output.insert(3, "Source File", uploaded_file.name)
    return prepare_accounting_fields(output), missing, detection_results


def save_upload(frame, username):
    conn = get_db()
    first = frame.iloc[0]
    cursor = conn.execute(
        "INSERT INTO uploads (filename, bank, account_name, account_code, uploaded_by, uploaded_at, row_count) VALUES (?, ?, ?, ?, ?, ?, ?)",
        (first["Source File"], first["Bank"], first["Account Name"], first["Account Code"], username, datetime.now().isoformat(timespec="seconds"), len(frame)),
    )
    upload_id = cursor.lastrowid
    rows = []
    for _, row in frame.iterrows():
        posting_date = "" if pd.isna(row["Posting Date"]) else row["Posting Date"].strftime("%Y-%m-%d")
        rows.append((upload_id, row["Bank"], row["Account Name"], row["Account Code"], row["Source File"], posting_date, row["Month"], row["Branch"], row["Description"], float(row["Debit"]), float(row["Credit"]), float(row["Running Balance"]), row["Check Number"]))
    conn.executemany(
        "INSERT INTO transactions (upload_id, bank, account_name, account_code, source_file, posting_date, month, branch, description, debit, credit, running_balance, check_number) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        rows,
    )
    conn.commit()
    conn.close()


def consolidated_df():
    return query_df(
        """
        SELECT bank AS Bank, account_name AS 'Account Name', account_code AS 'Account Code', source_file AS 'Source File', posting_date AS 'Posting Date', branch AS Branch, description AS Description, debit AS Debit, credit AS Credit, running_balance AS 'Running Balance', check_number AS 'Check Number', month AS Month
        FROM transactions
        ORDER BY posting_date, id
        """
    )


def uploads_df():
    return query_df("SELECT id AS ID, filename AS Filename, bank AS Bank, account_name AS 'Account Name', account_code AS 'Account Code', uploaded_by AS 'Uploaded By', uploaded_at AS 'Uploaded At', row_count AS Rows FROM uploads ORDER BY id DESC")


def filter_text(df, columns, query):
    if not query or df.empty:
        return df
    mask = pd.Series(False, index=df.index)
    query = str(query).strip().lower()
    for col in columns:
        if col in df.columns:
            mask = mask | df[col].fillna("").astype(str).str.lower().str.contains(query, na=False, regex=False)
    return df[mask]


def checks_df(query=""):
    df = consolidated_df()
    if df.empty:
        return df
    df = df[df["Check Number"].fillna("").astype(str).str.strip() != ""]
    return filter_text(df, ["Check Number", "Description", "Source File", "Bank", "Account Code"], query)


def payments_df(query="", account_code="All"):
    df = consolidated_df()
    if df.empty:
        return df
    df = df[df["Credit"] > 0].copy()
    if account_code != "All":
        df = df[df["Account Code"].astype(str) == str(account_code)]
    if query:
        amount_mask = df["Credit"].astype(str).str.contains(query, na=False, regex=False)
        text_result = filter_text(df, ["Description", "Source File", "Bank", "Account Name", "Account Code"], query)
        df = df[amount_mask | df.index.isin(text_result.index)]
    return df


def recon_df(bank="All", month="All"):
    df = consolidated_df()
    if df.empty:
        return pd.DataFrame()
    result = df.groupby(["Bank", "Account Name", "Account Code", "Month"], dropna=False).agg(
        Transactions=("Source File", "count"),
        **{"Total Debit": ("Debit", "sum"), "Total Credit": ("Credit", "sum"), "Ending Balance": ("Running Balance", "last")},
    ).reset_index()
    result["Net Movement"] = result["Total Credit"] - result["Total Debit"]
    if bank != "All":
        result = result[result["Bank"] == bank]
    if month != "All":
        result = result[result["Month"] == month]
    return result


def to_excel_bytes(sheets):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        money_fmt = workbook.add_format({"num_format": "#,##0.00"})
        for sheet_name, df in sheets.items():
            safe_sheet_name = sheet_name[:31]
            df.to_excel(writer, index=False, sheet_name=safe_sheet_name)
            worksheet = writer.sheets[safe_sheet_name]
            for col_num, col_name in enumerate(df.columns):
                worksheet.write(0, col_num, col_name, header_fmt)
                worksheet.set_column(col_num, col_num, max(12, min(35, len(str(col_name)) + 4)))
            for col in ["Debit", "Credit", "Running Balance", "Total Debit", "Total Credit", "Ending Balance", "Net Movement"]:
                if col in df.columns:
                    idx = df.columns.get_loc(col)
                    worksheet.set_column(idx, idx, 16, money_fmt)
    return buffer.getvalue()


def download_button(label, sheets, filename, primary=False):
    st.download_button(label, data=to_excel_bytes(sheets), file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary" if primary else "secondary", use_container_width=True)


def main():
    st.set_page_config(page_title="Accounting System", layout="wide")
    init_db()
    if not require_login():
        st.stop()

    st.title("Accounting System")
    with st.sidebar:
        page = st.radio("Menu", ["Dashboard", "Upload Bank Files", "Bank Consolidation", "Check No. Tracking", "Payment Verification", "Bank Recon", "Settings"])

    if page == "Dashboard":
        consolidated = consolidated_df()
        uploads = uploads_df()
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Uploads", len(uploads))
        col2.metric("Transactions", len(consolidated))
        col3.metric("Total Debit", f"{consolidated['Debit'].sum():,.2f}" if not consolidated.empty else "0.00")
        col4.metric("Total Credit", f"{consolidated['Credit'].sum():,.2f}" if not consolidated.empty else "0.00")
        st.subheader("Upload History")
        st.dataframe(uploads, use_container_width=True)
        if not consolidated.empty:
            download_button("Download full accounting workbook", {"Consolidated": consolidated, "Payment Verification": payments_df(), "Check Tracking": checks_df(), "Bank Recon": recon_df()}, "accounting_bank_workbook.xlsx", primary=True)

    elif page == "Upload Bank Files":
        st.subheader("Upload Bank Files")
        uploaded_files = st.file_uploader("Upload Excel or CSV bank files", type=["xlsx", "xls", "csv"], accept_multiple_files=True)
        if uploaded_files and st.button("Process and save uploaded files", type="primary"):
            saved_files = 0
            saved_rows = 0
            issues = []
            detection_summary = []
            for file in uploaded_files:
                try:
                    frame, missing, detection_results = consolidate_file(file)
                    save_upload(frame, st.session_state.get("username", "Accounting"))
                    saved_files += 1
                    saved_rows += len(frame)
                    detected_bank = frame["Bank"].iloc[0]
                    detection_summary.append({"file": file.name, "detected_bank": detected_bank, "account_name": frame["Account Name"].iloc[0], "account_code": frame["Account Code"].iloc[0], "matched_headers": next(item["score"] for item in detection_results if item["bank"] == detected_bank)})
                    if missing:
                        issues.append({"file": file.name, "warning": f"Missing source headers: {', '.join(missing)}"})
                except Exception as exc:
                    issues.append({"file": file.name, "error": str(exc)})
            if saved_files:
                st.success(f"Saved {saved_files} file(s) with {saved_rows:,} transaction row(s).")
            if detection_summary:
                st.dataframe(pd.DataFrame(detection_summary), use_container_width=True)
            if issues:
                st.warning("Review these issues:")
                st.json(issues)

    elif page == "Bank Consolidation":
        st.subheader("Bank Consolidation")
        q = st.text_input("Search description, source file, bank, account code, or check number")
        df = filter_text(consolidated_df(), ["Description", "Source File", "Bank", "Account Code", "Check Number"], q)
        st.caption(f"Showing {len(df):,} transaction(s).")
        st.dataframe(df, use_container_width=True)
        if not df.empty:
            download_button("Download consolidated result", {"Consolidated": df}, "bank_consolidated.xlsx")

    elif page == "Check No. Tracking":
        st.subheader("Check No. Tracking")
        q = st.text_input("Search by check number, description, source file, bank, or account code")
        df = checks_df(q)
        st.caption(f"Showing {len(df):,} check transaction(s).")
        st.dataframe(df, use_container_width=True)
        if not df.empty:
            download_button("Download check tracking result", {"Check Tracking": df}, "check_no_tracking.xlsx")

    elif page == "Payment Verification":
        st.subheader("Payment Verification")
        st.info("Credit-side transactions only. BDO MAIN = 709. METROBANK = 253. Codes are detected from filename when available.")
        all_data = consolidated_df()
        codes = ["All"]
        if not all_data.empty:
            codes += sorted([str(x) for x in all_data["Account Code"].dropna().unique().tolist() if str(x) != ""])
        account_code = st.selectbox("Account code", codes)
        q = st.text_input("Search amount, description, source file, bank, or account")
        df = payments_df(q, account_code)
        st.metric("Total credits found", f"{df['Credit'].sum():,.2f}" if not df.empty else "0.00")
        st.caption(f"Showing {len(df):,} credit transaction(s).")
        st.dataframe(df, use_container_width=True)
        if not df.empty:
            download_button("Download payment verification result", {"Payment Verification": df}, "payment_verification.xlsx")

    elif page == "Bank Recon":
        st.subheader("Bank Recon")
        all_data = consolidated_df()
        banks = ["All"]
        months = ["All"]
        if not all_data.empty:
            banks += sorted(all_data["Bank"].dropna().unique().tolist())
            months += sorted(all_data["Month"].dropna().unique().tolist())
        col1, col2 = st.columns(2)
        with col1:
            bank = st.selectbox("Bank", banks)
        with col2:
            month = st.selectbox("Month", months)
        df = recon_df(bank, month)
        st.caption(f"Showing {len(df):,} recon row(s).")
        st.dataframe(df, use_container_width=True)
        if not df.empty:
            download_button("Download bank recon result", {"Bank Recon": df}, "bank_recon.xlsx")

    elif page == "Settings":
        st.subheader("Settings")
        st.write("Secrets format for Streamlit Cloud:")
        st.code('APP_USERNAME = "Accounting"\nAPP_PASSWORD_SHA256 = "008c70392e3abfbd0fa47bbc2ed96aa99bd49e159727fcba0f2e6abeb3a9d601"', language="toml")


if __name__ == "__main__":
    main()
