import io
import hashlib
import hmac
import os

import pandas as pd
import streamlit as st

MAIN_HEADERS = [
    "Posting Date",
    "Branch",
    "Description",
    "Debit",
    "Credit",
    "Running Balance",
    "Check Number",
]

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

ACCOUNT_CODES = {
    "BDO MAIN": "709",
    "METROBANK": "253",
}

MIN_HEADER_MATCHES = 2


def canonical_header(value):
    return " ".join(str(value).strip().upper().replace("_", " ").replace("-", " ").split())


def read_uploaded_file(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
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
        detection_results.append(
            {
                "bank": bank,
                "score": len(matched_headers),
                "total_expected": len(expected_headers),
                "matched_headers": sorted(matched_headers),
            }
        )

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


def build_case_insensitive_mapping(df, bank):
    uploaded_lookup = {canonical_header(column): column for column in df.columns}
    mapping = {}
    missing_source_headers = []

    for source_header, target_header in BANK_MAPPINGS[bank].items():
        actual_column = uploaded_lookup.get(canonical_header(source_header))
        if actual_column:
            mapping[actual_column] = target_header
        else:
            missing_source_headers.append(source_header)

    return mapping, missing_source_headers


def consolidate_file(uploaded_file):
    raw_df = normalize_columns(read_uploaded_file(uploaded_file))
    bank, detection_results, detection_error = detect_bank_from_headers(raw_df)
    if detection_error:
        raise ValueError(f"{detection_error} File: {uploaded_file.name}")

    mapping, missing_source_headers = build_case_insensitive_mapping(raw_df, bank)
    renamed = raw_df.rename(columns=mapping)

    output = pd.DataFrame()
    for header in MAIN_HEADERS:
        output[header] = renamed[header] if header in renamed.columns else pd.NA

    account_name, account_code = detect_account_from_filename(uploaded_file.name, bank)
    output.insert(0, "Bank", bank)
    output.insert(1, "Account Name", account_name)
    output.insert(2, "Account Code", account_code)
    output.insert(3, "Source File", uploaded_file.name)
    return output, missing_source_headers, detection_results


def prepare_accounting_fields(df):
    df = df.copy()
    df["Posting Date"] = pd.to_datetime(df["Posting Date"], errors="coerce")
    for col in ["Debit", "Credit", "Running Balance"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df["Month"] = df["Posting Date"].dt.to_period("M").astype(str)
    df["Check Number"] = df["Check Number"].fillna("").astype(str).str.strip()
    df["Description"] = df["Description"].fillna("").astype(str)
    return df


def filter_text(df, columns, query):
    if not query:
        return df
    mask = pd.Series(False, index=df.index)
    query = str(query).strip().lower()
    for col in columns:
        if col in df.columns:
            mask = mask | df[col].fillna("").astype(str).str.lower().str.contains(query, na=False, regex=False)
    return df[mask]


def to_excel_bytes(sheets):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        money_fmt = workbook.add_format({"num_format": "#,##0.00"})
        date_fmt = workbook.add_format({"num_format": "yyyy-mm-dd"})

        for sheet_name, df in sheets.items():
            safe_sheet_name = sheet_name[:31]
            df.to_excel(writer, index=False, sheet_name=safe_sheet_name)
            worksheet = writer.sheets[safe_sheet_name]
            for col_num, col_name in enumerate(df.columns):
                worksheet.write(0, col_num, col_name, header_fmt)
                width = max(12, min(35, len(str(col_name)) + 4))
                worksheet.set_column(col_num, col_num, width)
            for col in ["Debit", "Credit", "Running Balance", "Total Debit", "Total Credit", "Ending Balance", "Net Movement"]:
                if col in df.columns:
                    idx = df.columns.get_loc(col)
                    worksheet.set_column(idx, idx, 16, money_fmt)
            if "Posting Date" in df.columns:
                idx = df.columns.get_loc("Posting Date")
                worksheet.set_column(idx, idx, 15, date_fmt)
    return buffer.getvalue()


def get_config_value(key, default=None):
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return os.getenv(key, default)


def check_password(username, password):
    configured_user = get_config_value("APP_USERNAME", "Accounting")
    configured_hash = get_config_value("APP_PASSWORD_SHA256", "008c70392e3abfbd0fa47bbc2ed96aa99bd49e159727fcba0f2e6abeb3a9d601")
    password_hash = hashlib.sha256(password.encode("utf-8")).hexdigest()
    return hmac.compare_digest(username, configured_user) and hmac.compare_digest(password_hash, configured_hash)


def require_login():
    if st.session_state.get("authenticated"):
        with st.sidebar:
            st.success(f"Signed in as {st.session_state.get('username', 'user')}")
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


st.set_page_config(page_title="Accounting System", layout="wide")

if not require_login():
    st.stop()

st.title("Accounting System")
st.write("Upload bank statement files once, then use Bank Consolidation, Payment Verification, and Bank Recon tabs.")

with st.expander("Current standard headers and bank mappings", expanded=False):
    st.write("Main headers:", MAIN_HEADERS)
    st.write("Payment verification account codes:", ACCOUNT_CODES)
    st.json(BANK_MAPPINGS)

uploaded_files = st.file_uploader(
    "Upload bank files", type=["xlsx", "xls", "csv"], accept_multiple_files=True
)

if uploaded_files:
    consolidated_frames = []
    issues = []
    detection_summary = []

    for file in uploaded_files:
        try:
            frame, missing, detection_results = consolidate_file(file)
            consolidated_frames.append(frame)
            detected_bank = frame["Bank"].iloc[0]
            detection_summary.append(
                {
                    "file": file.name,
                    "detected_bank": detected_bank,
                    "account_name": frame["Account Name"].iloc[0],
                    "account_code": frame["Account Code"].iloc[0],
                    "matched_headers": next(item["score"] for item in detection_results if item["bank"] == detected_bank),
                }
            )
            if missing:
                issues.append({"file": file.name, "missing source headers": missing})
        except Exception as exc:
            issues.append({"file": file.name, "error": str(exc)})

    if consolidated_frames:
        consolidated = prepare_accounting_fields(pd.concat(consolidated_frames, ignore_index=True))
        credit_summary = consolidated[consolidated["Credit"] > 0].copy()
        check_tracking = consolidated[consolidated["Check Number"] != ""].copy()
        recon_summary = (
            consolidated.groupby(["Bank", "Account Name", "Account Code", "Month"], dropna=False)
            .agg(
                Transactions=("Source File", "count"),
                Total_Debit=("Debit", "sum"),
                Total_Credit=("Credit", "sum"),
                Ending_Balance=("Running Balance", "last"),
            )
            .reset_index()
            .rename(columns={"Total_Debit": "Total Debit", "Total_Credit": "Total Credit", "Ending_Balance": "Ending Balance"})
        )
        recon_summary["Net Movement"] = recon_summary["Total Credit"] - recon_summary["Total Debit"]

        full_workbook = to_excel_bytes(
            {
                "Consolidated": consolidated,
                "Payment Verification": credit_summary,
                "Check Tracking": check_tracking,
                "Bank Recon": recon_summary,
            }
        )

        st.success(f"Processed {len(consolidated_frames)} file(s), {len(consolidated):,} row(s). Your accounting workbook is ready.")
        st.download_button(
            "⬇️ Download full accounting workbook",
            data=full_workbook,
            file_name="accounting_bank_workbook.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

        tab1, tab2, tab3, tab4 = st.tabs([
            "Bank Consolidation",
            "Check No. Tracking",
            "Payment Verification",
            "Bank Recon",
        ])

        with tab1:
            st.subheader("Bank Consolidation")
            st.dataframe(consolidated, use_container_width=True)
            st.download_button(
                "Download consolidated only",
                data=to_excel_bytes({"Consolidated": consolidated}),
                file_name="bank_consolidated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with tab2:
            st.subheader("Check No. Tracking")
            check_query = st.text_input("Search by check number, description, source file, or bank", key="check_query")
            check_result = filter_text(check_tracking, ["Check Number", "Description", "Source File", "Bank", "Account Code"], check_query)
            st.caption(f"Showing {len(check_result):,} check transaction(s).")
            st.dataframe(check_result, use_container_width=True)
            st.download_button(
                "Download check tracking result",
                data=to_excel_bytes({"Check Tracking": check_result}),
                file_name="check_no_tracking.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with tab3:
            st.subheader("Payment Verification")
            st.write("Credit-side transactions only. BDO MAIN uses account code 709; METROBANK uses account code 253, detected from the filename when available.")
            account_options = ["All"] + sorted([x for x in credit_summary["Account Code"].dropna().unique().tolist() if str(x) != ""])
            selected_account = st.selectbox("Account code", account_options)
            payment_query = st.text_input("Search credit transactions by amount, description, source file, bank, or account", key="payment_query")
            payment_result = credit_summary.copy()
            if selected_account != "All":
                payment_result = payment_result[payment_result["Account Code"].astype(str) == str(selected_account)]
            if payment_query:
                amount_mask = payment_result["Credit"].astype(str).str.contains(payment_query, na=False, regex=False)
                text_result = filter_text(payment_result, ["Description", "Source File", "Bank", "Account Name", "Account Code"], payment_query)
                payment_result = payment_result[amount_mask | payment_result.index.isin(text_result.index)]
            st.metric("Total credits found", f"{payment_result['Credit'].sum():,.2f}")
            st.caption(f"Showing {len(payment_result):,} credit transaction(s).")
            st.dataframe(payment_result, use_container_width=True)
            st.download_button(
                "Download payment verification result",
                data=to_excel_bytes({"Payment Verification": payment_result}),
                file_name="payment_verification.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with tab4:
            st.subheader("Bank Recon")
            st.write("Consolidated per bank and per month.")
            bank_options = ["All"] + sorted(consolidated["Bank"].dropna().unique().tolist())
            month_options = ["All"] + sorted(consolidated["Month"].dropna().unique().tolist())
            col1, col2 = st.columns(2)
            with col1:
                selected_bank = st.selectbox("Bank", bank_options)
            with col2:
                selected_month = st.selectbox("Month", month_options)
            recon_result = recon_summary.copy()
            if selected_bank != "All":
                recon_result = recon_result[recon_result["Bank"] == selected_bank]
            if selected_month != "All":
                recon_result = recon_result[recon_result["Month"] == selected_month]
            st.dataframe(recon_result, use_container_width=True)
            st.download_button(
                "Download bank recon result",
                data=to_excel_bytes({"Bank Recon": recon_result}),
                file_name="bank_recon.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with st.expander("Bank detection summary", expanded=False):
            st.json(detection_summary)

    if issues:
        st.warning("Review these mapping or detection issues:")
        st.json(issues)
else:
    st.info("Upload one or more bank statement files to start.")
