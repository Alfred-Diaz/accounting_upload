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

# Bank-specific header names mapped to the main headers above.
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

MIN_HEADER_MATCHES = 2


def canonical_header(value):
    """Normalize headers so matching is based on text only, not filename or exact casing."""
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
    """Detect bank format only from uploaded file headers."""
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

    output.insert(0, "Bank", bank)
    output.insert(1, "Source File", uploaded_file.name)
    return output, missing_source_headers, detection_results


def to_excel_bytes(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Consolidated")
        workbook = writer.book
        worksheet = writer.sheets["Consolidated"]
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        money_fmt = workbook.add_format({"num_format": "#,##0.00"})
        date_fmt = workbook.add_format({"num_format": "yyyy-mm-dd"})
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, header_fmt)
            width = max(12, min(35, len(str(col_name)) + 4))
            worksheet.set_column(col_num, col_num, width)
        for col in ["Debit", "Credit", "Running Balance"]:
            if col in df.columns:
                idx = df.columns.get_loc(col)
                worksheet.set_column(idx, idx, 16, money_fmt)
        if "Posting Date" in df.columns:
            idx = df.columns.get_loc("Posting Date")
            worksheet.set_column(idx, idx, 15, date_fmt)
    return buffer.getvalue()


def get_config_value(key, default=None):
    """Read configuration from Streamlit secrets first, then environment variables."""
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

    st.title("Bank Statement Consolidator")
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


st.set_page_config(page_title="Bank Statement Consolidator", layout="wide")

if not require_login():
    st.stop()

st.title("Bank Statement Consolidator")
st.write("Upload bank statement files. The app detects the bank format from headers only, consolidates the rows, and prepares the download.")

with st.expander("Current standard headers and bank mappings", expanded=False):
    st.write("Main headers:", MAIN_HEADERS)
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
                    "matched_headers": next(
                        item["score"] for item in detection_results if item["bank"] == detected_bank
                    ),
                }
            )
            if missing:
                issues.append({"file": file.name, "missing source headers": missing})
        except Exception as exc:
            issues.append({"file": file.name, "error": str(exc)})

    if consolidated_frames:
        consolidated = pd.concat(consolidated_frames, ignore_index=True)
        excel_bytes = to_excel_bytes(consolidated)

        st.success(f"Consolidated {len(consolidated_frames)} file(s), {len(consolidated):,} row(s). Your download is ready.")
        st.download_button(
            "⬇️ Download consolidated Excel",
            data=excel_bytes,
            file_name="consolidated_bank_statements.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

        with st.expander("Preview consolidated data", expanded=True):
            st.dataframe(consolidated, use_container_width=True)

        with st.expander("Bank detection summary", expanded=False):
            st.json(detection_summary)

    if issues:
        st.warning("Review these mapping or detection issues:")
        st.json(issues)
else:
    st.info("Upload one or more bank statement files to start.")
