import hashlib
import hmac
import io
import os
from datetime import datetime

import pandas as pd
from flask import Flask, redirect, render_template, request, send_file, session, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "change-this-local-secret-key")

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
DATA_STORE = {}
ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv"}


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def canonical_header(value):
    return " ".join(str(value).strip().upper().replace("_", " ").replace("-", " ").split())


def get_config_value(key, default=None):
    return os.getenv(key, default)


def check_password(username, password):
    configured_user = get_config_value("APP_USERNAME", "Accounting")
    configured_hash = get_config_value(
        "APP_PASSWORD_SHA256",
        "008c70392e3abfbd0fa47bbc2ed96aa99bd49e159727fcba0f2e6abeb3a9d601",
    )
    password_hash = hashlib.sha256(password.encode("utf-8")).hexdigest()
    return hmac.compare_digest(username, configured_user) and hmac.compare_digest(password_hash, configured_hash)


def login_required():
    return session.get("authenticated") is True


def read_uploaded_file(file_storage):
    filename = file_storage.filename.lower()
    file_storage.stream.seek(0)
    if filename.endswith(".csv"):
        return pd.read_csv(file_storage)
    return pd.read_excel(file_storage)


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


def consolidate_file(file_storage):
    filename = secure_filename(file_storage.filename)
    raw_df = normalize_columns(read_uploaded_file(file_storage))
    bank, detection_results, detection_error = detect_bank_from_headers(raw_df)
    if detection_error:
        raise ValueError(f"{detection_error} File: {filename}")

    mapping, missing_source_headers = build_case_insensitive_mapping(raw_df, bank)
    renamed = raw_df.rename(columns=mapping)

    output = pd.DataFrame()
    for header in MAIN_HEADERS:
        output[header] = renamed[header] if header in renamed.columns else pd.NA

    account_name, account_code = detect_account_from_filename(filename, bank)
    output.insert(0, "Bank", bank)
    output.insert(1, "Account Name", account_name)
    output.insert(2, "Account Code", account_code)
    output.insert(3, "Source File", filename)
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


def build_reports(consolidated):
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
        .rename(
            columns={
                "Total_Debit": "Total Debit",
                "Total_Credit": "Total Credit",
                "Ending_Balance": "Ending Balance",
            }
        )
    )
    recon_summary["Net Movement"] = recon_summary["Total Credit"] - recon_summary["Total Debit"]
    return credit_summary, check_tracking, recon_summary


def to_excel_bytes(sheets):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        money_fmt = workbook.add_format({"num_format": "#,##0.00"})
        date_fmt = workbook.add_format({"num_format": "yyyy-mm-dd"})

        for sheet_name, df in sheets.items():
            safe_sheet_name = sheet_name[:31]
            export_df = df.copy()
            if "Posting Date" in export_df.columns:
                export_df["Posting Date"] = pd.to_datetime(export_df["Posting Date"], errors="coerce")
            export_df.to_excel(writer, index=False, sheet_name=safe_sheet_name)
            worksheet = writer.sheets[safe_sheet_name]
            for col_num, col_name in enumerate(export_df.columns):
                worksheet.write(0, col_num, col_name, header_fmt)
                width = max(12, min(35, len(str(col_name)) + 4))
                worksheet.set_column(col_num, col_num, width)
            for col in ["Debit", "Credit", "Running Balance", "Total Debit", "Total Credit", "Ending Balance", "Net Movement"]:
                if col in export_df.columns:
                    idx = export_df.columns.get_loc(col)
                    worksheet.set_column(idx, idx, 16, money_fmt)
            if "Posting Date" in export_df.columns:
                idx = export_df.columns.get_loc("Posting Date")
                worksheet.set_column(idx, idx, 15, date_fmt)
    buffer.seek(0)
    return buffer


def df_to_records(df, limit=500):
    view = df.head(limit).copy()
    for col in view.columns:
        if pd.api.types.is_datetime64_any_dtype(view[col]):
            view[col] = view[col].dt.strftime("%Y-%m-%d").fillna("")
    return view.fillna("").to_dict(orient="records"), list(view.columns)


def current_dataset():
    dataset_id = session.get("dataset_id")
    if not dataset_id:
        return None
    return DATA_STORE.get(dataset_id)


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        if check_password(username, password):
            session["authenticated"] = True
            session["username"] = username
            return redirect(url_for("index"))
        error = "Invalid username or password."
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/", methods=["GET", "POST"])
def index():
    if not login_required():
        return redirect(url_for("login"))

    errors = []
    detection_summary = []

    if request.method == "POST":
        files = request.files.getlist("files")
        consolidated_frames = []
        for file_storage in files:
            if not file_storage or not file_storage.filename:
                continue
            if not allowed_file(file_storage.filename):
                errors.append({"file": file_storage.filename, "error": "Unsupported file type."})
                continue
            try:
                frame, missing, detection_results = consolidate_file(file_storage)
                consolidated_frames.append(frame)
                detected_bank = frame["Bank"].iloc[0]
                detection_summary.append(
                    {
                        "file": secure_filename(file_storage.filename),
                        "detected_bank": detected_bank,
                        "account_name": frame["Account Name"].iloc[0],
                        "account_code": frame["Account Code"].iloc[0],
                        "matched_headers": next(item["score"] for item in detection_results if item["bank"] == detected_bank),
                    }
                )
                if missing:
                    errors.append({"file": file_storage.filename, "warning": f"Missing source headers: {', '.join(missing)}"})
            except Exception as exc:
                errors.append({"file": file_storage.filename, "error": str(exc)})

        if consolidated_frames:
            consolidated = prepare_accounting_fields(pd.concat(consolidated_frames, ignore_index=True))
            credit_summary, check_tracking, recon_summary = build_reports(consolidated)
            dataset_id = datetime.now().strftime("%Y%m%d%H%M%S%f")
            DATA_STORE[dataset_id] = {
                "consolidated": consolidated,
                "credit_summary": credit_summary,
                "check_tracking": check_tracking,
                "recon_summary": recon_summary,
                "detection_summary": detection_summary,
                "errors": errors,
            }
            session["dataset_id"] = dataset_id
            return redirect(url_for("index"))

    dataset = current_dataset()
    consolidated = dataset["consolidated"] if dataset else pd.DataFrame()
    rows, columns = df_to_records(consolidated, 200) if dataset else ([], [])
    return render_template(
        "index.html",
        username=session.get("username"),
        dataset=dataset,
        rows=rows,
        columns=columns,
        errors=dataset.get("errors", errors) if dataset else errors,
        detection_summary=dataset.get("detection_summary", detection_summary) if dataset else detection_summary,
    )


@app.route("/checks")
def checks():
    if not login_required():
        return redirect(url_for("login"))
    dataset = current_dataset()
    if not dataset:
        return redirect(url_for("index"))
    query = request.args.get("q", "")
    result = filter_text(dataset["check_tracking"], ["Check Number", "Description", "Source File", "Bank", "Account Code"], query)
    rows, columns = df_to_records(result, 500)
    return render_template("table.html", title="Check No. Tracking", rows=rows, columns=columns, query=query, download_endpoint="download_checks")


@app.route("/payments")
def payments():
    if not login_required():
        return redirect(url_for("login"))
    dataset = current_dataset()
    if not dataset:
        return redirect(url_for("index"))
    query = request.args.get("q", "")
    account_code = request.args.get("account_code", "All")
    result = dataset["credit_summary"].copy()
    if account_code != "All":
        result = result[result["Account Code"].astype(str) == str(account_code)]
    if query:
        amount_mask = result["Credit"].astype(str).str.contains(query, na=False, regex=False)
        text_result = filter_text(result, ["Description", "Source File", "Bank", "Account Name", "Account Code"], query)
        result = result[amount_mask | result.index.isin(text_result.index)]
    account_codes = ["All"] + sorted([str(x) for x in dataset["credit_summary"]["Account Code"].dropna().unique().tolist() if str(x) != ""])
    rows, columns = df_to_records(result, 500)
    return render_template(
        "payments.html",
        rows=rows,
        columns=columns,
        query=query,
        account_code=account_code,
        account_codes=account_codes,
        total_credit=result["Credit"].sum(),
    )


@app.route("/recon")
def recon():
    if not login_required():
        return redirect(url_for("login"))
    dataset = current_dataset()
    if not dataset:
        return redirect(url_for("index"))
    bank = request.args.get("bank", "All")
    month = request.args.get("month", "All")
    result = dataset["recon_summary"].copy()
    if bank != "All":
        result = result[result["Bank"] == bank]
    if month != "All":
        result = result[result["Month"] == month]
    banks = ["All"] + sorted(dataset["consolidated"]["Bank"].dropna().unique().tolist())
    months = ["All"] + sorted(dataset["consolidated"]["Month"].dropna().unique().tolist())
    rows, columns = df_to_records(result, 500)
    return render_template("recon.html", rows=rows, columns=columns, bank=bank, month=month, banks=banks, months=months)


def download_excel(sheets, filename):
    output = to_excel_bytes(sheets)
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/download/full")
def download_full():
    dataset = current_dataset()
    if not dataset:
        return redirect(url_for("index"))
    return download_excel(
        {
            "Consolidated": dataset["consolidated"],
            "Payment Verification": dataset["credit_summary"],
            "Check Tracking": dataset["check_tracking"],
            "Bank Recon": dataset["recon_summary"],
        },
        "accounting_bank_workbook.xlsx",
    )


@app.route("/download/consolidated")
def download_consolidated():
    dataset = current_dataset()
    if not dataset:
        return redirect(url_for("index"))
    return download_excel({"Consolidated": dataset["consolidated"]}, "bank_consolidated.xlsx")


@app.route("/download/checks")
def download_checks():
    dataset = current_dataset()
    if not dataset:
        return redirect(url_for("index"))
    query = request.args.get("q", "")
    result = filter_text(dataset["check_tracking"], ["Check Number", "Description", "Source File", "Bank", "Account Code"], query)
    return download_excel({"Check Tracking": result}, "check_no_tracking.xlsx")


@app.route("/download/payments")
def download_payments():
    dataset = current_dataset()
    if not dataset:
        return redirect(url_for("index"))
    query = request.args.get("q", "")
    account_code = request.args.get("account_code", "All")
    result = dataset["credit_summary"].copy()
    if account_code != "All":
        result = result[result["Account Code"].astype(str) == str(account_code)]
    if query:
        amount_mask = result["Credit"].astype(str).str.contains(query, na=False, regex=False)
        text_result = filter_text(result, ["Description", "Source File", "Bank", "Account Name", "Account Code"], query)
        result = result[amount_mask | result.index.isin(text_result.index)]
    return download_excel({"Payment Verification": result}, "payment_verification.xlsx")


@app.route("/download/recon")
def download_recon():
    dataset = current_dataset()
    if not dataset:
        return redirect(url_for("index"))
    bank = request.args.get("bank", "All")
    month = request.args.get("month", "All")
    result = dataset["recon_summary"].copy()
    if bank != "All":
        result = result[result["Bank"] == bank]
    if month != "All":
        result = result[result["Month"] == month]
    return download_excel({"Bank Recon": result}, "bank_recon.xlsx")


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
