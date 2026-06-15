import hashlib
import hmac
import io
import os
import sqlite3
from datetime import datetime

import pandas as pd
from flask import Flask, redirect, render_template, request, send_file, session, url_for
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "change-this-local-secret-key")

DATABASE_PATH = os.getenv("DATABASE_PATH", "accounting.db")
ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv"}
MIN_HEADER_MATCHES = 2

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

ACCOUNT_CODES = {"BDO MAIN": "709", "METROBANK": "253"}


def get_db():
    conn = sqlite3.connect(DATABASE_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'admin',
            created_at TEXT NOT NULL
        );

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
    username = os.getenv("APP_USERNAME", "Accounting")
    password = os.getenv("APP_PASSWORD", "Password123")
    existing = conn.execute("SELECT id FROM users WHERE username = ?", (username,)).fetchone()
    if not existing:
        conn.execute(
            "INSERT INTO users (username, password_hash, role, created_at) VALUES (?, ?, ?, ?)",
            (username, generate_password_hash(password), "admin", datetime.now().isoformat(timespec="seconds")),
        )
    conn.commit()
    conn.close()


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def canonical_header(value):
    return " ".join(str(value).strip().upper().replace("_", " ").replace("-", " ").split())


def legacy_hash_ok(username, password):
    configured_user = os.getenv("APP_USERNAME", "Accounting")
    configured_hash = os.getenv("APP_PASSWORD_SHA256", "008c70392e3abfbd0fa47bbc2ed96aa99bd49e159727fcba0f2e6abeb3a9d601")
    password_hash = hashlib.sha256(password.encode("utf-8")).hexdigest()
    return hmac.compare_digest(username, configured_user) and hmac.compare_digest(password_hash, configured_hash)


def authenticate(username, password):
    conn = get_db()
    user = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
    conn.close()
    if user and check_password_hash(user["password_hash"], password):
        return True
    return legacy_hash_ok(username, password)


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
        detection_results.append({"bank": bank, "score": len(matched_headers), "matched_headers": sorted(matched_headers)})
    detection_results.sort(key=lambda item: item["score"], reverse=True)
    best = detection_results[0]
    tied_best = [item for item in detection_results if item["score"] == best["score"]]
    if best["score"] < MIN_HEADER_MATCHES:
        return None, detection_results, "Not enough matching headers to detect the bank."
    if len(tied_best) > 1:
        return None, detection_results, "Bank detection is ambiguous between: " + ", ".join(item["bank"] for item in tied_best)
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
    return prepare_accounting_fields(output), missing_source_headers, detection_results


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
        rows.append(
            (
                upload_id,
                row["Bank"],
                row["Account Name"],
                row["Account Code"],
                row["Source File"],
                posting_date,
                row["Month"],
                row["Branch"],
                row["Description"],
                float(row["Debit"]),
                float(row["Credit"]),
                float(row["Running Balance"]),
                row["Check Number"],
            )
        )
    conn.executemany(
        """
        INSERT INTO transactions
        (upload_id, bank, account_name, account_code, source_file, posting_date, month, branch, description, debit, credit, running_balance, check_number)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        rows,
    )
    conn.commit()
    conn.close()


def query_df(sql, params=()):
    conn = get_db()
    df = pd.read_sql_query(sql, conn, params=params)
    conn.close()
    return df


def consolidated_df():
    return query_df(
        """
        SELECT bank AS Bank, account_name AS 'Account Name', account_code AS 'Account Code', source_file AS 'Source File',
               posting_date AS 'Posting Date', branch AS Branch, description AS Description, debit AS Debit, credit AS Credit,
               running_balance AS 'Running Balance', check_number AS 'Check Number', month AS Month
        FROM transactions
        ORDER BY posting_date, id
        """
    )


def uploads_df():
    return query_df("SELECT id, filename, bank, account_name, account_code, uploaded_by, uploaded_at, row_count FROM uploads ORDER BY id DESC")


def checks_df(query=""):
    base = consolidated_df()
    base = base[base["Check Number"].fillna("").astype(str).str.strip() != ""]
    return filter_text(base, ["Check Number", "Description", "Source File", "Bank", "Account Code"], query)


def payments_df(query="", account_code="All"):
    df = consolidated_df()
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
        return df
    result = (
        df.groupby(["Bank", "Account Name", "Account Code", "Month"], dropna=False)
        .agg(Transactions=("Source File", "count"), **{"Total Debit": ("Debit", "sum"), "Total Credit": ("Credit", "sum"), "Ending Balance": ("Running Balance", "last")})
        .reset_index()
    )
    result["Net Movement"] = result["Total Credit"] - result["Total Debit"]
    if bank != "All":
        result = result[result["Bank"] == bank]
    if month != "All":
        result = result[result["Month"] == month]
    return result


def filter_text(df, columns, query):
    if not query or df.empty:
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
        for sheet_name, df in sheets.items():
            safe_sheet_name = sheet_name[:31]
            df.to_excel(writer, index=False, sheet_name=safe_sheet_name)
            worksheet = writer.sheets[safe_sheet_name]
            for col_num, col_name in enumerate(df.columns):
                worksheet.write(0, col_num, col_name, header_fmt)
                worksheet.set_column(col_num, col_num, max(12, min(35, len(str(col_name)) + 4)))
            for col in ["Debit", "Credit", "Running Balance", "Total Debit", "Total Credit", "Ending Balance", "Net Movement"]:
                if col in df.columns:
                    worksheet.set_column(df.columns.get_loc(col), df.columns.get_loc(col), 16, money_fmt)
    buffer.seek(0)
    return buffer


def df_to_records(df, limit=500):
    view = df.head(limit).copy()
    return view.fillna("").to_dict(orient="records"), list(view.columns)


def render_table_page(title, df, **kwargs):
    rows, columns = df_to_records(df, 500)
    return render_template("table.html", title=title, rows=rows, columns=columns, count=len(df), **kwargs)


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        if authenticate(username, password):
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
    messages = []
    errors = []
    if request.method == "POST":
        files = request.files.getlist("files")
        saved_count = 0
        saved_rows = 0
        for file_storage in files:
            if not file_storage or not file_storage.filename:
                continue
            if not allowed_file(file_storage.filename):
                errors.append(f"{file_storage.filename}: unsupported file type.")
                continue
            try:
                frame, missing, _ = consolidate_file(file_storage)
                save_upload(frame, session.get("username"))
                saved_count += 1
                saved_rows += len(frame)
                if missing:
                    errors.append(f"{file_storage.filename}: missing source headers: {', '.join(missing)}")
            except Exception as exc:
                errors.append(f"{file_storage.filename}: {exc}")
        if saved_count:
            messages.append(f"Saved {saved_count} file(s) with {saved_rows:,} transaction row(s).")
    stats = {
        "uploads": int(query_df("SELECT COUNT(*) AS c FROM uploads")["c"].iloc[0]),
        "transactions": int(query_df("SELECT COUNT(*) AS c FROM transactions")["c"].iloc[0]),
        "credits": float(query_df("SELECT COALESCE(SUM(credit),0) AS c FROM transactions")["c"].iloc[0]),
        "debits": float(query_df("SELECT COALESCE(SUM(debit),0) AS c FROM transactions")["c"].iloc[0]),
    }
    rows, columns = df_to_records(uploads_df(), 100)
    return render_template("index.html", username=session.get("username"), stats=stats, rows=rows, columns=columns, messages=messages, errors=errors)


@app.route("/consolidated")
def consolidated():
    if not login_required():
        return redirect(url_for("login"))
    q = request.args.get("q", "")
    df = filter_text(consolidated_df(), ["Description", "Source File", "Bank", "Account Code", "Check Number"], q)
    return render_table_page("Bank Consolidation", df, query=q, download_url=url_for("download_consolidated", q=q))


@app.route("/checks")
def checks():
    if not login_required():
        return redirect(url_for("login"))
    q = request.args.get("q", "")
    df = checks_df(q)
    return render_table_page("Check No. Tracking", df, query=q, download_url=url_for("download_checks", q=q))


@app.route("/payments")
def payments():
    if not login_required():
        return redirect(url_for("login"))
    q = request.args.get("q", "")
    account_code = request.args.get("account_code", "All")
    df = payments_df(q, account_code)
    codes = ["All"] + sorted([str(x) for x in consolidated_df()["Account Code"].dropna().unique().tolist() if str(x) != ""])
    rows, columns = df_to_records(df, 500)
    return render_template("payments.html", rows=rows, columns=columns, count=len(df), query=q, account_code=account_code, account_codes=codes, total_credit=df["Credit"].sum() if not df.empty else 0, download_url=url_for("download_payments", q=q, account_code=account_code))


@app.route("/recon")
def recon():
    if not login_required():
        return redirect(url_for("login"))
    bank = request.args.get("bank", "All")
    month = request.args.get("month", "All")
    base = consolidated_df()
    banks = ["All"] + sorted(base["Bank"].dropna().unique().tolist()) if not base.empty else ["All"]
    months = ["All"] + sorted(base["Month"].dropna().unique().tolist()) if not base.empty else ["All"]
    df = recon_df(bank, month)
    rows, columns = df_to_records(df, 500)
    return render_template("recon.html", rows=rows, columns=columns, count=len(df), bank=bank, month=month, banks=banks, months=months, download_url=url_for("download_recon", bank=bank, month=month))


def download_excel(sheets, filename):
    return send_file(to_excel_bytes(sheets), as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/download/full")
def download_full():
    if not login_required():
        return redirect(url_for("login"))
    return download_excel({"Consolidated": consolidated_df(), "Payment Verification": payments_df(), "Check Tracking": checks_df(), "Bank Recon": recon_df()}, "accounting_bank_workbook.xlsx")


@app.route("/download/consolidated")
def download_consolidated():
    if not login_required():
        return redirect(url_for("login"))
    q = request.args.get("q", "")
    return download_excel({"Consolidated": filter_text(consolidated_df(), ["Description", "Source File", "Bank", "Account Code", "Check Number"], q)}, "bank_consolidated.xlsx")


@app.route("/download/checks")
def download_checks():
    if not login_required():
        return redirect(url_for("login"))
    return download_excel({"Check Tracking": checks_df(request.args.get("q", ""))}, "check_no_tracking.xlsx")


@app.route("/download/payments")
def download_payments():
    if not login_required():
        return redirect(url_for("login"))
    return download_excel({"Payment Verification": payments_df(request.args.get("q", ""), request.args.get("account_code", "All"))}, "payment_verification.xlsx")


@app.route("/download/recon")
def download_recon():
    if not login_required():
        return redirect(url_for("login"))
    return download_excel({"Bank Recon": recon_df(request.args.get("bank", "All"), request.args.get("month", "All"))}, "bank_recon.xlsx")


if __name__ == "__main__":
    init_db()
    app.run(host="127.0.0.1", port=5000, debug=True)
