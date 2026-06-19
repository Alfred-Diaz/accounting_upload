import hashlib
import hmac
import io
import os
import sqlite3
from datetime import datetime

import pandas as pd
import streamlit as st

DB = "accounting.db"
PASSWORD_HASH = "008c70392e3abfbd0fa47bbc2ed96aa99bd49e159727fcba0f2e6abeb3a9d601"
HEADERS = ["Posting Date", "Branch", "Description", "Debit", "Credit", "Running Balance", "Check Number"]
MAPS = {
    "METROBANK": {"Posting Date": "Posting Date", "Branch/Channel": "Branch", "Transaction Description": "Description", "Debit Amount": "Debit", "Credit Amount": "Credit", "Balance": "Running Balance", "Check Number": "Check Number"},
    "EASTWEST": {"Value Date": "Posting Date", "Descript": "Branch", "Reference": "Description", "Debit": "Debit", "Credit": "Credit", "Closing Balance": "Running Balance", "Cheque Number": "Check Number"},
    "BDO": {"Posting Date": "Posting Date", "Branch": "Branch", "Description": "Description", "Debit": "Debit", "Credit": "Credit", "Running Balance": "Running Balance", "Check Number": "Check Number"},
}


def secret(key, default=""):
    try:
        return st.secrets.get(key, os.getenv(key, default))
    except Exception:
        return os.getenv(key, default)


def logged_in():
    if st.session_state.get("auth"):
        with st.sidebar:
            st.success(f"Signed in as {st.session_state.get('user', 'Accounting')}")
            if st.button("Log out"):
                st.session_state.clear()
                st.rerun()
        return True
    st.title("Accounting System")
    with st.form("login"):
        user = st.text_input("Username")
        pwd = st.text_input("Password", type="password")
        ok = st.form_submit_button("Sign in")
    if ok:
        user_ok = hmac.compare_digest(user, secret("APP_USERNAME", "Accounting"))
        pwd_hash = hashlib.sha256(pwd.encode()).hexdigest()
        pwd_ok = hmac.compare_digest(pwd_hash, secret("APP_PASSWORD_SHA256", PASSWORD_HASH))
        if user_ok and pwd_ok:
            st.session_state["auth"] = True
            st.session_state["user"] = user
            st.rerun()
        st.error("Invalid username or password.")
    return False


def db():
    con = sqlite3.connect(DB, check_same_thread=False)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    con = db()
    con.execute("CREATE TABLE IF NOT EXISTS uploads (id INTEGER PRIMARY KEY, filename TEXT, bank TEXT, account_name TEXT, account_code TEXT, uploaded_by TEXT, uploaded_at TEXT, row_count INTEGER, inserted_rows INTEGER DEFAULT 0, duplicate_rows INTEGER DEFAULT 0)")
    con.execute("CREATE TABLE IF NOT EXISTS transactions (id INTEGER PRIMARY KEY, upload_id INTEGER, bank TEXT, account_name TEXT, account_code TEXT, source_file TEXT, posting_date TEXT, month TEXT, branch TEXT, description TEXT, debit REAL, credit REAL, running_balance REAL, check_number TEXT, row_key TEXT)")
    for sql in ["ALTER TABLE uploads ADD COLUMN inserted_rows INTEGER DEFAULT 0", "ALTER TABLE uploads ADD COLUMN duplicate_rows INTEGER DEFAULT 0", "ALTER TABLE transactions ADD COLUMN row_key TEXT"]:
        try:
            con.execute(sql)
        except sqlite3.OperationalError:
            pass
    con.commit()
    con.close()


def qdf(sql, params=()):
    con = db()
    out = pd.read_sql_query(sql, con, params=params)
    con.close()
    return out


def canon(x):
    return " ".join(str(x).strip().upper().replace("_", " ").replace("-", " ").split())


def read_file(f):
    f.seek(0)
    if f.name.lower().endswith(".csv"):
        return pd.read_csv(f)
    return pd.read_excel(f)


def detect_bank(df):
    cols = {canon(c) for c in df.columns}
    scores = []
    for bank, mapping in MAPS.items():
        hits = cols.intersection({canon(c) for c in mapping})
        scores.append((bank, len(hits), sorted(hits)))
    scores.sort(key=lambda x: x[1], reverse=True)
    if scores[0][1] < 2:
        raise ValueError("Not enough matching headers to detect bank.")
    if len(scores) > 1 and scores[0][1] == scores[1][1]:
        raise ValueError("Bank detection is ambiguous.")
    return scores[0]


def acct_from_name(name, bank):
    up = name.upper()
    if "709" in up or "BDO MAIN" in up or bank == "BDO":
        return "BDO MAIN", "709"
    if "253" in up or "METROBANK" in up or bank == "METROBANK":
        return "METROBANK", "253"
    return bank, ""


def normalize_upload(f):
    raw = read_file(f)
    raw.columns = [str(c).strip() for c in raw.columns]
    bank, score, hits = detect_bank(raw)
    lookup = {canon(c): c for c in raw.columns}
    rename = {}
    missing = []
    for src, dest in MAPS[bank].items():
        actual = lookup.get(canon(src))
        if actual:
            rename[actual] = dest
        else:
            missing.append(src)
    temp = raw.rename(columns=rename)
    out = pd.DataFrame()
    for h in HEADERS:
        out[h] = temp[h] if h in temp.columns else pd.NA
    acct, code = acct_from_name(f.name, bank)
    out.insert(0, "Bank", bank)
    out.insert(1, "Account Name", acct)
    out.insert(2, "Account Code", code)
    out.insert(3, "Source File", f.name)
    out["Posting Date"] = pd.to_datetime(out["Posting Date"], errors="coerce")
    for c in ["Debit", "Credit", "Running Balance"]:
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0)
    out["Month"] = out["Posting Date"].dt.to_period("M").astype(str).replace("NaT", "")
    out["Branch"] = out["Branch"].fillna("").astype(str)
    out["Description"] = out["Description"].fillna("").astype(str)
    out["Check Number"] = out["Check Number"].fillna("").astype(str).str.strip()
    return out, missing, {"file": f.name, "detected_bank": bank, "matched_headers": score, "account_code": code}


def debit_key(row):
    date = "" if pd.isna(row["Posting Date"]) else row["Posting Date"].strftime("%Y-%m-%d")
    parts = [row["Bank"], row["Account Name"], row["Account Code"], date, row["Branch"], row["Description"], f"{float(row['Debit']):.2f}", f"{float(row['Running Balance']):.2f}", row["Check Number"]]
    return hashlib.sha256("|".join(str(p).strip() for p in parts).encode()).hexdigest()


def existing_debits():
    try:
        df = qdf("SELECT row_key FROM transactions WHERE debit > 0 AND row_key IS NOT NULL AND row_key != ''")
        return set(df["row_key"].dropna().astype(str))
    except Exception:
        return set()


def save_frame(frame, user):
    con = db()
    first = frame.iloc[0]
    cur = con.execute("INSERT INTO uploads (filename, bank, account_name, account_code, uploaded_by, uploaded_at, row_count, inserted_rows, duplicate_rows) VALUES (?, ?, ?, ?, ?, ?, ?, 0, 0)", (first["Source File"], first["Bank"], first["Account Name"], first["Account Code"], user, datetime.now().isoformat(timespec="seconds"), len(frame)))
    upload_id = cur.lastrowid
    seen = existing_debits()
    inserted = 0
    skipped = 0
    for _, row in frame.iterrows():
        post = "" if pd.isna(row["Posting Date"]) else row["Posting Date"].strftime("%Y-%m-%d")
        key = ""
        if float(row["Debit"]) > 0:
            key = debit_key(row)
            if key in seen:
                skipped += 1
                continue
            seen.add(key)
        con.execute("INSERT INTO transactions (upload_id, bank, account_name, account_code, source_file, posting_date, month, branch, description, debit, credit, running_balance, check_number, row_key) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", (upload_id, row["Bank"], row["Account Name"], row["Account Code"], row["Source File"], post, row["Month"], row["Branch"], row["Description"], float(row["Debit"]), float(row["Credit"]), float(row["Running Balance"]), row["Check Number"], key))
        inserted += 1
    con.execute("UPDATE uploads SET inserted_rows=?, duplicate_rows=? WHERE id=?", (inserted, skipped, upload_id))
    con.commit()
    con.close()
    return inserted, skipped


def all_txn():
    return qdf("SELECT bank AS Bank, account_name AS 'Account Name', account_code AS 'Account Code', source_file AS 'Source File', posting_date AS 'Posting Date', branch AS Branch, description AS Description, debit AS Debit, credit AS Credit, running_balance AS 'Running Balance', check_number AS 'Check Number', month AS Month FROM transactions ORDER BY posting_date, id")


def uploads():
    return qdf("SELECT id AS ID, filename AS Filename, bank AS Bank, account_name AS 'Account Name', account_code AS 'Account Code', uploaded_by AS 'Uploaded By', uploaded_at AS 'Uploaded At', row_count AS Rows, inserted_rows AS 'Inserted Rows', duplicate_rows AS 'Duplicate Debit Rows' FROM uploads ORDER BY id DESC")


def contains(df, cols, text):
    if not text or df.empty:
        return df
    mask = pd.Series(False, index=df.index)
    text = text.lower().strip()
    for col in cols:
        if col in df.columns:
            mask |= df[col].fillna("").astype(str).str.lower().str.contains(text, regex=False, na=False)
    return df[mask]


def checks(text=""):
    df = all_txn()
    if df.empty:
        return df
    df = df[(df["Debit"] > 0) | (df["Check Number"].fillna("").astype(str).str.strip() != "")].copy()
    cols = ["Posting Date", "Bank", "Account Name", "Account Code", "Check Number", "Debit", "Description", "Branch", "Running Balance", "Source File", "Month"]
    df = df[[c for c in cols if c in df.columns]]
    return contains(df, ["Check Number", "Description", "Source File", "Bank", "Account Code"], text)


def payments(text="", code="All"):
    df = all_txn()
    if df.empty:
        return df
    df = df[df["Credit"] > 0].copy()
    if code != "All":
        df = df[df["Account Code"].astype(str) == str(code)]
    if text:
        amt = df["Credit"].astype(str).str.contains(text, regex=False, na=False)
        txt = contains(df, ["Description", "Source File", "Bank", "Account Name", "Account Code"], text)
        df = df[amt | df.index.isin(txt.index)]
    return df


def recon(bank="All", month="All"):
    df = all_txn()
    if df.empty:
        return pd.DataFrame()
    out = df.groupby(["Bank", "Account Name", "Account Code", "Month"], dropna=False).agg(Transactions=("Source File", "count"), **{"Total Debit": ("Debit", "sum"), "Total Credit": ("Credit", "sum"), "Ending Balance": ("Running Balance", "last")}).reset_index()
    out["Net Movement"] = out["Total Credit"] - out["Total Debit"]
    if bank != "All":
        out = out[out["Bank"] == bank]
    if month != "All":
        out = out[out["Month"] == month]
    return out


def excel_bytes(sheets):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=name[:31])
    return buf.getvalue()


def dl(label, sheets, name, primary=False):
    st.download_button(label, excel_bytes(sheets), name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary" if primary else "secondary", use_container_width=True)


def main():
    st.set_page_config(page_title="Accounting System", layout="wide")
    init_db()
    if not logged_in():
        st.stop()
    st.title("Accounting System")
    with st.sidebar:
        page = st.radio("Menu", ["Dashboard", "Upload Bank Files", "Bank Consolidation", "Check No. Tracking", "Payment Verification", "Bank Recon", "Settings"])
    if page == "Dashboard":
        tx = all_txn(); up = uploads()
        a, b, c, d = st.columns(4)
        a.metric("Uploads", len(up)); b.metric("Transactions", len(tx))
        c.metric("Total Debit", f"{tx['Debit'].sum():,.2f}" if not tx.empty else "0.00")
        d.metric("Total Credit", f"{tx['Credit'].sum():,.2f}" if not tx.empty else "0.00")
        st.subheader("Upload History"); st.dataframe(up, use_container_width=True)
        if not tx.empty:
            dl("Download full accounting workbook", {"Consolidated": tx, "Payment Verification": payments(), "Check Tracking": checks(), "Bank Recon": recon()}, "accounting_bank_workbook.xlsx", True)
    elif page == "Upload Bank Files":
        st.subheader("Upload Bank Files")
        st.info("Duplicate checking applies only to debit/check rows. Credit rows are all saved, even when similar.")
        files = st.file_uploader("Upload Excel or CSV bank files", ["xlsx", "xls", "csv"], accept_multiple_files=True)
        if files and st.button("Process and save uploaded files", type="primary"):
            summary = []; issues = []; total_inserted = 0; total_skipped = 0
            for f in files:
                try:
                    frame, missing, info = normalize_upload(f)
                    ins, skip = save_frame(frame, st.session_state.get("user", "Accounting"))
                    total_inserted += ins; total_skipped += skip
                    info.update({"rows_in_file": len(frame), "inserted_rows": ins, "duplicate_debit_rows_skipped": skip})
                    summary.append(info)
                    if missing:
                        issues.append({"file": f.name, "warning": f"Missing source headers: {', '.join(missing)}"})
                except Exception as e:
                    issues.append({"file": f.name, "error": str(e)})
            st.success(f"Inserted {total_inserted:,} row(s). Skipped {total_skipped:,} duplicate debit/check row(s).")
            if summary: st.dataframe(pd.DataFrame(summary), use_container_width=True)
            if issues: st.json(issues)
    elif page == "Bank Consolidation":
        st.subheader("Bank Consolidation")
        q = st.text_input("Search description, source file, bank, account code, or check number")
        df = contains(all_txn(), ["Description", "Source File", "Bank", "Account Code", "Check Number"], q)
        st.caption(f"Showing {len(df):,} transaction(s). Download exports this exact table.")
        st.dataframe(df, use_container_width=True)
        dl("Download Bank Consolidation table", {"Bank Consolidation": df}, "bank_consolidation_table.xlsx", True)
    elif page == "Check No. Tracking":
        st.subheader("Check No. Tracking")
        q = st.text_input("Search by check number, description, source file, bank, or account code")
        df = checks(q)
        st.caption(f"Showing {len(df):,} debit/check transaction(s). Debit amount is included. Download exports this exact table.")
        st.dataframe(df, use_container_width=True)
        dl("Download Check No. Tracking table", {"Check No Tracking": df}, "check_no_tracking_table.xlsx", True)
    elif page == "Payment Verification":
        st.subheader("Payment Verification")
        tx = all_txn(); codes = ["All"] + (sorted([str(x) for x in tx["Account Code"].dropna().unique() if str(x) != ""]) if not tx.empty else [])
        code = st.selectbox("Account code", codes); q = st.text_input("Search amount, description, source file, bank, or account")
        df = payments(q, code); st.metric("Total credits found", f"{df['Credit'].sum():,.2f}" if not df.empty else "0.00")
        st.dataframe(df, use_container_width=True); dl("Download payment verification result", {"Payment Verification": df}, "payment_verification.xlsx")
    elif page == "Bank Recon":
        st.subheader("Bank Recon")
        tx = all_txn(); banks = ["All"] + (sorted(tx["Bank"].dropna().unique()) if not tx.empty else []); months = ["All"] + (sorted(tx["Month"].dropna().unique()) if not tx.empty else [])
        c1, c2 = st.columns(2); bank = c1.selectbox("Bank", banks); month = c2.selectbox("Month", months)
        df = recon(bank, month); st.dataframe(df, use_container_width=True); dl("Download bank recon result", {"Bank Recon": df}, "bank_recon.xlsx")
    else:
        st.subheader("Settings")
        st.code('APP_USERNAME = "Accounting"\nAPP_PASSWORD_SHA256 = "008c70392e3abfbd0fa47bbc2ed96aa99bd49e159727fcba0f2e6abeb3a9d601"', language="toml")


if __name__ == "__main__":
    main()
