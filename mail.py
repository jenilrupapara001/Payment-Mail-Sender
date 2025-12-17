import pandas as pd
import streamlit as st
import os
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
import xlsxwriter
import hashlib
import zipfile
import pyodbc
import webbrowser
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from datetime import datetime
import time  # For SMTP throttling!
import random

# Constants
JSON_PATH = Path("party_emails.json")
EXCEL_PATH = Path("Invoices.xlsx")
EMAIL_UPLOAD_PASSWORD = "Payment Mail Sender Dashboard"

connection_string = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=(localdb)\\MSSQLLocalDB;"
    "DATABASE=EasySell;"
    "Trusted_Connection=yes;"
)

def create_sample_excel():
    sample_payment = pd.DataFrame({
        "Party Name": ["Alpha Corp", "Beta Ltd"],
        "Inv. No.": ["INV001", "INV002"],
        "Pur. Date": ["2025-01-10", "2025-01-15"],
        "Total Inv. Amount": [10000, 20000],
        "Debit Amount": [1000, ""],
        "Net Amount": [9500, 20000],
        "Bank Payment": [9500, 20000],
        "Payment Date": ["2025-02-10", "2025-02-20"],
        "Amount": [9500, 20000],
    })
    sample_debit = pd.DataFrame({
        "Party Name": ["Alpha Corp"],
        "Date": ["2025-02-05"],
        "Return Invoice No.": ["DN001"],
        "Amount": [500]
    })
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        sample_payment.to_excel(writer, index=False, sheet_name="Payment Details")
        sample_debit.to_excel(writer, index=False, sheet_name="Debit Notes")
    return output.getvalue()

def create_sample_mail_excel():
    df = pd.DataFrame({
        'Party Code': ['PC123', 'PC456'],
        'Party Name': ['ABC Traders', 'XYZ Pvt Ltd'],
        'Email': ['abc@example.com,bcd@gmail.com', 'xyz@example.com']
    })
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output

EMAIL_TEMPLATE = """
<html>
  <body style="font-family: Arial, sans-serif; color: #333;">
    <p>Dear [Party Name],</p>
    <p>Please find below the summary of your recent transactions with us:</p>
    <h3>Purchase & Payment Details</h3>
    <table style="border-collapse: collapse;  width: 100%; margin-bottom: 20px;">
      <thead>
        <tr style="background-color: #f2f2f2; border: 2px solid #333;">
          <th style="border: 1px solid #333; padding: 8px; ">Purchase Bill</th>
          <th style="border: 1px solid #ddd; padding: 8px; ">Main Advised No.</th>
          <th style="border: 1px solid #ddd; padding: 8px; ">Seller Advised No.</th>
          <th style="border: 1px solid #ddd; padding: 8px; ">Transaction Type</th>
          <th style="border: 1px solid #ddd; padding: 8px; ">Pur. Date</th>
          <th style="border: 1px solid #ddd; padding: 8px; ">Credit (CR)</th>
          <th style="border: 1px solid #ddd; padding: 8px; ">Debit (DR)</th>
          <th style="border: 1px solid #ddd; padding: 8px; ">Balance</th>
        </tr>
      </thead>
      <tbody>
        <!-- Dynamic payment rows inserted here -->
      </tbody>
    </table>
  </body>
</html>
"""

def load_party_emails():
    if not JSON_PATH.exists():
        sample = [
            {"PartyName": "Alpha Corp", "Email": "alpha@example.com"},
            {"PartyName": "Beta Ltd", "Email": "beta@example.com"}
        ]
        with open(JSON_PATH, 'w') as f:
            json.dump(sample, f, indent=2)
    with open(JSON_PATH, 'r') as f:
        return json.load(f)

def save_party_emails(data):
    with open(JSON_PATH, 'w') as f:
        json.dump(data, f, indent=2)

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def check_password(input_pwd):
    return hash_password(input_pwd) == hash_password("Password")

def load_excel(file_path):
    wb = pd.ExcelFile(file_path)
    sheet_names = [s.strip() for s in wb.sheet_names]

    # Legacy two-sheet format: keep existing behavior
    if "Payment Details" in sheet_names and "Debit Notes" in sheet_names:
        payment_df = wb.parse("Payment Details")
        debit_df = wb.parse("Debit Notes")
        payment_df.columns = payment_df.columns.str.strip()
        debit_df.columns = debit_df.columns.str.strip()
        return payment_df, debit_df

    # New single-sheet format with columns highlighted in yellow
    # Expected headers (case-insensitive): Seller Name, Channel, Transaction Type, Category,
    # Bill No, Invoice Date, Quantity, Total Without Tax, Total Tax, Total With Tax,
    # Zoho Total Without Tax, Zoho Total Tax, Zoho Total With Tax, Balance Due,
    # Zoho Status, CR, DR, Balance
    sheet_name = sheet_names[0]

    # Detect merged summary rows and offset header (seen in vendor Payment Details.xlsx)
    raw_df_preview = wb.parse(sheet_name, header=None, nrows=5)
    header_row = 0
    first_cell = str(raw_df_preview.iloc[0, 0]) if not pd.isna(raw_df_preview.iloc[0, 0]) else ""
    if "Seller Name:" in first_cell and "Advised No" in first_cell:
        header_row = 2  # actual headers at row index 2 (0-based)

    raw_df = wb.parse(sheet_name, header=header_row)
    raw_df.columns = raw_df.columns.str.strip()

    def pick(col_candidates):
        lower_map = {c.lower(): c for c in raw_df.columns}
        for cand in col_candidates:
            if cand.lower() in lower_map:
                return lower_map[cand.lower()]
        return None

    col_seller = pick(["Seller Name", "Party Name"])
    col_bill = pick(["Bill No", "Invoice No", "Inv. No."])
    col_date = pick(["Invoice Date", "Date"])
    col_total_with_tax = pick(["Total With Tax", "Total With Tax ", "Total_with_tax"])
    col_total_with_tax_alt = pick(["Zoho Total With Tax", "Zoho total with tax"])
    col_main_advise_no = pick(["Main Advised No", "Main Advise No"])
    col_seller_advised_no = pick(["Seller Advised No", "Seller Advise No"])
    col_dr = pick(["DR", "Debit", "Debit Amount"])
    col_cr = pick(["CR", "Credit", "Credit Amount"])
    col_category = pick(["Category"])
    col_channel = pick(["Channel"])
    col_txn_type = pick(["Transaction Type", "Transaction", "Transacation Type"])
    col_quantity = pick(["Quantity", "Qty"])
    col_total_wo_tax = pick(["Total Without Tax", "Total Without Tax "])
    col_total_tax = pick(["Total Tax"])
    col_zoho_wo_tax = pick(["Zoho Total Without Tax"])
    col_zoho_tax = pick(["Zoho Total Tax"])
    col_zoho_with_tax = pick(["Zoho Total With Tax"])
    col_balance_due = pick(["Balance Due"])
    col_zoho_status = pick(["Zoho Status"])
    col_balance = pick(["Balance"])

    # Basic required columns
    missing_cols = []
    if col_seller is None:
        missing_cols.append("Seller Name")
    if col_bill is None:
        missing_cols.append("Bill No")
    if col_date is None:
        missing_cols.append("Invoice Date")
    if col_main_advise_no is None:
        missing_cols.append("Main Advised No")
    if col_seller_advised_no is None:
        missing_cols.append("Seller Advised No")
    # For amounts we allow fallbacks; collect missing for messaging only
    amt_missing = []
    if col_total_with_tax is None and col_total_with_tax_alt is None:
        amt_missing.append("Total With Tax")
    if col_dr is None:
        amt_missing.append("DR")
    if col_cr is None:
        amt_missing.append("CR")
    if missing_cols:
        raise ValueError(f"Missing required columns in the uploaded sheet: {', '.join(missing_cols)}. Expected at least Seller Name, Bill No, Invoice Date.")

    # Normalize numeric columns
    def num(series):
        return pd.to_numeric(series, errors="coerce").fillna(0)

    # Drop summary/empty rows
    raw_df = raw_df.dropna(how="all")
    if col_seller:
        raw_df = raw_df[~raw_df[col_seller].isna()]

    # Fallback order for totals
    if col_total_with_tax:
        total_with_tax_series = num(raw_df[col_total_with_tax])
    elif col_total_with_tax_alt:
        total_with_tax_series = num(raw_df[col_total_with_tax_alt])
    elif col_total_wo_tax:
        total_with_tax_series = num(raw_df[col_total_wo_tax])
    else:
        total_with_tax_series = pd.Series([0] * len(raw_df))

    dr_series = num(raw_df[col_dr]) if col_dr else pd.Series([0] * len(raw_df))
    cr_series = num(raw_df[col_cr]) if col_cr else pd.Series([0] * len(raw_df))

    # If there is no explicit total column but we do have CR/DR, derive a pseudo total
    if (col_total_with_tax is None and col_total_with_tax_alt is None and col_total_wo_tax is None) and (col_cr or col_dr):
        total_with_tax_series = cr_series + dr_series

    # Base series for seller/bill/date
    seller_series = raw_df[col_seller].fillna("").astype(str).str.strip()
    bill_series = raw_df[col_bill].fillna("").astype(str).str.strip()
    date_series = raw_df[col_date]

    # Filter out only total/blank rows (keep all rows with valid seller name)
    filtered_idx = ~(
        (bill_series.str.lower().isin(["", "total", "nan"])) 
        & (seller_series.str.strip() == "")
    )
    seller_series = seller_series[filtered_idx]
    bill_series = bill_series[filtered_idx]
    date_series = date_series[filtered_idx]
    raw_df = raw_df.loc[filtered_idx]

    # Derive Party Code from seller name where possible (e.g. "731-AUROMIN-Amazon" -> "731", "731s-AUROMIN-demo" -> "731")
    import re
    def derive_code(val: str) -> str:
        if not val:
            return ""
        m = re.match(r"(\d+)", val.strip())
        if m:
            return m.group(1)
        # fallback to chunk before first dash
        return val.split("-")[0].strip() if "-" in val else val.strip()
    party_code_series = seller_series.apply(derive_code)
    party_code_series = party_code_series.where(party_code_series != "", seller_series)

    payment_df = pd.DataFrame({
        "Party Name": seller_series,
        "Party Code": party_code_series,
        "Inv. No.": bill_series,
        "Main Advised No.": raw_df[col_main_advise_no] if col_main_advise_no else "",
        "Seller Advised No.": raw_df[col_seller_advised_no] if col_seller_advised_no else "",
        "Pur. Date": date_series,
        "Total Inv. Amount": total_with_tax_series,
        "Debit Amount": dr_series,
        # Net = Total - DR - CR (treat CR as credit note)
        "Net Amount": total_with_tax_series - dr_series - cr_series,
        # Bank Payment shows CR so existing email layout still reflects reduction
        "Bank Payment": cr_series,
        "Payment Date": date_series,
        # Provide a debit/credit note reference when present
        "Debit Note": bill_series.where(dr_series > 0, "").fillna(""),
        "Transaction Type": raw_df[col_txn_type] if col_txn_type else ""
    })

    # Trim to only the needed columns for mail logic
    keep_cols = [
        "Party Name",
        "Party Code",
        "Inv. No.",
        "Main Advised No.",
        "Seller Advised No.",
        "Pur. Date",
        "Total Inv. Amount",
        "Debit Amount",
        "Net Amount",
        "Bank Payment",
        "Payment Date",
        "Debit Note",
        "Transaction Type",
    ]
    payment_df = payment_df[keep_cols]

    # Build a synthetic Debit Notes sheet from DR amounts
    debit_rows = []
    for _, row in raw_df.iterrows():
        seller_val = str(row[col_seller]).strip() if pd.notna(row[col_seller]) else ""
        party_code_val = derive_code(seller_val)
        party_code_val = party_code_val or seller_val
        party_name_val = seller_val
        bill_no = str(row[col_bill]).strip() if pd.notna(row[col_bill]) else ""
        inv_date = row[col_date]
        dr_amt = pd.to_numeric(row[col_dr], errors="coerce")
        if pd.notna(dr_amt) and dr_amt > 0:
            debit_rows.append({
                "Party Name": party_name_val,
                "Party Code": party_code_val,
                "Date": inv_date,
                "Return Invoice No.": bill_no,
                "Amount": float(dr_amt),
            })
        cr_amt = pd.to_numeric(row[col_cr], errors="coerce") if col_cr else 0
        if pd.notna(cr_amt) and cr_amt > 0:
            debit_rows.append({
                "Party Name": party_name_val,
                "Party Code": party_code_val,
                "Date": inv_date,
                "Return Invoice No.": f"{bill_no} (CR)",
                "Amount": float(cr_amt) * -1.0,  # credit note reduces balance
            })
    debit_df = pd.DataFrame(debit_rows) if debit_rows else pd.DataFrame(columns=["Party Code", "Party Name", "Date", "Return Invoice No.", "Amount"])
    debit_df.columns = debit_df.columns.str.strip()
    return payment_df, debit_df

def match_data(payment_df, debit_df, party_emails):
    # Match strictly on Party Name (Seller Name) instead of Party Code
    email_map = {
        e["PartyName"].strip(): {
            "to": [email.strip() for email in e["Email"].split(",")],
            "cc": [cc.strip() for cc in e["CC"].split(",")] if "CC" in e and pd.notna(e["CC"]) else []
        }
        for e in party_emails
    }
    payment_df.columns = payment_df.columns.str.strip()
    debit_df.columns = debit_df.columns.str.strip()
    result = []
    mismatch_log_lines = []
    skip_log_lines = []
    parties_without_email = []
    
    # Prefer matching by Party Name (Seller Name)
    payment_party_col = None
    if 'Party Name' in payment_df.columns:
        payment_party_col = 'Party Name'
    elif 'Party Code' in payment_df.columns:
        payment_party_col = 'Party Code'
    
    debit_party_col = None
    if 'Party Name' in debit_df.columns:
        debit_party_col = 'Party Name'
    elif 'Party Code' in debit_df.columns:
        debit_party_col = 'Party Code'
    
    # First, identify parties in payment sheet that have no email
    parties_without_email = []
    if payment_party_col:
        payment_party_codes = set(payment_df[payment_party_col].astype(str).str.strip())
        for party_code in payment_party_codes:
            if party_code not in email_map or not email_map[party_code]["to"] or all(email.strip().lower() in ['nan', 'none', ''] for email in email_map[party_code]["to"]):
                party_name = party_code  # Default to party code if name not found
                # Try to find party name from party_emails
                for party in party_emails:
                    if party["PartyName"].strip() == party_code.strip():
                        party_name = party["PartyName"]  # keep same; already name
                        break
                parties_without_email.append({
                    "party_code": party_code,
                    "party_name": party_name,
                    "payment_count": len(payment_df[payment_df[payment_party_col].astype(str).str.strip() == party_code.strip()])
                })
    
    for party_code, email_data in email_map.items():
        if payment_party_col:
            party_payments = payment_df[payment_df[payment_party_col].astype(str).str.strip() == party_code.strip()]
        else:
            party_payments = pd.DataFrame()  # Empty DataFrame if no party column found
        
        if party_payments.empty:
            skip_log_lines.append(f"SKIPPED: {party_code} — No payment rows found in Payment Sheet")
            continue
        related_debits = debit_df[debit_df[debit_party_col].astype(str).str.strip() == party_code.strip()] if debit_party_col else pd.DataFrame()
        # Only compare positive debit notes against payment debit amounts; credits are negative and excluded from this check
        total_debit_amount = related_debits[related_debits['Amount'] > 0]['Amount'].sum() if not related_debits.empty else 0
        party_payments = party_payments.copy()
        party_payments['Debit Amount'] = party_payments['Debit Amount'].fillna(0)
        party_debit_sum = party_payments['Debit Amount'].sum()

        if abs(party_debit_sum - total_debit_amount) > 0.01:
            skip_log_lines.append(f"SKIPPED: {party_code} — Debit Amount mismatch between payment sheet and debit sheet")
            continue

        # Include ALL payment rows for this party (no filtering based on debit note matching)
        payment_issues = []
        for _, row in party_payments.iterrows():
            payment_issues.append(row.to_dict())
        if payment_issues:
            result.append({
                'party_code': party_code,
                'emails': email_data["to"],
                'cc_emails': email_data["cc"],
                'payments': payment_issues,
                'debits': related_debits.to_dict(orient='records') if not related_debits.empty else []
            })
        else:
            skip_log_lines.append(f"SKIPPED: {party_code} — All payment rows matched with debit notes correctly.")

    if skip_log_lines:
        with open('SkippedPartiesLog.txt', 'w') as f:
            for line in skip_log_lines:
                f.write(line + "\n")
    if mismatch_log_lines:
        with open('MismatchLog.txt', 'w') as f:
            for line in mismatch_log_lines:
                f.write(line + "\n")
    return result, skip_log_lines, parties_without_email

def generate_email_body(party_code, payment_rows, debit_rows):
    # party_code is actually PartyName since we match by name now
    party_name = next((e['PartyName'] for e in party_emails if e['PartyName'] == party_code), party_code if party_code else 'Unknown Party')
    template = EMAIL_TEMPLATE
    payment_html = ""
    total_credit = 0.0
    total_debit = 0.0
    running_balance = 0.0
    for row in payment_rows:
        # Raw numeric values for CR / DR
        debit_val_num = row.get('Debit Amount', 0)
        credit_val_num = row.get('Bank Payment', 0)
        try:
            dr = float(debit_val_num) if not pd.isna(debit_val_num) and debit_val_num != '' else 0.0
        except (ValueError, TypeError):
            dr = 0.0
        try:
            cr = float(credit_val_num) if not pd.isna(credit_val_num) and credit_val_num != '' else 0.0
        except (ValueError, TypeError):
            cr = 0.0

        total_credit += cr
        total_debit += dr
        running_balance += cr - dr

        # Handle NaN and missing values for display
        inv_no = row.get('Inv. No.', '')
        main_adv = row.get('Main Advised No.', '')
        seller_adv = row.get('Seller Advised No.', '')
        pur_date = row.get('Pur. Date', '')
        txn_type = row.get('Transaction Type', '')
        
        inv_no = '-' if pd.isna(inv_no) or inv_no == '' else str(inv_no)
        main_adv_display = '-' if pd.isna(main_adv) or main_adv == '' else str(main_adv)
        seller_adv_display = '-' if pd.isna(seller_adv) or seller_adv == '' else str(seller_adv)
        pur_date = '-' if pd.isna(pur_date) or pur_date == '' else str(pur_date)
        debit_val_display = '-' if pd.isna(dr) or dr == '' else f"{dr:.2f}"
        credit_val_display = '-' if pd.isna(cr) or cr == '' else f"{cr:.2f}"
        txn_type_display = '-' if pd.isna(txn_type) or txn_type == '' else str(txn_type)
        balance_display = f"{running_balance:.2f}"
        
        payment_html += f"""
        <tr style="text-align:center; border:1px solid #ccc;">
          <td style="border:1px solid #ccc;">{inv_no}</td>
          <td style="border:1px solid #ccc;">{main_adv_display}</td>
          <td style="border:1px solid #ccc;">{seller_adv_display}</td>
          <td style="border:1px solid #ccc;">{txn_type_display}</td>
          <td style="border:1px solid #ccc;">{pur_date}</td>
          <td style="border:1px solid #ccc;">{credit_val_display}</td>
          <td style="border:1px solid #ccc;">{debit_val_display}</td>
          <td style="border:1px solid #ccc;">{balance_display}</td>
        </tr>"""
    # Final balance = total credit - total debit (as in sheet Balance column)
    final_balance = total_credit - total_debit
    # Show a single summary row where Balance is the net (CR - DR)
    payment_html += f"""
    <tr style="text-align:center; font-weight:bold; background-color:#f9f9f9;">
      <td colspan="7" style="border:1px solid #ccc; text-align:right;">Bank Final Amount</td>
      <td style="border:1px solid #ccc;">{final_balance:.2f}</td>
    </tr>"""
    html_body = template.replace("[Party Name]", party_name)
    html_body = html_body.replace("<!-- Dynamic payment rows inserted here -->", payment_html)
    closing_note = """
    <br><br>
    <p><strong>🔔 Important Note:</strong> If you have any discrepancies or concerns regarding the above payment summary, please raise the issue within 7 days. No changes or claims will be entertained after this period.</p>
    <p>Thank you for your continued partnership.</p>
    <p>Best regards,<br><strong>Easy Sell Service Pvt. Ltd.</strong></p>
        """
    html_body = html_body.replace("</body>", f"{closing_note}</body>")
    return html_body

def send_email(gmail_user, app_password, to_emails, subject, html_body, cc=None):
    msg = MIMEMultipart('alternative')
    msg['From'] = gmail_user
    msg['To'] = ", ".join(to_emails)
    if cc:
        msg['Cc'] = ", ".join(cc)
    msg['Subject'] = subject
    msg.attach(MIMEText(html_body, 'html'))
    recipients = to_emails + (cc if cc else [])
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(gmail_user, app_password)
        server.sendmail(gmail_user, recipients, msg.as_string())

st.set_page_config(page_title="Payment Reconciliation", layout="wide")

if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    pwd = st.text_input("Enter Admin Password", type="password")
    if st.button("Login"):
        if check_password(pwd):
            st.session_state.auth = True
        else:
            st.error("Invalid password")
    st.stop()

st.title("📧 Payment Mail Sender Dashboard")
col1, col3 = st.columns(2)
with col1:
    st.download_button(
        label="📥 Download Payment Sample Excel",
        data=create_sample_excel(),
        file_name="SampleInvoices.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
with col3:
    st.download_button(
        label="📥 Download Mail Sample Excel",
        data=create_sample_mail_excel(),
        file_name="SampleMail.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.subheader("📁 Upload Party Emails via Excel (One Time Only)")
with st.expander("🔑 Protected Upload", expanded=False):
    upload_pass = st.text_input("Enter password to upload email list:", type="password")
    if upload_pass == EMAIL_UPLOAD_PASSWORD:
        email_upload = st.file_uploader("Upload Party Email Excel", type=["xlsx"], key="email_uploader")
        if email_upload:
            try:
                email_df = pd.read_excel(email_upload)
                if "Party Code" in email_df.columns and "Email" in email_df.columns:
                    updated_json = []
                    missing_emails = []
                    for _, row in email_df.iterrows():
                        party = str(row["Party Name"]).strip()
                        code = str(row['Party Code']).strip()
                        emails = str(row['Email']).strip()
                        cc = str(row['CC']).strip() if 'CC' in row else ''
                        updated_json.append({"PartyCode": code, "Email": emails, "PartyName": party, "CC": cc})
                        if not emails or emails.lower() in ['nan', 'none', '']:
                            missing_emails.append(f"{party} ({code})")
                    save_party_emails(updated_json)
                    st.success("✅ Party email list updated from Excel!")
                    if missing_emails:
                        st.warning(
                            "⚠️ The following vendors have no email addresses in your file:\n" +
                            "\n".join(missing_emails)
                        )
                else:
                    st.error("Excel must contain 'Party Code' 'Party Name' 'CC' and 'Email' columns.")
            except Exception as e:
                st.error(f"Error reading Excel: {e}")
    elif upload_pass:
        st.error("❌ Incorrect password!")

st.subheader("📁 Upload Payment Details Excel")
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
if uploaded_file:
    with open(EXCEL_PATH, "wb") as f:
        f.write(uploaded_file.getbuffer())
    st.success("Excel uploaded. Processing...")

    payment_df, debit_df = load_excel(EXCEL_PATH)
    st.subheader("Payment Details Sheet Columns")
    st.write(payment_df.columns.tolist())
    st.subheader("Debit Notes Sheet Columns")
    st.write(debit_df.columns.tolist())
    party_emails = load_party_emails()
    st.subheader("📬 Party Emails")
    party_names = [e['PartyCode'] for e in party_emails]
    selected_party = st.selectbox("Select Party to Edit Emails", [""] + party_names)
    if selected_party:
        idx = next((i for i, e in enumerate(party_emails) if e['PartyCode'] == selected_party), None)
        if idx is not None:
            new_email = st.text_input(f"Emails for {selected_party}", party_emails[idx]['Email'])
            pwd_confirm = st.text_input(f"Confirm Password to Update Emails for {selected_party}", type="password")
            if st.button("Update Emails"):
                if pwd_confirm == "password":
                    party_emails[idx]['Email'] = new_email
                    save_party_emails(party_emails)
                    st.success(f"Emails updated for {selected_party}")
                else:
                    st.error("Incorrect password. Emails not updated.")

    st.subheader("📧 Gmail Settings")
    gmail_user = st.text_input("Your Gmail")
    gmail_pwd = st.text_input("App Password (Use Gmail App Password)", type="password")

    if gmail_user and gmail_pwd:
        matched_results, skips, parties_without_email = match_data(payment_df, debit_df, party_emails)
        
        # Display parties without email addresses in card format
        if parties_without_email:
            st.subheader("⚠️ Parties Without Email Addresses")
            
            # Summary metrics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Parties", len(parties_without_email))
            with col2:
                total_payment_records = sum(party['payment_count'] for party in parties_without_email)
                st.metric("Total Payment Records", total_payment_records)
            with col3:
                avg_records = total_payment_records / len(parties_without_email) if parties_without_email else 0
                st.metric("Avg Records per Party", f"{avg_records:.1f}")
            
            st.markdown("---")
            
            # Show parties without email in card format
            st.subheader("📋 Parties Requiring Email Setup")
            
            # Create columns for better layout
            cols_per_row = 2
            email_columns = st.columns(cols_per_row)
            
            for i, party in enumerate(parties_without_email):
                col_idx = i % cols_per_row
                
                with email_columns[col_idx]:
                    # Create a card-like container
                    with st.container():
                        st.markdown(f"""
                        <div style="
                            background-color: #fff3cd; 
                            border: 1px solid #ffeaa7; 
                            border-radius: 8px; 
                            padding: 15px; 
                            margin: 5px 0;
                            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                        ">
                            <h4 style="color: #856404; margin: 0 0 8px 0;">🏢 {party['party_name']}</h4>
                            <p style="margin: 2px 0; color: #6c757d;"><strong>Code:</strong> {party['party_code']}</p>
                            <p style="margin: 2px 0; color: #6c757d;"><strong>Payment Records:</strong> {party['payment_count']}</p>
                            <div style="
                                background-color: #f8d7da; 
                                color: #721c24; 
                                padding: 5px 10px; 
                                border-radius: 4px; 
                                font-size: 0.85em; 
                                margin-top: 8px;
                                text-align: center;
                            ">
                                ⚠️ Email Required
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                
                # Add new row after every cols_per_row items
                if (i + 1) % cols_per_row == 0 and i < len(parties_without_email) - 1:
                    st.markdown("---")
            
            # Add download option for parties without email
            email_missing_df = pd.DataFrame(parties_without_email)
            csv_no_email = email_missing_df.to_csv(index=False)
            st.download_button(
                label="📥 Download Parties Without Email (CSV)",
                data=csv_no_email,
                file_name="parties_without_email.csv",
                mime="text/csv"
            )
            
            st.markdown("---")
            
            # Add action section
            st.subheader("🔧 Next Steps")
            st.info("""
            **To enable email sending for these parties:**
            1. Update the party email list via the protected upload section above
            2. Ensure each party has a valid email address
            3. Re-upload the payment Excel file to reprocess
            """)
        
        st.subheader("✅ Ready to Email")
        for entry in matched_results:
            with st.expander(entry['party_code']):
                st.json(entry)
        # Display skipped parties (minimal format)
        if skips:
            st.subheader("⏭️ Skipped Parties")
            
            # Count skip reasons
            skip_reasons = {}
            for line in skips:
                reason = line.split(" — ")[1] if " — " in line else "Unknown reason"
                skip_reasons[reason] = skip_reasons.get(reason, 0) + 1
            
            # Show summary
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Skipped", len(skips))
            with col2:
                st.metric("Processed", len(matched_results))
            with col3:
                st.metric("Success Rate", f"{(len(matched_results)/(len(matched_results)+len(skips))*100):.1f}%")
            
            # Show skip reasons in compact format
            for reason, count in skip_reasons.items():
                st.warning(f"**{count} parties**: {reason}")
            
            # Add download option for skip list
            skip_data = []
            for line in skips:
                if " — " in line:
                    party_info, reason = line.split(" — ", 1)
                    party_code = party_info.replace("SKIPPED: ", "").strip()
                    skip_data.append({"Party Code": party_code, "Skip Reason": reason})
            
            if skip_data:
                skip_df = pd.DataFrame(skip_data)
                csv = skip_df.to_csv(index=False)
                st.download_button(
                    label="📥 Download Skip List (CSV)",
                    data=csv,
                    file_name="skipped_parties.csv",
                    mime="text/csv"
                )

        # ------------- SMTP FIXED EMAIL LOOP ------------
        if st.button("Send Emails"):
            log_lines = []
            sent_count = 0
            failed_count = 0
            skips = []
            log_lines.append("=== Emails Sent Successfully ===")
            for entry in matched_results:
                party_code = entry['party_code']  # This is actually PartyName since we match by name
                party_name = next((e['PartyName'] for e in party_emails if e['PartyName'] == party_code), party_code if party_code else 'Unknown Party')
                cc_str = next((e.get('CC', '') for e in party_emails if e['PartyName'] == party_code), '')
                cc_emails = [email.strip() for email in cc_str.split(',')] if cc_str else []
                invoice_list = list({item['InvoiceNo'] for item in entry['payments'] + entry['debits'] if 'InvoiceNo' in item})
                html_body = generate_email_body(party_code, entry['payments'], entry['debits'])
                try:
                    send_email(
                        gmail_user,
                        gmail_pwd,
                        entry['emails'],
                        f"Payment Reconciliation for {party_code} - {party_name}",
                        html_body,
                        cc=cc_emails
                    )
                    st.success(f"✅ Email sent to {party_name} ({party_code})")
                    log_lines.append(f"Party Code: {party_code} | Party Name: {party_name} | Emails: {', '.join(entry['emails'])} | CC: {', '.join(cc_emails)}")
                    sent_count += 1
                except Exception as e:
                    st.error(f"❌ Failed for {party_code}: {e}")
                    log_lines.append(f"FAILED: {party_code} | Error: {e}")
                    failed_count += 1
                time.sleep(random.uniform(1, 5))  # Random delay to avoid SMTP connection refused errors
            log_lines.append("\n=== Skipped Parties ===")
            if skips:
                for line in skips:
                    log_lines.append(line)
            else:
                log_lines.append("None")
            with open("FinalEmailLog.txt", "w", encoding="utf-8") as log_file:
                for line in log_lines:
                    log_file.write(line + "\n")
            st.success(f"✅ Emails sent: {sent_count}, Failed: {failed_count}, Skipped: {len(skips)}")
        # ----------- END SMTP SENDING LOOP ------------

        st.subheader("📂 Download All Party-wise Sheets in One Excel File")
        if 'matched_results' in locals() and matched_results:
            partywise_output = BytesIO()
            with pd.ExcelWriter(partywise_output, engine='xlsxwriter') as writer:
                for party in matched_results:
                    party_code = party['party_code']
                    df = pd.DataFrame(party['payments'])
                    df_debit = pd.DataFrame(party['debits'])
                    sheet_name_payment = f"{party_code[:28]}_Pay"
                    sheet_name_debit = f"{party_code[:28]}_Debit"
                    df.to_excel(writer, index=False, sheet_name=sheet_name_payment)
                    if not df_debit.empty:
                        df_debit.to_excel(writer, index=False, sheet_name=sheet_name_debit)
            partywise_output.seek(0)
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            filename = f"All_Partywise_Payments_{timestamp}.xlsx"
            st.download_button(
                label="📥 Download All Party-wise Payments (Excel)",
                data=partywise_output,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with open("FinalEmailLog.txt", "rb") as log_file:
            st.download_button(
                label="📄 Download Final Email Log",
                data=log_file,
                file_name="FinalEmailLog.txt",
                mime="text/plain"
            )
st.subheader("📊 Convert Final Email Log to Excel")
if os.path.exists("FinalEmailLog.txt"):
    with open("FinalEmailLog.txt", "r", encoding="utf-8") as f:
        lines = f.readlines()
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("Email Log")
    headers = ["Status", "Party Code", "Party Name", "Emails / Error"]
    for col, header in enumerate(headers):
        worksheet.write(0, col, header)
    row_num = 1
    for line in lines:
        line = line.strip()
        if line.startswith("Party Code:"):
            parts = line.replace("Party Code:", "").split("|")
            party_code = parts[0].strip()
            party_name = parts[1].replace("Party Name:", "").strip() if len(parts) > 1 else ""
            emails = parts[2].replace("Emails:", "").strip() if len(parts) > 2 else ""
            worksheet.write_row(row_num, 0, ["SENT", party_code, party_name, emails])
            row_num += 1
        elif line.startswith("FAILED:"):
            parts = line.replace("FAILED:", "").split("|")
            party_code = parts[0].strip()
            error = parts[1].replace("Error:", "").strip() if len(parts) > 1 else ""
            worksheet.write_row(row_num, 0, ["FAILED", party_code, "", error])
            row_num += 1
        elif line.startswith("SKIPPED:"):
            worksheet.write_row(row_num, 0, ["SKIPPED", "", "", line])
            row_num += 1
    workbook.close()
    output.seek(0)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"FinalEmailLog_{timestamp}.xlsx"
    st.download_button(
        label="📥 Download Log as Excel",
        data=output,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def create_partywise_zip(send_data):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for party in send_data:
            party_code = str(party['party_code']).strip()
            df = pd.DataFrame(party['payments'])
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name="Payments")
                writer.save()
            excel_buffer.seek(0)
            zip_file.writestr(f"{party_code}.xlsx", excel_buffer.read())
    zip_buffer.seek(0)
    return zip_buffer
