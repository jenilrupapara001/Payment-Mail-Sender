import streamlit as st
import pandas as pd
import os
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
import xlsxwriter
import hashlib
import pyodbc
import webbrowser
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from datetime import datetime

# Constants
JSON_PATH = Path("party_emails.json")
EXCEL_PATH = Path("Invoices.xlsx")
EMAIL_UPLOAD_PASSWORD = "PaymentMailSenderDashboard"


connection_string =(
    "DRIVER= {ODBC Driver 17 for SQL Server};"
    "SERVER=(localdb)\\MSSQLLocalDB;"
    "DATABASE= Payment_Email;"
    "Trusted_Connection=yes;"
)

def get_db_connection():
    return pyodbc.connect(connection_string)

def already_sent(PartyCode, InvoiceNo):
    conn = get_db_connection()
    cursor = conn.cursor()
    query = "SELECT * FROM dbo.EmailLogs WHERE PartyCode = ? AND InvoiceNo = ?"
    cursor.execute(query, (PartyCode,InvoiceNo))
    result = cursor.fetchone()
    cursor.close()
    conn.close()
    return result is not None

def record_sent_email(PartyCode, PartyName, InvoiceNo):
    conn = get_db_connection()
    cursor = conn.cursor()
    query = "INSERT INTO dbo.EmailLogs (PartyCode, PartyName, InvoiceNo, TimeStamp) VALUES (?, ?, ?, ?)"
    cursor.execute(query, (PartyCode, PartyName, InvoiceNo, datetime.now()))
    conn.commit()
    cursor.close()
    conn.close()


    
# Sample Excel creation function (for download button)
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
        "Amount": [9500, 20000],  # for matching with debit note
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

# Party Mail Sample Excel 
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

# Embedded HTML email template
EMAIL_TEMPLATE = """
<html>
  <body style="font-family: Arial, sans-serif; color: #333;">
    <h2 style="color: #2c3e50;">Payment Reconciliation Summary</h2>
    <p>Dear [Party Name],</p>
    <p>Please find below the summary of your recent transactions with us:</p>

    <h3>Purchase & Payment Details</h3>
    <table style="border-collapse: collapse;  width: 100%; margin-bottom: 20px;">
      <thead>
        <tr style="background-color: #f2f2f2; border: 2px solid #333;">
          <th style="border: 1px solid #333; padding: 8px; ">Purchase Bill</th>
          <th style="border: 1px solid #ddd; padding: 8px; ">Pur. Date</th>
          <th style="border: 1px solid #ddd; padding: 8px; ">Amount Rs.</th>
          <th style="border: 1px solid #ddd; padding: 8px; ">Debit Note</th>
          <th style="border: 1px solid #ddd; padding: 8px; ">Total Payment</th>
          <th style="border: 1px solid #ddd; padding: 8px; ">Bank Payment</th>
          <th style="border: 1px solid #ddd; padding: 8px; ">Payment Date</th>
        </tr>
      </thead>
      <tbody>
        <!-- Dynamic payment rows inserted here -->
      </tbody>
    </table>

    <!-- Debit Note section will only appear if replaced -->

   
  
  </body>
</html>
"""

# Helper: Load JSON party-email mapping
def load_party_emails():
    if not JSON_PATH.exists():
        # Create a sample JSON file if not exists
        sample = [
            {"PartyName": "Alpha Corp", "Email": "alpha@example.com"},
            {"PartyName": "Beta Ltd", "Email": "beta@example.com"}
        ]
        with open(JSON_PATH, 'w') as f:
            json.dump(sample, f, indent=2)
    with open(JSON_PATH, 'r') as f:
        return json.load(f)

# Helper: Save updated party-email mapping
def save_party_emails(data):
    with open(JSON_PATH, 'w') as f:
        json.dump(data, f, indent=2)

# Helper: Hash and check password
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def check_password(input_pwd):
    # Here default password is 'password'
    return hash_password(input_pwd) == hash_password("password")

# Helper: Load Excel Data
def load_excel(file_path):
    wb = pd.ExcelFile(file_path)
    payment_df = wb.parse("Payment Details")
    debit_df = wb.parse("Debit Notes")
    # Strip column names
    payment_df.columns = payment_df.columns.str.strip()
    debit_df.columns = debit_df.columns.str.strip()
    return payment_df, debit_df

# Match data function with debit note logic
# Match data function with debit note logic
def match_data(payment_df, debit_df, party_emails):
    email_map = {e["PartyCode"].strip(): [email.strip() for email in e["Email"].split(",")] for e in party_emails}

    payment_df.columns = payment_df.columns.str.strip()
    debit_df.columns = debit_df.columns.str.strip()

    result = []
    mismatch_log_lines = []
    skip_log_lines = []

    for party_code, emails in email_map.items():
        party_payments = payment_df[payment_df['Party Code'].astype(str).str.strip() == party_code.strip()]
        if party_payments.empty:
            skip_log_lines.append(f"SKIPPED: {party_code} — No payment rows found in Payment Sheet")
            continue

        related_debits = debit_df[debit_df['Party Code'].astype(str).str.strip() == party_code.strip()] if 'Party Code' in debit_df.columns else pd.DataFrame()
        total_debit_amount = related_debits['Amount'].sum() if not related_debits.empty else 0

        # Match debit amount with total
        party_payments['Debit Amount'] = party_payments['Debit Amount'].fillna(0)
        party_debit_sum = party_payments['Debit Amount'].sum()

        if abs(party_debit_sum - total_debit_amount) > 0.01:  # check if mismatch
            skip_log_lines.append(f"SKIPPED: {party_code} — Debit Amount mismatch between payment sheet and debit sheet")
            continue

        payment_issues = []

        for _, row in party_payments.iterrows():
            debit_note = row.get('Debit Note') if 'Debit Note' in row else None

            if debit_note is None or (pd.isna(debit_note) or debit_note == ''):
                # No debit note means include payment row as usual
                payment_issues.append(row.to_dict())
            else:
                matched_debit_rows = related_debits[related_debits['Return Invoice No.'] == debit_note] if 'Return Invoice No.' in related_debits.columns else pd.DataFrame()
                if matched_debit_rows.empty:
                    payment_issues.append(row.to_dict())
                else:
                    amount_in_debit_sheet = matched_debit_rows.iloc[0]['Amount']
                    if abs(row['Net Amount'] - amount_in_debit_sheet) < 0.01:
                        payment_issues.append(row.to_dict())
                    else:
                        mismatch_log_lines.append(
                            f"Mismatch DebitNote: {debit_note} | Party: {party_code} | Payment Sheet Amount: {row['Net Amount']} | Debit Sheet Amount: {amount_in_debit_sheet}"
                        )

        if payment_issues:
            result.append({
                'party_code': party_code,
                'emails': emails,
                'payments': payment_issues,
                'debits': related_debits.to_dict(orient='records') if not related_debits.empty else []
            })
        else:
            skip_log_lines.append(f"SKIPPED: {party_code} — All payment rows matched with debit notes correctly.")

    # Write logs
    if skip_log_lines:
        with open('SkippedPartiesLog.txt', 'w') as f:
            for line in skip_log_lines:
                f.write(line + "\n")

    if mismatch_log_lines:
        with open('MismatchLog.txt', 'w') as f:
            for line in mismatch_log_lines:
                f.write(line + "\n")

    return result, skip_log_lines


# Generate email HTML body
def generate_email_body(party_code, payment_rows, debit_rows):
    # Fetch PartyName from the party_emails JSON using party_code
    party_name = next((e['PartyName'] for e in party_emails if e['PartyCode'] == party_code), 'Unknown Party')

    template = EMAIL_TEMPLATE
    payment_html = ""

    # Totals for the summary row
    total_inv_amount = 0
    total_net_amount = 0
    total_bank_payment = 0

    for row in payment_rows:
        # Format Payment Date (remove timestamp)
        payment_date = row.get('Payment Date', '')
        try:
            payment_date_str = pd.to_datetime(payment_date).strftime('%Y-%m-%d')
        except Exception:
            payment_date_str = str(payment_date).split(" ")[0]

        # Format Bank Payment
        bank_payment = row.get('Bank Payment', '')
        if isinstance(bank_payment, pd.Timestamp) or (' ' in str(bank_payment)):
            bank_payment = str(bank_payment).split(' ')[0]

        debit_note_val = row.get('Debit Amount', 'N/A')

        payment_html += f"""
        <tr style="text-align:center; border:1px solid #ccc;">
          <td style="border:1px solid #ccc;">{row.get('Inv. No.', '')}</td>
          <td style="border:1px solid #ccc;">{row.get('Pur. Date', '')}</td>
          <td style="border:1px solid #ccc;">{row.get('Total Inv. Amount', '')}</td>
          <td style="border:1px solid #ccc;">{debit_note_val}</td>
          <td style="border:1px solid #ccc;">{row.get('Net Amount', '')}</td>
          <td style="border:1px solid #ccc;">{bank_payment}</td>
          <td style="border:1px solid #ccc;">{payment_date_str}</td>
        </tr>"""

        # Accumulate totals
        total_inv_amount += float(row.get('Total Inv. Amount', 0) or 0)
        total_net_amount += float(row.get('Net Amount', 0) or 0)
        total_bank_payment += float(row.get('Bank Payment', 0) or 0)

    # Add total row under payments
    payment_html += f"""
    <tr style="text-align:center; font-weight:bold; background-color:#f9f9f9;">
      <td colspan="2" style="border:1px solid #ccc;">Total</td>
      <td style="border:1px solid #ccc;">{total_inv_amount:.2f}</td>
      <td style="border:1px solid #ccc;">-</td>
      <td style="border:1px solid #ccc;">{total_net_amount:.2f}</td>
      <td style="border:1px solid #ccc;">{total_bank_payment:.2f}</td>
      <td style="border:1px solid #ccc;">-</td>
    </tr>"""

    # Inject payment HTML
    html_body = template.replace("[Party Name]", party_name)
    html_body = html_body.replace("<!-- Dynamic payment rows inserted here -->", payment_html)

    # Add debit section if needed
    if debit_rows:
        debit_html = """
        <h3>Return/Debit Details</h3>
        <table style="border-collapse: collapse; width: auto;text-align:center">
          <thead>
            <tr style="background-color: #f2f2f2; border: 2px solid #333;">
              <th style="border: 2px solid #333; padding: 8px;">Date</th>
              <th style="border: 2px solid #333; padding: 8px;">Return Invoice No.</th>
              <th style="border: 2px solid #333; padding: 8px;">Amount</th>
            </tr>
          </thead>
          <tbody>
        """

        total_debit_amount = 0

        for row in debit_rows:
            date_str = row.get('Date', '')
            try:
                date_str = pd.to_datetime(date_str).strftime("%Y-%m-%d")
            except Exception:
                pass

            amount = float(row.get('Amount', 0) or 0)
            total_debit_amount += amount

            debit_html += f"""
            <tr style="border: 1px solid #ccc; text-align: center;">
              <td style="border:1px solid #ccc;">{date_str}</td>
              <td style="border:1px solid #ccc;">{row.get('Return Invoice No.', '')}</td>
              <td style="border:1px solid #ccc;">{amount:.2f}</td>
            </tr>"""

        # Add total row to debit section
        debit_html += f"""
        <tr style="border:1px solid #ccc; background-color: #f9f9f9;">
          <td colspan="2" style="text-align:right; font-weight:bold; border:1px solid #ccc;">Total Debit Amount:</td>
          <td style="border:1px solid #ccc; text-align:center; font-weight:bold;">{total_debit_amount:.2f}</td>
        </tr>
        </tbody>
        </table>
        """
        html_body = html_body.replace("</body>", debit_html + "</body>")

    # Closing message
    html_body = html_body.replace(
    '<p>Let us know if you have any questions regarding this reconciliation.</p>',
    '<p>If you have any concerns or questions regarding this summary, please raise them within 7 days of receiving this email. After this period, no further action will be taken.</p>'
)

   

    return html_body






# Send email via Gmail SMTP
def send_email(gmail_user, app_password, to_emails, subject, html_body):
    msg = MIMEMultipart('alternative')
    msg['From'] = gmail_user
    msg['To'] = ", ".join(to_emails)
    msg['Subject'] = subject
    msg.attach(MIMEText(html_body, 'html'))
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(gmail_user, app_password)
        server.sendmail(gmail_user, to_emails, msg.as_string())

# --------------- STREAMLIT UI ---------------
# --------------- STREAMLIT UI ---------------
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

# Download sample Excel button
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
                    for _, row in email_df.iterrows():
                        party = str(row["Party Name"]).strip()
                        code = str(row['Party Code']).strip()
                        emails = str(row['Email']).strip()
                        updated_json.append({"PartyCode": code, "Email": emails, "PartyName": party})
                    save_party_emails(updated_json)
                    st.success("✅ Party email list updated from Excel!")
                else:
                    st.error("Excel must contain'Party Code' 'Party Name' and 'Email' columns.")
            except Exception as e:
                st.error(f"Error reading Excel: {e}")
    elif upload_pass:
        st.error("❌ Incorrect password!")

# Upload Excel
st.subheader("📁 Upload Payment Details Excel")
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
if uploaded_file:
    with open(EXCEL_PATH, "wb") as f:
        f.write(uploaded_file.getbuffer())
    st.success("Excel uploaded. Processing...")

    payment_df, debit_df = load_excel(EXCEL_PATH)

    # Show columns of both sheets (debug + info)
    st.subheader("Payment Details Sheet Columns")
    st.write(payment_df.columns.tolist())

    st.subheader("Debit Notes Sheet Columns")
    st.write(debit_df.columns.tolist())

    # Load party emails json
    party_emails = load_party_emails()

    st.subheader("📬 Party Emails")
    party_names = [e['PartyCode'] for e in party_emails]

    # Dropdown to select party to edit emails
    selected_party = st.selectbox("Select Party to Edit Emails", [""] + party_names)

    if selected_party:
        # Find the entry
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
        matched_results, skips = match_data(payment_df, debit_df, party_emails)

        st.subheader("✅ Ready to Email")

        for entry in matched_results:
            with st.expander(entry['party_code']):
                st.json(entry)

        if skips:
            st.warning(f"Skipped Parties: {len(skips)}")
            for line in skips:
                st.text(line)




        if st.button("Send Emails"):
            log_lines = []
            sent_count = 0
            failed_count = 0
            skip_count = 0

    log_lines.append("=== Emails Sent Successfully ===")

    for entry in matched_results:
        party_code = entry['party_code']
        party_name = next(
            (e['PartyName'] for e in party_emails if e['PartyCode'] == party_code),
            'Unknown Party'
        )
        invoice_nos = [row.get('Inv. No.', '') for row in entry['payments'] if row.get('Inv. No.', '')]

        # Skip if any invoice was already emailed
        if any(already_sent(party_code, inv) for inv in invoice_nos):
            st.warning(f"SKIPPED (Already Sent): {party_code} - {party_name}")
            log_lines.append(f"SKIPPED: {party_code} — Already emailed for at least one invoice")
            skip_count += 1
            continue

        html_body = generate_email_body(party_code, entry['payments'], entry['debits'])

        try:
            send_email(
                gmail_user,
                gmail_pwd,
                entry['emails'],
                f"Payment Reconciliation for {party_name}",
                html_body
            )

            st.success(f"✅ Email sent to {party_name} ({party_code})")
            log_lines.append(f"Party Code: {party_code} | Party Name: {party_name} | Emails: {', '.join(entry['emails'])}")
            sent_count += 1

            # Log each invoice sent to DB
            for inv in invoice_nos:
                record_sent_email(party_code, party_name, inv)

        except Exception as e:
            st.error(f"❌ Failed for {party_code}: {e}")
            log_lines.append(f"FAILED: {party_code} | Error: {e}")
            failed_count += 1

    # Add skipped logs
        log_lines.append("\n=== Skipped Parties ===")
        if skips:
            for line in skips:
                log_lines.append(line)
                skip_count += 1
        else:
            log_lines.append("None")

        # Write to log file
        with open("FinalEmailLog.txt", "w", encoding="utf-8") as log_file:
            for line in log_lines:
                log_file.write(line + "\n")

    # Final summary
    st.success(f"✅ Emails sent: {sent_count}, ❌ Failed: {failed_count}, ⏭️ Skipped: {skip_count}")

            # Download log
st.subheader("📊 Convert Final Email Log to Excel")

try:
    if os.path.exists("FinalEmailLog.txt"):
        with open("FinalEmailLog.txt", "r", encoding="utf-8") as f:
            lines = f.readlines()

        # Create Excel in memory
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet("Email Log")

        # Header row
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
    else:
        st.info("📄 No FinalEmailLog.txt found yet. Please send emails first.")
except Exception as e:
    st.error(f"Error reading or creating Excel log: {e}")

    
