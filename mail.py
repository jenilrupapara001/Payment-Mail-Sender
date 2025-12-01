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
    payment_df = wb.parse("Payment Details")
    debit_df = wb.parse("Debit Notes")
    payment_df.columns = payment_df.columns.str.strip()
    debit_df.columns = debit_df.columns.str.strip()
    return payment_df, debit_df

def match_data(payment_df, debit_df, party_emails):
    email_map = {
        e["PartyCode"].strip(): {
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
    
    # Check for different column names that might represent party code/name
    payment_party_col = None
    if 'Party Code' in payment_df.columns:
        payment_party_col = 'Party Code'
    elif 'Party Name' in payment_df.columns:
        payment_party_col = 'Party Name'
    
    debit_party_col = None
    if 'Party Code' in debit_df.columns:
        debit_party_col = 'Party Code'
    elif 'Party Name' in debit_df.columns:
        debit_party_col = 'Party Name'
    
    # First, identify parties in payment sheet that have no email
    parties_without_email = []
    if payment_party_col:
        payment_party_codes = set(payment_df[payment_party_col].astype(str).str.strip())
        for party_code in payment_party_codes:
            if party_code not in email_map or not email_map[party_code]["to"] or all(email.strip().lower() in ['nan', 'none', ''] for email in email_map[party_code]["to"]):
                party_name = party_code  # Default to party code if name not found
                # Try to find party name from party_emails
                for party in party_emails:
                    if party["PartyCode"].strip() == party_code.strip():
                        party_name = party["PartyName"]
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
        total_debit_amount = related_debits['Amount'].sum() if not related_debits.empty else 0
        party_payments['Debit Amount'] = party_payments['Debit Amount'].fillna(0)
        party_debit_sum = party_payments['Debit Amount'].sum()

        if abs(party_debit_sum - total_debit_amount) > 0.01:
            skip_log_lines.append(f"SKIPPED: {party_code} — Debit Amount mismatch between payment sheet and debit sheet")
            continue

        payment_issues = []
        for _, row in party_payments.iterrows():
            debit_note = row.get('Debit Note') if 'Debit Note' in row else None
            if debit_note is None or (pd.isna(debit_note) or debit_note == ''):
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
    party_name = next((e['PartyName'] for e in party_emails if e['PartyCode'] == party_code), 'Unknown Party')
    template = EMAIL_TEMPLATE
    payment_html = ""
    total_inv_amount = 0
    total_net_amount = 0
    total_bank_payment = 0            
    for row in payment_rows:
        payment_date = row.get('Payment Date', '')
        try:
            payment_date_str = pd.to_datetime(payment_date).strftime('%Y-%m-%d')
        except Exception:
            payment_date_str = str(payment_date).split(" ")[0]
        bank_payment = row.get('Bank Payment', '')
        if isinstance(bank_payment, pd.Timestamp) or (' ' in str(bank_payment)):
            bank_payment = str(bank_payment).split(' ')[0]
        # Handle NaN and missing values
        inv_no = row.get('Inv. No.', '')
        pur_date = row.get('Pur. Date', '')
        total_inv_amount = row.get('Total Inv. Amount', '')
        debit_note_val = row.get('Debit Amount', '')
        net_amount = row.get('Net Amount', '')
        
        # Replace NaN and empty values with '-'
        inv_no = '-' if pd.isna(inv_no) or inv_no == '' else str(inv_no)
        pur_date = '-' if pd.isna(pur_date) or pur_date == '' else str(pur_date)
        total_inv_amount = '-' if pd.isna(total_inv_amount) or total_inv_amount == '' else str(total_inv_amount)
        debit_note_val = '-' if pd.isna(debit_note_val) or debit_note_val == '' else str(debit_note_val)
        net_amount = '-' if pd.isna(net_amount) or net_amount == '' else str(net_amount)
        bank_payment = '-' if pd.isna(bank_payment) or bank_payment == '' else str(bank_payment)
        
        payment_html += f"""
        <tr style="text-align:center; border:1px solid #ccc;">
          <td style="border:1px solid #ccc;">{inv_no}</td>
          <td style="border:1px solid #ccc;">{pur_date}</td>
          <td style="border:1px solid #ccc;">{total_inv_amount}</td>
          <td style="border:1px solid #ccc;">{debit_note_val}</td>
          <td style="border:1px solid #ccc;">{net_amount}</td>
          <td style="border:1px solid #ccc;">{bank_payment}</td>
          <td style="border:1px solid #ccc;">{payment_date_str}</td>
        </tr>"""
        # Handle NaN values for calculations
        total_inv_val = row.get('Total Inv. Amount', 0)
        net_amount_val = row.get('Net Amount', 0)
        bank_payment_val = row.get('Bank Payment', 0)
        
        # Convert to float, replacing NaN with 0
        try:
            total_inv_amount += float(total_inv_val) if not pd.isna(total_inv_val) and total_inv_val != '' else 0
        except (ValueError, TypeError):
            total_inv_amount += 0
            
        try:
            total_net_amount += float(net_amount_val) if not pd.isna(net_amount_val) and net_amount_val != '' else 0
        except (ValueError, TypeError):
            total_net_amount += 0
            
        try:
            total_bank_payment += float(bank_payment_val) if not pd.isna(bank_payment_val) and bank_payment_val != '' else 0
        except (ValueError, TypeError):
            total_bank_payment += 0
    payment_html += f"""
    <tr style="text-align:center; font-weight:bold; background-color:#f9f9f9;">
      <td colspan="2" style="border:1px solid #ccc;">Total</td>
      <td style="border:1px solid #ccc;">{total_inv_amount:.2f}</td>
      <td style="border:1px solid #ccc;">-</td>
      <td style="border:1px solid #ccc;">{total_net_amount:.2f}</td>
      <td style="border:1px solid #ccc;">{total_bank_payment:.2f}</td>
      <td style="border:1px solid #ccc;">-</td>
    </tr>"""
    html_body = template.replace("[Party Name]", party_name)
    html_body = html_body.replace("<!-- Dynamic payment rows inserted here -->", payment_html)
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
            return_inv_no = row.get('Return Invoice No.', '')
            amount_val = row.get('Amount', 0)
            
            # Handle NaN and empty values
            date_str = '-' if pd.isna(date_str) or date_str == '' else str(date_str)
            return_inv_no = '-' if pd.isna(return_inv_no) or return_inv_no == '' else str(return_inv_no)
            
            try:
                date_str = pd.to_datetime(date_str).strftime("%Y-%m-%d") if date_str != '-' else '-'
            except Exception:
                pass
            
            try:
                amount = float(amount_val) if not pd.isna(amount_val) and amount_val != '' else 0
                total_debit_amount += amount
            except (ValueError, TypeError):
                amount = 0
            
            debit_html += f"""
            <tr style="border: 1px solid #ccc; text-align: center;">
              <td style="border:1px solid #ccc;">{date_str}</td>
              <td style="border:1px solid #ccc;">{return_inv_no}</td>
              <td style="border:1px solid #ccc;">{amount:.2f}</td>
            </tr>"""
        debit_html += f"""
        <tr style="border:1px solid #ccc; background-color: #f9f9f9;">
          <td colspan="2" style="text-align:right; font-weight:bold; border:1px solid #ccc;">Total Debit Amount:</td>
          <td style="border:1px solid #ccc; text-align:center; font-weight:bold;">{total_debit_amount:.2f}</td>
        </tr>
        </tbody>
        </table>
        """
        html_body = html_body.replace("</body>", debit_html + "</body>")
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

st.set_page_config(
    page_title="EasySell - Payment Reconciliation Portal", 
    page_icon="🏢", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for professional styling
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #1e3a8a 0%, #3b82f6 100%);
        padding: 20px;
        border-radius: 10px;
        margin-bottom: 30px;
        color: white;
        text-align: center;
    }
    .business-card {
        background-color: #f8fafc;
        border: 2px solid #e2e8f0;
        border-radius: 12px;
        padding: 20px;
        margin: 10px 0;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .metric-container {
        background-color: white;
        border: 1px solid #d1d5db;
        border-radius: 8px;
        padding: 15px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
    }
    .success-box {
        background-color: #ecfdf5;
        border: 1px solid #10b981;
        border-radius: 8px;
        padding: 15px;
        color: #065f46;
    }
    .warning-box {
        background-color: #fef3c7;
        border: 1px solid #f59e0b;
        border-radius: 8px;
        padding: 15px;
        color: #92400e;
    }
    .error-box {
        background-color: #fef2f2;
        border: 1px solid #ef4444;
        border-radius: 8px;
        padding: 15px;
        color: #991b1b;
    }
    .footer {
        position: fixed;
        bottom: 0;
        left: 0;
        right: 0;
        background-color: #1f2937;
        color: white;
        padding: 10px;
        text-align: center;
        font-size: 0.8em;
    }
</style>
""", unsafe_allow_html=True)

if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    # Professional login screen
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div class="main-header">
            <h1>🏢 EasySell Financial Portal</h1>
            <h3>Payment Reconciliation & Communication System</h3>
            <p>Secure Access for Authorized Personnel Only</p>
        </div>
        """, unsafe_allow_html=True)
        
        pwd = st.text_input("🔐 Administrator Password", type="password", placeholder="Enter secure access password")
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("🔓 Secure Login", use_container_width=True):
                if check_password(pwd):
                    st.session_state.auth = True
                    st.rerun()
                else:
                    st.error("❌ Access Denied: Invalid credentials")
        
        st.markdown("""
        <div style="text-align: center; margin-top: 30px; color: #6b7280;">
            <p>© 2024 EasySell Service Pvt. Ltd. | All Rights Reserved</p>
            <p>For technical support, contact IT Department</p>
        </div>
        """, unsafe_allow_html=True)
    st.stop()

# Main dashboard header
st.markdown("""
<div class="main-header">
    <h1>🏢 EasySell Financial Portal</h1>
    <h3>Payment Reconciliation & Vendor Communication System</h3>
    <p>Streamlined Financial Operations Management</p>
</div>
""", unsafe_allow_html=True)
# Sidebar navigation
with st.sidebar:
    st.markdown("""
    <div style="text-align: center; padding: 20px; background-color: #1e3a8a; color: white; border-radius: 10px; margin-bottom: 20px;">
        <h3>🏢 EasySell Portal</h3>
        <p>Financial Operations</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 📋 Navigation Menu")
    selected_section = st.selectbox("Select Operation", [
        "📊 Dashboard Overview",
        "📁 Data Management", 
        "📧 Communication Center",
        "📈 Reports & Analytics",
        "⚙️ System Settings"
    ])
    
    st.markdown("---")
    st.markdown("### 🔐 Session Info")
    st.info(f"**User:** Administrator\n\n**Last Login:** {datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n**Status:** Active")
    
    if st.button("🚪 Secure Logout"):
        st.session_state.auth = False
        st.rerun()

# Main content based on navigation
if selected_section == "📊 Dashboard Overview":
    # Key Performance Indicators
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown("""
        <div class="metric-container">
            <h3 style="color: #1e3a8a;">📊 Total Parties</h3>
            <h2 style="color: #3b82f6; margin: 0;">150+</h2>
            <p style="color: #6b7280; margin: 0;">Active Vendors</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="metric-container">
            <h3 style="color: #1e3a8a;">📧 Emails Sent</h3>
            <h2 style="color: #10b981; margin: 0;">1,250+</h2>
            <p style="color: #6b7280; margin: 0;">This Month</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
        <div class="metric-container">
            <h3 style="color: #1e3a8a;">✅ Success Rate</h3>
            <h2 style="color: #10b981; margin: 0;">98.5%</h2>
            <p style="color: #6b7280; margin: 0;">Delivery Rate</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown("""
        <div class="metric-container">
            <h3 style="color: #1e3a8a;">💰 Amount Processed</h3>
            <h2 style="color: #3b82f6; margin: 0;">₹2.5Cr</h2>
            <p style="color: #6b7280; margin: 0;">This Quarter</p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")

elif selected_section == "📁 Data Management":
    st.subheader("📋 Document Templates & Sample Files")
    st.markdown("""
    <div class="business-card">
        <h4>📊 Standardized Templates</h4>
        <p>Download official Excel templates for payment reconciliation and vendor communication setup.</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="📊 Download Payment Reconciliation Template",
            data=create_sample_excel(),
            file_name="Payment_Reconciliation_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.markdown("*Required format for payment details processing*")
    
    with col2:
        st.download_button(
            label="👥 Download Vendor Database Template",
            data=create_sample_mail_excel(),
            file_name="Vendor_Database_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.markdown("*Required format for vendor email management*")

elif selected_section == "📧 Communication Center":
    st.markdown("""
    <div class="business-card">
        <h3>📧 Vendor Communication Management</h3>
        <p>Streamlined payment reconciliation communication with automated email processing and tracking.</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Step 1: Vendor Database Management
    st.subheader("👥 Step 1: Vendor Database Management")
    st.markdown("""
    <div class="warning-box">
        <strong>🔒 Restricted Access:</strong> Only authorized personnel can modify vendor contact information.
    </div>
    """, unsafe_allow_html=True)
    
    with st.expander("🔐 Database Administration Panel", expanded=False):
        upload_pass = st.text_input("🔑 Administrator Authorization Code", type="password", placeholder="Enter secure access code")
        if upload_pass == EMAIL_UPLOAD_PASSWORD:
            email_upload = st.file_uploader("📊 Upload Vendor Database (Excel Format)", type=["xlsx"], key="email_uploader")
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
                        st.markdown('<div class="success-box">✅ Vendor database successfully updated!</div>', unsafe_allow_html=True)
                        if missing_emails:
                            st.warning(f"⚠️ {len(missing_emails)} vendors missing email addresses")
                    else:
                        st.error("❌ Invalid file format. Required columns: Party Code, Party Name, Email, CC")
                except Exception as e:
                    st.error(f"❌ File processing error: {e}")
        elif upload_pass:
            st.error("❌ Access Denied: Invalid authorization code")

    # Step 2: Payment Data Processing
    st.subheader("📊 Step 2: Payment Data Processing")
    uploaded_file = st.file_uploader("📁 Upload Payment Reconciliation File", type=["xlsx"], help="Upload Excel file containing payment details and debit notes")
    
    if uploaded_file:
        with open(EXCEL_PATH, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        st.markdown('<div class="success-box">✅ File uploaded successfully. Processing data...</div>', unsafe_allow_html=True)

        payment_df, debit_df = load_excel(EXCEL_PATH)
        
        # Data validation summary
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("""
            <div class="business-card">
                <h4>📊 Payment Details Sheet</h4>
                <p><strong>Records:</strong> {}</p>
                <p><strong>Columns:</strong> {}</p>
            </div>
            """.format(len(payment_df), len(payment_df.columns)), unsafe_allow_html=True)
        
        with col2:
            st.markdown("""
            <div class="business-card">
                <h4>📋 Debit Notes Sheet</h4>
                <p><strong>Records:</strong> {}</p>
                <p><strong>Columns:</strong> {}</p>
            </div>
            """.format(len(debit_df), len(debit_df.columns)), unsafe_allow_html=True)
        
        party_emails = load_party_emails()
        
        # Step 3: Email Configuration
        st.subheader("📧 Step 3: Email Configuration")
        gmail_user = st.text_input("📧 Corporate Email Address", placeholder="your.email@company.com")
        gmail_pwd = st.text_input("🔒 Email Application Password", type="password", help="Use Gmail App Password for security")

        if gmail_user and gmail_pwd:
            with st.spinner("🔄 Processing payment reconciliation data..."):
                matched_results, skips, parties_without_email = match_data(payment_df, debit_df, party_emails)
            
            # Data Quality Assessment
            st.subheader("📈 Data Quality Assessment")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("✅ Processed Parties", len(matched_results))
            with col2:
                st.metric("⏭️ Skipped Parties", len(skips))
            with col3:
                st.metric("⚠️ Missing Email", len(parties_without_email))
            with col4:
                success_rate = (len(matched_results) / (len(matched_results) + len(skips)) * 100) if (len(matched_results) + len(skips)) > 0 else 0
                st.metric("📊 Success Rate", f"{success_rate:.1f}%")
            
            st.markdown("---")
            
            # Quality Issues Display
            if parties_without_email or skips:
                st.subheader("⚠️ Quality Issues Requiring Attention")
                
                # Parties without email addresses
                if parties_without_email:
                    st.markdown("### 📧 Vendors Missing Email Addresses")
                    
                    cols_per_row = 2
                    email_columns = st.columns(cols_per_row)
                    
                    for i, party in enumerate(parties_without_email):
                        col_idx = i % cols_per_row
                        
                        with email_columns[col_idx]:
                            st.markdown(f"""
                            <div style="
                                background-color: #fef3cd; 
                                border: 1px solid #fbbf24; 
                                border-radius: 8px; 
                                padding: 15px; 
                                margin: 5px 0;
                            ">
                                <h4 style="color: #92400e; margin: 0;">🏢 {party['party_name']}</h4>
                                <p style="margin: 5px 0;"><strong>Code:</strong> {party['party_code']}</p>
                                <p style="margin: 5px 0;"><strong>Records:</strong> {party['payment_count']}</p>
                                <div style="background-color: #dc2626; color: white; padding: 5px; border-radius: 4px; text-align: center;">
                                    ⚠️ Email Configuration Required
                                </div>
                            </div>
                            """, unsafe_allow_html=True)

                # Skipped parties
                if skips:
                    st.markdown("### ⏭️ Parties Excluded from Processing")
                    
                    skip_reasons = {}
                    for line in skips:
                        reason = line.split(" — ")[1] if " — " in line else "Unknown reason"
                        skip_reasons[reason] = skip_reasons.get(reason, 0) + 1
                    
                    if len(skip_reasons) > 1:
                        st.markdown("#### 📊 Exclusion Reasons Analysis")
                        for reason, count in skip_reasons.items():
                            st.info(f"**{count} parties:** {reason}")
                
                st.markdown("---")
            
            # Communication Preview
            if matched_results:
                st.subheader("✅ Communication Ready Parties")
                st.markdown(f"**{len(matched_results)} parties** are ready for payment reconciliation communication.")
                
                # Bulk Actions
                col1, col2, col3 = st.columns(3)
                with col1:
                    if st.button("🚀 Initiate Bulk Communication", use_container_width=True, type="primary"):
                        # Email sending logic here
                        pass
                with col2:
                    if st.button("📊 Generate Detailed Report", use_container_width=True):
                        # Report generation logic
                        pass
                with col3:
                    if st.button("📁 Export Processing Summary", use_container_width=True):
                        # Export logic
                        pass
                
                # Party Preview
                with st.expander("🔍 Preview Communication Data", expanded=False):
                    for i, entry in enumerate(matched_results[:3]):  # Show first 3 as preview
                        st.markdown(f"#### Party {i+1}: {entry['party_code']}")
                        st.json(entry)
                        if i < 2:  # Don't show separator after last item
                            st.markdown("---")

elif selected_section == "📈 Reports & Analytics":
    st.markdown("""
    <div class="business-card">
        <h3>📈 Financial Analytics & Reporting</h3>
        <p>Comprehensive analysis of payment reconciliation performance and vendor communication metrics.</p>
    </div>
    """, unsafe_allow_html=True)
    
    # KPI Dashboard
    st.subheader("📊 Key Performance Indicators")
    
    # Mock data for demonstration
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown("""
        <div class="metric-container">
            <h4 style="color: #1e3a8a;">📧 Total Communications</h4>
            <h2 style="color: #3b82f6;">1,247</h2>
            <p style="color: #6b7280;">This Quarter</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="metric-container">
            <h4 style="color: #1e3a8a;">✅ Success Rate</h4>
            <h2 style="color: #10b981;">97.8%</h2>
            <p style="color: #6b7280;">Delivery Rate</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
        <div class="metric-container">
            <h4 style="color: #1e3a8a;">💰 Amount Processed</h4>
            <h2 style="color: #3b82f6;">₹2.5Cr</h2>
            <p style="color: #6b7280;">Quarter Total</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown("""
        <div class="metric-container">
            <h4 style="color: #1e3a8a;">⏱️ Avg Response Time</h4>
            <h2 style="color: #10b981;">2.3hrs</h2>
            <p style="color: #6b7280;">Processing Time</p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Report Generation
    st.subheader("📋 Generate Custom Reports")
    col1, col2 = st.columns(2)
    
    with col1:
        report_type = st.selectbox("📊 Report Type", [
            "Payment Reconciliation Summary",
            "Vendor Communication Status",
            "Financial Performance Analysis",
            "System Usage Analytics",
            "Compliance Report"
        ])
        
        date_range = st.date_input("📅 Date Range", value=(datetime.now() - pd.Timedelta(days=30), datetime.now()))
        
    with col2:
        format_type = st.selectbox("📄 Export Format", ["PDF", "Excel", "CSV", "PowerPoint"])
        
        if st.button("🔄 Generate Report", use_container_width=True):
            st.info("🔄 Generating report... This may take a few moments.")
            # Report generation logic would go here

elif selected_section == "⚙️ System Settings":
    st.markdown("""
    <div class="business-card">
        <h3>⚙️ System Administration</h3>
        <p>Configure system parameters, security settings, and operational preferences.</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Security Settings
    st.subheader("🔒 Security Configuration")
    
    with st.expander("🔐 Access Control Settings", expanded=False):
        st.markdown("#### Current Security Status")
        col1, col2 = st.columns(2)
        with col1:
            st.success("✅ Authentication: Active")
            st.success("✅ Session Management: Enabled")
        with col2:
            st.success("✅ Data Encryption: SSL/TLS")
            st.success("✅ Audit Logging: Active")
    
    # System Information
    st.subheader("ℹ️ System Information")
    
    system_info = {
        "System Version": "v2.1.0",
        "Last Updated": "2024-12-01",
        "Database Status": "🟢 Connected",
        "Email Service": "🟢 Gmail API Active",
        "Storage Used": "2.3 GB / 10 GB",
        "Active Sessions": "1"
    }
    
    for key, value in system_info.items():
        col1, col2 = st.columns([1, 2])
        with col1:
            st.text(key)
        with col2:
            st.text(value)

# Professional Footer
st.markdown("""
<div class="footer">
    <p>© 2024 EasySell Service Pvt. Ltd. | Financial Operations Portal | Version 2.1.0</p>
    <p>For technical support: it-support@easysell.in | +91-XXX-XXX-XXXX</p>
</div>
""", unsafe_allow_html=True)

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
