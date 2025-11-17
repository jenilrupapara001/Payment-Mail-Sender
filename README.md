# 🧾 Payment Mail Sender Dashboard

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.0+-red.svg)](https://streamlit.io/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

A modern, user-friendly Streamlit-based Python application designed to automate the process of sending payment reconciliation emails to parties. It matches payment and debit note data from uploaded Excel files, validates transactions, and sends personalized HTML emails via Gmail SMTP.

## ✨ Features

- **📊 Excel Data Processing**: Upload and process payment details and debit notes from Excel files
- **🔍 Automated Matching**: Automatically match payments with debit notes per party
- **📧 Personalized Email Generation**: Create and send customized HTML emails with transaction summaries
- **🔐 Secure Authentication**: Password-protected interface for sensitive operations
- **📈 Real-time Validation**: Validate payment and debit amounts across sheets
- **📋 Comprehensive Logging**: Track sent, failed, and skipped emails with downloadable logs
- **📥 Sample Data Downloads**: Download sample Excel templates for easy setup
- **📊 Export Capabilities**: Export final email logs and party-wise payment summaries to Excel
- **🎨 Modern UI**: Clean, responsive Streamlit interface with expandable sections

## 🚀 Installation

### Prerequisites

- Python 3.8 or higher
- Gmail account with App Password (for email sending)

### Setup Steps

1. **Clone the repository:**
   ```bash
   git clone https://github.com/yourusername/payment-mail-sender.git
   cd payment-mail-sender
   ```

2. **Create a virtual environment (recommended):**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the application:**
   ```bash
   streamlit run mail.py
   ```

5. **Access the app:**
   Open your browser and navigate to `http://localhost:8501`

## 📖 Usage

### 1. Initial Setup

- **Login**: Enter the admin password to access the dashboard
- **Upload Party Emails**: Use the protected upload section to import party email lists via Excel (one-time setup)

### 2. Data Preparation

- **Download Samples**: Get sample Excel templates for payment details and party emails
- **Prepare Excel Files**:
  - **Payment Details Sheet**: Contains columns like Party Code, Invoice No., Purchase Date, Total Amount, Debit Amount, Net Amount, Bank Payment, Payment Date
  - **Debit Notes Sheet**: Contains return invoice details with amounts

### 3. Email Configuration

- Enter your Gmail address and App Password
- The app uses Gmail SMTP for sending emails

### 4. Process and Send

- Upload your payment Excel file
- Review matched data and validation results
- Send emails to all eligible parties
- Download logs and summaries

### 5. Monitoring

- View real-time status of email sending
- Download comprehensive logs in text and Excel formats
- Export party-wise payment summaries

## 📋 Requirements

- **Python**: 3.8+
- **Dependencies**:
  - streamlit
  - pandas
  - numpy
  - openpyxl
  - xlsxwriter
  - pyodbc

## 🔧 Configuration

### Gmail Setup

1. Enable 2-Factor Authentication on your Gmail account
2. Generate an App Password:
   - Go to Google Account settings
   - Security → 2-Step Verification → App passwords
   - Generate password for "Mail"
3. Use the App Password (not your regular password) in the app

### Excel Format Requirements

**Payment Details Sheet:**
- Party Code
- Inv. No.
- Pur. Date
- Total Inv. Amount
- Debit Amount
- Net Amount
- Bank Payment
- Payment Date

**Debit Notes Sheet:**
- Party Code
- Date
- Return Invoice No.
- Amount

**Party Email Sheet:**
- Party Code
- Party Name
- Email (comma-separated for multiple)
- CC (optional, comma-separated)

## 🛠️ Development

### Project Structure

```
payment-mail-sender/
├── mail.py                 # Main Streamlit application
├── party_emails.json       # Party email database
├── requirements.txt        # Python dependencies
├── README.md              # This file
└── .devcontainer/         # Development container config
```

### Key Components

- **Data Processing**: Pandas-based Excel parsing and validation
- **Email Generation**: HTML template system for professional emails
- **SMTP Integration**: Secure Gmail SMTP with throttling
- **Logging System**: Comprehensive error and success tracking

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## ⚠️ Disclaimer

This application is designed for business use in payment reconciliation processes. Ensure compliance with email marketing regulations and obtain necessary consents before sending bulk emails. The developers are not responsible for misuse of this tool.

## 🆘 Support

For issues or questions:
- Check the logs for error details
- Ensure all Excel columns match the required format
- Verify Gmail App Password configuration
- Confirm internet connectivity for SMTP

---

**Built with ❤️ using Streamlit and Python**
