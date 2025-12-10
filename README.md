# 🐍 Python Automation — Personalized Email Sender with Auto Cover Letter

This **Python automation project** enables you to send fully personalized emails to recruiters or contacts **without sending each one manually**.  
It dynamically fills in recipient details, attaches resumes & project PDFs, and even generates a **customized cover letter PDF** for every recipient automatically! 🚀  

---

## ✨ Overview

This project automates **job outreach and networking emails**. It reads recipient data from a CSV file, personalizes email templates (HTML + Word), converts them to PDFs, attaches relevant files, and sends them via Gmail — all with one script.  

It’s ideal for:
- Job seekers sending tailored applications 📄  
- Networkers reaching out to recruiters 💼  
- Professionals managing bulk personalized emails 📬  

---

## ⚙️ Key Features

| ✅ Feature | 💡 Description |
|------------|----------------|
| **CSV-Driven Personalization** | Reads `Name`, `Email`, `Company` from CSV |
| **Dynamic Email Template** | Auto-fills placeholders inside HTML |
| **Auto DOCX → PDF Conversion** | Creates company-specific cover letters |
| **Inline Images Support** | Embeds banners or signatures in the email |
| **Smart Attachments** | Adds Resume, Projects PDF, and generated Cover Letter |
| **Excel Logging** | Logs all emails with timestamp & status in `email_log.xlsx` |
| **Bounce Monitoring (Optional)** | Launches `bounce_handler.py` for failed email tracking |
| **Secure Gmail Login** | Uses App Password authentication via SSL |

---


🧠 How It Works

Load Recipient Data

Name,Email,Company
John Doe,john@example.com,Microsoft
Jane Smith,jane@abc.com,Azure


Replace Placeholders
The script replaces {{Name}} and {{Company}} inside:

email_template.html (Email Body)

base_doc.docx (Cover Letter)

Convert Cover Letter
DOCX → PDF per recipient using win32com.

Attach Files
Resume + Projects + Generated Cover Letter.

Send Email
Securely sends via Gmail’s SMTP (SSL, Port 465).

Log Deliveries
Each email is logged in email_log.xlsx with time, name, company, and status.

Monitor Bounces (Optional)
Launches bounce_handler.py to monitor failed sends.

💻 Setup Instructions
🧰 Step 1: Install Dependencies
pip install python-docx openpyxl pywin32

🔐 Step 2: Configure Gmail App Password

Enable 2-Step Verification in Gmail.

Generate an App Password for “Mail”.

Add credentials inside the script:

GMAIL_USER = 'your_email@gmail.com'
GMAIL_PASS = 'your_app_password'

📋 Step 3: Create contacts.csv
Name,Email,Company
Sahil,sahil@example.com,Azure
Mangesh,mangesh@example.com,Microsoft

📝 Step 4: Prepare Templates

email_template.html → contains placeholders like {{Name}} and {{Company}}

base_doc.docx → personalized cover letter Word template

▶️ Step 5: Run the Script
python send_mails.py

📩 Example Output
📨 Sending email to John Doe at john@example.com (Microsoft)
✅ Email sent successfully!

📨 Sending email to Jane Smith at jane@abc.com (Azure)
✅ Email sent successfully!

🎉 All emails prepared and sent successfully!
├── image1.png ... image4.png  # Inline images (logos/banners)
├── email_log.xlsx             # Generated log of sent emails
├── bounce_handler.py          # Optional bounce monitor script
└── README.md
