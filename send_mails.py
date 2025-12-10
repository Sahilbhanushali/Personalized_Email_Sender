#!/usr/bin/env python3
"""
job_mailer_gsheets.py

Google Sheets-based job application mailer.

Features:
- Loads rows from a Google Sheet (header row required).
- Validates email addresses (email-validator).
- Removes invalid email rows and writes the cleaned sheet back.
- Selects local resume files (resumes/ folder) based on JobRole mappings.
- Generates a short, human-like cover letter (.txt attachment).
- Sends emails via SMTP with retry logic.
- Updates Status column in the Google Sheet (SENT / FAILED / SKIPPED_...).
- Detailed logging to rotating file.

Setup (summary):
1) Install dependencies:
   pip install google-auth google-auth-oauthlib google-api-python-client pandas email-validator python-dotenv

2) Create OAuth client JSON (client_secrets.json) using your clientId/clientSecret.

3) Enable Google Sheets API for your project and share the target sheet with the client email if required.

4) Create a .env file with SMTP credentials and configuration (or use env vars directly).

Notes:
- The Google Sheets API will open a browser on first run to authorize.
- Ensure resume files are present in the configured directory.
"""

from __future__ import annotations
import os
import sys
import time
import io
import json
import random
import logging
from logging.handlers import RotatingFileHandler
import mimetypes
from typing import Optional, Tuple, Dict, List, Set

import pandas as pd
from email.message import EmailMessage
import smtplib
from email_validator import validate_email, EmailNotValidError
from dotenv import load_dotenv  


from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request  

# ---------------------------
# Load environment
# ---------------------------
load_dotenv()  

# ---------------------------
# Config - change via env vars or edit here
# ---------------------------
CONFIG = {
  
    "sheet_id": os.environ.get("GSHEET_ID", ""),
    "sheet_range": os.environ.get("GSHEET_RANGE", "contacts.csv"),

   
    "client_secrets_file": os.environ.get("GOOGLE_CLIENT_SECRETS", "client_secrets.json"),
    "token_file": os.environ.get("GOOGLE_TOKEN_FILE", "token.json"),

  
    "scopes": ["https://www.googleapis.com/auth/spreadsheets"],



    "resumes_folder": os.environ.get("RESUMES_FOLDER", "resumes"),
    "default_resume": os.environ.get("DEFAULT_RESUME", "resumes/Sahil_Bhanushali-M.pdf"),


    "resume_map": {
        "backend developer": "resumes/Sahil_Bhanushali-M.pdf",
        "frontend developer": "resumes/Sahil_bhanushali_Resume.pdf",
        "mern stack developer": "resumes/fullstack_resume.pdf",
    },

    "cover_letter_map": {
        "backend developer": "cover_letters/backend_backenddev.pdf",
        "frontend developer": "cover_letters/frontend_cover.pdf",
        "mern stack developer": "cover_letters/mern_cover.pdf",
    },
    "default_cover_letter_txt": "cover_letters/default_cover.txt",

    # Sender / SMTP
    "smtp_host": os.environ.get("SMTP_HOST", "smtp.gmail.com"),
    "smtp_port": int(os.environ.get("SMTP_PORT", "587")),
    "smtp_use_ssl": bool(int(os.environ.get("SMTP_USE_SSL", "0"))),  
    "smtp_username": os.environ.get("SMTP_USERNAME", ""), 
    "smtp_password_envvar": os.environ.get("SMTP_PASSWORD_ENVVAR", "SMTP_PASSWORD"),
    "sender_name": os.environ.get("SENDER_NAME", "Sahil Bhanushali"),
    "sender_email": os.environ.get("SENDER_EMAIL", os.environ.get("SMTP_USERNAME", "")),

    # Retry settings
    "max_retries": int(os.environ.get("MAX_RETRIES", "3")),
    "retry_delay_seconds": int(os.environ.get("RETRY_DELAY_SECONDS", "30")),

    # Logging
    "log_file": os.environ.get("LOG_FILE", "job_mailer_gsheets.log"),
    "log_max_bytes": 5 * 1024 * 1024,
    "log_backup_count": 3,

    # Where the candidate saw the job (used in email body)
    "where_found": os.environ.get("WHERE_FOUND", "LinkedIn or company website"),
}

# ---------------------------
# Setup logging
# ---------------------------
def setup_logging(path: str) -> logging.Logger:
    logger = logging.getLogger("job_mailer_gsheets")
    logger.setLevel(logging.DEBUG)
    if not logger.handlers:
        fh = RotatingFileHandler(path, maxBytes=CONFIG["log_max_bytes"], backupCount=CONFIG["log_backup_count"])
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))
        logger.addHandler(fh)

        ch = logging.StreamHandler(sys.stdout)
        ch.setLevel(logging.INFO)
        ch.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"))
        logger.addHandler(ch)
    return logger

logger = setup_logging(CONFIG["log_file"])

# ---------------------------
# Prepare resume & CoverLetter map normalized
# ---------------------------
def normalize_key(s: str) -> str:
    return (s or "").strip().lower()

RESUME_MAP: Dict[str, str] = {normalize_key(k): v for k, v in CONFIG["resume_map"].items()}

COVER_LETTER_MAP: Dict[str, str] = {normalize_key(k): v for k, v in CONFIG["cover_letter_map"].items()}


# ---------------------------
# Google Sheets helpers
# ---------------------------
def get_sheets_service(client_secrets_file: str, token_file: str, scopes: List[str]):
    """
    Create or load OAuth credentials and build a Google Sheets service client.
    Saves token_file for later runs.
    """
    creds = None
    
    if os.path.exists(token_file):
        try:
            creds = Credentials.from_authorized_user_file(token_file, scopes)
        except Exception as e:
            logger.warning("Failed to load token file '%s': %s", token_file, e)


    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                logger.warning("Failed to refresh token: %s", e)
                creds = None

        if not creds:
            if not os.path.exists(client_secrets_file):
                logger.error("Google client secrets file not found: %s", client_secrets_file)
                raise FileNotFoundError("client_secrets.json not found; place your client credentials there.")
            flow = InstalledAppFlow.from_client_secrets_file(client_secrets_file, scopes)
            creds = flow.run_local_server(port=0)
            # Save for next run
            with open(token_file, "w") as tf:
                tf.write(creds.to_json())
            logger.info("Saved OAuth token to %s", token_file)

    try:
        service = build("sheets", "v4", credentials=creds)
    except Exception as e:
        logger.exception("Failed to build Google Sheets service: %s", e)
        raise
    return service

def read_sheet_to_dataframe(service, spreadsheet_id: str, sheet_name_or_range: str) -> pd.DataFrame:
    """
    Read the entire sheet (or named range) into a pandas DataFrame.
    Assumes the first row is the header.
    """
    logger.info("Reading sheet '%s' from spreadsheet %s", sheet_name_or_range, spreadsheet_id)
    sheets = service.spreadsheets()
    try:
        result = sheets.values().get(spreadsheetId=spreadsheet_id, range=sheet_name_or_range).execute()
    except HttpError as e:
        logger.exception("Google Sheets API error getting values: %s", e)
        raise
    values = result.get("values", [])
    if not values:
        logger.warning("Sheet is empty.")
        return pd.DataFrame()

    
    header = [h.strip() for h in values[0]]
    rows = values[1:]
    df = pd.DataFrame(rows, columns=header)
   
    if "Status" not in df.columns:
        df["Status"] = ""
    logger.info("Loaded %d rows from sheet.", len(df))
    return df

def write_dataframe_to_sheet(service, spreadsheet_id: str, sheet_name: str, df: pd.DataFrame) -> None:
    """
    Write entire DataFrame back to sheet, replacing the sheet contents.
    This will overwrite the data in the target sheet.
    """
    logger.info("Writing %d rows back to sheet %s", len(df), sheet_name)
   
    header = list(df.columns)
    values = [header] + df.fillna("").astype(str).values.tolist()
    body = {"values": values}
    try:
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=sheet_name,
            valueInputOption="RAW",
            body=body
        ).execute()
    except HttpError:
        logger.exception("Failed to write back data to sheet.")
        raise

# ---------------------------
# Email & validation helpers
# ---------------------------
def validate_email_address(address: str) -> Tuple[bool, Optional[str]]:
    """
    Validate email using email-validator. Returns (is_valid, normalized_email_or_None)
    """
    if not address or not isinstance(address, str) or not address.strip():
        return False, None
    try:
        v = validate_email(address, check_deliverability=True)  
        return True, v["email"]
    except EmailNotValidError as e:
        logger.warning("Invalid email '%s': %s", address, e)
        return False, None
    except Exception as e:
       
        try:
            v = validate_email(address, check_deliverability=False)
            logger.info("Deliverability check failed for '%s' but syntax ok: %s", address, e)
            return True, v["email"]
        except Exception as e2:
            logger.warning("Email syntax invalid '%s': %s", address, e2)
            return False, None

def select_resume(job_role: str, resume_folder: str, default_resume: Optional[str]) -> Optional[str]:
    """
    Map job_role to a local resume path. Returns path or None.
    Prefers exact mapped keys so 'backend developer' never picks the frontend resume.
    """
    key = normalize_key(job_role)

    # 1) exact mapping first
    if key in RESUME_MAP:
        p = RESUME_MAP[key]
        if os.path.isfile(p):
            return p
        else:
            logger.warning("Mapped resume missing for '%s': %s", job_role, p)

    # 2) conservative heuristic search in folder (whole name/words)
    if os.path.isdir(resume_folder):
        for fname in os.listdir(resume_folder):
            fpath = os.path.join(resume_folder, fname)
            if not os.path.isfile(fpath):
                continue
            name = os.path.splitext(fname)[0].lower()
            words = name.replace("_", " ").split()
            joined = " ".join(words)
            if key == joined or key == name:
                return fpath

    # 3) fallback to default resume
    if default_resume and os.path.isfile(default_resume):
        logger.warning("Using default resume for '%s': %s", job_role, default_resume)
        return default_resume

    logger.error("No resume found for job role '%s'", job_role)
    return None



def select_cover_letter(job_role: str) -> Tuple[Optional[str], bool]:
    """
    Returns (path, is_pdf) for the cover letter.
    Uses mapped PDF if available; otherwise a default txt file.
    """
    key = normalize_key(job_role)

    # mapped cover letter
    if key in COVER_LETTER_MAP:
        p = COVER_LETTER_MAP[key]
        if os.path.isfile(p):
            return p, p.lower().endswith(".pdf")

    # default txt
    txt = CONFIG.get("default_cover_letter_txt")
    if txt and os.path.isfile(txt):
        return txt, False

    return None, False


HTML_TEMPLATES = [
    "templates/email_template.html",
 
]


def render_html_template(job_role: str, company: str, name: Optional[str]) -> str:
    """
    Load one random HTML template and replace simple placeholders.
    """
    path = random.choice(HTML_TEMPLATES)
    with open(path, "r", encoding="utf-8") as f:
        html = f.read()
    html = html.replace("{{Name}}", name or "Hiring Manager")
    html = html.replace("{{Company}}", company or "your company")
    html = html.replace("{{JobRole}}", job_role or "the role")
    return html


def build_email_message(
    to_email: str,
    to_name: Optional[str],
    job_role: str,
    company: str,
    resume_path: str,
    cover_path: Optional[str],
    cover_is_pdf: bool,
    sender_name: str,
    sender_email: str,
    where_found: str
) -> EmailMessage:
    # plain‑text fallback
    greeting = f"Hello {to_name}," if to_name else random.choice(["Hi,", "Hello,"])
    text_body = (
        f"{greeting}\n\n"
        f"I'm reaching out about the {job_role} role at {company}. I saw the opening on {where_found}. "
        f"I've attached my resume and cover letter.\n\n"
        f"Thanks for your time,\n{sender_name}"
    )

    # HTML body from random template
    html_body = render_html_template(job_role, company, to_name)

    msg = EmailMessage()
    msg["From"] = f"{sender_name} <{sender_email}>"
    msg["To"] = to_email
    msg["Subject"] = f"Application for {job_role} - {sender_name}"

    msg.set_content(text_body)
    msg.add_alternative(html_body, subtype="html")

    # ---- inline images for cid:image1, cid:image2, cid:image4 ----
    inline_images = {
        "image1": "image1.png",   # Chat App screenshot
        "image2": "image2.png",   # Job Importer screenshot
        "image4": "image4.png",   # Task Management screenshot
    }

    html_part = msg.get_payload()[-1]  # the text/html part

    for cid, path in inline_images.items():
        if os.path.isfile(path):
            with open(path, "rb") as img:
                img_data = img.read()
            ctype, _ = mimetypes.guess_type(path)
            if ctype is None:
                ctype = "image/png"
            maintype, subtype = ctype.split("/", 1)
            html_part.add_related(
                img_data,
                maintype=maintype,
                subtype=subtype,
                cid=f"<{cid}>",
            )
        else:
            logger.warning("Inline image not found at path: %s", path)
    # --------------------------------------------------------------

    # attach resume
    if not os.path.isfile(resume_path):
        raise FileNotFoundError(f"Resume not found: {resume_path}")
    with open(resume_path, "rb") as rf:
        data = rf.read()
    ctype, _ = mimetypes.guess_type(resume_path)
    if ctype is None:
        ctype = "application/octet-stream"
    maintype, subtype = ctype.split("/", 1)
    msg.add_attachment(
        data,
        maintype=maintype,
        subtype=subtype,
        filename=os.path.basename(resume_path),
    )

    # attach cover letter (PDF or txt)
    if cover_path:
        with open(cover_path, "rb") as cf:
            cdata = cf.read()
        ctype, _ = mimetypes.guess_type(cover_path)
        if ctype is None:
            ctype = "application/octet-stream"
        maintype, subtype = ctype.split("/", 1)
        msg.add_attachment(
            cdata,
            maintype=maintype,
            subtype=subtype,
            filename=os.path.basename(cover_path),
        )

    return msg

        

def send_email(message: EmailMessage, smtp_conf: dict) -> None:
    host = smtp_conf["host"]
    port = smtp_conf["port"]
    user = smtp_conf["username"]
    pwd = smtp_conf["password"]
    use_ssl = smtp_conf.get("use_ssl", False)

    if not user:
        raise RuntimeError("SMTP username is empty; set SMTP_USERNAME in your environment.")
    if not pwd:
        raise RuntimeError("SMTP password is empty; set SMTP_PASSWORD (or matching env) in your environment.")

    logger.info("Connecting to SMTP %s:%s ssl=%s", host, port, use_ssl)
    if use_ssl:
        with smtplib.SMTP_SSL(host, port, timeout=60) as server:
            server.login(user, pwd)
            server.send_message(message)
    else:
        with smtplib.SMTP(host, port, timeout=60) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(user, pwd)
            server.send_message(message)
    logger.info("Sent message to %s", message["To"])

# ---------------------------
# Main orchestration
# ---------------------------
def process_sheet_and_send():
    

    service = get_sheets_service(CONFIG["client_secrets_file"], CONFIG["token_file"], CONFIG["scopes"])

  
    df = read_sheet_to_dataframe(service, CONFIG["sheet_id"], CONFIG["sheet_range"])
    if df.empty:
        logger.info("No data found - exiting.")
        return

    
    required = {"JobRole", "CompanyName", "Email"}
    if not required.issubset(set(df.columns)):
        logger.error("Sheet missing required columns. Must contain at least: %s", required)
        return

   
    if "Status" not in df.columns:
        df["Status"] = ""

   
    invalid_rows = []
    valid_indices = []
    for i, row in df.iterrows():
        email = str(row.get("Email", "")).strip()
        ok, normalized = validate_email_address(email)
        if not ok:
            invalid_rows.append((i, email))
        else:
            df.at[i, "Email"] = normalized
            valid_indices.append(i)

    if invalid_rows:
        logger.warning("Found %d invalid emails - they will be removed and logged.", len(invalid_rows))
        with open("invalid_emails.log", "a", encoding="utf-8") as invf:
            for i, e in invalid_rows:
                invf.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - Row {i} - {e}\n")
       
        df = df.drop([i for i, _ in invalid_rows]).reset_index(drop=True)
        write_dataframe_to_sheet(service, CONFIG["sheet_id"], CONFIG["sheet_range"], df)
        logger.info("Removed invalid rows and updated sheet. Continuing with remaining rows.")

  
    processed_keys: Set[Tuple[str, str, str]] = set()
    smtp_password = os.environ.get(CONFIG["smtp_password_envvar"], "")
    if not smtp_password:
        logger.warning("SMTP password environment variable '%s' not set.", CONFIG["smtp_password_envvar"])

    smtp_conf = {
        "host": CONFIG["smtp_host"],
        "port": CONFIG["smtp_port"],
        "username": CONFIG["smtp_username"],
        "password": smtp_password,
        "use_ssl": CONFIG["smtp_use_ssl"],
    }

    # Work on a copy for iteration
    for idx in range(len(df)):
        row = df.loc[idx]
        job = str(row.get("JobRole", "")).strip()
        company = str(row.get("CompanyName", "")).strip()
        contact = str(row.get("ContactName", "")).strip() if "ContactName" in df.columns else ""
        email = str(row.get("Email", "")).strip()
        location = str(row.get("Location", "")).strip() if "Location" in df.columns else ""
        status = str(row.get("Status", "")).strip()

        logger.info("Row %d: Job='%s' Company='%s' Email='%s' Status='%s'", idx, job, company, email, status)

        if status.upper() == "SENT":
            logger.info("Skipping row %d - already SENT", idx)
            continue
        if status and status.upper() not in ("", "PENDING", "TO_SEND"):
            logger.info("Skipping row %d - status '%s'", idx, status)
            continue

        key = (email.lower(), job.lower(), company.lower())
        if key in processed_keys:
            logger.info("Skipping duplicate in-run for row %d", idx)
            df.at[idx, "Status"] = "SKIPPED_DUPLICATE"
            continue

        resume_path = select_resume(job, CONFIG["resumes_folder"], CONFIG["default_resume"])
        if not resume_path:
            df.at[idx, "Status"] = "SKIPPED_NO_RESUME"
            logger.error("Skipping row %d - no resume for job '%s'", idx, job)
            
            write_dataframe_to_sheet(service, CONFIG["sheet_id"], CONFIG["sheet_range"], df)
            processed_keys.add(key)
            continue
     
        cover_path, cover_is_pdf = select_cover_letter(job)


        try:
            msg = build_email_message(
                to_email=email,
                to_name=contact if contact else None,
                job_role=job,
                company=company,
                resume_path=resume_path,
                cover_path=cover_path,
                cover_is_pdf=cover_is_pdf,
                sender_name=CONFIG["sender_name"],
                sender_email=CONFIG["sender_email"] or CONFIG["smtp_username"],
                where_found=CONFIG["where_found"],
            )
        except Exception:
            logger.exception("Failed to build message for row %d", idx)
            df.at[idx, "Status"] = "FAILED_BUILD_MESSAGE"
            write_dataframe_to_sheet(service, CONFIG["sheet_id"], CONFIG["sheet_range"], df)
            processed_keys.add(key)
            continue


        # send with retries
        attempt = 0
        sent = False
        while attempt < CONFIG["max_retries"] and not sent:
            attempt += 1
            try:
                logger.info("Sending row %d attempt %d/%d to %s", idx, attempt, CONFIG["max_retries"], email)
                send_email(msg, smtp_conf)
                sent = True
            except Exception as e:
                logger.error("Attempt %d failed for row %d: %s", attempt, idx, e)
                if attempt < CONFIG["max_retries"]:
                    logger.info("Sleeping %d seconds before retry...", CONFIG["retry_delay_seconds"])
                    time.sleep(CONFIG["retry_delay_seconds"])
                else:
                    logger.error("All retries failed for row %d", idx)

        if sent:
            df.at[idx, "Status"] = "SENT"
            logger.info("Row %d - SENT", idx)
        else:
            df.at[idx, "Status"] = "FAILED"

        
        try:
            write_dataframe_to_sheet(service, CONFIG["sheet_id"], CONFIG["sheet_range"], df)
        except Exception:
            logger.exception("Failed to persist status update after row %d", idx)

        processed_keys.add(key)

    logger.info("Processing complete.")

if __name__ == "__main__":
    try:
        process_sheet_and_send()
    except Exception as e:
        logger.exception("Unhandled exception: %s", e)
        sys.exit(1)
    