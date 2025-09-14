import os
import sys
import time
import csv
import logging
import smtplib
from pathlib import Path
from datetime import datetime, date
from typing import List, Tuple, Dict

import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.utils import formataddr, make_msgid
from email import encoders

from dotenv import load_dotenv
load_dotenv()


class BirthdayEmailSystem:
    def __init__(self):
        self.setup_logging()
        self.load_configuration()

        # ------------ COMPANY IMAGE PATHS (EDIT THESE) -------------------
        self.company_images = {
            'Company1': r"C:\path\to\your\assets\Company1BdayCard.jpg",
            'Company2': r"C:\path\to\your\assets\Company2BdayCard.jpg",
            'Company3': r"C:\path\to\your\assets\Company3BdayCard.jpg",
            'Company4': r"C:\path\to\your\assets\Company4BdayCard.jpg",
            'Company': ""  # fallback (optional)
        }

        # ------------ COMPANY WEBSITE LINE (OPTIONAL) --------------------
        self.company_sites = {
            'Company1': "WWW.Company1.com",
            'Company2': "WWW.Company2.com",
            'Company3': "WWW.Company3.com",
            'Company4': "WWW.Company4.com",
            'Company': ""
        }

    def setup_logging(self):
        log_file = Path("birthday_emails.log")
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[logging.FileHandler(log_file, mode='a', encoding='utf-8'),
                      logging.StreamHandler(sys.stdout)]
        )
        self.logger = logging.getLogger(__name__)
        self.logger.info("")
        self.logger.info("")
        self.logger.info("=" * 80)
        self.logger.info("NEW PROGRAM RUN STARTED")
        self.logger.info("=" * 80)
        self.logger.info("")
        self.logger.info("Logging system initialized - using birthday_emails.log")
        self.logger.info(f"Current time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # ------------------------ CONFIG ------------------------------------

    def load_configuration(self):
        """Load common & per-company config from environment variables."""
        # Common (behavior, pacing, generic fallbacks — NOT SMTP)
        self.config = {
            'attach_path': os.getenv('ATTACH_PATH'),
            'dry_run': os.getenv('DRY_RUN', 'false').lower() == 'true',
            'delay_between_sends': float(os.getenv('DELAY_BETWEEN_SENDS', 0.5)),
            # NEW: per-company switch delay (simple anti-spam pacing)
            'delay_between_companies': float(os.getenv('DELAY_BETWEEN_COMPANIES', 2.5)),
            'p_status_filter': os.getenv('P_STATUS_FILTER', 'A').upper(),  # A, T, BOTH
            'smtp_timeout': int(os.getenv('SMTP_TIMEOUT', 30)),
            'fallback_email_cc': self._parse_list(os.getenv('EMAIL_CC', '')),
            'fallback_email_bcc': self._parse_list(os.getenv('EMAIL_BCC', '')),
            'fallback_team_name_template': os.getenv('TEAM_NAME_TEMPLATE', '{company} HR Team'),
            'fallback_subject_template': os.getenv('SUBJECT_TEMPLATE', '🎉 Happy Birthday, {first_name}! - {company} Team'),
            'fallback_email_reputation_domain': os.getenv('EMAIL_REPUTATION_DOMAIN', ''),
        }

        # Build per-company configs
        self.company_configs: Dict[str, dict] = {
            'Company1': self._build_company_config('COMPANY1', 'Company1'),
            'Company2':  self._build_company_config('COMPANY2',  'Company2'),
            'Company3': self._build_company_config('COMPANY3', 'Company3'),
            'Company4':  self._build_company_config('COMPANY4',  'Company4'),
            'Company': self._build_company_config('', 'Company')  # fallback
        }

        self.logger.info(f"Configuration loaded. Dry run: {self.config['dry_run']}")
        self.logger.info(f"P_Status filter: {self.config['p_status_filter']}")
        for name, cfg in self.company_configs.items():
            masked_user = (cfg.get('smtp_user') or '')[:1] + '***' if cfg.get('smtp_user') else '(not set)'
            self.logger.info(f"Company cfg [{name}] host={cfg.get('smtp_host','')} port={cfg.get('smtp_port','')} user={masked_user} rep_domain={cfg.get('email_reputation_domain','')}")

    def _parse_list(self, raw: str) -> List[str]:
        return [e.strip() for e in raw.split(',') if e.strip()]

    def _build_company_config(self, prefix: str, company_label: str) -> dict:
        """
        Build a single company config with overrides. If prefix == '', we use global SMTP_* as fallback.
        """
        def gv(key: str, default: str = '') -> str:
            if prefix:
                return os.getenv(f'{prefix}_{key}', default)
            # fallback to non-prefixed env (global)
            return os.getenv(key, default)

        cfg = {
            'company': company_label,
            'smtp_host': gv('SMTP_HOST', ''),
            'smtp_port': int(gv('SMTP_PORT', os.getenv('SMTP_PORT', '587') or '587')),
            'smtp_user': gv('SMTP_USER', ''),
            'smtp_pass': gv('SMTP_PASS', ''),
            'email_cc': self._parse_list(gv('EMAIL_CC', os.getenv('EMAIL_CC', ''))),
            'email_bcc': self._parse_list(gv('EMAIL_BCC', os.getenv('EMAIL_BCC', ''))),
            'team_name_template': gv('TEAM_NAME_TEMPLATE', self.config['fallback_team_name_template']),
            'subject_template': gv('SUBJECT_TEMPLATE', self.config['fallback_subject_template']),
            'email_reputation_domain': gv('EMAIL_REPUTATION_DOMAIN', self.config['fallback_email_reputation_domain'] or (company_label.lower() + '.com')),
            'use_authentication': True,
        }
        cfg['connection_security'] = 'SSL' if cfg['smtp_port'] == 465 else 'STARTTLS'
        return cfg

    def get_company_config(self, company: str) -> dict:
        """Return the config for the detected company, else fallback to 'Company'."""
        cfg = self.company_configs.get(company) or self.company_configs['Company']
        # Validate minimal SMTP fields
        missing = [k for k in ('smtp_host', 'smtp_port', 'smtp_user', 'smtp_pass') if not cfg.get(k)]
        if missing:
            raise ValueError(f"Missing SMTP settings for {company}: {', '.join(missing)}")
        return cfg

    # --- Data loading & normalization -----------------------------------

    def detect_company_from_path(self, file_path: str) -> str:
        """
        Detect company name from Excel filename (case-insensitive).
        Handles: Company1, Company2, Company3, Company4
        """
        name = Path(file_path).name.upper()
        mapping = {
            'COMPANY1': 'Company1',
            'COMPANY2': 'Company2',
            'COMPANY3': 'Company3',
            'COMPANY4': 'Company4',
        }
        for token, proper in mapping.items():
            if token in name:
                self.logger.info(f"Detected company name from filename: {proper}")
                return proper
        self.logger.info("No known company token detected in filename; defaulting to generic 'Company'")
        return 'Company'

    def normalize_column_names(self, df: pd.DataFrame) -> pd.DataFrame:
        df.columns = df.columns.str.strip()
        mapping = {}
        for col in df.columns:
            lc = col.lower().replace('_', '').strip()
            if ('emp' in lc) and ('id' in lc):
                mapping[col] = 'Emp_Id'
            elif 'firstname' in lc or ('first' in lc and 'name' in lc):
                mapping[col] = 'First_Name'
            elif 'lastname' in lc or ('last' in lc and 'name' in lc):
                mapping[col] = 'Last_Name'
            elif col.strip().upper() == 'DOB':
                mapping[col] = 'DOB'
            elif col.strip() == 'P_Email1':
                mapping[col] = 'P_Email1'
            elif 'pstatus' in lc or 'p_status' in col.lower():
                mapping[col] = 'P_Status'
        return df.rename(columns=mapping)

    def load_and_validate_data(self, excel_path: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """Load Confidential, Contact Details, Employee Status."""
        try:
            confidential_df = pd.read_excel(excel_path, sheet_name='Confidential')
            contact_df = pd.read_excel(excel_path, sheet_name='Contact Details')
            status_df = pd.read_excel(excel_path, sheet_name='Employee Status')

            confidential_df = self.normalize_column_names(confidential_df)
            contact_df = self.normalize_column_names(contact_df)
            status_df = self.normalize_column_names(status_df)

            req_conf = ['Emp_Id', 'First_Name', 'Last_Name', 'DOB']
            req_contact = ['Emp_Id', 'First_Name', 'Last_Name', 'P_Email1']
            req_status = ['Emp_Id', 'First_Name', 'Last_Name', 'P_Status']

            for col in req_conf:
                if col not in confidential_df.columns:
                    raise ValueError(f"Required column '{col}' not found in Confidential")
            for col in req_contact:
                if col not in contact_df.columns:
                    raise ValueError(f"Required column '{col}' not found in Contact Details")
            for col in req_status:
                if col not in status_df.columns:
                    raise ValueError(f"Required column '{col}' not found in Employee Status")

            confidential_df['DOB_Parsed'] = pd.to_datetime(confidential_df['DOB'], errors='coerce')
            confidential_df = confidential_df[confidential_df['DOB_Parsed'].notna()]

            self.logger.info(f"Loaded {len(confidential_df)} records from Confidential sheet")
            self.logger.info(f"Loaded {len(contact_df)} records from Contact Details sheet")
            self.logger.info(f"Loaded {len(status_df)} records from Employee Status sheet")

            return confidential_df, contact_df, status_df

        except Exception as e:
            self.logger.error(f"Error loading data from {excel_path}: {str(e)}")
            raise

    def filter_todays_birthdays(self, confidential_df: pd.DataFrame, status_df: pd.DataFrame) -> pd.DataFrame:
        """Filter for today's birthdays and apply P_Status filter."""
        today = date.today()
        mask = (
            (confidential_df['DOB_Parsed'].dt.month == today.month) &
            (confidential_df['DOB_Parsed'].dt.day == today.day)
        )
        df = confidential_df[mask].copy()

        if df.empty:
            self.logger.info(f"No birthdays on {today.strftime('%Y-%m-%d')}. Exiting.")
            return pd.DataFrame()

        df = df.merge(status_df[['Emp_Id', 'P_Status']], on='Emp_Id', how='left')

        p = self.config['p_status_filter']
        if p == 'A':
            self.logger.info("Filtering for P_Status = 'A' (Active) only")
            df = df[df['P_Status'] == 'A']
        elif p == 'T':
            self.logger.info("Filtering for P_Status = 'T' (Terminated) only")
            df = df[df['P_Status'] == 'T']
        elif p == 'BOTH':
            self.logger.info("Including all P_Status values (A, T, and others)")
        else:
            self.logger.warning(f"Invalid P_STATUS_FILTER value: {p}. Using A (Active only).")
            df = df[df['P_Status'] == 'A']

        self.logger.info(f"Found {len(df)} birthdays on {today.strftime('%Y-%m-%d')}")
        for _, person in df.iterrows():
            self.logger.info(
                f"Birthday Person: Emp_Id={person.get('Emp_Id','N/A')}, "
                f"Name={person.get('First_Name','')} {person.get('Last_Name','')}, "
                f"DOB={person.get('DOB','')}, P_Status={person.get('P_Status','')}"
            )
        return df

    def join_email_data(self, birthdays_df: pd.DataFrame, contact_df: pd.DataFrame) -> pd.DataFrame:
        """Join with contacts and validate/deduplicate emails."""
        joined = birthdays_df.merge(
            contact_df[['Emp_Id', 'First_Name', 'Last_Name', 'P_Email1']],
            on='Emp_Id', how='left', suffixes=('', '_contact')
        )
        joined['Greeting_Name'] = joined['First_Name_contact'].fillna(joined['First_Name']).astype(str).str.title()
        joined['Email'] = joined['P_Email1']

        self.logger.info("People with birthdays and their email addresses:")
        for _, row in joined.iterrows():
            self.logger.info(
                f"   Emp_Id={row.get('Emp_Id','N/A')}, "
                f"Name={row.get('Greeting_Name','N/A')}, "
                f"Email={row.get('Email','N/A')}"
            )

        email_re = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        valid = joined['Email'].notna() & joined['Email'].astype(str).str.match(email_re)
        result = joined[valid].copy()

        invalid = joined[~valid]
        if not invalid.empty:
            self.logger.warning("People with invalid email addresses:")
            for _, row in invalid.iterrows():
                self.logger.warning(
                    f"   Emp_Id={row.get('Emp_Id','N/A')}, "
                    f"Name={row.get('Greeting_Name','N/A')}, "
                    f"Email={row.get('Email','N/A')}"
                )

        result = result.drop_duplicates(subset=['Email'], keep='first')

        self.logger.info(f"After email validation and deduplication: {len(result)} recipients")
        self.logger.info("Final recipients for birthday emails:")
        for _, row in result.iterrows():
            self.logger.info(
                f"   Emp_Id={row.get('Emp_Id','N/A')}, "
                f"Name={row.get('Greeting_Name','N/A')}, "
                f"Email={row.get('Email','N/A')}"
            )
        return result[['Greeting_Name', 'Email']]

    # --- Logging of send attempts ----------------------------------------------------

    def log_send_attempt(self, recipient: str, first_name: str, source_file: str,
                         company: str, status: str, response: str = "",
                         message_id: str = "", spam_score: str = "N/A"):
        today_str = datetime.now().strftime("%Y-%m-%d")
        log_file = Path("C:/logs/birthday_sends_{today_str}.csv")
        log_file.parent.mkdir(parents=True, exist_ok=True)
        write_header = not log_file.exists()

        try:
            with open(log_file, 'a', newline='', encoding='utf-8') as f:
                fieldnames = ['Timestamp', 'Recipient', 'First_Name', 'Source_File',
                              'Company', 'Status', 'Response', 'Message_ID', 'Spam_Score']
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                if write_header:
                    writer.writeheader()
                writer.writerow({
                    'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'Recipient': recipient,
                    'First_Name': first_name,
                    'Source_File': Path(source_file).name,
                    'Company': company,
                    'Status': status,
                    'Response': response,
                    'Message_ID': message_id,
                    'Spam_Score': spam_score
                })
        except Exception as e:
            self.logger.error(f"Error writing to log file: {e}")

    # --- Email compose & send --------------------------------------------------------

    def create_email_message(self, recipient: str, first_name: str,
                              source_filename: str, company: str, cfg: dict):
        """
        Create the email with company-specific From/Reply-To, CC/BCC, subject, body, and attach image.
        Also set Message-ID domain from EMAIL_REPUTATION_DOMAIN when provided.
        """
        team_name = cfg['team_name_template'].format(company=company)
        site_text = self.company_sites.get(company, "") or ""

        msg = MIMEMultipart('mixed')  # allow attachments
        msg['From'] = formataddr((team_name, cfg['smtp_user']))
        msg['To'] = recipient
        subject = cfg.get('subject_template', '🎉 Happy Birthday, {first_name}! - {company} Team').format(
            first_name=first_name, company=company
        )
        msg['Subject'] = subject
        msg['Reply-To'] = cfg['smtp_user']

        # Reputation / alignment
        rep_domain = cfg.get('email_reputation_domain')
        if not rep_domain:
            try:
                rep_domain = cfg['smtp_user'].split('@', 1)[1]
            except Exception:
                rep_domain = 'localhost'

        message_id = make_msgid(domain=rep_domain)
        msg['Message-ID'] = message_id
        msg['Return-Path'] = cfg['smtp_user']
        msg['Sender'] = cfg['smtp_user']
        msg['MIME-Version'] = '1.0'

        # --- Simplified headers (removed non-essential headers) ---
        # msg['X-Mailer'] = f'{company} Birthday System v1.0'
        # msg['X-Priority'] = '3'
        # msg['Importance'] = 'Normal'
        # msg['X-MSMail-Priority'] = 'Normal'
        # msg['List-Unsubscribe'] = f'<mailto:{cfg["smtp_user"]}?subject=Unsubscribe>'

        if cfg['email_cc']:
            msg['Cc'] = ', '.join(cfg['email_cc'])

        # Alternative part for text + HTML
        alt = MIMEMultipart('alternative')
        msg.attach(alt)

        # --- BODY (your wording, dynamic company) ---
        text_body = f"""Dear {first_name},

Here's wishing you a very Happy Birthday on behalf of our {company} Team. Hope you are having a Blast !!

Be Blessed! Have a Great day!

{team_name}
{site_text}""".rstrip()

        html_body = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Happy Birthday from {company}</title>
</head>
<body style="font-family: Arial, Helvetica, sans-serif; line-height: 1.6; color: #333333; max-width: 600px; margin: 0 auto; padding: 20px; background-color: #ffffff;">
    <div style="border: 1px solid #e0e0e0; padding: 30px; border-radius: 8px; background-color: #fefefe;">
        <p style="margin: 0 0 15px 0;">Dear {first_name},</p>
        
        <p style="margin: 0 0 15px 0;">Here's wishing you a very Happy Birthday on behalf of our <strong>{company} Team</strong>. Hope you are having a Blast !!</p>
        
        <p style="margin: 0 0 15px 0;">Be Blessed! Have a Great day!</p>
        
        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #e0e0e0;">
            <p style="margin: 0; font-weight: bold;">{team_name}</p>
            {f'<p style="margin: 5px 0 0 0; color: #666666;">{site_text}</p>' if site_text else ''}
        </div>

        <div style="margin-top: 30px; font-size: 12px; color: #888888; text-align: center;">
             <p style="font-size: 12px; color: #888888; text-align: center;">
                 Sent with warm wishes from the {company} HR Team.  
                 Contact: {cfg['smtp_user']}
             </p>
        </div>
    </div>
</body>
</html>""".rstrip()

        alt.attach(MIMEText(text_body, 'plain', 'utf-8'))
        alt.attach(MIMEText(html_body, 'html', 'utf-8'))

        # Attach company-specific image as file attachment
        img_path = self.company_images.get(company, "") or ""
        if img_path and Path(img_path).exists():
            try:
                ext = Path(img_path).suffix.lower()
                if ext in ['.jpg', '.jpeg']:
                    maintype, subtype = 'image', 'jpeg'
                elif ext == '.png':
                    maintype, subtype = 'image', 'png'
                elif ext == '.gif':
                    maintype, subtype = 'image', 'gif'
                else:
                    maintype, subtype = 'application', 'octet-stream'

                with open(img_path, 'rb') as f:
                    data = f.read()

                part = MIMEBase(maintype, subtype)
                part.set_payload(data)
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{Path(img_path).name}"')
                part.add_header('Content-ID', f'<birthday_card_{company.lower()}>')
                msg.attach(part)

                self.logger.info(f"Attached birthday card for {company}")
            except Exception as e:
                self.logger.warning(f"Could not attach image for {company}: {e}")
        else:
            if img_path:
                self.logger.warning(f"Attachment image path not found for {company}: {img_path}")

        # Optional small generic attachment (<=200KB)
        if self.config['attach_path']:
            attach_path = Path(self.config['attach_path'])
            if attach_path.exists():
                try:
                    if attach_path.stat().st_size <= 200 * 1024:
                        with open(attach_path, 'rb') as f:
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(f.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename="{attach_path.name}"')
                        msg.attach(part)
                except Exception as e:
                    self.logger.warning(f"Could not attach generic file: {e}")

        return msg, message_id

    def _connect_smtp(self, cfg: dict) -> smtplib.SMTP:
        """Create and return an SMTP connection using the provided company config."""
        if cfg['smtp_port'] == 465:
            server = smtplib.SMTP_SSL(cfg['smtp_host'], cfg['smtp_port'], timeout=self.config['smtp_timeout'])
        else:
            server = smtplib.SMTP(cfg['smtp_host'], cfg['smtp_port'], timeout=self.config['smtp_timeout'])
            server.starttls()
        if cfg['use_authentication']:
            server.login(cfg['smtp_user'], cfg['smtp_pass'])
        server.noop()
        self.logger.info(f"Connected to SMTP server for {cfg['company']} ({cfg['smtp_host']}:{cfg['smtp_port']} - {'SSL' if cfg['smtp_port']==465 else 'STARTTLS'})")
        return server

    def send_emails(self, recipients_df: pd.DataFrame, source_file: str, company: str, cfg: dict):
        """Send birthday emails sequentially with a small delay between sends."""
        if self.config['dry_run']:
            self.logger.info("DRY RUN MODE - No emails will be sent")

        self.logger.info(f"[{company}] Using CC: {cfg['email_cc']}")
        self.logger.info(f"[{company}] Using BCC: {cfg['email_bcc']}")

        server = None
        try:
            if not self.config['dry_run']:
                try:
                    server = self._connect_smtp(cfg)
                except Exception as e:
                    self.logger.error(f"SMTP connection failed for {company}: {str(e)}")
                    raise

            sent_count = 0
            failed_count = 0

            for _, row in recipients_df.iterrows():
                recipient = row['Email']
                first_name = row['Greeting_Name']

                try:
                    msg, message_id = self.create_email_message(recipient, first_name, Path(source_file).name, company, cfg)
                    all_recipients = [recipient] + cfg['email_cc'] + cfg['email_bcc']

                    if self.config['dry_run']:
                        status, response = "Sent (Dry Run)", "Dry run mode"
                        self.logger.info(f"DRY RUN: Would send to {recipient} ({first_name}) [{company}]")
                        self.log_send_attempt(recipient, first_name, source_file, company, status, response,
                                               message_id=message_id, spam_score="N/A")
                    else:
                        send_result = server.send_message(msg, from_addr=cfg['smtp_user'], to_addrs=all_recipients)
                        if send_result:
                            status, response = "Partial Failure", str(send_result)
                            failed_count += 1
                        else:
                            status, response = "Sent", "Success"
                            sent_count += 1

                        self.logger.info(f"Sent birthday email to {recipient} ({first_name}) [{company}]")
                        self.log_send_attempt(recipient, first_name, source_file, company, status, response,
                                               message_id=message_id, spam_score="N/A")

                        # pacing between individual sends
                        time.sleep(self.config['delay_between_sends'])

                except Exception as e:
                    self.logger.error(f"Failed to send email to {recipient}: {str(e)}")
                    self.log_send_attempt(recipient, first_name, source_file, company, "Failed", str(e),
                                           message_id="", spam_score="N/A")
                    failed_count += 1

            if not self.config['dry_run'] and server is not None:
                try:
                    server.quit()
                except Exception:
                    pass
                self.logger.info("Disconnected from SMTP server")

            self.logger.info(f"[{company}] Email sending summary: Sent={sent_count} Failed={failed_count}")

        except Exception as e:
            self.logger.error(f"SMTP connection error for {company}: {str(e)}")
            raise

    # --- Orchestration ---------------------------------------------------------------

    def process_file(self, excel_path: str):
        """Process a single Excel file for birthday emails."""
        try:
            self.logger.info(f"Processing file: {excel_path}")

            company = self.detect_company_from_path(excel_path)
            self.logger.info(f"Company for this file: {company}")
            company_cfg = self.get_company_config(company)

            confidential_df, contact_df, status_df = self.load_and_validate_data(excel_path)
            birthdays_df = self.filter_todays_birthdays(confidential_df, status_df)
            if birthdays_df.empty:
                return

            recipients_df = self.join_email_data(birthdays_df, contact_df)
            if recipients_df.empty:
                self.logger.warning("No valid email addresses found for birthday recipients")
                return

            self.send_emails(recipients_df, excel_path, company, company_cfg)

        except Exception as e:
            self.logger.error(f"Error processing file {excel_path}: {str(e)}")
            raise

    def run(self, excel_files: List[str]):
        self.logger.info("Starting Birthday Email System (No-Batch)")
        for idx, excel_file in enumerate(excel_files):
            if not Path(excel_file).exists():
                self.logger.error(f"File not found: {excel_file}")
                continue
            try:
                self.process_file(excel_file)
            except Exception as e:
                self.logger.error(f"Failed to process {excel_file}: {str(e)}")
                continue

            # NEW: small delay between switching companies/files
            if idx < len(excel_files) - 1:
                self.logger.info(f"Pacing between company files for {self.config['delay_between_companies']} seconds...")
                time.sleep(self.config['delay_between_companies'])

        self.logger.info("Birthday Email System completed")


def main():
    # Example usage - update these as needed
    excel_files = [
        r"C:\path\to\your\excel_files\COMPANY1_MASTER_EXCEL_HR_2025.xlsx",
        r"C:\path\to\your\excel_files\COMPANY2_MASTER_EXCEL_HR_2025.xlsx",
        r"C:\path\to\your\excel_files\COMPANY3_MASTER_EXCEL_HR_2025.xlsx",
        r"C:\path\to\your\excel_files\COMPANY4_MASTER_EXCEL_HR_2025.xlsx"
    ]
    try:
        app = BirthdayEmailSystem()
        app.run(excel_files)
    except Exception as e:
        logging.error(f"Fatal error: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
