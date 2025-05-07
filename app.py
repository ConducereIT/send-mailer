import os
import smtplib
import time
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Dict, List, Optional, Set

import pandas as pd
import requests
from dotenv import load_dotenv
from loguru import logger


class Settings:
    def __init__(self):
        load_dotenv()
        self.template_path = os.getenv('TEMPLATE_PATH')
        self.sheet_id = os.getenv('GOOGLE_SHEET_ID')
        self.send_mail_host = os.getenv('SEND_MAIL_HOST')
        self.send_mail_service = os.getenv('SEND_MAIL_SERVICE')
        self.send_mail_user = os.getenv('SEND_MAIL_USER')
        self.send_mail_pass = os.getenv('SEND_MAIL_PASS')
        self.send_mail_subject = os.getenv('SEND_MAIL_SUBJECT')
        self.replace_array_names = os.getenv('REPLACE_ARRAY_NAMES', '').split(',')
        
        required_vars = {
            'TEMPLATE_PATH': self.template_path,
            'GOOGLE_SHEET_ID': self.sheet_id,
            'SEND_MAIL_HOST': self.send_mail_host,
            'SEND_MAIL_SERVICE': self.send_mail_service,
            'SEND_MAIL_USER': self.send_mail_user,
            'SEND_MAIL_PASS': self.send_mail_pass,
            'SEND_MAIL_SUBJECT': self.send_mail_subject
        }
        
        missing_vars = [var for var, value in required_vars.items() if not value]
        if missing_vars:
            raise ValueError(f"Missing required environment variables: {', '.join(missing_vars)}")
        
        logger.info("Settings initialized")

settings = Settings()

class EmailSender:
    def __init__(self, config: Settings):
        self.config = config
        self.server: Optional[smtplib.SMTP] = None
        self.last_error_time: Optional[datetime] = None
        self.cooldown_minutes = 20  # Cooldown time in minutes
        self.failed_emails = []  # List to track failed emails
        self.connect()

    def connect(self) -> bool:
        try:
            self.server = smtplib.SMTP(self.config.send_mail_host, 587)
            self.server.starttls()
            self.server.login(self.config.send_mail_user, self.config.send_mail_pass)
            logger.info("Successfully connected to email server")
            return True
        except Exception as e:
            logger.error(f"Failed to connect to email server: {str(e)}")
            self.last_error_time = datetime.now()
            return False

    def _check_cooldown(self) -> bool:
        if self.last_error_time is None:
            return True
        
        cooldown_end = self.last_error_time + timedelta(minutes=self.cooldown_minutes)
        if datetime.now() < cooldown_end:
            wait_time = (cooldown_end - datetime.now()).total_seconds()
            logger.warning(f"Rate limit reached. Waiting for {wait_time:.0f} seconds due to cooldown period")
            time.sleep(wait_time)
            return self.connect()
        return True

    def send_email(self, to_email: str, html_content: str, row_number: int) -> bool:
        if not self._check_cooldown():
            self.failed_emails.append({
                'email': to_email,
                'row_number': row_number,
                'error': 'Rate limit cooldown',
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            })
            return False

        try:
            msg = MIMEMultipart('alternative')
            msg['From'] = self.config.send_mail_user
            msg['To'] = to_email
            msg['Subject'] = self.config.send_mail_subject

            # Attach HTML content
            msg.attach(MIMEText(html_content, 'html'))

            self.server.send_message(msg)
            logger.info(f"Email sent successfully to {to_email} (Row {row_number})")
            return True
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Failed to send email to {to_email} (Row {row_number}): {error_msg}")
            self.last_error_time = datetime.now()
            self.failed_emails.append({
                'email': to_email,
                'row_number': row_number,
                'error': error_msg,
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            })
            return False

    def get_failed_emails_report(self) -> str:
        if not self.failed_emails:
            return "No failed emails to report."
        
        report = ["Failed Emails Report:", "==================="]
        for entry in self.failed_emails:
            report.append(
                f"Row {entry['row_number']}: {entry['email']}\n"
                f"Error: {entry['error']}\n"
                f"Time: {entry['timestamp']}\n"
                f"-------------------"
            )
        return "\n".join(report)

    def disconnect(self):
        if self.server:
            try:
                self.server.quit()
                logger.info("Disconnected from email server")
            except:
                pass

class GoogleSheetReader:
    def __init__(self, sheet_id: str):
        self.sheet_id = sheet_id
        self.base_url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv'

    def fetch_data(self) -> List[Dict]:
        try:
            response = requests.get(self.base_url)
            response.raise_for_status()
            
            df = pd.read_csv(pd.io.common.StringIO(response.text))
            data = df.to_dict('records')
            
            logger.info(f"Successfully fetched {len(data)} rows from Google Sheets")
            return data
        except Exception as e:
            logger.error(f"Error fetching data from Google Sheets: {str(e)}")
            return []

    @staticmethod
    def find_duplicate_emails(data: List[Dict]) -> Set[str]:
        emails = set()
        duplicates = set()
        
        for row in data:
            email = row.get('email', '').strip().lower()
            if email:
                if email in emails:
                    duplicates.add(email)
                else:
                    emails.add(email)
        
        return duplicates

class EmailTemplate:
    def __init__(self, template_path: str, replace_array_names: List[str]):
        self.template_path = template_path
        self.template = self._load_template()
        self.replace_array_names = replace_array_names
        
    def _load_template(self) -> str:
        try:
            with open(self.template_path, 'r', encoding='utf-8') as file:
                return file.read()
        except Exception as e:
            logger.error(f"Error loading email template: {str(e)}")
            raise
            
    def render(self, row: Dict) -> str:
        """
        Replace placeholders in the template with values from the row.
        Uses replace_array_names to know which fields to replace.
        """
        result = self.template
        for field in self.replace_array_names:
            value = row.get(field, '').strip()
            if isinstance(value, str):
                try:
                    value = value.encode('latin1').decode('utf-8')
                except:
                    pass
            result = result.replace(f'{{{{{field}}}}}', str(value))
        return result

def main():
    logger.add("app.log", rotation="1 day", retention="7 days")
    logger.info("Starting Google Sheets reader")

    try:
        email_template = EmailTemplate(template_path=settings.template_path, replace_array_names=settings.replace_array_names)
        email_sender = EmailSender(config=settings)
        
        sheet_reader = GoogleSheetReader(sheet_id=settings.sheet_id)
        data = sheet_reader.fetch_data()
        
        if not data:
            logger.error("No data found in Google Sheets")
            return

        duplicate_emails = GoogleSheetReader.find_duplicate_emails(data=data)
        if duplicate_emails:
            logger.error(f"Found {len(duplicate_emails)} duplicate emails:")
            for email in duplicate_emails:
                logger.error(f"- {email}")
            raise ValueError("Duplicate emails found")
        else:
            logger.info("No duplicate emails found")

        logger.info("Processing email templates:")
        success_count = 0
        fail_count = 0
        
        for index, row in enumerate(data, 1):
            email = row.get('email', '').strip()
            if email:
                personalized_content = email_template.render(row=row)
                logger.info(f"{index} - {email}")
                
                if email_sender.send_email(
                    to_email=email,
                    html_content=personalized_content,
                    row_number=index
                ):
                    success_count += 1
                else:
                    fail_count += 1
                
                # Add a small delay between emails to avoid rate limiting
                time.sleep(2)

        logger.info(f"Email sending completed. Success: {success_count}, Failed: {fail_count}")
        
        if fail_count > 0:
            logger.warning("\n" + email_sender.get_failed_emails_report())

    except Exception as e:
        logger.error(f"An error occurred: {str(e)}")
    finally:
        if 'email_sender' in locals():
            email_sender.disconnect()

if __name__ == "__main__":
    main()
