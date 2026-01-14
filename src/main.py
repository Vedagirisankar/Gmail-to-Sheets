"""
Gmail to Google Sheets Automation
Main execution script
"""

import sys
import os
# Add parent directory to path to find config
current_dir = os.path.dirname(os.path.abspath(__file__))
# Add current directory to path to find config and other modules
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)
import logging
from datetime import datetime
from gmail_service import GmailService
from sheets_service import SheetsService
from email_parser import EmailParser
from config import SPREADSHEET_ID, LOG_FILE

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


def main():
    """Main execution function"""
    try:
        logger.info("=== Starting Gmail to Sheets sync ===")
        
        # Initialize services
        logger.info("Initializing Gmail service...")
        gmail_service = GmailService()
        
        logger.info("Initializing Sheets service...")
        sheets_service = SheetsService(SPREADSHEET_ID)
        
        # Fetch unread emails
        logger.info("Fetching unread emails...")
        emails = gmail_service.get_unread_emails()
        
        if not emails:
            logger.info("No unread emails found.")
            return
        
        logger.info(f"Found {len(emails)} unread email(s)")
        
        # Get existing email IDs from sheet to prevent duplicates
        logger.info("Checking for existing emails in sheet...")
        existing_ids = sheets_service.get_existing_email_ids()
        
        # Parse and filter emails
        parser = EmailParser()
        new_emails = []
        
        for email in emails:
            email_id = email['id']
            
            # Skip if already processed
            if email_id in existing_ids:
                logger.info(f"Skipping duplicate email ID: {email_id}")
                continue
            
            # Parse email details
            parsed_email = parser.parse_email(email)
            if parsed_email:
                parsed_email['email_id'] = email_id
                new_emails.append(parsed_email)
        
        if not new_emails:
            logger.info("No new emails to process (all are duplicates)")
            return
        
        # Append to Google Sheets
        logger.info(f"Appending {len(new_emails)} new email(s) to sheet...")
        sheets_service.append_emails(new_emails)
        
        # Mark emails as read
        logger.info("Marking emails as read...")
        email_ids = [email['email_id'] for email in new_emails]
        gmail_service.mark_as_read(email_ids)
        
        logger.info(f"=== Successfully processed {len(new_emails)} email(s) ===")
        
    except Exception as e:
        logger.error(f"Error in main execution: {str(e)}", exc_info=True)
        raise


if __name__ == "__main__":
    main()