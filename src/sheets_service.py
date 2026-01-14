"""
Google Sheets Service Module
Handles Sheets API operations
"""

import os
import sys
import pickle
# Add parent directory to path to find config
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from config import SHEETS_SCOPES, TOKEN_FILE, CREDENTIALS_FILE


class SheetsService:
    """Handles Google Sheets API operations"""
    
    def __init__(self, spreadsheet_id):
        """
        Initialize Sheets service
        
        Args:
            spreadsheet_id: The ID of the Google Sheet
        """
        self.spreadsheet_id = spreadsheet_id
        self.service = self._authenticate()
        self._ensure_headers()
    
    def _authenticate(self):
        """
        Authenticate using OAuth 2.0
        Reuses the same token as Gmail service
        """
        creds = None
        
        # Load existing token
        if os.path.exists(TOKEN_FILE):
            with open(TOKEN_FILE, 'rb') as token:
                creds = pickle.load(token)
        
        # Refresh or create new credentials
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not os.path.exists(CREDENTIALS_FILE):
                    raise FileNotFoundError(
                        f"Credentials file not found: {CREDENTIALS_FILE}"
                    )
                
                # Combine scopes for both Gmail and Sheets
                all_scopes = list(set(SHEETS_SCOPES))
                flow = InstalledAppFlow.from_client_secrets_file(
                    CREDENTIALS_FILE, all_scopes
                )
                creds = flow.run_local_server(port=0)
            
            # Save credentials
            with open(TOKEN_FILE, 'wb') as token:
                pickle.dump(creds, token)
        
        return build('sheets', 'v4', credentials=creds)
    
    def _ensure_headers(self):
        """Ensure the sheet has proper headers"""
        try:
            # Check if headers exist
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range='A1:E1'
            ).execute()
            
            values = result.get('values', [])
            
            # Add headers if sheet is empty
            if not values:
                headers = [['From', 'Subject', 'Date', 'Content', 'Email ID']]
                self.service.spreadsheets().values().update(
                    spreadsheetId=self.spreadsheet_id,
                    range='A1:E1',
                    valueInputOption='RAW',
                    body={'values': headers}
                ).execute()
        except Exception as e:
            raise Exception(f"Error ensuring headers: {str(e)}")
    
    def get_existing_email_ids(self):
        """
        Get all email IDs already in the sheet
        This prevents duplicate entries
        
        Returns:
            Set of email IDs
        """
        try:
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range='E:E'  # Email ID column
            ).execute()
            
            values = result.get('values', [])
            
            # Skip header and extract IDs
            email_ids = set()
            for row in values[1:]:  # Skip header row
                if row:  # Check if row is not empty
                    email_ids.add(row[0])
            
            return email_ids
            
        except Exception as e:
            raise Exception(f"Error fetching existing email IDs: {str(e)}")
    
    def append_emails(self, emails):
        """
        Append emails to the sheet
        
        Args:
            emails: List of parsed email dictionaries
        """
        try:
            # Prepare rows
            rows = []
            for email in emails:
                row = [
                    email.get('from', ''),
                    email.get('subject', ''),
                    email.get('date', ''),
                    email.get('content', ''),
                    email.get('email_id', '')
                ]
                rows.append(row)
            
            # Append to sheet
            body = {'values': rows}
            self.service.spreadsheets().values().append(
                spreadsheetId=self.spreadsheet_id,
                range='A:E',
                valueInputOption='RAW',
                insertDataOption='INSERT_ROWS',
                body=body
            ).execute()
            
        except Exception as e:
            raise Exception(f"Error appending emails to sheet: {str(e)}")