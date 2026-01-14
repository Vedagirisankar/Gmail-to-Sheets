"""
Gmail Service Module
Handles Gmail API authentication and operations
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

# Import config after path is set
import config


class GmailService:
    """Handles Gmail API operations"""
    
    def __init__(self):
        """Initialize Gmail service with OAuth authentication"""
        self.service = self._authenticate()
    
    def _authenticate(self):
        """
        Authenticate using OAuth 2.0
        Stores token.pickle for subsequent runs
        """
        creds = None
        
        # Load existing token
        if os.path.exists(config.TOKEN_FILE):
            with open(config.TOKEN_FILE, 'rb') as token:
                creds = pickle.load(token)
        
        # Refresh or create new credentials
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not os.path.exists(config.CREDENTIALS_FILE):
                    raise FileNotFoundError(
                        f"Credentials file not found: {config.CREDENTIALS_FILE}\n"
                        "Download it from Google Cloud Console"
                    )
                
                flow = InstalledAppFlow.from_client_secrets_file(
                    config.CREDENTIALS_FILE, config.GMAIL_SCOPES
                )
                creds = flow.run_local_server(port=0)
            
            # Save credentials for next run
            with open(config.TOKEN_FILE, 'wb') as token:
                pickle.dump(creds, token)
        
        return build('gmail', 'v1', credentials=creds)
    
    def get_unread_emails(self):
        """
        Fetch unread emails from inbox
        Returns list of email message objects
        """
        try:
            # Query for unread emails in inbox
            results = self.service.users().messages().list(
                userId='me',
                q='is:unread in:inbox',
                maxResults=100
            ).execute()
            
            messages = results.get('messages', [])
            
            # Fetch full email details
            emails = []
            for msg in messages:
                email = self.service.users().messages().get(
                    userId='me',
                    id=msg['id'],
                    format='full'
                ).execute()
                emails.append(email)
            
            return emails
            
        except Exception as e:
            raise Exception(f"Error fetching emails: {str(e)}")
    
    def mark_as_read(self, email_ids):
        """
        Mark emails as read
        
        Args:
            email_ids: List of email IDs to mark as read
        """
        try:
            for email_id in email_ids:
                self.service.users().messages().modify(
                    userId='me',
                    id=email_id,
                    body={'removeLabelIds': ['UNREAD']}
                ).execute()
        except Exception as e:
            raise Exception(f"Error marking emails as read: {str(e)}")