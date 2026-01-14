"""
Configuration file for Gmail to Sheets automation
"""

import os

# API Scopes
GMAIL_SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify'
]

SHEETS_SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify'
]

# File paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_DIR = os.path.join(BASE_DIR, 'credentials')
CREDENTIALS_FILE = os.path.join(CREDENTIALS_DIR, 'credentials.json')
TOKEN_FILE = os.path.join(CREDENTIALS_DIR, 'token.pickle')

# Google Sheets configuration
# IMPORTANT: Replace with your actual spreadsheet ID
# The spreadsheet ID is found in the URL:
# https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit
SPREADSHEET_ID = 'https://docs.google.com/spreadsheets/d/1_KMTgnmk_8u8y1yHy5B8DcXDismuE69TnKPg7-wLyP0/edit?gid=0#gid=0'

# Logging
LOG_FILE = os.path.join(BASE_DIR, 'gmail_sheets.log')

# Create credentials directory if it doesn't exist
os.makedirs(CREDENTIALS_DIR, exist_ok=True)