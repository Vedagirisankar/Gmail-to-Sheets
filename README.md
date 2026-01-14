# Gmail to Google Sheets Automation

**Author:** [Your Full Name Here]

## 📋 Project Overview

This Python automation system connects to Gmail and Google Sheets APIs to automatically fetch unread emails and log them into a spreadsheet. The system prevents duplicates, marks processed emails as read, and maintains state across multiple runs.

## 🏗️ Architecture Diagram

```
┌─────────────────────────────────────────────────────────────┐
│                         main.py                              │
│                    (Orchestration Layer)                     │
└────────┬───────────────────────┬────────────────────────────┘
         │                       │
         │                       │
    ┌────▼─────────┐      ┌─────▼──────────┐
    │              │      │                 │
    │ gmail_       │      │  sheets_        │
    │ service.py   │      │  service.py     │
    │              │      │                 │
    │ - OAuth Auth │      │  - OAuth Auth   │
    │ - Fetch      │      │  - Read existing│
    │   unread     │      │    email IDs    │
    │ - Mark read  │      │  - Append rows  │
    │              │      │                 │
    └────┬─────────┘      └─────┬───────────┘
         │                      │
         │     ┌────────────────┘
         │     │
    ┌────▼─────▼───────┐
    │                  │
    │  email_parser.py │
    │                  │
    │  - Extract       │
    │    headers       │
    │  - Parse body    │
    │  - HTML→Text     │
    │                  │
    └──────────────────┘
         
         ┌──────────────┐
         │              │
         │  config.py   │
         │              │
         │  Settings &  │
         │  Constants   │
         │              │
         └──────────────┘

State Persistence:
┌───────────────────────┐
│  token.pickle         │ ← OAuth tokens (reused across runs)
└───────────────────────┘
┌───────────────────────┐
│  Google Sheet         │ ← Email IDs in column E (duplicate check)
│  (Column E)           │
└───────────────────────┘
```

## 🔐 OAuth Flow Explanation

This project uses **OAuth 2.0 Authorization Code Flow**:

1. **First Run:**
   - User runs the script
   - Browser opens to Google consent screen
   - User grants permissions for Gmail and Sheets access
   - OAuth token is generated and saved to `token.pickle`

2. **Subsequent Runs:**
   - Script loads existing token from `token.pickle`
   - If token is expired, it's automatically refreshed
   - No browser interaction needed

3. **Security:**
   - Tokens stored locally in `credentials/` folder
   - `.gitignore` prevents committing sensitive files
   - Token has refresh capability for long-term use

## 🚫 Duplicate Prevention Logic

The system uses a **two-layer approach** to prevent duplicates:

### Layer 1: Gmail-level filtering
- Only fetches **unread emails** from inbox
- After processing, emails are marked as **read**
- Already-read emails won't appear in subsequent runs

### Layer 2: Sheet-level verification
- Before appending, fetches all email IDs from **Column E** of the sheet
- Compares incoming email IDs against existing ones
- Only processes emails with IDs not already in the sheet

**Why this approach?**
- Robust: Works even if marking as read fails
- Idempotent: Safe to run script multiple times
- Traceable: Email IDs provide audit trail

## 💾 State Persistence Method

**Method:** Combination of OAuth token caching + Google Sheet as database

### OAuth Token (`token.pickle`)
- Stores authentication credentials
- Prevents re-authentication on every run
- Automatically refreshed when expired

### Google Sheet (Column E)
- Stores unique email IDs
- Acts as persistent "processed emails" database
- Survives script restarts and crashes

**Why not a local database?**
- Cloud-based: Works across different machines
- Simple: No additional database setup needed
- Transparent: Easy to audit in the sheet itself

## 🛠️ Setup Instructions

### Prerequisites
- Python 3.7+
- Google Cloud Project with Gmail and Sheets APIs enabled
- Google account

### Step 1: Clone the Repository
```bash
git clone <your-repo-url>
cd gmail-to-sheets
```

### Step 2: Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 3: Set Up Google Cloud Credentials

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project (or select existing)
3. Enable **Gmail API** and **Google Sheets API**
4. Go to **Credentials** → **Create Credentials** → **OAuth 2.0 Client ID**
5. Choose **Desktop App** as application type
6. Download the JSON file
7. Rename it to `credentials.json`
8. Place it in `credentials/` folder

### Step 4: Create Google Sheet

1. Go to [Google Sheets](https://sheets.google.com/)
2. Create a new blank spreadsheet
3. Copy the spreadsheet ID from the URL:
   ```
   https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit
   ```
4. Open `config.py` and update:
   ```python
   SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID'
   ```

### Step 5: Run the Script

```bash
cd src
python main.py
```

On first run:
- Browser will open for OAuth consent
- Grant required permissions
- Token will be saved for future runs

### Step 6: Verify

1. Check your Google Sheet - new rows should appear
2. Check Gmail - processed emails should be marked as read
3. Check `gmail_sheets.log` for execution logs

## 🔄 Running Multiple Times

The script is designed to be run repeatedly:

```bash
# Run manually
python main.py

# Or set up a cron job (Linux/Mac)
# Run every 15 minutes:
*/15 * * * * cd /path/to/gmail-to-sheets/src && python main.py
```

Each run will:
- Only process **new unread** emails
- Skip emails already in the sheet
- Maintain clean logs

## 💡 Challenge Faced & Solution

### Challenge: HTML Email Parsing

**Problem:** Many emails contain HTML content rather than plain text. The Gmail API returns HTML body, which includes tags, styling, and formatting that shouldn't appear in the sheet.

**Initial Approach:** Used regex to strip HTML tags with `re.sub('<[^<]+?>', '', html)`, but this left behind:
- Excess whitespace
- JavaScript code snippets
- CSS styles
- Broken formatting

**Solution Implemented:**
Created a custom `HTMLToTextParser` class inheriting from Python's `HTMLParser`:
- Properly parses HTML structure
- Extracts only visible text content
- Handles nested tags correctly
- Falls back to regex if parsing fails

**Code Location:** `email_parser.py` → `_html_to_text()` method

**Result:** Clean, readable email content in Google Sheets without HTML artifacts.

## ⚠️ Limitations

1. **Rate Limits:**
   - Gmail API: 250 quota units per user per second
   - Sheets API: 60 requests per minute per project
   - If exceeded, script will fail (no retry logic in basic version)

2. **Email Size:**
   - Email body truncated to 5000 characters to avoid sheet cell limits
   - Attachments are not downloaded or processed

3. **Filtering:**
   - Processes ALL unread emails in inbox
   - No subject-based or sender-based filtering (basic version)

4. **Error Recovery:**
   - If marking as read fails, emails may be reprocessed
   - No automatic retry for API failures

5. **Authentication:**
   - Requires manual OAuth consent on first run
   - Token expires if not used for extended periods

6. **Concurrency:**
   - Not designed for multiple instances running simultaneously
   - Race conditions possible if run in parallel

## 📂 Project Structure

```
gmail-to-sheets/
│
├── src/
│   ├── gmail_service.py       # Gmail API operations
│   ├── sheets_service.py      # Sheets API operations
│   ├── email_parser.py        # Email parsing logic
│   ├── main.py                # Main orchestration
│   └── config.py              # Configuration
│
├── credentials/
│   ├── credentials.json       # OAuth client secret (DO NOT COMMIT)
│   └── token.pickle           # OAuth token (DO NOT COMMIT)
│
├── proof/                     # Screenshots and video (for submission)
│   ├── gmail_inbox.png
│   ├── google_sheet.png
│   ├── oauth_consent.png
│   └── demo_video.mp4
│
├── .gitignore                 # Excludes sensitive files
├── requirements.txt           # Python dependencies
├── README.md                  # This file
└── gmail_sheets.log           # Execution logs (generated)
```

## 🎬 Proof of Execution

See `/proof/` folder for:
- ✅ Screenshots of Gmail inbox with unread emails
- ✅ Google Sheet with processed emails
- ✅ OAuth consent screen
- ✅ Screen recording demonstrating the full workflow

## 🚀 Bonus Features (Implemented)

- ✅ **HTML → Plain text conversion:** Properly parses HTML emails
- ✅ **Logging with timestamps:** All operations logged to file
- ⬜ Subject-based filtering (not yet implemented)
- ⬜ Retry logic (not yet implemented)
- ⬜ Docker setup (not yet implemented)

## 📝 Post-Submission Modifications

Ready to implement requested changes within 24 hours, such as:
- Date-range filtering
- Additional columns
- Sender filtering
- Label extraction

## 📞 Support

For questions or issues:
- Review logs in `gmail_sheets.log`
- Check Google Cloud Console for API quota
- Verify credentials are correctly placed

## 📄 License

This project is for educational purposes as part of an internship assignment.

---

**Submission Date:** [To be filled]  
**Repository:** [Your GitHub/GitLab URL]