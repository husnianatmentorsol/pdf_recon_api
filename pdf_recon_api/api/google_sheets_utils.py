from django.conf import settings
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# 1️⃣ Sheets API service create karne ka function
def get_sheets_service():
    creds = Credentials.from_service_account_file(
        settings.GOOGLE_SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    service = build('sheets', 'v4', credentials=creds)
    return service

# 2️⃣ New tab create karne ka function
def create_new_tab_only(sheets_service):
    # Existing sheet ka ID
    spreadsheet_id = settings.MASTER_SHEET_ID
    
    # Tab ka naam (timestamp ke saath)
    sheet_title = datetime.now().strftime("%Y-%m-%d_%H%M")

    # Sheet add karne ka request body
    body = {
        "requests": [{
            "addSheet": {
                "properties": {
                    "title": sheet_title
                }
            }
        }]
    }

    # Execute request
    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=body
    ).execute()

    return sheet_title
