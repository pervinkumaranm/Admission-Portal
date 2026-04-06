import gspread
from google.oauth2.service_account import Credentials
import os

CREDENTIALS_FILE = os.path.join(os.getcwd(), 'credentials.json')
SHEET_NAME = "SSEC_ADMISSION DATABASE_2026-27"

def test_connection():
    try:
        print(f"Testing connection with {CREDENTIALS_FILE}...")
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
        client = gspread.authorize(creds)
        sheet = client.open(SHEET_NAME).sheet1
        print("✅ SUCCESS: Connected to Google Sheets!")
        print(f"Sheet Title: {sheet.title}")
    except Exception as e:
        print(f"❌ FAILURE: {e}")

if __name__ == "__main__":
    test_connection()
