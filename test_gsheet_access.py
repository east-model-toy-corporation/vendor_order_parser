# Save as test_gsheet_access.py
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Replace with your own sheet ID
SHEET_ID = "1J4E0YrcsVkrY70IhvxZ-LnQaCc6muvfYjMC1ECqf70w"

# Your JSON file should be in the same folder
scope = ["https://www.googleapis.com/auth/spreadsheets"]
creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(SHEET_ID)

print("Access successful! Sheet title:", sheet.title)
