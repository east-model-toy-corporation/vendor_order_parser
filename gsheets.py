import re
import json
import os
from typing import Optional
import gspread
from google.oauth2.service_account import Credentials

# Scope for Google Sheets and Drive
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']


def extract_sheet_id_from_url(url: str) -> Optional[str]:
    """Extracts the Google Sheet ID from various URL formats."""
    if not url:
        return None
    m = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    if m:
        return m.group(1)
    # fallback: sometimes sheets url is just the id
    if re.match(r"^[a-zA-Z0-9-_]{20,}$", url):
        return url
    return None


class GSheetsClient:
    def __init__(self, creds_json_path: str = None, creds_dict: dict = None):
        """Initialize with either path to a service account JSON or the dict contents.

        Note: user must provide a service account JSON with proper permissions to edit the target sheet.
        """
        if creds_dict:
            # create Credentials from a dict (service account info)
            self.creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        elif creds_json_path and os.path.exists(creds_json_path):
            # create Credentials from a JSON keyfile
            self.creds = Credentials.from_service_account_file(creds_json_path, scopes=SCOPES)
        else:
            raise ValueError("Service account credentials required (creds_json_path or creds_dict).")
        # gspread accepts google-auth credentials
        self.client = gspread.authorize(self.creds)

    def append_dataframe(self, sheet_id: str, df, logger):
        """Append rows from DataFrame to the first sheet of the spreadsheet specified by sheet_id.

        The target Google Sheet is expected to have the header row already set up with the 30 columns
        (ERP, GD, 平台前導, ... rest of columns). This function will append rows after the last non-empty row.

        Note: Preserving dropdowns/data-validation formatting depends on how the sheet was preconfigured. If
        the target column has a data validation rule applied to the whole column or to a range that includes
        the appended rows, the validation will apply. The code here appends values only.
        """
        try:
            sh = self.client.open_by_key(sheet_id)
            worksheet = sh.sheet1

            # find first empty row
            last_row = len(worksheet.get_all_values())
            start_row = last_row + 1

            # prepare rows: target sheet expects 30 columns, with first 3 are ERP, GD, 平台前導
            expected_cols = 30
            rows = []
            for _, r in df.iterrows():
                # convert to list in the order of df columns (should match ERP_COLUMNS)
                vals = [str(r.get(c, '')) for c in df.columns]
                # if fewer than expected, pad with empty strings
                if len(vals) < expected_cols:
                    vals = vals + [''] * (expected_cols - len(vals))
                else:
                    vals = vals[:expected_cols]
                rows.append(vals)

            # append_rows will add them after last row
            worksheet.append_rows(rows, value_input_option='USER_ENTERED')
            logger(f"Appended {len(rows)} rows to Google Sheet (ID: {sheet_id}) starting at row {start_row}.")
        except Exception as e:
            # Log the exception with full traceback to help debugging, then re-raise
            import traceback
            tb = traceback.format_exc()
            logger(f"Error appending to Google Sheet: {e}\n{tb}")
            raise
