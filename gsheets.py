import re
import os
from typing import Optional
import gspread
from google.oauth2.service_account import Credentials

try:
    from googleapiclient.discovery import build
except Exception:
    build = None

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

            # Prefer a sheet named '究極進化' first (explicit user request)
            worksheet = None
            try:
                worksheet = sh.worksheet('究極進化')
                logger("Selected worksheet '究極進化' by name.")
            except Exception:
                # Try to choose the correct worksheet by header match, then name 'ERP', else first sheet
                worksheets = sh.worksheets()
                logger(f"Spreadsheet '{sh.title}' has sheets: {[ws.title for ws in worksheets]}")
                for ws in worksheets:
                    vals = ws.get_all_values()
                    if not vals:
                        continue
                    header = [c.strip() for c in vals[0]]
                    if len(header) >= 3 and header[0:3] == ['ERP', 'GD', '平台前導']:
                        worksheet = ws
                        logger(f"Auto-detected worksheet '{ws.title}' by header match.")
                        break

                if worksheet is None:
                    try:
                        worksheet = sh.worksheet('ERP')
                        logger("Selected worksheet 'ERP' by name.")
                    except Exception:
                        worksheet = sh.sheet1
                        logger(f"Falling back to the first worksheet: '{worksheet.title}'.")

            # find first empty row by locating the last non-empty '條碼' cell (preferred)
            values_before = worksheet.get_all_values()
            start_row = 1
            last_row = 0
            barcode_col_index = None
            if values_before:
                header = [c.strip() for c in values_before[0]]
                # find index of '條碼' in header
                try:
                    barcode_col_index = header.index('條碼')
                except ValueError:
                    barcode_col_index = None

            if barcode_col_index is not None:
                # scan the barcode column to find last non-empty barcode
                for idx, row in enumerate(values_before, start=1):
                    # ensure row has enough cols
                    val = ''
                    if len(row) > barcode_col_index:
                        val = row[barcode_col_index]
                    if val is not None and str(val).strip() != '':
                        last_row = idx
                start_row = last_row + 1
                logger(f"Determined start_row by '條碼' column at index {barcode_col_index} (last non-empty at {last_row}).")
            else:
                # fallback: robust scan for any non-empty cell per row
                last_row = 0
                for idx, row in enumerate(values_before, start=1):
                    if any((cell is not None and str(cell).strip() != '') for cell in row):
                        last_row = idx
                start_row = last_row + 1
                logger(f"'條碼' column not found; fallback determined start_row by any non-empty cell (last non-empty at {last_row}).")

            # prepare rows based on the SHEET HEADER order to avoid misalignment
            sheet_header = header if values_before else []
            if not sheet_header:
                # If header unavailable, fall back to DataFrame columns
                sheet_header = list(df.columns)
            expected_cols = len(sheet_header)
            rows = []
            for _, r in df.iterrows():
                # Build row by mapping df values to each header column name
                vals = [str(r.get(col_name, '')) for col_name in sheet_header]
                # ensure correct length
                if len(vals) < expected_cols:
                    vals = vals + [''] * (expected_cols - len(vals))
                elif len(vals) > expected_cols:
                    vals = vals[:expected_cols]
                rows.append(vals)

            # write explicitly to the computed range
            end_row = start_row + len(rows) - 1
            # compute end column letter dynamically
            def col_letter(n):
                s = ""
                while n > 0:
                    n, rem = divmod(n - 1, 26)
                    s = chr(65 + rem) + s
                return s
            end_col = col_letter(expected_cols)
            verify_range = f"A{start_row}:{end_col}{end_row}"
            worksheet.update(verify_range, rows, value_input_option='USER_ENTERED')

            # read back to verify
            written = worksheet.get_values(verify_range)
            if not written or all(all(cell == '' for cell in row) for row in written):
                logger(f"Warning: After write, the read-back range {verify_range} appears empty or blank. Please verify the target worksheet and permissions.")
            else:
                logger(f"Appended {len(rows)} rows to Google Sheet (ID: {sheet_id}) starting at row {start_row} (sheet '{worksheet.title}'). Sample written row: {written[0][:6]}...")
        except Exception as e:
            # Log the exception with full traceback to help debugging, then re-raise
            import traceback
            tb = traceback.format_exc()
            logger(f"Error appending to Google Sheet: {e}\n{tb}")
            raise

    def ensure_month_sheet(self, year: int, month: int, logger=None, base_folder_name: str = '究極進化版', base_folder_id: str = None) -> str:
        """Ensure a monthly sheet exists and return its spreadsheet ID.

        Logic:
        - Find folder named `base_folder_name` on Drive.
        - Try to find a subfolder named as the year (e.g., '2025') under that folder.
        - In the year folder, look for a spreadsheet named '究極進化-YYYY年MM月結單'. If found, return its id.
        - Otherwise, in the base folder look for a template file whose name contains '複製用範本-究極進化' and copy it, renaming to '究極進化-YYYY年MM月結單'. Return new id.

        Requires Drive API access (googleapiclient)."""
        if build is None:
            raise RuntimeError('googleapiclient is required for Drive operations')

        drive = build('drive', 'v3', credentials=self.creds)

        target_name = f"究極進化-{year}年{int(month):02d}月結單"

        # determine base folder id: use provided base_folder_id, otherwise find by name
        if base_folder_id:
            # assume provided id is valid
            pass
        else:
            q = f"name='{base_folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
            resp = drive.files().list(q=q, spaces='drive', fields='files(id,name)', pageSize=10).execute()
            files = resp.get('files', [])
            if not files:
                if logger:
                    logger(f"Base folder '{base_folder_name}' not found on Drive.")
                return None
            base_folder_id = files[0]['id']

        # find or create year folder under base folder
        q_year = f"name='{year}' and mimeType='application/vnd.google-apps.folder' and '{base_folder_id}' in parents and trashed=false"
        resp = drive.files().list(q=q_year, spaces='drive', fields='files(id,name)', pageSize=10).execute()
        year_files = resp.get('files', [])
        year_folder_id = year_files[0]['id'] if year_files else None

        if year_folder_id:
            if logger:
                logger(f"Found year folder '{year}' (id: {year_folder_id}).")
        else:
            # create the year folder under base_folder
            try:
                folder_body = {
                    'name': str(year),
                    'mimeType': 'application/vnd.google-apps.folder',
                    'parents': [base_folder_id]
                }
                created = drive.files().create(body=folder_body, fields='id,name').execute()
                year_folder_id = created.get('id')
                if logger:
                    logger(f"Created year folder '{year}' (id: {year_folder_id}) under base folder.")
            except Exception as e:
                if logger:
                    logger(f"Failed to create year folder '{year}': {e}")
                raise

        # if year folder exists, search for target spreadsheet inside it
        if year_folder_id:
            q_sheet = f"name='{target_name}' and mimeType='application/vnd.google-apps.spreadsheet' and '{year_folder_id}' in parents and trashed=false"
            resp = drive.files().list(q=q_sheet, spaces='drive', fields='files(id,name)', pageSize=5).execute()
            found = resp.get('files', [])
            if found:
                if logger:
                    logger(f"Found existing monthly sheet '{target_name}' in folder '{year}'.")
                return found[0]['id']

        # not found in year folder: look for template in base folder
        q_template = f"name contains '複製用範本-究極進化' and mimeType='application/vnd.google-apps.spreadsheet' and '{base_folder_id}' in parents and trashed=false"
        resp = drive.files().list(q=q_template, spaces='drive', fields='files(id,name)', pageSize=5).execute()
        templates = resp.get('files', [])
        if not templates:
            if logger:
                logger("No template spreadsheet named like '複製用範本-究極進化' found in base folder.")
            return None

        template_id = templates[0]['id']
        # copy template into year folder with new name
        copy_body = {'name': target_name, 'parents': [year_folder_id or base_folder_id]}
        try:
            new_file = drive.files().copy(fileId=template_id, body=copy_body, fields='id,name').execute()
            if logger:
                logger(f"Copied template to create monthly sheet: {new_file.get('name')} (id: {new_file.get('id')}).")
            return new_file.get('id')
        except Exception as e:
            # If it's a Drive HttpError caused by storage quota, provide a clearer message and owner info
            err_str = str(e)
            from googleapiclient.errors import HttpError
            if isinstance(e, HttpError) and 'storageQuotaExceeded' in err_str:
                # try to fetch template owners to help identify whose Drive is full
                try:
                    meta = drive.files().get(fileId=template_id, fields='id,name,owners').execute()
                    owners = meta.get('owners', [])
                    owner_emails = [o.get('emailAddress') or o.get('displayName') for o in owners]
                    if logger:
                        logger(f"Drive storage quota exceeded when copying template (template id: {template_id}). Template owners: {owner_emails}")
                        logger("Suggestion: free up storage for the owner above, use a template located in a Drive with available quota, or grant the service account access to a folder in a Drive with space.")
                except Exception:
                    if logger:
                        logger("Drive storage quota exceeded and failed to fetch template owner info.")
                # Re-raise to let caller handle fallback
                if logger:
                    logger(f"Failed to copy template for monthly sheet due to storage quota: {e}")
                raise
            else:
                if logger:
                    logger(f"Failed to copy template for monthly sheet: {e}")
                raise
