import os
import pandas as pd
import re
from datetime import datetime, date, timedelta

ERP_COLUMNS = [
    # ERP layout: first three are ERP / GD / 平台前導
    'ERP', 'GD', '平台前導',
    '寄件廠商', '暫代條碼', '型號', '貨號', '預計發售月份', '上架日期',
    '結單日期', '廠商結單日期', '條碼', '品名', '品牌', '國際條碼', '起始進價',
    '建議售價', '廠商', '類1', '類2', '類3', '類4', '顏色', '季別',
    '尺1', '尺寸名稱', '特價', '批價', '建檔', '備註', '規格'
]

def convert_excel_to_csv(file_path, logger):
    """Reads an Excel file and converts its first sheet to a CSV string."""
    try:
        df = pd.read_excel(file_path, sheet_name=0, header=None)
        df = df.astype(str)
        df.replace('nan', '', inplace=True)
        logger(f"Successfully converted '{os.path.basename(file_path)}' to CSV format.")
        return df.to_csv(index=False, header=False)
    except Exception as e:
        logger(f"Error reading or converting Excel file {os.path.basename(file_path)}: {e}")
        return None


def extract_order_date_from_filename(file_path, logger=None):
    """Attempt to extract an order date (MMDD) from the start of the filename.

    Rules:
    - If the filename starts with 4 digits like '0126', treat as MMDD.
    - Choose the nearest future date (strictly greater than today). If current year's MMDD is in the future, use that; otherwise use next year, etc.
    - Returns a string formatted as 'YYYY/MM/DD' or None if not found/invalid.
    """
    filename = os.path.basename(file_path)
    m = re.match(r'^(\d{3,4})', filename)
    if not m:
        return None

    mmdd = m.group(1)
    try:
        if len(mmdd) == 4:
            month = int(mmdd[:2])
            day = int(mmdd[2:])
        else:
            # len == 3, treat as MDD (e.g. '126' -> 1/26)
            month = int(mmdd[0])
            day = int(mmdd[1:])
    except ValueError:
        return None

    today = date.today()
    # try this year first, then increment year until a valid future date is found (limit to 5 years)
    for add_years in range(0, 6):
        year = today.year + add_years
        try:
            candidate = date(year, month, day)
        except ValueError:
            # invalid date for this year (e.g., Feb 29 on non-leap year)
            continue
        # must be strictly greater than today
        if candidate > today:
            return candidate.strftime('%Y/%m/%d')

    if logger:
        logger(f"無法從檔名產生有效的未來結單日期: {filename}")
    return None


def adjust_order_date(date_str, logger=None):
    """Adjust order date by: subtracting one day, then if the result falls on weekend (Sat/Sun)
    move backwards to the nearest previous non-weekend day.

    Input: date_str in formats parseable by pandas.to_datetime or 'YYYY/MM/DD'.
    Output: string 'YYYY/MM/DD' or None on failure.
    """
    if not date_str:
        return None

    try:
        dt = pd.to_datetime(date_str).date()
    except Exception:
        if logger:
            logger(f"adjust_order_date: 無法解析日期字串: {date_str}")
        return None

    def back_to_weekday(d):
        # move backwards until it's Mon-Fri
        while d.weekday() >= 5:
            d = d - timedelta(days=1)
        return d

    # If the initial date is a holiday (weekend), first move back to the nearest previous weekday,
    # then subtract one more day, and finally ensure result is a weekday.
    if dt.weekday() >= 5:
        prev_weekday = back_to_weekday(dt)
        result = prev_weekday - timedelta(days=1)
        result = back_to_weekday(result)
        return result.strftime('%Y/%m/%d')

    # Else, initial date is not holiday: subtract one day first, then if that day is holiday move back to previous weekday.
    candidate = dt - timedelta(days=1)
    if candidate.weekday() >= 5:
        candidate = back_to_weekday(candidate)
    return candidate.strftime('%Y/%m/%d')

def generate_erp_excel(all_products, output_path, logger):
    """Generates the final ERP Excel file from the combined list of processed products."""
    if not all_products:
        logger("No products to process for the final ERP output.")
        return

    logger("Formatting data for the final ERP output...")

    # build the DataFrame using helper
    final_df = build_final_df(all_products, logger)

    try:
        final_df.to_excel(output_path, index=False, sheet_name='ERP')
        logger(f"Success! Final report saved to:\n{output_path}")
    except Exception as e:
        logger(f"Error saving final Excel file: {e}")


def build_final_df(all_products, logger):
    """Builds and returns the final ERP DataFrame from processed products.

    This helper is used by both Excel output and Google Sheets append.
    """
    processed_rows = []
    for p_info in all_products:
        p = p_info['product_data']
        global_info = p_info['global_info']

        release_month = p.get('預計發售月份', '')
        # Normalize release_month whether it's a datetime-like, a string with time, or a simple string
        if release_month is None:
            release_month = ''
        else:
            # If it's not a string, try to parse with pandas (covers Timestamp/datetime)
            if not isinstance(release_month, str):
                try:
                    ts = pd.to_datetime(release_month, errors='coerce')
                    if not pd.isna(ts):
                        release_month = f"{ts.year}{int(ts.month):02d}"
                    else:
                        release_month = str(release_month)
                except Exception:
                    release_month = str(release_month)
            else:
                # it's a string: try parsing as datetime first (handles '2025-11-01 00:00:00')
                try:
                    ts = pd.to_datetime(release_month, errors='coerce')
                    if not pd.isna(ts):
                        release_month = f"{ts.year}{int(ts.month):02d}"
                    else:
                        match = re.search(r'(\d{4})[/\\-年.]?(\d{1,2})', release_month)
                        if match:
                            year, month = match.groups()
                            release_month = f"{year}{int(month):02d}"
                except Exception:
                    match = re.search(r'(\d{4})[/\\-年.]?(\d{1,2})', release_month)
                    if match:
                        year, month = match.groups()
                        release_month = f"{year}{int(month):02d}"

        # 廠商結單日期：AI 抓到的原始日期（不做週末調整；僅嘗試統一格式）
        vendor_raw_date = global_info.get('廠商結單日期', '')
        if isinstance(vendor_raw_date, str) and vendor_raw_date:
            try:
                vendor_raw_date_norm = pd.to_datetime(vendor_raw_date).strftime('%Y/%m/%d')
            except Exception:
                vendor_raw_date_norm = vendor_raw_date
        else:
            vendor_raw_date_norm = ''

        # 結單日期：可被檔名覆寫後再做週末避開調整
        order_date = global_info.get('結單日期', '')
        if isinstance(order_date, str) and order_date:
            try:
                # normalize first
                normalized = pd.to_datetime(order_date).strftime('%Y/%m/%d')
                # then apply business rule: move back one day and skip weekends
                adjusted = adjust_order_date(normalized, logger=logger)
                if adjusted:
                    order_date = adjusted
                else:
                    order_date = normalized
            except (ValueError, TypeError):
                logger(f"Could not parse date '{order_date}', leaving as is.")

        # 上架日期：填入今天日期，格式 YYYY/MM/DD
        try:
            shelf_date = datetime.now().strftime('%Y/%m/%d')
        except Exception:
            shelf_date = ''

    # actual data starts at row 3 (no need to compute excel_row here)

        # formula to extract brand token from 品名 (now at column M after inserting 廠商結單日期) and lookup brand code from
        # '品牌對照資料查詢' sheet (A:A contains 品牌代號, C:C contains 品名開頭 to match)
        # Use ROW() + INDIRECT so formula is independent of absolute row when appended.
        # Example formula (works when appended anywhere):
        # =IFERROR(INDEX('品牌對照資料查詢'!A:A, MATCH(IFERROR(TRIM(LEFT(INDIRECT("M"&ROW()),FIND("|",INDIRECT("M"&ROW()))-1)),TRIM(INDIRECT("M"&ROW()))), '品牌對照資料查詢'!C:C, 0)), "")
        brand_formula = (
            "=IFERROR(INDEX('品牌對照資料查詢'!A:A, "
            "MATCH(IFERROR(TRIM(LEFT(INDIRECT(\"M\"&ROW()),FIND(\"|\",INDIRECT(\"M\"&ROW()))-1)),TRIM(INDIRECT(\"M\"&ROW()))), '品牌對照資料查詢'!C:C, 0)), \"\")"
        )

        # formula to lookup 廠商代碼 from '廠商基本資料' sheet by matching 寄件廠商 in column D
        # Use INDIRECT("D"&ROW()) so the formula matches the same row regardless of where appended.
        # =IFERROR(INDEX('廠商基本資料'!A:A, MATCH(INDIRECT("D"&ROW()), '廠商基本資料'!D:D, 0)), "")
        vendor_formula = (
            "=IFERROR(INDEX('廠商基本資料'!A:A, MATCH(INDIRECT(\"D\"&ROW()), '廠商基本資料'!D:D, 0)), \"\")"
        )

        new_row = {
            # first three columns required by Google Sheet template
            'ERP': '待匯',
            'GD': '',
            '平台前導': '',
            '寄件廠商': global_info.get('寄件廠商', ''),
            '暫代條碼': '',
            '型號': p.get('國際條碼', ''),
            '貨號': p.get('貨號', ''),
            '預計發售月份': release_month,
            '上架日期': shelf_date,
            '結單日期': order_date,
            '廠商結單日期': vendor_raw_date_norm,
            '條碼': p.get('國際條碼', ''),
            '品名': p.get('品名', ''),
            # insert formula so spreadsheet computes 品牌 based on 品名 and 品牌對照資料查詢 sheet
            '品牌': brand_formula,
            '國際條碼': '',
            '起始進價': p.get('起始進價', ''),
            '建議售價': p.get('建議售價', ''),
            # insert formula so spreadsheet computes 廠商 (廠商代碼) based on 寄件廠商 and 廠商基本資料 sheet
            '廠商': vendor_formula, '類1': '', '類2': '', '類3': '', '類4': '',
            '顏色': '', '季別': '',
            '尺1': 'F', '尺寸名稱': 'F',
            '特價': '', '批價': '', '建檔': '',
            '備註': p.get('備註', ''),
            '規格': '',
        }
        processed_rows.append(new_row)

    final_df = pd.DataFrame(processed_rows)
    final_df = final_df.reindex(columns=ERP_COLUMNS).fillna('')
    return final_df
