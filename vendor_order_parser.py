import os
import pandas as pd
from openai import OpenAI
import numpy as np
from datetime import datetime
import re

# 1. 設定你的 OpenAI API 金鑰
# 建議將金鑰儲存為環境變數
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY", "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"))

# 2. 定義你的檔案路徑
input_file_path = "C:/Users/東海/Desktop/Project/11【商品上架完整流程】/廠商單子/Python抓廠商訂單資料用/input_data.xlsx"
output_file_path = "C:/Users/東海/Desktop/Project/11【商品上架完整流程】/廠商單子/Python抓廠商訂單資料用/ERP.csv"

# 定義目標欄位與可能的來源關鍵字
COLUMN_MAPPING = {
    '國際條碼': ['國際條碼', 'jan code', '條碼'],
    '貨號': ['sku', '貨號', '商品貨號'],
    '品名': ['品名', '商品名', '品項', '商品', '中文品名'],
    '結單日期': ['結單日', '結單日期', '訂購截止日, 訂單截止日', '預定截止日', '最後回單日'],
    '預計發售月份': ['發售日', '預定到貨', '預計上市日', '發貨日'],
    '備註': ['備註', '備考', '附註', '註'],
    '起始進價': ['東海成本'],
    '建議售價': ['東海售價']
}

# 定義最終輸出的27個ERP欄位
ERP_COLUMNS = [
    '寄件廠商', '暫代條碼', '型號', '貨號', '預計發售月份',
    '上架日期', '結單日期', '條碼', '品名', '品牌', '國際條碼', '起始進價',
    '建議售價', '廠商', '類1', '類2', '類3', '類4', '顏色', '季別', '尺1',
    '尺寸名稱', '特價', '批價', '建檔', '備註', '規格'
]

def read_excel_content(file_path):
    """
    讀取 Excel 檔案並轉換為 JSON 字串。
    """
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Error: The file '{file_path}' was not found.")
            
        df = pd.read_excel(file_path, sheet_name=0)
        # 將 DataFrame 轉換為 JSON 格式，以便傳送給 AI 模型
        return df.to_json(orient='records', force_ascii=False)
        
    except Exception as e:
        print(f"An error occurred while reading the file: {e}")
        return None

# 3. 完整的提示詞，將包含所有處理規則
def get_processing_prompt(json_data):
    """
    根據您提供的規格，產生給 AI 模型的完整提示詞。
    """
    return f"""
你是一個精通資料處理的助手，你的任務是將一個未經整理的訂單資料轉換成一個固定的 ERP 格式。請嚴格按照以下所有規則和步驟執行：

一、輸入資料格式：
- 我將會提供你一個 JSON 格式的資料，它是由一份 Excel 檔案轉換而來。
- 檔案中可能包含多種欄位，但你只會需要根據我的規則來對映和使用它們。

二、處理與轉換規則：
1. 資料篩選：
   - 在處理之前，請先篩選資料。你只需要保留那些「東海成本」和「東海售價」兩個欄位都有數值的商品資料。
   - 任何一個欄位為空或沒有值的商品都應該被捨棄。

2. 輸出格式要求：
   - 輸出必須是一個**單一的 CSV 格式字串**。
   - **不要**包含任何額外的文字、解釋、說明或程式碼區塊（例如 ```csv...```）。
   - 輸出必須嚴格包含以下 27 個固定欄位，順序必須完全一致：
     寄件廠商,暫代條碼,型號,預計發售月份,上架日期,結單日期,條碼,品名,品牌,國際條碼,起始進價,建議售價,廠商,類1,類2,類3,類4,顏色,季別,尺1,尺寸名稱,特價,批價,建檔,備註,規格

3. 欄位填值規則：
   - **欄位**：
     - `寄件廠商`：從"廠商名單.xlsx"讀取到的"寄件廠商" list 查找，若有完全符合的值則填寫"廠商名單.xlsx"中"寄件廠商"的值，若無則留空。
     - `暫代條碼`：留空。
     - `型號`：在來源資料中尋找欄位名稱包含 '國際條碼', 'jan code', '條碼' 關鍵字的欄位值。
     - '貨號'：在來源資料中尋找欄位名稱包含 'sku', '貨號', '商品貨號' 關鍵字的欄位值。
     - `預計發售月份`：在來源資料中尋找欄位名稱包含 '發售日', '預定到貨', '預計上市日', '發貨日' 等關鍵字的欄位值，請將來源值轉換為 YYYYMM（六位數）。例如，如果來源是 "2025-08", "2025/8", "2025年8月"，請統一轉換成 "202508"。
     - `上架日期`：留空。
     - `結單日期`：在來源資料中尋找欄位名稱包含 '結單日', '結單日期', '訂購截止日, 訂單截止日', '預定截止日', '最後回單日' 等關鍵字的欄位值。
     - `條碼`：請輸入`型號`欄位的值。
     - `品名`：在來源資料中尋找欄位名稱包含 '品名', '商品名', '品項', '商品', '中文品名' 等關鍵字的欄位值。
     - `品牌`：留空。
     - `國際條碼`：留空。
     - `起始進價`：在來源資料中尋找欄位名稱包含 "東海成本" 關鍵字的欄位值。
     - `建議售價`：在來源資料中尋找欄位名稱包含 "東海售價" 關鍵字的欄位值。
     - `廠商`：留空。
     - `類1`：留空。
     - `類2`：留空。
     - `類3`：留空。
     - `類4`：留空。
     - `顏色`：留空。
     - `季別`：留空。
     - `尺1`：固定填入 "F"
     - `尺寸名稱`：固定填入 "F"
     - `特價`：留空。
     - `批價`：留空。
     - `建檔`：留空。
     - `備註`：在來源資料中尋找欄位名稱包含 '備註', '備考', '附註', '註' 等關鍵字的欄位值。
     - `規格`：留空。



三、輸出要求：
- 請確保最終輸出的 CSV 檔案中，所有欄位的順序和名稱都與「輸出格式要求」中定義的一致。
- 檔案中的資料必須是經過篩選和轉換後的結果。
- 請不要在 CSV 內容的開頭或結尾加上任何額外的說明，直接以 CSV 標頭開始。

這是需要你處理的原始資料：
{json_data}
"""

def call_ai_api(json_data):
    """
    呼叫 OpenAI API，將檔案內容和完整需求傳送給模型。
    """
    if not json_data:
        print("No data to process.")
        return None

    user_prompt = get_processing_prompt(json_data)
    
    try:
        response = client.chat.completions.create(
            model="gpt-4o",  # gpt-4o 或 gpt-4-turbo-2024-04-09 都能很好地處理此類任務
            messages=[
                {"role": "system", "content": "你是一個專業的資料處理和格式轉換助手。"},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0,  # 設定為 0 確保結果的確定性
        )
        
        # 從回應中提取模型的輸出文字
        result_text = response.choices[0].message.content
        return result_text
        
    except Exception as e:
        print(f"An error occurred while calling the OpenAI API: {e}")
        return None

def save_to_csv(csv_text, file_path):
    """
    將 AI 模型輸出的 CSV 格式文字儲存成檔案。
    """
    if not csv_text:
        print("No CSV data to save.")
        return False
        
    try:
        # 使用 utf-8-sig 編碼以支援中文和 Excel 讀取
        with open(file_path, 'w', encoding='utf-8-sig') as f:
            f.write(csv_text)
        print(f"Success: The result has been saved to {file_path}")
        return True
    except Exception as e:
        print(f"An error occurred while saving the file: {e}")
        return False

def find_column_by_keywords(df_columns, keywords):
    """根據關鍵字列表，從DataFrame的欄位中找到對應的欄位名稱。"""
    df_columns_lower = {col.lower().strip(): col for col in df_columns}
    for keyword in keywords:
        for col_lower, col_original in df_columns_lower.items():
            if keyword in col_lower:
                return col_original
    return None

def map_source_columns(df):
    """將來源DataFrame的欄位對映到標準化的字典結構。"""
    mapped_cols = {}
    for target_col, keywords in COLUMN_MAPPING.items():
        found_col = find_column_by_keywords(df.columns, keywords)
        if found_col:
            mapped_cols[target_col] = found_col
    return mapped_cols

def format_release_month(date_val):
    """將不同格式的發售月份正規化為 YYYYMM。"""
    if pd.isna(date_val):
        return ''
    s = str(date_val)
    match = re.search(r'(\d{4})[/\-年.]?(\d{1,2})', s)
    if match:
        year, month = match.groups()
        return f"{year}{int(month):02d}"
    if isinstance(date_val, (int, float)) and 200000 < date_val < 220000:
         return str(int(date_val))
    return ''

def find_shipper_in_row(row, shipper_list):
    """在單一資料列中尋找符合的寄件廠商。"""
    for item in row:
        if isinstance(item, str) and item in shipper_list:
            return item
    return ''

def process_order_file(file_path):
    """
    處理單一廠商訂單Excel檔案的核心函式。
    實現步驟3（讀取）和步驟4（對映與清洗）。
    """
    # 讀取廠商名單
    try:
        # 假設 '廠商名單.xlsx' 與執行的腳本在同一個目錄
        script_dir = os.path.dirname(os.path.abspath(__file__))
        shipper_list_path = os.path.join(script_dir, '廠商名單.xlsx')
        shipper_df = pd.read_excel(shipper_list_path)
        shipper_list = set(shipper_df['寄件廠商'].dropna().unique())
    except FileNotFoundError:
        print("警告: '廠商名單.xlsx' 不存在，無法對映'寄件廠商'。")
        shipper_list = set()
    except Exception as e:
        print(f"讀取 '廠商名單.xlsx' 失敗: {e}")
        shipper_list = set()
        
    # 步驟 3: 核心資料讀取
    try:
        df = pd.read_excel(file_path, sheet_name=0)
    except Exception as e:
        print(f"讀取Excel檔案失敗 {file_path}: {e}")
        return None

    df.dropna(how='all', inplace=True)
    df.reset_index(drop=True, inplace=True)

    # 步驟 4: 欄位對映
    column_map = map_source_columns(df)
    
    # 根據規格書，篩選 '東海成本' 和 '東海售價' 有值的商品
    cost_col_name = column_map.get('起始進價')
    price_col_name = column_map.get('建議售價')
    
    if cost_col_name and price_col_name:
        df_filtered = df[
            pd.to_numeric(df[cost_col_name], errors='coerce').notna() &
            pd.to_numeric(df[price_col_name], errors='coerce').notna()
        ].copy()
        if df_filtered.empty:
            print("找不到同時具有有效'東海成本'與'東海售價'的資料列。")
            return None
    else:
        print("警告: 找不到'東海成本'或'東海售價'欄位，將處理所有資料列。")
        df_filtered = df.copy()

    # 步驟 4: 資料清洗與轉換
    erp_df = pd.DataFrame(index=df_filtered.index)

    # 1. 處理對映欄位
    for target_col, source_col in column_map.items():
        if source_col in df_filtered.columns:
            erp_df[target_col] = df_filtered[source_col]

    # 1.1. 處理寄件廠商 (新邏輯)
    if shipper_list:
        erp_df['寄件廠商'] = df_filtered.apply(lambda row: find_shipper_in_row(row.values, shipper_list), axis=1)
    elif '寄件廠商' not in column_map: # 如果沒有廠商名單，且原本也沒對映到，則留空
        erp_df['寄件廠商'] = ''

    # 2. 填入固定值
    erp_df['上架日期'] = datetime.now().strftime('%Y/%m/%d')
    erp_df['尺1'] = 'F'
    erp_df['尺寸名稱'] = 'F'

    # 3. 處理衍生值
    if '國際條碼' in erp_df.columns:
        erp_df['型號'] = erp_df['國際條碼']
        erp_df['條碼'] = erp_df['國際條碼']

    # 4. 日期格式化
    if '結單日期' in erp_df.columns:
        erp_df['結單日期'] = pd.to_datetime(erp_df['結單日期'], errors='coerce').dt.strftime('%Y/%m/%d')

    if '預計發售月份' in erp_df.columns:
        erp_df['預計發售月份'] = erp_df['預計發售月份'].apply(format_release_month)

    # 建立符合最終規格的DataFrame
    final_df = pd.DataFrame(columns=ERP_COLUMNS)
    for col in ERP_COLUMNS:
        if col in erp_df.columns:
            final_df[col] = erp_df[col]
        else:
            final_df[col] = '' 

    return final_df

# 4. 主要執行流程
if __name__ == "__main__":
    json_data = read_excel_content(input_file_path)
    
    if json_data:
        # 儲存原始的 JSON 資料以供檢查
        json_output_path = os.path.join(os.path.dirname(output_file_path), "raw_data.json")
        try:
            with open(json_output_path, 'w', encoding='utf-8') as f:
                f.write(json_data)
            print(f"原始 JSON 資料已儲存至: {json_output_path}")
        except Exception as e:
            print(f"儲存 JSON 檔案時發生錯誤: {e}")

        csv_result = call_ai_api(json_data)

        if csv_result:
            save_to_csv(csv_result, output_file_path)