import os
import pandas as pd
from openai import OpenAI

# 1. 設定你的 OpenAI API 金鑰
# 建議將金鑰儲存為環境變數
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY", "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"))

# 2. 定義你的檔案路徑
input_file_path = "C:/Users/東海/Desktop/Project/11【商品批次上架三平台】/程式/Python抓廠商訂單/input_data.xlsx"
output_file_path = "C:/Users/東海/Desktop/Project/11【商品批次上架三平台】/程式/Python抓廠商訂單/ERP.csv"

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
   - 輸出必須嚴格包含以下 31 個固定欄位，順序必須完全一致：
     ERP,GD,平台前導,寄件廠商,暫代條碼,型號,預計發售月份,上架日期,結單日期,條碼,品名,品牌,國際條碼,起始進價,建議售價,廠商,類1,類2,類3,類4,顏色,季別,尺1,尺寸名稱,特價,原價,原價幣別,批價,建檔,備註,規格

3. 欄位填值規則：
   - **固定留空欄位**：
     ERP, GD, 平台前導, 暫代條碼, 國際條碼, 類2, 類3, 類4, 顏色, 季別, 特價, 批價
   - **固定值欄位**：
     - `尺1`：固定填入 "F"
     - `尺寸名稱`：固定填入 "F"
     - `上架日期`：填入你執行任務當天的日期，格式為 YYYY/MM/DD
   - **來源欄位對映與推導欄位**：
     - `寄件廠商`：在來源資料中尋找欄位名稱包含 "廠商", "代理", "供應" 等關鍵字的欄位值。
     - `型號`：在來源資料中尋找欄位名稱包含 "條碼" 或 "國際條碼" 關鍵字的欄位值。
     - `預計發售月份`：在來源資料中尋找欄位名稱包含 "發售月", "預定到貨", "預計上市", "預定月份" 等關鍵字的欄位值。
       - **格式轉換**：請將來源值轉換為 YYYYMM（六位數）。例如，如果來源是 "2025-08", "2025/8", "2025年8月"，請統一轉換成 "202508"。
     - `結單日期`：在來源資料中尋找欄位名稱包含 "結單", "截止", "截單" 等關鍵字的欄位值。
       - **格式轉換**：請將來源值轉換為 YYYY/MM/DD 格式。無論來源是 Excel 日期數字或各種文字格式（例如 "2025.08.13", "2025-8-13", "2025年8月13日"），都必須正規化。
     - `條碼`：在來源資料中尋找欄位名稱包含 "條碼" 或 "國際條碼" 關鍵字的欄位值。
     - `品名`：在來源資料中尋找欄位名稱包含 "品名", "名稱", "商品名", "品項", "商品" 等關鍵字的欄位值。
     - `品牌`：**根據「品名」欄位的值推導出品牌代號**。由於你沒有品牌代號的對照表，請先將此欄位留空。
     - `起始進價`：在來源資料中尋找欄位名稱包含 "東海成本" 關鍵字的欄位值。
     - `建議售價`：在來源資料中尋找欄位名稱包含 "東海售價" 關鍵字的欄位值。
     - `廠商`：**根據「品名」欄位的值推導出廠商代號**。請先將此欄位留空。
     - `類1`：**根據「品名」欄位的值推導出類別代號**。請先將此欄位留空。
     - `原價`：在來源資料中尋找欄位名稱包含 "原價", "定價", "售價" 等關鍵字的欄位值。
     - `原價幣別`：在來源資料中尋找欄位名稱包含 "幣別", "貨幣", "JPY", "USD", "TWD", "日幣", "美元" 等關鍵字的欄位值。若未明確，請留空。
     - `建檔`：在來源資料中尋找欄位名稱包含 "建立", "建檔", "編號" 等關鍵字的欄位值。
     - `備註`：在來源資料中尋找欄位名稱包含 "備註", "備考", "附註" 等關鍵字的欄位值。
     - `規格`：在來源資料中尋找欄位名稱包含 "規格", "說明", "描述" 等關鍵字的欄位值。

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

# 4. 主要執行流程
if __name__ == "__main__":
    json_data = read_excel_content(input_file_path)
    
    if json_data:
        csv_result = call_ai_api(json_data)

        if csv_result:
            save_to_csv(csv_result, output_file_path)