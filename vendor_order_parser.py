import os
import pandas as pd
import openai
import json
import re
from datetime import datetime
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

# --- Constants ---
CONFIG_FILE = 'config.json'
ERP_COLUMNS = [
    '寄件廠商', '暫代條碼', '型號', '貨號', '預計發售月份', '上架日期',
    '結單日期', '條碼', '品名', '品牌', '國際條碼', '起始進價',
    '建議售價', '廠商', '類1', '類2', '類3', '類4', '顏色', '季別',
    '尺1', '尺寸名稱', '特價', '批價', '建檔', '備註', '規格'
]

# --- AI Processing Module (ai_processor.py) ---

def get_extraction_prompt(csv_data, shipper_list):
    """Generates the prompt for the AI to extract data from a CSV string."""
    return f"""
You are an expert data extraction AI. Your task is to analyze the following CSV data, which represents an entire Excel sheet, and extract structured information based on the rules provided.

**Rules:**

1.  **Identify Global Information:**
    *   `寄件廠商`: Scan the entire CSV. Find a cell that **exactly matches** one of the names in the provided "Valid Shipper List". This is the global shipper.
    *   `結單日期`: Scan the entire CSV. Find a cell containing keywords like '結單日', '結單日期', '訂購截止日', '最後回單日' and extract its corresponding date value.

2.  **Identify Product Rows:**
    *   First, determine the header row. The header contains titles like '品名', '貨號', '條碼', '東海成本'.
    *   Process each row below the header as a potential product.

3.  **Extract Product Data:** For each product row, extract the following based on the header and keywords:
    *   `國際條碼`: Keywords ['國際條碼', 'jan code', '條碼']
    *   `貨號`: Keywords ['sku', '貨號', '商品貨號']
    *   `品名`: Keywords ['品名', '商品名', '品項', '商品', '中文品名']
    *   `預計發售月份`: Keywords ['發售日', '預定到貨', '預計上市日', '發貨日']
    *   `備註`: Keywords ['備註', '備考', '附註', '註']
    *   `起始進價`: Keywords ['東海成本']
    *   `建議售價`: Keywords ['東海售價']

4.  **Filter Data:**
    *   Only keep product entries where both `起始進價` and `建議售價` have valid, non-empty values. Discard all others.

5.  **Output Format:**
    *   Return a single JSON object.
    *   The JSON object must have two top-level keys:
        1.  `global_info`: An object containing the `寄件廠商` and `結單日期` you found.
        2.  `products`: An array of objects, where each object is a product that passed the filtering rule.
    *   **CRITICAL**: Your entire response must be ONLY the JSON object, with no other text, explanations, or markdown formatting.

**Valid Shipper List:**
{shipper_list}

**CSV Data to Process:**
```csv
{csv_data}
```
"""

def call_ai_to_extract_data(client, csv_data, shipper_list, logger):
    """Calls the AI to extract structured JSON from CSV data."""
    if not client:
        logger("OpenAI client not configured. Please set your OPENAI_API_KEY environment variable.")
        return None
    if not csv_data:
        logger("No CSV data to process.")
        return None

    prompt = get_extraction_prompt(csv_data, shipper_list)
    logger("Calling OpenAI API for data extraction...")

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are an AI assistant that extracts structured JSON from raw CSV data according to user-provided rules."},
                {"role": "user", "content": prompt},
            ],
            temperature=0,
            response_format={"type": "json_object"},
        )
        logger("Successfully received response from AI.")
        return response.choices[0].message.content
    except Exception as e:
        logger(f"Error calling OpenAI API: {e}")
        return None

# --- Excel Formatting Module (excel_formatter.py) ---

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

def generate_erp_excel(all_products, output_path, logger):
    """Generates the final ERP Excel file from the combined list of processed products."""
    if not all_products:
        logger("No products to process for the final Excel file.")
        return

    logger("Formatting data for the final ERP Excel file...")
    
    processed_rows = []
    for p_info in all_products:
        p = p_info['product_data']
        global_info = p_info['global_info']
        
        release_month = p.get('預計發售月份', '')
        if isinstance(release_month, str) and release_month:
            match = re.search(r'(\d{4})[/\-年.]?(\d{1,2})', release_month)
            if match:
                year, month = match.groups()
                release_month = f"{year}{int(month):02d}"
        
        order_date = global_info.get('結單日期', '')
        if isinstance(order_date, str) and order_date:
            try:
                order_date = pd.to_datetime(order_date).strftime('%Y/%m/%d')
            except (ValueError, TypeError):
                logger(f"Could not parse date '{order_date}', leaving as is.")

        new_row = {
            '寄件廠商': global_info.get('寄件廠商', ''),
            '暫代條碼': '',
            '型號': p.get('國際條碼', ''),
            '貨號': p.get('貨號', ''),
            '預計發售月份': release_month,
            '上架日期': '',
            '結單日期': order_date,
            '條碼': p.get('國際條碼', ''),
            '品名': p.get('品名', ''),
            '品牌': '',
            '國際條碼': '',
            '起始進價': p.get('起始進價', ''),
            '建議售價': p.get('建議售價', ''),
            '廠商': '', '類1': '', '類2': '', '類3': '', '類4': '',
            '顏色': '', '季別': '',
            '尺1': 'F', '尺寸名稱': 'F',
            '特價': '', '批價': '', '建檔': '',
            '備註': p.get('備註', ''),
            '規格': '',
        }
        processed_rows.append(new_row)

    final_df = pd.DataFrame(processed_rows)
    final_df = final_df.reindex(columns=ERP_COLUMNS).fillna('')

    try:
        final_df.to_excel(output_path, index=False, sheet_name='ERP')
        logger(f"Success! Final report saved to:\n{output_path}")
    except Exception as e:
        logger(f"Error saving final Excel file: {e}")

# --- Main Application (main.py) ---

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("究極進化版 ERP 合併標準化工具")
        self.root.geometry("700x500")
        self.client = None

        self.api_key_label = tk.Label(root, text="OpenAI API Key:")
        self.api_key_label.pack(pady=(10,0))
        
        self.api_key_entry = tk.Entry(root, width=60, show="*")
        self.api_key_entry.pack(pady=5)

        self.select_button = tk.Button(root, text="開始處理", command=self.run_processing_thread, height=2, width=20)
        self.select_button.pack(pady=20)

        self.log_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=80, height=20)
        self.log_area.pack(pady=10, padx=10, expand=True, fill=tk.BOTH)
        
        self.log("歡迎使用 Zoe 的魔法工具！")
        self.log("請先輸入您的 OpenAI API Key，然後點擊「開始處理」。")
        
        self.load_api_key() # Load key after log area is created

    def log(self, message):
        self.log_area.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log_area.see(tk.END)
        self.root.update_idletasks()

    def load_api_key(self):
        """Loads API key from config file if it exists."""
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r') as f:
                    config = json.load(f)
                    api_key = config.get('OPENAI_API_KEY')
                    if api_key:
                        self.api_key_entry.insert(0, api_key)
                        self.log("成功從 config.json 讀取 API Key。")
        except Exception as e:
            self.log(f"讀取設定檔時發生錯誤: {e}")

    def save_api_key(self, api_key):
        """Saves API key to the config file."""
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump({'OPENAI_API_KEY': api_key}, f)
            self.log("API Key 已儲存至 config.json 供下次使用。")
        except Exception as e:
            self.log(f"儲存設定檔時發生錯誤: {e}")

    def run_processing_thread(self):
        thread = threading.Thread(target=self.process_files)
        thread.daemon = True
        thread.start()

    def process_files(self):
        self.select_button.config(state=tk.DISABLED)
        
        try:
            api_key = self.api_key_entry.get()
            if not api_key:
                messagebox.showerror("錯誤", "請輸入您的 OpenAI API Key。")
                return
            
            self.client = openai.OpenAI(api_key=api_key)
            self.log("OpenAI API Key 已設定。")
            self.save_api_key(api_key) # Save key on successful use

            input_files = filedialog.askopenfilenames(
                title="請選擇 1 到 5 個訂單檔案",
                filetypes=[("Excel files", "*.xlsx *.xls")]
            )
            if not input_files:
                self.log("操作取消：未選擇任何檔案。")
                return
            if len(input_files) > 5:
                messagebox.showerror("錯誤", "您選擇了超過 5 個檔案，請重新選擇。")
                self.log("錯誤：選擇的檔案超過 5 個。")
                return
            
            self.log(f"已選擇 {len(input_files)} 個檔案。")

            output_file = filedialog.asksaveasfilename(
                title="請指定輸出檔案位置",
                defaultextension=".xlsx",
                initialfile="究極進化版用.xlsx",
                filetypes=[("Excel file", "*.xlsx")]
            )
            if not output_file:
                self.log("操作取消：未指定輸出檔案。")
                return
            
            self.log(f"輸出檔案將儲存至: {output_file}")

            script_dir = os.path.dirname(os.path.abspath(__file__))
            shipper_file = os.path.join(script_dir, '廠商名單.xlsx')
            shipper_list = []
            if os.path.exists(shipper_file):
                shipper_df = pd.read_excel(shipper_file)
                shipper_list = shipper_df['寄件廠商'].dropna().tolist()
                self.log(f"成功讀取 {len(shipper_list)} 個寄件廠商。")
            else:
                self.log("警告: '廠商名單.xlsx' 不存在。")

            all_processed_products = []
            output_dir = os.path.dirname(output_file)

            for i, file_path in enumerate(input_files):
                self.log(f"\n--- 處理檔案 {i+1}/{len(input_files)}: {os.path.basename(file_path)} ---")
                
                csv_content = convert_excel_to_csv(file_path, self.log)
                if not csv_content: continue

                ai_json_str = call_ai_to_extract_data(self.client, csv_content, shipper_list, self.log)
                if not ai_json_str:
                    self.log(f"AI 提取失敗，跳過檔案 {os.path.basename(file_path)}。")
                    continue

                # raw_data_filename = f"{os.path.splitext(os.path.basename(file_path))[0]}_raw_data.json"
                # raw_data_path = os.path.join(output_dir, raw_data_filename)
                try:
                    parsed_json = json.loads(ai_json_str)
                    # with open(raw_data_path, 'w', encoding='utf-8') as f:
                    #     json.dump(parsed_json, f, ensure_ascii=False, indent=4)
                    # self.log(f"AI 提取的中介 JSON 已儲存至: {raw_data_path}")
                    
                    global_info = parsed_json.get('global_info', {})
                    products = parsed_json.get('products', [])
                    for product in products:
                        all_processed_products.append({"global_info": global_info, "product_data": product})
                except json.JSONDecodeError:
                    self.log(f"錯誤: AI 回傳的不是有效的 JSON。")
                    # Fallback to save the raw string if JSON parsing fails
                    raw_data_filename = f"{os.path.splitext(os.path.basename(file_path))[0]}_invalid_response.txt"
                    raw_data_path = os.path.join(output_dir, raw_data_filename)
                    with open(raw_data_path, 'w', encoding='utf-8') as f: f.write(ai_json_str)
                    self.log(f"無效的 AI 回應已儲存至: {raw_data_path}")
                    continue
            
            if all_processed_products:
                generate_erp_excel(all_processed_products, output_file, self.log)
            else:
                self.log("所有檔案處理完畢，但沒有找到任何有效的商品資料可供輸出。")
            
            messagebox.showinfo("完成", "所有檔案處理完畢！")

        except Exception as e:
            self.log(f"發生未預期的錯誤: {e}")
            messagebox.showerror("嚴重錯誤", f"發生未預期的錯誤: {e}")
        finally:
            self.select_button.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()