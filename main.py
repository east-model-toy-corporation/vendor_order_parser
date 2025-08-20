import os
import json
import pandas as pd
import openai
from tkinter import filedialog, messagebox
import tkinter as tk

from gui import App
from ai_api import call_ai_to_extract_data
from data_processor import convert_excel_to_csv, generate_erp_excel, extract_order_date_from_filename
from data_processor import build_final_df

def process_files_main(app, api_key, input_files, output_file):
    try:
        client = openai.OpenAI(api_key=api_key)
        app.log("OpenAI API Key 已設定。")
        app.save_api_key(api_key)
        # input_files and output_file are provided by the GUI
        if not input_files:
            app.log("操作取消：未選擇任何檔案。")
            app.select_button.config(state=tk.NORMAL)
            return
        if len(input_files) > 10:
            messagebox.showerror("錯誤", "您選擇了超過 10 個檔案，請重新選擇。")
            app.log("錯誤：選擇的檔案超過 10 個。")
            app.select_button.config(state=tk.NORMAL)
            return

        app.log(f"已選擇 {len(input_files)} 個檔案。")
        if not output_file:
            app.log("操作取消：未指定輸出檔案。")
            app.select_button.config(state=tk.NORMAL)
            return
        app.log(f"輸出檔案將儲存至: {output_file}")

        script_dir = os.path.dirname(os.path.abspath(__file__))
        shipper_file = os.path.join(script_dir, '廠商名單.xlsx')
        shipper_list = []
        if os.path.exists(shipper_file):
            try:
                shipper_df = pd.read_excel(shipper_file)
                # If explicit header exists, use it. Otherwise prefer column D (index 3) then C (index 2).
                raw_list = []
                if '寄件廠商' in shipper_df.columns:
                    raw_list = shipper_df['寄件廠商'].dropna().astype(str).tolist()
                else:
                    for idx in (3, 2):
                        try:
                            col = shipper_df.iloc[:, idx].dropna().astype(str)
                            if not col.empty:
                                raw_list = col.tolist()
                                app.log(f"注意: 未找到 '寄件廠商' 標題，改從第 {idx+1} 欄讀取。")
                                break
                        except Exception:
                            continue

                # Expand comma-separated values in each cell into individual tokens
                expanded = []
                for cell in raw_list:
                    for part in [p.strip() for p in str(cell).split(',')]:
                        if part:
                            expanded.append(part)

                # Remove duplicates while preserving order
                seen = set()
                ordered = []
                for v in expanded:
                    if v not in seen:
                        seen.add(v)
                        ordered.append(v)

                shipper_list = ordered
                app.log(f"成功讀取 {len(shipper_list)} 個寄件廠商（已展開逗號分隔值）。")
            except Exception as e:
                app.log(f"讀取 '廠商名單.xlsx' 時發生錯誤: {e}")
        else:
            app.log("警告: '廠商名單.xlsx' 不存在。")

        all_processed_products = []
        output_dir = os.path.dirname(output_file)

        for i, file_path in enumerate(input_files):
            app.log(f"\n--- 處理檔案 {i+1}/{len(input_files)}: {os.path.basename(file_path)} ---")
            
            csv_content = convert_excel_to_csv(file_path, app.log)
            if not csv_content: continue
            # 優先嘗試從檔名擷取結單日期
            filename_order_date = extract_order_date_from_filename(file_path, app.log)

            # 新規則：先嘗試從檔名找寄件廠商（只要檔名包含任一 shipper token 即視為命中），
            # 若檔名命中則只把該 token 傳給 AI（優先），否則再以檔案內容搭配完整 shipper_list 搜尋。
            filename_based_shipper = None
            basename = os.path.basename(file_path)
            if shipper_list:
                lowname = basename.lower()
                for s in shipper_list:
                    try:
                        if str(s).strip() and str(s).lower() in lowname:
                            filename_based_shipper = s
                            break
                    except Exception:
                        continue

            if filename_based_shipper:
                app.log(f"在檔名找到寄件廠商: {filename_based_shipper}（將優先使用此值）")
                shipper_list_for_call = [filename_based_shipper]
            else:
                shipper_list_for_call = shipper_list

            ai_json_str = call_ai_to_extract_data(client, csv_content, shipper_list_for_call, app.log)
            if not ai_json_str:
                app.log(f"AI 提取失敗，跳過檔案 {os.path.basename(file_path)}。")
                continue

            try:
                parsed_json = json.loads(ai_json_str)
                global_info = parsed_json.get('global_info', {})
                # 如果檔名有可用的結單日期，優先覆蓋 AI 回傳的值
                if filename_order_date:
                    global_info['結單日期'] = filename_order_date
                products = parsed_json.get('products', [])
                for product in products:
                    all_processed_products.append({"global_info": global_info, "product_data": product})
            except json.JSONDecodeError:
                app.log(f"錯誤: AI 回傳的不是有效的 JSON。")
                raw_data_filename = f"{os.path.splitext(os.path.basename(file_path))[0]}_invalid_response.txt"
                raw_data_path = os.path.join(output_dir, raw_data_filename)
                with open(raw_data_path, 'w', encoding='utf-8') as f: f.write(ai_json_str)
                app.log(f"無效的 AI 回應已儲存至: {raw_data_path}")
                continue
        
        if all_processed_products:
            # Build final DataFrame first
            final_df = build_final_df(all_processed_products, app.log)

            # Check if GUI provided a Google Sheet URL
            try:
                sheet_url = app.get_sheet_url()
            except Exception:
                sheet_url = None

            sheet_id = None
            if sheet_url:
                # lightweight local extraction to avoid importing gsheets just for parsing
                import re
                m = re.search(r"/d/([a-zA-Z0-9-_]+)", sheet_url)
                if m:
                    sheet_id = m.group(1)
                elif re.match(r"^[a-zA-Z0-9-_]{20,}$", sheet_url):
                    sheet_id = sheet_url

            if sheet_id:
                # try to import Google Sheets client lazily and append
                try:
                    from gsheets import GSheetsClient
                    script_dir = os.path.dirname(os.path.abspath(__file__))
                    creds_path = os.path.join(script_dir, 'service_account.json')
                    gs = GSheetsClient(creds_json_path=creds_path)
                    gs.append_dataframe(sheet_id, final_df, app.log)
                except Exception as e:
                    # log and fall back to Excel output
                    app.log(f"Google Sheets append failed or unavailable: {e}. Falling back to Excel output.")
                    generate_erp_excel(all_processed_products, output_file, app.log)
            else:
                generate_erp_excel(all_processed_products, output_file, app.log)
        else:
            app.log("所有檔案處理完畢，但沒有找到任何有效的商品資料可供輸出。")
        
        messagebox.showinfo("完成", "所有檔案處理完畢！")

    except Exception as e:
        app.log(f"發生未預期的錯誤: {e}")
        messagebox.showerror("嚴重錯誤", f"發生未預期的錯誤: {e}")
    finally:
        app.select_button.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()