import os
import json
import pandas as pd
import openai
from tkinter import filedialog, messagebox
import tkinter as tk

from gui import App
from ai_api import call_ai_to_extract_data
from data_processor import convert_excel_to_csv, generate_erp_excel, extract_order_date_from_filename

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
            shipper_df = pd.read_excel(shipper_file)
            shipper_list = shipper_df['寄件廠商'].dropna().tolist()
            app.log(f"成功讀取 {len(shipper_list)} 個寄件廠商。")
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

            ai_json_str = call_ai_to_extract_data(client, csv_content, shipper_list, app.log)
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