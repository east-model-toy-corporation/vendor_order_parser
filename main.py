import os
import sys
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

        # Resolve base directory robustly so external files placed next to the exe are found
        if getattr(sys, 'frozen', False):
            # Running as PyInstaller one-file executable
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))

        shipper_file = os.path.join(base_dir, '廠商名單.xlsx')
        app.log(f"Using base directory: {base_dir} (looking for 廠商名單.xlsx, service_account.json, config.json here)")
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
                ai_date = global_info.get('結單日期')
                # 新命名規則：J=內部結單日期（內部調整用），K=結單日期（來源/外部）
                chosen = filename_order_date if filename_order_date else ai_date
                if chosen:
                    # K 欄：來源日期（檔名優先，否則 AI）
                    global_info['結單日期'] = chosen
                    # J 欄：內部結單日期（與來源相同，後續在 data_processor 中再避開週末）
                    global_info['內部結單日期'] = chosen
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

            # Check if GUI provided a Google Sheet/Drive URL
            try:
                sheet_url = app.get_sheet_url()
            except Exception:
                sheet_url = None

            sheet_id = None
            base_folder_id = None
            if sheet_url:
                # lightweight local extraction to avoid importing gsheets just for parsing
                import re
                m = re.search(r"/d/([a-zA-Z0-9-_]+)", sheet_url)
                if m:
                    sheet_id = m.group(1)
                else:
                    m2 = re.search(r"/folders/([a-zA-Z0-9-_]+)", sheet_url)
                    if m2:
                        base_folder_id = m2.group(1)
                    elif re.match(r"^[a-zA-Z0-9-_]{20,}$", sheet_url):
                        # ambiguous ID: treat as sheet id by default
                        sheet_id = sheet_url

            # If user provided either a sheet id/url or a Drive folder url/id, attempt Google upload
            if sheet_id or base_folder_id:
                try:
                    from gsheets import GSheetsClient
                    creds_path = os.path.join(base_dir, 'service_account.json')
                    if not os.path.exists(creds_path):
                        app.log(f"Google Sheets append skipped: service_account.json not found at {creds_path}. Falling back to Excel output.")
                        generate_erp_excel(all_processed_products, output_file, app.log)
                    else:
                        gs = GSheetsClient(creds_json_path=creds_path)
                        # group rows by 內部結單日期 year-month（J 欄）
                        final_df['__ym'] = final_df['內部結單日期'].apply(lambda d: '' if not d else pd.to_datetime(d, errors='coerce').strftime('%Y-%m'))
                        groups = final_df.groupby('__ym')
                        for ym, group in groups:
                            subdf = group.drop(columns=['__ym']).reset_index(drop=True)
                            if not ym:
                                # No 內部結單日期: if user provided a sheet_id, append there; otherwise fall back to Excel
                                if sheet_id:
                                    target_sheet_id = sheet_id
                                else:
                                    app.log("Rows without 內部結單日期 cannot be routed to a monthly sheet when only a Drive folder was provided. Writing these rows to Excel fallback.")
                                    group_out = os.path.splitext(output_file)[0] + f"_nogroup.xlsx"
                                    try:
                                        subdf.to_excel(group_out, index=False, sheet_name='ERP')
                                        app.log(f"Rows without 內部結單日期 saved to Excel fallback: {group_out}")
                                    except Exception as ee:
                                        app.log(f"Failed to write Excel fallback for rows without 內部結單日期: {ee}")
                                    continue
                            else:
                                year, mon = ym.split('-')
                                try:
                                    target_sheet_id = gs.ensure_month_sheet(int(year), int(mon), logger=app.log, base_folder_id=base_folder_id)
                                    if not target_sheet_id:
                                        app.log(f"Could not resolve monthly sheet for {ym}; falling back to main sheet if available.")
                                        target_sheet_id = sheet_id
                                except Exception as e:
                                    app.log(f"Error ensuring month sheet for {ym}: {e}; falling back to main sheet if available.")
                                    target_sheet_id = sheet_id

                            if target_sheet_id:
                                try:
                                    gs.append_dataframe(target_sheet_id, subdf, app.log)
                                except Exception as e:
                                    app.log(f"Append to sheet {target_sheet_id} failed: {e}. Writing this group's data to Excel fallback.")
                                    group_out = os.path.splitext(output_file)[0] + f"_{ym.replace('-', '')}.xlsx"
                                    try:
                                        subdf.to_excel(group_out, index=False, sheet_name='ERP')
                                        app.log(f"Group data for {ym} saved to Excel fallback: {group_out}")
                                    except Exception as ee:
                                        app.log(f"Failed to write Excel fallback for group {ym}: {ee}")
                            else:
                                # No target sheet resolved; write fallback
                                app.log(f"No target sheet resolved for group {ym}. Writing to Excel fallback.")
                                group_out = os.path.splitext(output_file)[0] + f"_{ym.replace('-', '')}.xlsx"
                                try:
                                    subdf.to_excel(group_out, index=False, sheet_name='ERP')
                                    app.log(f"Group data for {ym} saved to Excel fallback: {group_out}")
                                except Exception as ee:
                                    app.log(f"Failed to write Excel fallback for group {ym}: {ee}")
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