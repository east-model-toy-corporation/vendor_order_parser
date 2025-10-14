import os
import sys
import json
import pandas as pd
import openai
from tkinter import filedialog, messagebox
import tkinter as tk

from gui import App
from ai_api import call_ai_to_extract_data
from data_processor import convert_excel_to_csv, generate_erp_excel, extract_order_date_from_filename, ERP_COLUMNS
from data_processor import build_final_df

def process_files_main(app, api_key, input_files, output_file):
    # --- DEBUG FLAG ---
    # Set to True to save the prompt and AI response for each file.
    SAVE_DEBUG_FILES = False
    # --- END DEBUG FLAG ---
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

        # Load brand map from '品牌對照資料查詢.xlsx' (Cols A, B, D)
        brand_map = {}
        brand_keywords = []
        brand_ref_file = os.path.join(base_dir, '品牌對照資料查詢.xlsx')
        if os.path.exists(brand_ref_file):
            try:
                # Use headers=None and iloc to be robust against missing headers
                brand_df = pd.read_excel(brand_ref_file, sheet_name=0, header=None)
                brand_keywords = brand_df.iloc[:, 0].dropna().astype(str).tolist()
                
                for _, row in brand_df.iterrows():
                    keyword = row.iloc[0] # Column A
                    code = row.iloc[1]    # Column B
                    
                    # Column D for display name, check if it exists
                    display_name = None
                    if brand_df.shape[1] > 3:
                        display_name = row.iloc[3]

                    if pd.notna(keyword) and pd.notna(code):
                        brand_map[str(keyword).lower()] = {
                            'code': str(code),
                            'display_name': str(display_name) if pd.notna(display_name) else None
                        }
                app.log(f"成功讀取 {len(brand_map)} 個品牌關鍵字與對照資料。")
            except Exception as e:
                app.log(f"讀取 '品牌對照資料查詢.xlsx' 時發生錯誤: {e}")
        else:
            app.log("警告: '品牌對照資料查詢.xlsx' 不存在，將無法自動偵測品牌。")

        # Load category map from '類別1資料查詢.xlsx'
        category1_map = {}
        category1_keywords_sorted = []
        category1_ref_file = os.path.join(base_dir, '類別1資料查詢.xlsx')
        if os.path.exists(category1_ref_file):
            try:
                cat1_df = pd.read_excel(category1_ref_file, sheet_name=0, header=None)
                # Sort keywords by length descending to match specific terms first
                keywords = cat1_df.iloc[:, 0].dropna().astype(str).tolist()
                category1_keywords_sorted = sorted(keywords, key=len, reverse=True)

                for _, row in cat1_df.iterrows():
                    keyword = row.iloc[0]
                    cat1_val = row.iloc[1] if cat1_df.shape[1] > 1 else None
                    suffix = row.iloc[3] if cat1_df.shape[1] > 3 else None
                    command = row.iloc[5] if cat1_df.shape[1] > 5 else None # Column F

                    if pd.notna(keyword):
                        category1_map[str(keyword)] = {
                            '類1': str(cat1_val) if pd.notna(cat1_val) else '',
                            'suffix': str(suffix) if pd.notna(suffix) else '',
                            'command': str(command) if pd.notna(command) else ''
                        }
                app.log(f"成功讀取 {len(category1_map)} 個類別關鍵字與對照資料。")
            except Exception as e:
                app.log(f"讀取 '類別1資料查詢.xlsx' 時發生錯誤: {e}")
        else:
            app.log("警告: '類別1資料查詢.xlsx' 不存在，將無法自動處理類別與品名重構。")

        all_processed_products = []
        output_dir = os.path.dirname(output_file)

        for i, file_path in enumerate(input_files):
            app.log(f"\n--- 處理檔案 {i+1}/{len(input_files)}: {os.path.basename(file_path)} ---")
            
            # --- Begin new brand scanning logic ---
            file_brand_override = None
            all_cells = None
            try:
                raw_df = pd.read_excel(file_path, sheet_name=0, header=None).astype(str)
                found_brands = set()
                
                # Flatten all cell values into a single series for efficient searching
                all_cells = raw_df.unstack().dropna().astype(str).str.lower()
                
                # Case-insensitive search for each keyword
                for keyword in brand_keywords:
                    if keyword and str(keyword).strip(): # Ensure keyword is not empty
                        # Check if keyword exists in any cell
                        if all_cells.str.contains(keyword.lower(), na=False, regex=False).any():
                            found_brands.add(keyword)
                
                app.log(f"在檔案中掃描到 {len(found_brands)} 個品牌: {found_brands if found_brands else '無'}")

                if len(found_brands) == 1:
                    single_brand = found_brands.pop()
                    brand_info = brand_map.get(single_brand.lower())
                    if brand_info:
                        file_brand_override = brand_info.get('code')
                        app.log(f"啟用單一品牌覆寫模式，將使用品牌代碼: {file_brand_override}")

            except Exception as e:
                app.log(f"掃描檔案品牌時發生錯誤: {e}")
            # --- End new brand scanning logic ---

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

            debug_path_prefix = None
            if SAVE_DEBUG_FILES:
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                debug_path_prefix = os.path.normpath(os.path.join(output_dir, base_name))

            ai_json_str = call_ai_to_extract_data(
                client, csv_content, shipper_list_for_call, brand_keywords, category1_keywords_sorted, app.log,
                debug_path_prefix=debug_path_prefix
            )
            if not ai_json_str:
                app.log(f"AI 提取失敗，跳過檔案 {os.path.basename(file_path)}。")
                continue

            if SAVE_DEBUG_FILES and debug_path_prefix:
                # Save the raw AI response for user inspection
                try:
                    ai_response_filename = f"{debug_path_prefix}_ai_response.json"
                    with open(ai_response_filename, 'w', encoding='utf-8') as f:
                        # Try to pretty-print if it's valid JSON, otherwise save as is
                        try:
                            parsed = json.loads(ai_json_str)
                            json.dump(parsed, f, ensure_ascii=False, indent=4)
                        except json.JSONDecodeError:
                            f.write(ai_json_str)
                    app.log(f"AI 原始回應已儲存至: {ai_response_filename}")
                except Exception as e:
                    app.log(f"儲存 AI 原始回應時發生錯誤: {e}")

            try:
                parsed_json = json.loads(ai_json_str)
                global_info = parsed_json.get('global_info', {})

                # Overwrite shipper with filename-based one if it exists, ensuring priority.
                if filename_based_shipper:
                    app.log(f"覆寫AI結果：強制使用檔名找到的寄件廠商 '{filename_based_shipper}'。")
                    global_info['寄件廠商'] = filename_based_shipper

                ai_date = global_info.get('結單日期')
                # 新命名規則：J=內部結單日期（內部調整用），K=結單日期（來源/外部）
                chosen = filename_order_date if filename_order_date else ai_date
                if chosen:
                    # K 欄：來源日期（檔名優先，否則 AI）
                    global_info['結單日期'] = chosen
                    # J 欄：內部結單日期（與來源相同，後續在 data_processor 中再避開週末）
                    global_info['內部結單日期'] = chosen
                products = parsed_json.get('products', [])

                # --- Begin Python-side validation ---
                validated_products = []
                for p in products:
                    start_price = p.get('起始進價')
                    rec_price = p.get('建議售價')
                    # Ensure both prices exist, are not empty strings, and are not just 'nan' or similar placeholders.
                    if start_price and str(start_price).strip() and rec_price and str(rec_price).strip():
                        validated_products.append(p)
                    else:
                        app.log(f"過濾掉商品 '{p.get('品名', 'N/A')[:20]}...' 因為缺少有效的價格資訊。")
                
                products = validated_products # Replace original products with the validated list
                # --- End Python-side validation ---

                # Analyze detected brands for this file and apply overrides
                # New logic: If a file-level brand override is set, apply it to all products.
                if file_brand_override:
                    app.log(f"套用檔案級別的品牌覆寫: {file_brand_override}")
                    for p in products:
                        p['final_brand_info'] = {'name': single_brand, 'code': file_brand_override}

                # Append products to the master list
                for p in products:
                    all_processed_products.append({"global_info": global_info, "product_data": p})

            except json.JSONDecodeError:
                app.log(f"錯誤: AI 回傳的不是有效的 JSON。")
                raw_data_filename = f"{os.path.splitext(os.path.basename(file_path))[0]}_invalid_response.txt"
                raw_data_path = os.path.join(output_dir, raw_data_filename)
                with open(raw_data_path, 'w', encoding='utf-8') as f: f.write(ai_json_str)
                app.log(f"無效的 AI 回應已儲存至: {raw_data_path}")
                continue
        
        if all_processed_products:
            # Build final DataFrame first
            final_df = build_final_df(all_processed_products, brand_map, category1_map, category1_keywords_sorted, app.log)

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
                    brand_ref_path = os.path.join(base_dir, '品牌對照資料查詢.xlsx')
                    vendor_ref_path = os.path.join(base_dir, '廠商基本資料.xlsx')

                    if not os.path.exists(creds_path):
                        app.log(f"Google Sheets append skipped: service_account.json not found at {creds_path}. Falling back to Excel output.")
                        generate_erp_excel(all_processed_products, output_file, app.log)
                    else:
                        gs = GSheetsClient(creds_json_path=creds_path)
                        # group rows by 內部結單日期 year-month（J 欄）
                        final_df['__ym'] = pd.to_datetime(final_df['內部結單日期'], errors='coerce').dt.strftime('%Y-%m').fillna('')

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
                                        subdf = subdf.astype(str)
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
                                    # Ensure reference sheets exist before appending
                                    # gs.ensure_reference_sheet(target_sheet_id, '品牌對照資料查詢', brand_ref_path, ERP_COLUMNS.index('品牌'), app.log)
                                    # gs.ensure_reference_sheet(target_sheet_id, '廠商基本資料', vendor_ref_path, ERP_COLUMNS.index('廠商'), app.log)
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
                    generate_erp_excel(final_df, output_file, app.log)
            else:
                generate_erp_excel(final_df, output_file, app.log)
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