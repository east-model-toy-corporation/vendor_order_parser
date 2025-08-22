import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog
import threading
import os
import json
from datetime import datetime

CONFIG_FILE = 'config.json'


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("究極進化版 ERP 合併標準化工具")
        self.root.geometry("700x500")

        # API key
        self.api_key_label = tk.Label(self.root, text="OpenAI API Key:")
        self.api_key_label.pack(pady=(10, 0))
        self.api_key_entry = tk.Entry(self.root, width=60, show="*")
        self.api_key_entry.pack(pady=5)

        # Google Drive URL input
        self.sheet_label = tk.Label(self.root, text="究極進化版資料夾的 Google Drive 網址:")
        self.sheet_label.pack(pady=(10, 0))
        self.sheet_entry = tk.Entry(self.root, width=80)
        self.sheet_entry.pack(pady=5)

        # Buttons for import files and start processing
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)

        self.import_button = tk.Button(btn_frame, text="匯入檔案", command=self.import_files, height=2, width=12)
        self.import_button.grid(row=0, column=0, padx=5)

        # removed dedicated output button per new UX; output is chosen when starting
        self.select_button = tk.Button(btn_frame, text="開始處理", command=self.run_processing_thread, height=2, width=12)
        self.select_button.grid(row=0, column=1, padx=5)

        # Selected input files listbox (show filenames only)
        self.input_files_listbox = tk.Listbox(self.root, height=6, width=100)
        self.input_files_listbox.pack(pady=(5, 10), padx=10)

        # Output file label (will be set when user chooses at start)
        self.output_label = tk.Label(self.root, text="輸出檔案: (尚未選擇)", anchor='w')
        self.output_label.pack(fill=tk.X, padx=10)

        # internal state
        self.input_files = []  # full paths
        self.output_file = None

        self.log_area = scrolledtext.ScrolledText(self.root, wrap=tk.WORD, width=80, height=20)
        self.log_area.pack(pady=10, padx=10, expand=True, fill=tk.BOTH)

        self.log("歡迎使用 Zoe 的魔法工具！")
        self.log("請先輸入您的 OpenAI API Key，使用「匯入檔案」選擇檔案，最後按「開始處理」並指定輸出檔名。")

        self.load_api_key()

    def log(self, message):
        self.log_area.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log_area.see(tk.END)
        self.root.update_idletasks()

    def load_api_key(self):
        """Loads API key and Drive URL from config file if it exists."""
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r') as f:
                    config = json.load(f)
                    api_key = config.get('OPENAI_API_KEY')
                    sheet = config.get('DRIVE_URL')
                    if api_key:
                        self.api_key_entry.insert(0, api_key)
                        self.log("成功從 config.json 讀取 API Key。")
                    if sheet:
                        self.sheet_entry.insert(0, sheet)
                        self.log("成功從 config.json 讀取 Google Drive 連結。")
        except Exception as e:
            self.log(f"讀取設定檔時發生錯誤: {e}")

    def import_files(self):
        """Opens file dialog to let user select input Excel files and displays them (filenames only)."""
        files = filedialog.askopenfilenames(
            title="請選擇 1 到 10 個訂單檔案",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not files:
            self.log("未選擇任何檔案。")
            return
        if len(files) > 10:
            messagebox.showerror("錯誤", "您選擇了超過 10 個檔案，請重新選擇。")
            return
        self.input_files = list(files)
        self.input_files_listbox.delete(0, tk.END)
        for f in self.input_files:
            self.input_files_listbox.insert(tk.END, os.path.basename(f))
        self.log(f"已匯入 {len(self.input_files)} 個檔案。")

    def select_output_file(self):
        """Opens save-as dialog to let user choose output file and returns path."""
        out = filedialog.asksaveasfilename(
            title="請指定輸出檔案位置",
            defaultextension=".xlsx",
            initialfile="究極進化版用.xlsx",
            filetypes=[("Excel file", "*.xlsx")]
        )
        return out

    def save_api_key(self, api_key):
        """Saves API key and Drive URL to the config file."""
        try:
            cfg = {'OPENAI_API_KEY': api_key}
            try:
                sheet = self.sheet_entry.get().strip()
                if sheet:
                    cfg['DRIVE_URL'] = sheet
            except Exception:
                pass
            with open(CONFIG_FILE, 'w') as f:
                json.dump(cfg, f)
            self.log("API Key 與 Google Drive 設定已儲存至 config.json 供下次使用。")
        except Exception as e:
            self.log(f"儲存設定檔時發生錯誤: {e}")

    def get_sheet_url(self):
        try:
            return self.sheet_entry.get().strip()
        except Exception:
            return ''

    def run_processing_thread(self):
        api_key = self.api_key_entry.get()
        if not api_key:
            messagebox.showerror("錯誤", "請輸入您的 OpenAI API Key。")
            return

        # validate selections
        if not getattr(self, 'input_files', None):
            messagebox.showerror("錯誤", "尚未匯入任何輸入檔案，請先按「匯入檔案」。")
            return

        # ask for output file now (per new requirement)
        out = self.select_output_file()
        if not out:
            self.log("未指定輸出檔案，已中止處理。")
            return
        self.output_file = out
        self.output_label.config(text=f"輸出檔案: {self.output_file}")
        self.log(f"輸出檔案已設定為: {self.output_file}")

        # import here to avoid circular import at module load time
        try:
            from main import process_files_main
        except Exception as e:
            messagebox.showerror("錯誤", f"無法載入處理函式: {e}")
            return

        self.select_button.config(state=tk.DISABLED)
        thread = threading.Thread(target=process_files_main, args=(self, api_key, self.input_files, self.output_file))
        thread.daemon = True
        thread.start()


