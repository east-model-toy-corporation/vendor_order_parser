import tkinter as tk
from tkinter import messagebox, scrolledtext
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
        
        self.load_api_key()

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
        api_key = self.api_key_entry.get()
        if not api_key:
            messagebox.showerror("錯誤", "請輸入您的 OpenAI API Key。")
            return

        # import here to avoid circular import at module load time
        try:
            from main import process_files_main
        except Exception as e:
            messagebox.showerror("錯誤", f"無法載入處理函式: {e}")
            return

        self.select_button.config(state=tk.DISABLED)
        thread = threading.Thread(target=process_files_main, args=(self, api_key))
        thread.daemon = True
        thread.start()
