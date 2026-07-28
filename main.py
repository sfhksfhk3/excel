import pandas as pd
import os
import re
import tempfile
import shutil
import traceback
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import threading

# ==================== 核心邏輯 ====================

def format_date_to_excel_ready(date_val):
    if pd.isna(date_val) or str(date_val).strip().lower() == "nan":
        return ""
    date_str = str(date_val).strip()
    if date_str.endswith('.0'):
        date_str = date_str[:-2]
    if len(date_str) == 8 and date_str.isdigit():
        return f"{date_str[4:6]}/{date_str[6:8]}/{date_str[0:4]}"
    return date_str

def parse_filename_to_batch(filename):
    base_name = os.path.splitext(os.path.basename(filename))[0]
    machine_code = ""
    batch_num = ""

    if "ED-8382" in base_name:
        machine_code = "A"
    elif "ID11832" in base_name:
        machine_code = "B"
    else:
        match_machine = re.search(r'_([A-Z0-9\-]+)_', base_name)
        if match_machine:
            machine_code = match_machine.group(1).replace("-", "")

    match_num = re.search(r'_(\d{3,4})$', base_name)
    if match_num:
        batch_num = match_num.group(1)

    if machine_code and batch_num:
        return f"{machine_code}{batch_num}"
    return base_name

def clean_key(text):
    s = str(text).upper().replace(" ", "")
    if not s:
        return s
    try:
        num = float(s)
        if num.is_integer():
            return str(int(num))
        return str(num)
    except ValueError:
        return s

def to_number_if_possible(s):
    if not isinstance(s, str):
        return s
    s = s.strip()
    if not s or s.lower() == "nan":
        return ""
    try:
        if '.' in s or 'e' in s.lower():
            num = float(s)
            if num.is_integer():
                return int(num)
            return num
        else:
            return int(s)
    except ValueError:
        return s

def normalize_col(col_name):
    return str(col_name).lower().replace(" ", "")

def find_column(df, target_names):
    cols_normalized = {normalize_col(col): col for col in df.columns}
    for name in target_names:
        norm_name = normalize_col(name)
        if norm_name in cols_normalized:
            return cols_normalized[norm_name]
    return None

def parse_serial(serial_str):
    s = str(serial_str).strip()
    match = re.match(r'^([A-Za-z]+)(\d+)$', s)
    if match:
        prefix = match.group(1).upper()
        num_part = match.group(2)
        return prefix, num_part, len(num_part)
    if s.isdigit():
        return '', s, len(s)
    last_non_digit = -1
    for i, ch in enumerate(s):
        if not ch.isdigit():
            last_non_digit = i
    if last_non_digit == -1:
        return '', s, len(s)
    prefix = s[:last_non_digit+1]
    num_part = s[last_non_digit+1:]
    if num_part.isdigit():
        return prefix, num_part, len(num_part)
    return s, '', 0

def generate_serials(start_serial, count=50):
    prefix, num_str, length = parse_serial(start_serial)
    if length == 0:
        return [start_serial]
    start_num = int(num_str)
    serials = []
    for i in range(count):
        new_num = str(start_num + i).zfill(length)
        serials.append(f"{prefix}{new_num}")
    return serials

def get_next_start(start_serial, count=50):
    prefix, num_str, length = parse_serial(start_serial)
    if length == 0:
        return start_serial
    start_num = int(num_str)
    next_num = str(start_num + count).zfill(length)
    return f"{prefix}{next_num}"

# ==================== GUI 應用程式 ====================

class SerialProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("吉配對")
        self.root.geometry("800x750")
        self.root.resizable(True, True)

        self.source_files = []
        self.target_file = None
        self.current_serials = [] 
        # 預設為手動輸入清單
        self.mode_var = tk.StringVar(value="manual")
        self.is_merged_var = tk.BooleanVar(value=True) # 預設打勾
        self.temp_dir = None

        self.btn_source = None
        self.lbl_source = None
        self.btn_clear_source = None
        self.chk_merged = None

        self.btn_target = None
        self.lbl_target = None
        self.btn_clear_target = None

        self.entry_start = None
        self.btn_generate = None
        self.btn_clear_list = None
        self.btn_export_list = None # 新增輸出按鈕
        self.listbox_serials = None

        self.btn_process = None
        self.progress = None
        self.progress_label = None
        self.status_text = None

        self.target_frame = None
        self.manual_frame = None

        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="吉配對", font=("微軟正黑體", 16, "bold")).pack(pady=(0, 10))

        step1 = ttk.LabelFrame(main_frame, text="步驟 1：來源檔案", padding=10)
        step1.pack(fill=tk.X, pady=5)

        f1 = ttk.Frame(step1)
        f1.pack(fill=tk.X)
        self.btn_source = ttk.Button(f1, text="選擇來源檔案 (可多選)", command=self.select_source_files)
        self.btn_source.pack(side=tk.LEFT, padx=5)
        self.lbl_source = ttk.Label(f1, text="尚未選擇", foreground="gray")
        self.lbl_source.pack(side=tk.LEFT, padx=5)
        self.btn_clear_source = ttk.Button(f1, text="清除", command=self.clear_source_files)
        self.btn_clear_source.pack(side=tk.LEFT, padx=5)

        self.chk_merged = ttk.Checkbutton(step1, text="來源為匯整總表（支援多Sheet、自動抓取飄移欄位）", variable=self.is_merged_var)
        self.chk_merged.pack(anchor=tk.W, pady=(5, 0))

        step2 = ttk.LabelFrame(main_frame, text="步驟 2：目標模式", padding=10)
        step2.pack(fill=tk.BOTH, expand=True, pady=5)

        mode_frame = ttk.Frame(step2)
        mode_frame.pack(fill=tk.X)
        ttk.Label(mode_frame, text="模式：").pack(side=tk.LEFT)
        ttk.Radiobutton(mode_frame, text="載入目標檔案", variable=self.mode_var, value="target", command=self.on_mode_change).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(mode_frame, text="手動產生清單", variable=self.mode_var, value="manual", command=self.on_mode_change).pack(side=tk.LEFT, padx=5)

        self.target_frame = ttk.Frame(step2)
        self.btn_target = ttk.Button(self.target_frame, text="選擇目標檔案", command=self.select_target_file)
        self.btn_target.pack(side=tk.LEFT, padx=5)
        self.lbl_target = ttk.Label(self.target_frame, text="尚未選擇", foreground="gray")
        self.lbl_target.pack(side=tk.LEFT, padx=5)
        self.btn_clear_target = ttk.Button(self.target_frame, text="清除", command=self.clear_target_file)
        self.btn_clear_target.pack(side=tk.LEFT, padx=5)

        self.manual_frame = ttk.Frame(step2)
        
        left_ctrl = ttk.Frame(self.manual_frame)
        left_ctrl.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        
        ttk.Label(left_ctrl, text="輸入起始編號：").pack(anchor=tk.W, pady=(10, 2))
        self.entry_start = ttk.Entry(left_ctrl, width=20, font=("Consolas", 11))
        self.entry_start.pack(anchor=tk.W, pady=2)
        self.entry_start.bind("<Return>", self.on_enter_generate) 
        
        ttk.Label(left_ctrl, text="(按 Enter 鍵可連續產生 50 筆)", foreground="gray", font=("微軟正黑體", 9)).pack(anchor=tk.W, pady=(0, 10))
        
        self.btn_generate = ttk.Button(left_ctrl, text="產生 50 筆", command=self.on_enter_generate)
        self.btn_generate.pack(anchor=tk.W, fill=tk.X, pady=2)
        
        self.btn_clear_list = ttk.Button(left_ctrl, text="🗑️ 清空清單", command=self.clear_manual_list)
        self.btn_clear_list.pack(anchor=tk.W, fill=tk.X, pady=(20, 2))

        # ====== 新增：輸出清單按鈕 ======
        self.btn_export_list = ttk.Button(left_ctrl, text=" 輸出清單Excel檔案", command=self.export_manual_list)
        self.btn_export_list.pack(anchor=tk.W, fill=tk.X, pady=2)

        right_list = ttk.Frame(self.manual_frame)
        right_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        ttk.Label(right_list, text="目標編號清單 (雙擊修改 / Delete 刪除)：").pack(anchor=tk.W, pady=(0, 2))
        
        scroll_y = tk.Scrollbar(right_list)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox_serials = tk.Listbox(right_list, font=("Consolas", 11), yscrollcommand=scroll_y.set)
        self.listbox_serials.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.config(command=self.listbox_serials.yview)
        
        self.listbox_serials.bind('<Double-Button-1>', self.edit_serial)
        self.listbox_serials.bind('<Delete>', self.delete_serial)

        self.on_mode_change()

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        self.btn_process = ttk.Button(btn_frame, text="開始處理 (套用所有清單)", command=self.start_processing, state=tk.DISABLED)
        self.btn_process.pack(side=tk.LEFT, padx=5)
        self.progress = ttk.Progressbar(btn_frame, mode='indeterminate', length=200)
        self.progress.pack(side=tk.LEFT, padx=5)
        self.progress_label = ttk.Label(btn_frame, text="", foreground="gray")
        self.progress_label.pack(side=tk.LEFT, padx=5)

        status_frame = ttk.LabelFrame(main_frame, text="處理狀態", padding=10)
        status_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.status_text = tk.Text(status_frame, height=8, wrap=tk.WORD, font=("Consolas", 9))
        self.status_text.pack(fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(self.status_text, command=self.status_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=scrollbar.set)

        self.check_ready()

    def clear_source_files(self):
        self.source_files = []
        self.lbl_source.config(text="尚未選擇", foreground="gray")
        self.check_ready()

    def clear_target_file(self):
        self.target_file = None
        self.lbl_target.config(text="尚未選擇", foreground="gray")
        self.check_ready()

    def clear_manual_list(self):
        if messagebox.askyesno("確認", "確定要清空目前產生的所有清單嗎？"):
            self.current_serials = []
            self.listbox_serials.delete(0, tk.END)
            self.entry_start.delete(0, tk.END)
            self.check_ready()

    def export_manual_list(self):
        """將目前產生的清單輸出為可供『載入目標檔案』讀取的格式（寫入C欄）"""
        if not self.current_serials:
            messagebox.showwarning("提示", "目前沒有清單可以輸出！")
            return
            
        file_path = filedialog.asksaveasfilename(
            title="輸出清單檔案",
            defaultextension=".xlsx",
            filetypes=[("Excel 檔案", "*.xlsx")]
        )
        if not file_path:
            return
            
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "目標清單"
            
            # 從第一列開始將編號寫入 C 欄 (column=3)，相容目標檔案的讀取邏輯
            for i, serial in enumerate(self.current_serials, start=1):
                cell = ws.cell(row=i, column=3, value=serial)
                cell.font = Font(name="新細明體", size=16)
                
            wb.save(file_path)
            messagebox.showinfo("成功", f"清單已成功輸出至：\n{file_path}\n此檔案可直接由『載入目標檔案』模式讀取。")
        except Exception as e:
            messagebox.showerror("錯誤", f"輸出失敗：{e}")

    def on_mode_change(self):
        mode = self.mode_var.get()
        if mode == "target":
            self.target_frame.pack(fill=tk.X, pady=5)
            self.manual_frame.pack_forget()
        else:
            self.target_frame.pack_forget()
            self.manual_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.check_ready()

    def select_source_files(self):
        files = filedialog.askopenfilenames(
            title="選擇來源 Excel 檔案",
            filetypes=[("Excel 檔案", "*.xlsx *.xls"), ("所有檔案", "*.*")]
        )
        if files:
            self.source_files = list(files)
            self.lbl_source.config(text=f"已選取 {len(files)} 個", foreground="green")
            self.check_ready()

    def select_target_file(self):
        file = filedialog.askopenfilename(
            title="選擇目標 Excel 檔案",
            filetypes=[("Excel 檔案", "*.xlsx *.xls"), ("所有檔案", "*.*")]
        )
        if file:
            self.target_file = file
            self.lbl_target.config(text=os.path.basename(file), foreground="green")
            self.check_ready()

    def on_enter_generate(self, event=None):
        start = self.entry_start.get().strip()
        if not start:
            if event is None: 
                messagebox.showwarning("提示", "請輸入起始編號")
            return
        try:
            new_serials = generate_serials(start, 50)
        except Exception as e:
            messagebox.showerror("錯誤", f"無法解析起始編號：{e}")
            return
        
        self.current_serials.extend(new_serials)
        for s in new_serials:
            self.listbox_serials.insert(tk.END, s)
            
        self.listbox_serials.see(tk.END)
        
        next_start = get_next_start(start, 50)
        self.entry_start.delete(0, tk.END)
        self.entry_start.insert(0, next_start)
        
        self.check_ready()

    def edit_serial(self, event):
        selection = self.listbox_serials.curselection()
        if not selection:
            return
        idx = selection[0]
        old_val = self.listbox_serials.get(idx)
        new_val = simpledialog.askstring("編輯編號", f"修改第 {idx+1} 個：", initialvalue=old_val)
        if new_val and new_val.strip():
            self.listbox_serials.delete(idx)
            self.listbox_serials.insert(idx, new_val.strip())
            self.current_serials[idx] = new_val.strip()

    def delete_serial(self, event):
        selection = self.listbox_serials.curselection()
        if not selection:
            return
        for idx in reversed(selection):
            self.listbox_serials.delete(idx)
            del self.current_serials[idx]
        self.check_ready()

    def check_ready(self):
        if self.btn_process is None:
            return
        ready = False
        if self.source_files:
            if self.mode_var.get() == "target" and self.target_file:
                ready = True
            elif self.mode_var.get() == "manual" and self.current_serials:
                ready = True
        
        if ready:
            self.btn_process.config(state=tk.NORMAL)
            if self.mode_var.get() == "manual":
                self.btn_process.config(text=f"開始處理 (共 {len(self.current_serials)} 筆)")
            else:
                self.btn_process.config(text="開始處理")
        else:
            self.btn_process.config(state=tk.DISABLED)
            self.btn_process.config(text="開始處理")

    def log(self, message):
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()

    def start_processing(self):
        self.btn_process.config(state=tk.DISABLED)
        self.btn_source.config(state=tk.DISABLED)
        self.btn_clear_source.config(state=tk.DISABLED)
        if self.mode_var.get() == "target":
            self.btn_target.config(state=tk.DISABLED)
            self.btn_clear_target.config(state=tk.DISABLED)
        else:
            self.btn_generate.config(state=tk.DISABLED)
            self.btn_clear_list.config(state=tk.DISABLED)
            self.btn_export_list.config(state=tk.DISABLED)
            self.entry_start.config(state=tk.DISABLED)
            
        self.progress.start(10)
        self.progress_label.config(text="處理中...")
        self.status_text.delete(1.0, tk.END)

        thread = threading.Thread(target=self.process_data)
        thread.daemon = True
        thread.start()

    def process_data(self):
        try:
            self.log("=" * 50)
            self.log(f"處理開始：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            mode = self.mode_var.get()
            is_merged = self.is_merged_var.get()
            self.log(f"目標模式：{'檔案' if mode == 'target' else '手動'}")
            self.log(f"來源類型：{'匯整總表 (多Sheet動態欄位)' if is_merged else '個別檔案'}")
            self.log("=" * 50)

            self.temp_dir = tempfile.mkdtemp(prefix="batch_proc_")
            temp_sources = []
            for src in self.source_files:
                dst = os.path.join(self.temp_dir, os.path.basename(src))
                shutil.copy2(src, dst)
                temp_sources.append(dst)

            data_mapping = {}

            self.log("\n 讀取來源資料...")
            for fpath in temp_sources:
                self.log(f"  → {os.path.basename(fpath)}")
                try:
                    if is_merged:
                        sheet_dict = pd.read_excel(fpath, sheet_name=None, header=None)
                        count_new = 0
                        count_dup = 0
                        
                        for sheet_name, df in sheet_dict.items():
                            self.log(f"    掃描工作表：{sheet_name}")
                            seal_idx, date_idx, cem_idx, file_idx = -1, -1, -1, 0
                            
                            for _, row in df.iterrows():
                                is_header_row = False
                                
                                for idx, val in enumerate(row):
                                    if pd.isna(val): continue
                                    val_str = str(val).lower().replace(" ", "")
                                    
                                    if 'seal1' in val_str or 'sealno' in val_str:
                                        seal_idx = idx
                                        is_header_row = True
                                    elif 'testdate' in val_str or val_str == 'date':
                                        date_idx = idx
                                        is_header_row = True
                                    elif 'cemmeter' in val_str or 'cemno' in val_str or 'meterno' in val_str:
                                        cem_idx = idx
                                        is_header_row = True
                                    elif 'source.name' in val_str or 'filename' in val_str:
                                        file_idx = idx
                                        is_header_row = True
                                        
                                if is_header_row:
                                    continue
                                    
                                if seal_idx == -1 or date_idx == -1 or cem_idx == -1:
                                    continue
                                    
                                fname = str(row.iloc[file_idx]).strip() if pd.notna(row.iloc[file_idx]) else ""
                                if not fname or fname.lower() == "nan":
                                    continue
                                    
                                current_batch = parse_filename_to_batch(fname)
                                if not current_batch:
                                    continue
                                    
                                seal_val = str(row.iloc[seal_idx]).strip() if pd.notna(row.iloc[seal_idx]) else ""
                                if not seal_val or seal_val.lower() in ["nan", "seal1", "seal 1", "seal no"]:
                                    continue
                                    
                                date_val = row.iloc[date_idx]
                                cem_raw = str(row.iloc[cem_idx]).strip() if pd.notna(row.iloc[cem_idx]) else ""

                                key = clean_key(seal_val)
                                cem_num = to_number_if_possible(cem_raw)
                                record = {
                                    "original_seal": seal_val,
                                    "date_str": format_date_to_excel_ready(date_val),
                                    "batch": current_batch,
                                    "cem_num": cem_num
                                }
                                
                                if key not in data_mapping:
                                    data_mapping[key] = []
                                    count_new += 1
                                else:
                                    count_dup += 1
                                data_mapping[key].append(record)
                                
                        self.log(f"    工作表讀取完畢：新增獨立 {count_new}，附加重複 {count_dup}")
                    
                    else:
                        batch = parse_filename_to_batch(fpath)
                        df = pd.read_excel(fpath, header=0)
                        seal_col = find_column(df, ['Seal1', 'Seal 1', 'seal1', 'SEAL1', 'Seal No'])
                        date_col = find_column(df, ['Test Date', 'Test date', 'test date', 'TestDate', 'Date'])
                        cem_col = find_column(df, ['CEM Meter Number', 'CEM meter number', 'cem meter number', 'CEM No', 'Meter No'])

                        if not all([seal_col, date_col, cem_col]):
                            missing = []
                            if not seal_col: missing.append("Seal1")
                            if not date_col: missing.append("Test Date")
                            if not cem_col: missing.append("CEM Meter Number")
                            self.log(f"    缺少欄位 {missing}，跳過")
                            continue

                        count_new = 0
                        count_dup = 0
                        for _, row in df.iterrows():
                            seal_val = str(row[seal_col]).strip() if pd.notna(row[seal_col]) else ""
                            if not seal_val or seal_val.lower() == "nan":
                                continue
                            date_val = row[date_col]
                            cem_raw = str(row[cem_col]).strip() if pd.notna(row[cem_col]) else ""

                            key = clean_key(seal_val)
                            cem_num = to_number_if_possible(cem_raw)
                            record = {
                                "original_seal": seal_val,
                                "date_str": format_date_to_excel_ready(date_val),
                                "batch": batch,
                                "cem_num": cem_num
                            }
                            
                            if key not in data_mapping:
                                data_mapping[key] = []
                                count_new += 1
                            else:
                                count_dup += 1
                            data_mapping[key].append(record)
                        
                        self.log(f"    新增獨立 {count_new}，附加重複 {count_dup}")
                except Exception as e:
                    self.log(f"     錯誤：{e}")
                    continue

            if not data_mapping:
                self.log("\n❌ 無有效資料")
                self.finish_processing()
                return

            total_duplicate_groups = sum(1 for v in data_mapping.values() if len(v) > 1)
            self.log(f"\n 有效來源配對池：{len(data_mapping)} 筆，其中重複資料：{total_duplicate_groups} 組")

            target_list = []
            if mode == "target":
                self.log(f"\n 讀取目標檔案：{os.path.basename(self.target_file)}")
                wb_target = load_workbook(self.target_file)
                ws_target = wb_target.active
                for row in ws_target.iter_rows(min_row=1, max_row=ws_target.max_row, min_col=3, max_col=3, values_only=True):
                    val = str(row[0]).strip() if row[0] is not None else ""
                    if val and val.lower() != "nan":
                        target_list.append(val)
                self.log(f"  目標筆數：{len(target_list)}")
            else:
                target_list = self.current_serials
                self.log(f"\n 手動產生清單，總計要匹配筆數：{len(target_list)}")

            if mode == "target":
                wb = wb_target
                ws = wb.active
                for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
                    for cell in row:
                        cell.font = Font(name="新細明體", size=16)
            else:
                wb = Workbook()
                ws = wb.active
                ws.title = "配對結果"
                
                # 手動模式也改為與目標檔案相同的 C, G, I, J 欄位排版
                headers_info = [(3, "目標編號"), (7, "批次"), (9, "日期"), (10, "CEM 編號")]
                for col, h in headers_info:
                    cell = ws.cell(row=1, column=col, value=h)
                    cell.font = Font(name="新細明體", size=12, bold=True)
                    cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                    cell.alignment = Alignment(horizontal='center')

            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            red_font = Font(name="新細明體", size=16, color="FF0000")
            default_font = Font(name="新細明體", size=16)

            match_count = 0
            not_found_count = 0
            matched_keys = set()
            start_row = 2 if mode == "manual" else 1

            for i, serial in enumerate(target_list):
                row = start_row + i
                target_key = clean_key(serial)

                # 不論 target 或 manual，全部統一輸出在 C, G, I, J 欄位
                c_col, g_col, i_col, j_col = 3, 7, 9, 10

                c_cell = ws.cell(row=row, column=c_col)
                if mode == "manual":
                    c_cell.value = serial
                
                c_cell.alignment = Alignment(horizontal='center')

                if target_key in data_mapping:
                    records = data_mapping[target_key]
                    info = records[0] 
                    is_duplicate = len(records) > 1 
                    
                    c_cell.font = red_font if is_duplicate else default_font

                    ws.cell(row=row, column=g_col, value=info["batch"]).font = default_font
                    
                    i_cell = ws.cell(row=row, column=i_col)
                    if info["date_str"]:
                        i_cell.value = pd.to_datetime(info["date_str"]).date()
                        i_cell.number_format = 'mm/dd/yyyy'
                    else:
                        i_cell.value = ""
                    i_cell.font = default_font

                    j_cell = ws.cell(row=row, column=j_col)
                    cem_val = info["cem_num"]
                    if cem_val == "" or (isinstance(cem_val, float) and pd.isna(cem_val)):
                        j_cell.value = ""
                    else:
                        j_cell.value = cem_val
                    j_cell.font = default_font

                    match_count += 1
                    matched_keys.add(target_key)
                else:
                    # 找不到時，清空 G, I, J 欄位並標黃，C 欄位保留編號並標黃
                    for col in [g_col, i_col, j_col]:
                        cell = ws.cell(row=row, column=col)
                        cell.value = ""
                        cell.fill = yellow_fill
                    c_cell.fill = yellow_fill
                    c_cell.font = default_font
                    not_found_count += 1

            # 警示區 P~S
            alert_start = 1
            alert_headers = ["狀態", "原始 Seal1", "批次", "CEM 編號"]
            for idx, h in enumerate(alert_headers):
                cell = ws.cell(row=alert_start, column=16+idx, value=h)
                cell.font = Font(name="新細明體", size=12, bold=True)
                cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                cell.alignment = Alignment(horizontal='center')

            alert_row = alert_start + 1

            dup_records_printed = 0
            for key, records in data_mapping.items():
                if len(records) > 1:
                    for info in records:
                        ws.cell(row=alert_row, column=16, value="重複").font = red_font
                        ws.cell(row=alert_row, column=17, value=info["original_seal"]).font = red_font
                        ws.cell(row=alert_row, column=18, value=info["batch"]).font = red_font
                        cem_disp = info["cem_num"] if not (isinstance(info["cem_num"], float) and pd.isna(info["cem_num"])) else "無"
                        ws.cell(row=alert_row, column=19, value=cem_disp).font = red_font
                        alert_row += 1
                        dup_records_printed += 1

            unmatched = {k: v for k, v in data_mapping.items() if k not in matched_keys and len(v) == 1}
            for key, records in unmatched.items():
                info = records[0]
                ws.cell(row=alert_row, column=16, value="未配對").font = default_font
                ws.cell(row=alert_row, column=17, value=info["original_seal"]).font = default_font
                ws.cell(row=alert_row, column=18, value=info["batch"]).font = default_font
                cem_disp = info["cem_num"] if not (isinstance(info["cem_num"], float) and pd.isna(info["cem_num"])) else "無"
                ws.cell(row=alert_row, column=19, value=cem_disp).font = default_font
                alert_row += 1

            out_dir = os.path.dirname(self.source_files[0]) if self.source_files else os.getcwd()
            if mode == "target":
                out_name = f"處理完成_{os.path.basename(self.target_file)}"
            else:
                out_name = f"配對結果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            out_path = os.path.join(out_dir, out_name)
            wb.save(out_path)

            self.log(f"\n 輸出：{os.path.basename(out_path)}")
            self.log(f" 成功填入：{match_count} 筆")
            self.log(f" 找不到資料被標黃：{not_found_count} 筆")
            self.log(f" 來源重複警示：共 {total_duplicate_groups} 組，印出 {dup_records_printed} 列")
            self.log(f" 來源未被使用：{len(unmatched)} 筆")
            self.log("\n 完成！")

            self.root.after(100, lambda: self.ask_open_folder(out_dir))

        except Exception as e:
            self.log(f"\n 錯誤：{e}")
            self.log(traceback.format_exc())
        finally:
            if self.temp_dir and os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir, ignore_errors=True)
            self.finish_processing()

    def ask_open_folder(self, path):
        if messagebox.askyesno("完成", "處理完成，是否開啟輸出資料夾？"):
            os.startfile(path)

    def finish_processing(self):
        self.progress.stop()
        self.progress_label.config(text="")
        
        self.btn_process.config(state=tk.NORMAL)
        self.btn_source.config(state=tk.NORMAL)
        self.btn_clear_source.config(state=tk.NORMAL)
        
        if self.mode_var.get() == "target":
            self.btn_target.config(state=tk.NORMAL)
            self.btn_clear_target.config(state=tk.NORMAL)
        else:
            self.btn_generate.config(state=tk.NORMAL)
            self.btn_clear_list.config(state=tk.NORMAL)
            self.btn_export_list.config(state=tk.NORMAL)
            self.entry_start.config(state=tk.NORMAL)
            
        self.check_ready()
        self.log("🧹 暫存已清除")

def main():
    root = tk.Tk()
    app = SerialProcessorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
