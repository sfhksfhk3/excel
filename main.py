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
        self.root.title("批次配對與編號生成工具 v5.0")
        self.root.geometry("800x700")
        self.root.resizable(True, True)

        self.source_files = []
        self.target_file = None
        self.current_serials = []
        self.next_start_var = tk.StringVar()
        self.mode_var = tk.StringVar(value="target")
        self.is_merged_var = tk.BooleanVar(value=False)
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
        self.listbox_serials = None
        self.entry_next = None
        self.btn_use_next = None

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

        ttk.Label(main_frame, text="批次配對與編號生成工具", font=("微軟正黑體", 16, "bold")).pack(pady=(0, 10))

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

        self.chk_merged = ttk.Checkbutton(step1, text="來源為匯整總表（第一欄為檔案名）", variable=self.is_merged_var)
        self.chk_merged.pack(anchor=tk.W, pady=(5, 0))

        step2 = ttk.LabelFrame(main_frame, text="步驟 2：目標模式", padding=10)
        step2.pack(fill=tk.X, pady=5)

        mode_frame = ttk.Frame(step2)
        mode_frame.pack(fill=tk.X)
        ttk.Label(mode_frame, text="模式：").pack(side=tk.LEFT)
        ttk.Radiobutton(mode_frame, text="載入目標檔案", variable=self.mode_var, value="target", command=self.on_mode_change).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(mode_frame, text="手動輸入起始編號", variable=self.mode_var, value="manual", command=self.on_mode_change).pack(side=tk.LEFT, padx=5)

        self.target_frame = ttk.Frame(step2)
        self.btn_target = ttk.Button(self.target_frame, text="選擇目標檔案", command=self.select_target_file)
        self.btn_target.pack(side=tk.LEFT, padx=5)
        self.lbl_target = ttk.Label(self.target_frame, text="尚未選擇", foreground="gray")
        self.lbl_target.pack(side=tk.LEFT, padx=5)
        self.btn_clear_target = ttk.Button(self.target_frame, text="清除", command=self.clear_target_file)
        self.btn_clear_target.pack(side=tk.LEFT, padx=5)

        self.manual_frame = ttk.Frame(step2)

        f_manual = ttk.Frame(self.manual_frame)
        f_manual.pack(fill=tk.X, pady=5)
        ttk.Label(f_manual, text="起始編號：").pack(side=tk.LEFT)
        self.entry_start = ttk.Entry(f_manual, width=20)
        self.entry_start.pack(side=tk.LEFT, padx=5)
        self.btn_generate = ttk.Button(f_manual, text="產生 50 筆", command=self.generate_serials)
        self.btn_generate.pack(side=tk.LEFT, padx=5)

        list_frame = ttk.Frame(self.manual_frame)
        list_frame.pack(fill=tk.X, pady=5)
        ttk.Label(list_frame, text="目標編號清單 (雙擊編輯)：").pack(anchor=tk.W)
        self.listbox_serials = tk.Listbox(list_frame, height=6, font=("Consolas", 10))
        self.listbox_serials.pack(fill=tk.X, pady=2)
        self.listbox_serials.bind('<Double-Button-1>', self.edit_serial)

        next_frame = ttk.Frame(self.manual_frame)
        next_frame.pack(fill=tk.X, pady=5)
        ttk.Label(next_frame, text="下一組起始：").pack(side=tk.LEFT)
        self.entry_next = ttk.Entry(next_frame, textvariable=self.next_start_var, width=20)
        self.entry_next.pack(side=tk.LEFT, padx=5)
        self.btn_use_next = ttk.Button(next_frame, text="套用", command=self.use_next_start)
        self.btn_use_next.pack(side=tk.LEFT, padx=5)

        self.on_mode_change()

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        self.btn_process = ttk.Button(btn_frame, text="▶ 開始處理", command=self.start_processing, state=tk.DISABLED)
        self.btn_process.pack(side=tk.LEFT, padx=5)
        self.progress = ttk.Progressbar(btn_frame, mode='indeterminate', length=200)
        self.progress.pack(side=tk.LEFT, padx=5)
        self.progress_label = ttk.Label(btn_frame, text="", foreground="gray")
        self.progress_label.pack(side=tk.LEFT, padx=5)

        status_frame = ttk.LabelFrame(main_frame, text="處理狀態", padding=10)
        status_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.status_text = tk.Text(status_frame, height=12, wrap=tk.WORD, font=("Consolas", 9))
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

    def on_mode_change(self):
        mode = self.mode_var.get()
        if mode == "target":
            self.target_frame.pack(fill=tk.X, pady=5)
            self.manual_frame.pack_forget()
        else:
            self.target_frame.pack_forget()
            self.manual_frame.pack(fill=tk.X, pady=5)
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

    def generate_serials(self):
        start = self.entry_start.get().strip()
        if not start:
            messagebox.showwarning("缺少起始編號", "請輸入起始編號")
            return
        try:
            self.current_serials = generate_serials(start, 50)
        except Exception as e:
            messagebox.showerror("錯誤", f"無法解析起始編號：{e}")
            return
        self.listbox_serials.delete(0, tk.END)
        for s in self.current_serials:
            self.listbox_serials.insert(tk.END, s)
        next_start = get_next_start(start, 50)
        self.next_start_var.set(next_start)
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

    def use_next_start(self):
        next_start = self.next_start_var.get().strip()
        if next_start:
            self.entry_start.delete(0, tk.END)
            self.entry_start.insert(0, next_start)
            self.generate_serials()

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
        else:
            self.btn_process.config(state=tk.DISABLED)

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
            self.log(f"來源類型：{'匯整總表' if is_merged else '個別檔案'}")
            self.log("=" * 50)

            self.temp_dir = tempfile.mkdtemp(prefix="batch_proc_")
            temp_sources = []
            for src in self.source_files:
                dst = os.path.join(self.temp_dir, os.path.basename(src))
                shutil.copy2(src, dst)
                temp_sources.append(dst)

            data_mapping = {}
            duplicate_records = []

            self.log("\n📦 讀取來源資料...")
            for fpath in temp_sources:
                self.log(f"  → {os.path.basename(fpath)}")
                try:
                    if is_merged:
                        df = pd.read_excel(fpath, header=0)
                        seal_col = find_column(df, ['Seal1', 'Seal 1', 'seal1', 'SEAL1', 'Seal No'])
                        date_col = find_column(df, ['Test Date', 'Test date', 'test date', 'TestDate', 'Date'])
                        cem_col = find_column(df, ['CEM Meter Number', 'CEM meter number', 'cem meter number', 'CEM No', 'Meter No'])
                        file_col = df.columns[0]

                        if not all([seal_col, date_col, cem_col]):
                            missing = []
                            if not seal_col: missing.append("Seal1")
                            if not date_col: missing.append("Test Date")
                            if not cem_col: missing.append("CEM Meter Number")
                            self.log(f"    ⚠️ 缺少欄位 {missing}，跳過")
                            continue

                        current_batch = None
                        count = 0
                        dup = 0
                        for _, row in df.iterrows():
                            fname = str(row[file_col]).strip() if pd.notna(row[file_col]) else ""
                            if fname and fname.lower() != "nan":
                                current_batch = parse_filename_to_batch(fname)
                            if not current_batch:
                                continue
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
                                "batch": current_batch,
                                "cem_num": cem_num
                            }
                            if key in data_mapping:
                                duplicate_records.append(record)
                                dup += 1
                            else:
                                data_mapping[key] = record
                                count += 1
                        self.log(f"    ✅ 新增 {count}，重複 {dup}")
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
                            self.log(f"    ⚠️ 缺少欄位 {missing}，跳過")
                            continue

                        count = 0
                        dup = 0
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
                            if key in data_mapping:
                                duplicate_records.append(record)
                                dup += 1
                            else:
                                data_mapping[key] = record
                                count += 1
                        self.log(f"    ✅ 新增 {count}，重複 {dup}")
                except Exception as e:
                    self.log(f"    ❌ 錯誤：{e}")
                    continue

            if not data_mapping and not duplicate_records:
                self.log("\n❌ 無有效資料")
                self.finish_processing()
                return

            self.log(f"\n📊 有效配對：{len(data_mapping)}，重複 Seal1：{len(duplicate_records)}")

            target_list = []
            if mode == "target":
                self.log(f"\n🎯 讀取目標檔案：{os.path.basename(self.target_file)}")
                wb_target = load_workbook(self.target_file)
                ws_target = wb_target.active
                for row in ws_target.iter_rows(min_row=1, max_row=ws_target.max_row, min_col=3, max_col=3, values_only=True):
                    val = str(row[0]).strip() if row[0] is not None else ""
                    if val and val.lower() != "nan":
                        target_list.append(val)
                self.log(f"  目標筆數：{len(target_list)}")
            else:
                target_list = self.current_serials
                self.log(f"\n🎯 手動清單筆數：{len(target_list)}")

            # 準備輸出
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
                # 【修改】刪除 E Column 的狀態欄位，統一生成至 A,B,C,D
                headers = ["目標編號", "批次", "日期", "CEM 編號"]
                for i, h in enumerate(headers, 1):
                    cell = ws.cell(row=1, column=i, value=h)
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

            # 【修改】目標清單為嚴格按照順序的迴圈
            for i, serial in enumerate(target_list):
                row = start_row + i
                target_key = clean_key(serial)

                if mode == "target":
                    c_col, g_col, i_col, j_col = 3, 7, 9, 10
                else:
                    # 【修改】配對成功的由 A Column 開始 (A=1, B=2, C=3, D=4)
                    c_col, g_col, i_col, j_col = 1, 2, 3, 4

                c_cell = ws.cell(row=row, column=c_col)
                if mode == "manual":
                    c_cell.value = serial
                c_cell.font = default_font
                c_cell.alignment = Alignment(horizontal='center')

                if target_key in data_mapping:
                    info = data_mapping[target_key]
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
                    # 未找到：清空並標黃
                    if mode == "target":
                        for col in [g_col, i_col, j_col]:
                            cell = ws.cell(row=row, column=col)
                            cell.value = ""
                            cell.fill = yellow_fill
                        c_cell.fill = yellow_fill
                    else:
                        # 【修改】保留空格並標黃色，不再寫入 E 欄的 "未找到" 狀態
                        for col in [2, 3, 4]:
                            cell = ws.cell(row=row, column=col)
                            cell.value = ""  # 保留空格
                            cell.fill = yellow_fill
                        c_cell.fill = yellow_fill
                    not_found_count += 1

            # 警示區 P~S 從第 1 列開始 (由 P Column 開始：16)
            alert_start = 1
            alert_headers = ["狀態", "原始 Seal1", "批次", "CEM 編號"]
            for idx, h in enumerate(alert_headers):
                cell = ws.cell(row=alert_start, column=16+idx, value=h)
                cell.font = Font(name="新細明體", size=12, bold=True)
                cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                cell.alignment = Alignment(horizontal='center')

            alert_row = alert_start + 1

            # 未配對
            unmatched = {k: v for k, v in data_mapping.items() if k not in matched_keys}
            for key, info in unmatched.items():
                ws.cell(row=alert_row, column=16, value="未配對").font = default_font
                ws.cell(row=alert_row, column=17, value=info["original_seal"]).font = default_font
                ws.cell(row=alert_row, column=18, value=info["batch"]).font = default_font
                cem_disp = info["cem_num"] if not (isinstance(info["cem_num"], float) and pd.isna(info["cem_num"])) else "無"
                ws.cell(row=alert_row, column=19, value=cem_disp).font = default_font
                alert_row += 1

            # 重複（紅色）
            for info in duplicate_records:
                ws.cell(row=alert_row, column=16, value="重複").font = red_font
                ws.cell(row=alert_row, column=17, value=info["original_seal"]).font = red_font
                ws.cell(row=alert_row, column=18, value=info["batch"]).font = red_font
                cem_disp = info["cem_num"] if not (isinstance(info["cem_num"], float) and pd.isna(info["cem_num"])) else "無"
                ws.cell(row=alert_row, column=19, value=cem_disp).font = red_font
                alert_row += 1

            # 儲存
            out_dir = os.path.dirname(self.source_files[0]) if self.source_files else os.getcwd()
            if mode == "target":
                out_name = f"處理完成_{os.path.basename(self.target_file)}"
            else:
                out_name = f"配對結果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            out_path = os.path.join(out_dir, out_name)
            wb.save(out_path)

            self.log(f"\n💾 輸出：{os.path.basename(out_path)}")
            self.log(f"✅ 配對成功：{match_count} 筆")
            self.log(f"⚠️ 未找到：{not_found_count} 筆")
            self.log(f"🔴 重複警示：{len(duplicate_records)} 筆")
            self.log(f"📋 未配對來源：{len(unmatched)} 筆")
            self.log("\n✨ 完成！")

            self.root.after(100, lambda: self.ask_open_folder(out_dir))

        except Exception as e:
            self.log(f"\n❌ 錯誤：{e}")
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
        self.check_ready()
        self.log("🧹 暫存已清除")

def main():
    root = tk.Tk()
    app = SerialProcessorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
