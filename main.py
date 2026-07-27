import pandas as pd
import os
import re
import sys
import traceback
import tempfile
import shutil
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import threading

# ==================== 核心邏輯 ====================

def format_date_to_excel_ready(date_val):
    """將日期值轉為 MM/DD/YYYY 字串"""
    if pd.isna(date_val) or str(date_val).strip().lower() == "nan":
        return ""
    date_str = str(date_val).strip()
    if date_str.endswith('.0'):
        date_str = date_str[:-2]
    if len(date_str) == 8 and date_str.isdigit():
        return f"{date_str[4:6]}/{date_str[6:8]}/{date_str[0:4]}"
    return date_str

def parse_filename_to_batch(filename):
    """根據檔名解析批次，例如 ED-8382 → A, ID11832 → B"""
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
    """標準化比對鍵（大寫、去空格、數字格式統一）"""
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
    """嘗試將字串轉為數字（int/float）"""
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
    """標準化欄位名稱（小寫、去空格）"""
    return str(col_name).lower().replace(" ", "")

def find_column(df, target_names):
    """在 DataFrame 中尋找欄位，支援多種可能名稱"""
    cols_normalized = {normalize_col(col): col for col in df.columns}
    for name in target_names:
        norm_name = normalize_col(name)
        if norm_name in cols_normalized:
            return cols_normalized[norm_name]
    return None

def parse_serial(serial_str):
    """
    解析編號，回傳 (prefix, number_part, number_length)
    例如 'R0026751' -> ('R', '0026751', 7)
         '104901'   -> ('', '104901', 6)
    """
    s = str(serial_str).strip()
    match = re.match(r'^([A-Za-z]+)(\d+)$', s)
    if match:
        prefix = match.group(1).upper()
        num_part = match.group(2)
        return prefix, num_part, len(num_part)
    if s.isdigit():
        return '', s, len(s)
    # 嘗試處理混合格式：找出最後的非數字字符位置
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
    # 無法解析則回傳原字串，長度 0 表示無法遞增
    return s, '', 0

def generate_serials(start_serial, count=50):
    """生成遞增序號清單"""
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
    """計算下一組的起始編號"""
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
        self.root.title("批次配對與編號生成工具 v3.0")
        self.root.geometry("800x680")
        self.root.resizable(True, True)

        self.default_font = ("微軟正黑體", 11)
        self.title_font = ("微軟正黑體", 16, "bold")

        self.source_files = []
        self.current_serials = []
        self.next_start_var = tk.StringVar()
        self.temp_dir = None

        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 標題
        ttk.Label(main_frame, text="🔢 批次配對與編號生成工具", font=self.title_font).pack(pady=(0, 15))

        # 步驟 1：來源檔案
        step1_frame = ttk.LabelFrame(main_frame, text="步驟 1：選擇來源檔案（可多選）", padding=10)
        step1_frame.pack(fill=tk.X, pady=(0, 10))

        btn_frame1 = ttk.Frame(step1_frame)
        btn_frame1.pack(fill=tk.X)
        self.btn_source = ttk.Button(btn_frame1, text="📂 選擇來源檔案", command=self.select_source_files)
        self.btn_source.pack(side=tk.LEFT, padx=(0, 10))
        self.lbl_source = ttk.Label(btn_frame1, text="尚未選擇檔案", foreground="gray")
        self.lbl_source.pack(side=tk.LEFT)

        # 步驟 2：目標編號
        step2_frame = ttk.LabelFrame(main_frame, text="步驟 2：設定目標編號（自動生成 50 筆遞增清單）", padding=10)
        step2_frame.pack(fill=tk.X, pady=(0, 10))

        row1 = ttk.Frame(step2_frame)
        row1.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(row1, text="起始編號：").pack(side=tk.LEFT)
        self.entry_start = ttk.Entry(row1, width=25)
        self.entry_start.pack(side=tk.LEFT, padx=5)
        self.btn_generate = ttk.Button(row1, text="生成編號清單", command=self.generate_serials)
        self.btn_generate.pack(side=tk.LEFT, padx=5)

        # 編號清單顯示
        list_frame = ttk.Frame(step2_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        ttk.Label(list_frame, text="當前目標編號（50 筆，雙擊可編輯單一編號）：").pack(anchor=tk.W)

        self.listbox_serials = tk.Listbox(list_frame, height=6, font=("Consolas", 10))
        self.listbox_serials.pack(fill=tk.BOTH, expand=True, pady=5)
        self.listbox_serials.bind('<Double-Button-1>', self.edit_serial)

        # 下一組起始編號
        next_frame = ttk.Frame(step2_frame)
        next_frame.pack(fill=tk.X, pady=(5, 0))
        ttk.Label(next_frame, text="下一組起始（自動遞增 50）：").pack(side=tk.LEFT)
        self.entry_next = ttk.Entry(next_frame, textvariable=self.next_start_var, width=25)
        self.entry_next.pack(side=tk.LEFT, padx=5)
        self.btn_update_next = ttk.Button(next_frame, text="套用此起始", command=self.use_next_start)
        self.btn_update_next.pack(side=tk.LEFT, padx=5)

        # 處理按鈕
        self.btn_process = ttk.Button(main_frame, text="▶️ 開始處理", command=self.start_processing, state=tk.DISABLED)
        self.btn_process.pack(pady=10)

        # 進度條
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=(5, 5))

        # 狀態顯示
        status_frame = ttk.LabelFrame(main_frame, text="處理狀態", padding=10)
        status_frame.pack(fill=tk.BOTH, expand=True)

        self.status_text = tk.Text(status_frame, height=10, wrap=tk.WORD, font=("Consolas", 10))
        self.status_text.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(self.status_text)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.status_text.yview)

        ttk.Label(main_frame, text="© 2025 | 來源檔案唯讀，自動生成目標清單，結果另存新檔", foreground="gray").pack(pady=(10, 0))

    def select_source_files(self):
        files = filedialog.askopenfilenames(
            title="選擇來源檔案",
            filetypes=[("Excel 檔案", "*.xlsx *.xls"), ("所有檔案", "*.*")]
        )
        if files:
            self.source_files = list(files)
            self.lbl_source.config(text=f"已選擇 {len(files)} 個檔案", foreground="green")
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
        # 自動計算下一組起始
        next_start = get_next_start(start, 50)
        self.next_start_var.set(next_start)
        self.check_ready()

    def edit_serial(self, event):
        selection = self.listbox_serials.curselection()
        if not selection:
            return
        idx = selection[0]
        old_val = self.listbox_serials.get(idx)
        new_val = simpledialog.askstring("編輯編號", f"修改第 {idx+1} 個編號：", initialvalue=old_val)
        if new_val and new_val.strip():
            self.listbox_serials.delete(idx)
            self.listbox_serials.insert(idx, new_val.strip())
            self.current_serials[idx] = new_val.strip()

    def use_next_start(self):
        """將下一組起始填入並生成清單"""
        next_start = self.next_start_var.get().strip()
        if next_start:
            self.entry_start.delete(0, tk.END)
            self.entry_start.insert(0, next_start)
            self.generate_serials()

    def check_ready(self):
        if self.source_files and self.current_serials:
            self.btn_process.config(state=tk.NORMAL)
        else:
            self.btn_process.config(state=tk.DISABLED)

    def log(self, message):
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()

    def start_processing(self):
        self.btn_process.config(state=tk.DISABLED)
        self.btn_source.config(state=tk.DISABLED)
        self.btn_generate.config(state=tk.DISABLED)
        self.progress.start()
        self.status_text.delete(1.0, tk.END)

        thread = threading.Thread(target=self.process_data)
        thread.daemon = True
        thread.start()

    def process_data(self):
        try:
            self.log("=" * 50)
            self.log(f"處理開始：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            self.log("=" * 50)

            # 建立暫存目錄
            self.temp_dir = tempfile.mkdtemp(prefix="batch_proc_")
            self.log(f"📁 暫存目錄：{self.temp_dir}")

            # 複製來源檔案到暫存區
            temp_source_files = []
            for src in self.source_files:
                dst = os.path.join(self.temp_dir, os.path.basename(src))
                shutil.copy2(src, dst)
                temp_source_files.append(dst)
                self.log(f"  ➜ 複製：{os.path.basename(src)}")

            # 讀取所有來源資料
            self.log("\n📦 讀取來源資料...")
            data_mapping = {}

            for filename in temp_source_files:
                self.log(f"  → {os.path.basename(filename)}")
                batch_value = parse_filename_to_batch(filename)

                try:
                    df = pd.read_excel(filename, header=0)
                    seal_col = find_column(df, ['Seal1', 'Seal 1', 'seal1', 'SEAL1', 'Seal No'])
                    date_col = find_column(df, ['Test Date', 'Test date', 'test date', 'TestDate', 'Date'])
                    cem_col = find_column(df, ['CEM Meter Number', 'CEM meter number', 'cem meter number', 'CEM No', 'Meter No'])

                    missing = []
                    if not seal_col: missing.append("Seal1")
                    if not date_col: missing.append("Test Date")
                    if not cem_col: missing.append("CEM Meter Number")
                    if missing:
                        self.log(f"    ⚠️ 缺少欄位 {missing}，跳過")
                        continue

                    count = 0
                    for _, row in df.iterrows():
                        seal_val = str(row[seal_col]).strip() if pd.notna(row[seal_col]) else ""
                        if not seal_val or seal_val.lower() == "nan":
                            continue
                        date_val = row[date_col]
                        cem_raw = str(row[cem_col]).strip() if pd.notna(row[cem_col]) else ""

                        key = clean_key(seal_val)
                        cem_num = to_number_if_possible(cem_raw)
                        data_mapping[key] = {
                            "date_str": format_date_to_excel_ready(date_val),
                            "batch": batch_value,
                            "cem_num": cem_num,
                            "original_seal": seal_val
                        }
                        count += 1
                    self.log(f"    ✅ 讀取 {count} 筆")
                except Exception as e:
                    self.log(f"    ❌ 錯誤：{e}")

            if not data_mapping:
                self.log("\n❌ 無有效來源資料，中斷")
                self.finish_processing()
                return

            self.log(f"\n📊 來源資料共 {len(data_mapping)} 筆")

            # 建立輸出 Excel
            self.log("\n🎯 建立輸出檔案...")
            wb = Workbook()
            ws = wb.active
            ws.title = "配對結果"

            # 標題列
            headers = ["目標編號 (C)", "批次 (G)", "日期 (I)", "CEM 編號 (J)", "狀態"]
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_idx, value=header)
                cell.font = Font(name="新細明體", size=12, bold=True)
                cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                cell.alignment = Alignment(horizontal='center')

            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            global_font = Font(name="新細明體", size=16)

            match_count = 0
            not_found_count = 0
            matched_keys = set()

            for i, serial in enumerate(self.current_serials):
                row = i + 2
                # C 欄（顯示為第1欄）
                c_cell = ws.cell(row=row, column=1, value=serial)
                c_cell.font = global_font
                c_cell.alignment = Alignment(horizontal='center')

                target_key = clean_key(serial)
                if target_key in data_mapping:
                    info = data_mapping[target_key]
                    # G 欄（第2欄）
                    ws.cell(row=row, column=2, value=info["batch"]).font = global_font
                    # I 欄（第3欄）日期
                    date_str = info["date_str"]
                    if date_str:
                        i_cell = ws.cell(row=row, column=3)
                        i_cell.value = pd.to_datetime(date_str).date()
                        i_cell.number_format = 'mm/dd/yyyy'
                        i_cell.font = global_font
                    else:
                        ws.cell(row=row, column=3, value="").font = global_font
                    # J 欄（第4欄）CEM
                    cem_val = info["cem_num"]
                    if cem_val == "" or (isinstance(cem_val, float) and pd.isna(cem_val)):
                        ws.cell(row=row, column=4, value="").font = global_font
                    else:
                        ws.cell(row=row, column=4, value=cem_val).font = global_font
                    # 狀態
                    ws.cell(row=row, column=5, value="✓ 已配對").font = global_font
                    match_count += 1
                    matched_keys.add(target_key)
                else:
                    # 找不到，黃底
                    for col in range(1, 5):
                        cell = ws.cell(row=row, column=col)
                        cell.fill = yellow_fill
                        cell.font = global_font
                    ws.cell(row=row, column=2, value="").font = global_font
                    ws.cell(row=row, column=3, value="").font = global_font
                    ws.cell(row=row, column=4, value="").font = global_font
                    ws.cell(row=row, column=5, value="⚠ 未找到").font = global_font
                    not_found_count += 1

            # 未配對來源記錄寫入另一工作表
            unmatched = {k: v for k, v in data_mapping.items() if k not in matched_keys}
            if unmatched:
                self.log(f"\nℹ️ 未使用來源記錄 {len(unmatched)} 筆，寫入「未配對記錄」工作表")
                ws2 = wb.create_sheet("未配對記錄")
                ws2.append(["原始 Seal1", "批次", "CEM 編號"])
                for key, info in unmatched.items():
                    cem_display = info["cem_num"]
                    if isinstance(cem_display, (int, float)):
                        cem_display = str(int(cem_display)) if isinstance(cem_display, float) and cem_display.is_integer() else str(cem_display)
                    if cem_display == "" or cem_display is None:
                        cem_display = "無"
                    ws2.append([info["original_seal"], info["batch"], cem_display])
                # 設定字型
                for row in ws2.iter_rows(min_row=1, max_row=ws2.max_row, max_col=3):
                    for cell in row:
                        cell.font = Font(name="新細明體", size=12)

            # 儲存
            output_dir = os.path.dirname(self.source_files[0]) if self.source_files else os.getcwd()
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = os.path.join(output_dir, f"配對結果_{timestamp}.xlsx")
            wb.save(output_filename)

            self.log(f"\n💾 輸出檔案：{output_filename}")
            self.log(f"✅ 配對成功：{match_count} 筆")
            self.log(f"⚠️ 未找到：{not_found_count} 筆")
            self.log("\n✨ 處理完成！")

            self.root.after(100, lambda: self.ask_open_folder(output_dir))

        except Exception as e:
            self.log(f"\n❌ 錯誤：{e}")
            self.log(traceback.format_exc())
        finally:
            # 清除暫存目錄
            if self.temp_dir and os.path.exists(self.temp_dir):
                try:
                    shutil.rmtree(self.temp_dir)
                    self.log("🧹 暫存檔案已清除")
                except Exception as e:
                    self.log(f"⚠️ 清除暫存失敗：{e}")
            self.finish_processing()

    def ask_open_folder(self, folder_path):
        if messagebox.askyesno("處理完成", "處理完成！\n\n是否開啟輸出檔案所在的資料夾？"):
            os.startfile(folder_path)

    def finish_processing(self):
        self.progress.stop()
        self.btn_process.config(state=tk.NORMAL)
        self.btn_source.config(state=tk.NORMAL)
        self.btn_generate.config(state=tk.NORMAL)
        self.check_ready()

def main():
    root = tk.Tk()
    app = SerialProcessorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
