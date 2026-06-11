"""Excel Batch Toolkit - Excel 批量处理工具
功能：合并多个Excel、拆分工作表、数据清洗、格式统一
Author: Codex Agent
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinterdnd2 import DND_FILES, TkinterDnD
import threading
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import glob

VERSION = "1.0.0"

class ExcelBatchToolkit:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Excel 批量处理工具 v{VERSION}")
        self.root.geometry("800x600")
        self.root.minsize(700, 500)
        
        # 文件列表
        self.file_list = []
        
        self.setup_ui()
        
    def setup_ui(self):
        # 标题
        title_frame = ttk.Frame(self.root)
        title_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(title_frame, text="Excel 批量处理工具", 
                 font=("微软雅黑", 16, "bold")).pack(side="left")
        ttk.Label(title_frame, text=f"v{VERSION}", 
                 font=("微软雅黑", 9)).pack(side="left", padx=5)
        
        # 主功能选项卡
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Tab 1: 合并Excel
        self.tab_merge = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_merge, text=" 合并Excel ")
        self.setup_merge_tab()
        
        # Tab 2: 拆分工作表
        self.tab_split = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_split, text=" 拆分工作表 ")
        self.setup_split_tab()
        
        # Tab 3: 数据清洗
        self.tab_clean = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_clean, text=" 数据清洗 ")
        self.setup_clean_tab()
        
        # Tab 4: 格式统一
        self.tab_format = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_format, text=" 格式统一 ")
        self.setup_format_tab()
        
        # 日志输出
        log_frame = ttk.LabelFrame(self.root, text="运行日志")
        log_frame.pack(fill="both", expand=False, padx=10, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, 
                                                   font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.log("程序已启动，请选择功能")
        
    def log(self, msg):
        self.log_text.insert(tk.END, f"{msg}\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    # ========== 合并 Tab ==========
    def setup_merge_tab(self):
        frame = ttk.Frame(self.tab_merge)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 拖拽区域
        drop_frame = ttk.LabelFrame(frame, text="拖拽 Excel 文件到此处，或点击选择")
        drop_frame.pack(fill="x", pady=5)
        
        self.merge_drop_label = tk.Label(drop_frame, text="📁 拖拽文件到此处", 
                                         bg="#e8e8e8", height=3, 
                                         font=("微软雅黑", 11))
        self.merge_drop_label.pack(fill="x", padx=5, pady=5)
        self.merge_drop_label.bind("<Button-1>", lambda e: self.select_merge_files())
        
        # 文件列表
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill="both", expand=True, pady=5)
        
        ttk.Label(list_frame, text="已选文件：").pack(anchor="w")
        
        self.merge_listbox = tk.Listbox(list_frame, height=6)
        self.merge_listbox.pack(fill="both", expand=True, side="left")
        
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", 
                                  command=self.merge_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.merge_listbox.config(yscrollcommand=scrollbar.set)
        
        # 操作按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x", pady=5)
        
        ttk.Button(btn_frame, text="选择文件", 
                  command=self.select_merge_files).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="清空列表", 
                  command=self.clear_merge_list).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="全部合并", 
                  command=self.start_merge).pack(side="right", padx=2)
        
        # 选项
        opt_frame = ttk.Frame(frame)
        opt_frame.pack(fill="x", pady=5)
        
        self.merge_header_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_frame, text="第一行为表头", 
                       variable=self.merge_header_var).pack(side="left", padx=5)
        
        self.merge_sheet_name = tk.StringVar(value="Sheet1")
        ttk.Label(opt_frame, text="工作表名：").pack(side="left", padx=5)
        ttk.Entry(opt_frame, textvariable=self.merge_sheet_name, 
                 width=15).pack(side="left")
        
    def select_merge_files(self):
        files = filedialog.askopenfilenames(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        for f in files:
            if f not in self.merge_listbox.get(0, tk.END):
                self.merge_listbox.insert(tk.END, f)
                
    def clear_merge_list(self):
        self.merge_listbox.delete(0, tk.END)
        
    def start_merge(self):
        files = list(self.merge_listbox.get(0, tk.END))
        if not files:
            messagebox.showwarning("提示", "请先选择要合并的 Excel 文件")
            return
            
        output = filedialog.asksaveasfilename(
            title="保存合并后的文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if not output:
            return
            
        threading.Thread(target=self.merge_files, 
                        args=(files, output), daemon=True).start()
        
    def merge_files(self, files, output):
        try:
            self.log(f"开始合并 {len(files)} 个文件...")
            sheet = self.merge_sheet_name.get()
            header = 0 if self.merge_header_var.get() else None
            
            dfs = []
            for f in files:
                self.log(f"  读取: {os.path.basename(f)}")
                df = pd.read_excel(f, sheet_name=sheet, header=header)
                dfs.append(df)
                
            result = pd.concat(dfs, ignore_index=True)
            result.to_excel(output, index=False)
            self.log(f"✅ 合并完成！共 {len(result)} 行数据")
            self.log(f"   保存至: {output}")
            messagebox.showinfo("完成", f"合并成功！\n共 {len(result)} 行数据\n保存至：{output}")
        except Exception as e:
            self.log(f"❌ 合并失败: {str(e)}")
            messagebox.showerror("错误", f"合并失败:\n{str(e)}")
            
    # ========== 拆分 Tab ==========
    def setup_split_tab(self):
        frame = ttk.Frame(self.tab_split)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        ttk.Label(frame, text="选择要拆分的 Excel 文件：").pack(anchor="w")
        
        file_frame = ttk.Frame(frame)
        file_frame.pack(fill="x", pady=5)
        
        self.split_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.split_path).pack(side="left", fill="x", expand=True)
        ttk.Button(file_frame, text="浏览", command=self.select_split_file).pack(side="right", padx=5)
        
        # 拆分方式
        method_frame = ttk.LabelFrame(frame, text="拆分方式")
        method_frame.pack(fill="x", pady=10)
        
        self.split_method = tk.StringVar(value="sheet")
        ttk.Radiobutton(method_frame, text="按工作表拆分（每个sheet存为一个文件）", 
                       variable=self.split_method, value="sheet").pack(anchor="w", pady=2)
        ttk.Radiobutton(method_frame, text="按行数拆分", 
                       variable=self.split_method, value="rows").pack(anchor="w", pady=2)
        
        rows_frame = ttk.Frame(method_frame)
        rows_frame.pack(fill="x", pady=5, padx=20)
        ttk.Label(rows_frame, text="每份行数：").pack(side="left")
        self.split_rows = tk.StringVar(value="1000")
        ttk.Entry(rows_frame, textvariable=self.split_rows, width=10).pack(side="left")
        
        ttk.Button(frame, text="开始拆分", command=self.start_split).pack(pady=10)
        
    def select_split_file(self):
        f = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if f:
            self.split_path.set(f)
            
    def start_split(self):
        path = self.split_path.get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "请选择有效的 Excel 文件")
            return
            
        out_dir = filedialog.askdirectory(title="选择保存目录")
        if not out_dir:
            return
            
        threading.Thread(target=self.split_file, 
                        args=(path, out_dir), daemon=True).start()
        
    def split_file(self, path, out_dir):
        try:
            base = os.path.splitext(os.path.basename(path))[0]
            self.log(f"开始拆分: {os.path.basename(path)}")
            
            if self.split_method.get() == "sheet":
                xl = pd.ExcelFile(path)
                for sheet in xl.sheet_names:
                    self.log(f"  导出工作表: {sheet}")
                    df = pd.read_excel(path, sheet_name=sheet)
                    out = os.path.join(out_dir, f"{base}_{sheet}.xlsx")
                    df.to_excel(out, index=False)
                self.log(f"✅ 拆分完成！共 {len(xl.sheet_names)} 个工作表")
                
            else:
                rows = int(self.split_rows.get())
                df = pd.read_excel(path)
                total = len(df)
                parts = (total // rows) + (1 if total % rows else 0)
                
                for i in range(parts):
                    start = i * rows
                    end = min((i + 1) * rows, total)
                    self.log(f"  导出第 {i+1}/{parts} 份 ({start+1}-{end} 行)")
                    chunk = df.iloc[start:end]
                    out = os.path.join(out_dir, f"{base}_part{i+1}.xlsx")
                    chunk.to_excel(out, index=False)
                self.log(f"✅ 拆分完成！共 {parts} 份文件")
                
            messagebox.showinfo("完成", "拆分成功！")
        except Exception as e:
            self.log(f"❌ 拆分失败: {str(e)}")
            messagebox.showerror("错误", f"拆分失败:\n{str(e)}")
            
    # ========== 数据清洗 Tab ==========
    def setup_clean_tab(self):
        frame = ttk.Frame(self.tab_clean)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        ttk.Label(frame, text="上传或拖拽 Excel 文件：").pack(anchor="w")
        
        file_frame = ttk.Frame(frame)
        file_frame.pack(fill="x", pady=5)
        self.clean_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.clean_path).pack(side="left", fill="x", expand=True)
        ttk.Button(file_frame, text="浏览", command=self.select_clean_file).pack(side="right", padx=5)
        
        # 清洗选项
        opt_frame = ttk.LabelFrame(frame, text="清洗选项")
        opt_frame.pack(fill="x", pady=10)
        
        self.clean_dup = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_frame, text="删除重复行", 
                       variable=self.clean_dup).pack(anchor="w", pady=2)
        
        self.clean_na = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_frame, text="删除全空行", 
                       variable=self.clean_na).pack(anchor="w", pady=2)
        
        self.trim_spaces = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_frame, text="去除首尾空格", 
                       variable=self.trim_spaces).pack(anchor="w", pady=2)
        
        ttk.Button(frame, text="开始清洗", command=self.start_clean).pack(pady=10)
        
    def select_clean_file(self):
        f = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if f:
            self.clean_path.set(f)
            
    def start_clean(self):
        path = self.clean_path.get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "请选择有效的 Excel 文件")
            return
            
        output = filedialog.asksaveasfilename(
            title="保存清洗后的文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if not output:
            return
            
        threading.Thread(target=self.clean_file, 
                        args=(path, output), daemon=True).start()
        
    def clean_file(self, path, output):
        try:
            self.log(f"开始清洗: {os.path.basename(path)}")
            df = pd.read_excel(path)
            before = len(df)
            
            if self.clean_dup.get():
                dup = len(df) - len(df.drop_duplicates())
                df = df.drop_duplicates()
                self.log(f"  删除重复行: {dup} 行")
                
            if self.clean_na.get():
                na = df.isna().all(axis=1).sum()
                df = df.dropna(how='all')
                self.log(f"  删除全空行: {na} 行")
                
            if self.trim_spaces.get():
                for col in df.select_dtypes(include='object').columns:
                    df[col] = df[col].astype(str).str.strip()
                self.log(f"  已去除所有单元格首尾空格")
                
            df.to_excel(output, index=False)
            after = len(df)
            self.log(f"✅ 清洗完成！{before} 行 → {after} 行 (减少 {before-after} 行)")
            messagebox.showinfo("完成", f"清洗成功！\n{before} 行 → {after} 行")
        except Exception as e:
            self.log(f"❌ 清洗失败: {str(e)}")
            messagebox.showerror("错误", f"清洗失败:\n{str(e)}")
            
    # ========== 格式统一 Tab ==========
    def setup_format_tab(self):
        frame = ttk.Frame(self.tab_format)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 文件选择
        file_frame = ttk.Frame(frame)
        file_frame.pack(fill="x", pady=5)
        self.format_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.format_path).pack(side="left", fill="x", expand=True)
        ttk.Button(file_frame, text="浏览", command=self.select_format_file).pack(side="right", padx=5)
        
        # 格式选项
        opt_frame = ttk.LabelFrame(frame, text="格式选项")
        opt_frame.pack(fill="x", pady=10)
        
        self.format_header = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_frame, text="表头加粗 + 背景色", 
                       variable=self.format_header).pack(anchor="w", pady=2)
        
        self.format_border = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_frame, text="添加边框", 
                       variable=self.format_border).pack(anchor="w", pady=2)
        
        self.format_auto_width = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_frame, text="自动调整列宽", 
                       variable=self.format_auto_width).pack(anchor="w", pady=2)
        
        ttk.Button(frame, text="开始格式化", command=self.start_format).pack(pady=10)
        
    def select_format_file(self):
        f = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if f:
            self.format_path.set(f)
            
    def start_format(self):
        path = self.format_path.get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "请选择有效的 Excel 文件")
            return
            
        output = filedialog.asksaveasfilename(
            title="保存格式化后的文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if not output:
            return
            
        threading.Thread(target=self.format_file, 
                        args=(path, output), daemon=True).start()
        
    def format_file(self, path, output):
        try:
            self.log(f"开始格式化: {os.path.basename(path)}")
            
            wb = load_workbook(path)
            
            for ws in wb.worksheets:
                self.log(f"  处理工作表: {ws.title}")
                
                if self.format_header.get() and ws.max_row > 0:
                    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                    header_font = Font(bold=True, color="FFFFFF", size=11)
                    for cell in ws[1]:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        
                if self.format_border.get():
                    thin_border = Border(
                        left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin')
                    )
                    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, 
                                           max_col=ws.max_column):
                        for cell in row:
                            cell.border = thin_border
                            
                if self.format_auto_width.get():
                    for col in ws.columns:
                        max_length = 0
                        col_letter = col[0].column_letter
                        for cell in col:
                            try:
                                cell_len = len(str(cell.value)) if cell.value else 0
                                if cell_len > max_length:
                                    max_length = cell_len
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        ws.column_dimensions[col_letter].width = max(adjusted_width, 8)
                        
            wb.save(output)
            self.log(f"✅ 格式化完成！保存至: {output}")
            messagebox.showinfo("完成", f"格式化成功！\n保存至：{output}")
        except Exception as e:
            self.log(f"❌ 格式化失败: {str(e)}")
            messagebox.showerror("错误", f"格式化失败:\n{str(e)}")


if __name__ == "__main__":
    root = TkinterDnD.Tk() if 'TkinterDnD' in dir() else tk.Tk()
    app = ExcelBatchToolkit(root)
    root.mainloop()