#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
表格汇总工具
支持两种汇总模式：
1. 按地区行汇总：底表 + 各地区更新表 -> 新汇总表
   【新增】支持底表/收表多Sheet，按Sheet名匹配汇总
   【新增】收表完整性检查、变更记录Sheet、自动打开文件夹
   【新增】输出文件占用检测、自动备份旧文件
2. 按A字段汇总：按指定字段分组汇总数字列，地区列自动前向填充
"""

import os
import sys
import platform
import subprocess
import warnings
from datetime import datetime

warnings.filterwarnings('ignore')

# ========== 【新增】资源文件路径适配（兼容 PyInstaller 打包环境） ==========
def resource_path(relative_path):
    """
    获取资源文件的绝对路径。
    开发环境：使用脚本所在目录
    PyInstaller 打包后：使用 _MEIPASS 临时目录
    """
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 创建的单文件 exe/elf 会将资源解压到 _MEIPASS 临时目录
        base_path = sys._MEIPASS
    else:
        # 正常 Python 运行环境
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# ===== 提前检查依赖 =====
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
except ImportError as e:
    print("未安装 tkinter，请用以下命令安装依赖：")
    print("Linux: sudo apt-get update && sudo apt-get install python3-tk")
    print("MacOS: brew install python-tk")
    sys.exit(1)
try:
    import pandas as pd
    import numpy as np
except ImportError as e:
    print("缺少 pandas 或 numpy，请先执行：pip3 install pandas numpy openpyxl xlrd")
    sys.exit(1)
try:
    import openpyxl, xlrd
except ImportError as e:
    print("缺少 openpyxl 或 xlrd，请先执行：pip3 install openpyxl xlrd")
    sys.exit(1)
from pathlib import Path


class TableMergeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("表格汇总工具")
        self.root.geometry("850x800")
        self.root.resizable(True, True)

        # 【新增】设置窗口图标（兼容打包后环境）
        try:
            icon_path = resource_path("tables_merge.png")
            if os.path.exists(icon_path):
                # Linux 下 iconbitmap 对 png 支持不好，改用 iconphoto
                self.root.iconphoto(True, tk.PhotoImage(file=icon_path))
        except Exception:
            pass  # 静默忽略，不影响功能

        # 模式1：底表 + 收表
        self.base_file = None
        self.update_files = []
        # 模式2：数据表
        self.mode2_files = []

        # 样式
        syst = platform.system()
        if syst == "Windows":
            default_font = ('微软雅黑', 10)
        elif syst == "Darwin":
            default_font = ('Heiti SC', 12)
        else:
            default_font = ('Arial', 10)

        self.style = ttk.Style()
        self.style.configure('TButton', font=default_font)
        self.style.configure('TLabel', font=default_font)
        self.style.configure('TEntry', font=default_font)

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)

        row = 0

        # ========== 模式选择（最前面） ==========
        mode_frame = ttk.LabelFrame(main_frame, text="【第1步】选择汇总模式", padding="10")
        mode_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=5)
        self.mode_var = tk.IntVar(value=1)
        ttk.Radiobutton(mode_frame, text="模式1：底表 + 收表更新（支持多Sheet，后覆盖前）",
                        variable=self.mode_var, value=1, command=self.on_mode_change).grid(row=0, column=0, sticky=tk.W,
                                                                                           pady=2)
        ttk.Radiobutton(mode_frame, text="模式2：按A字段汇总（按任务批次/月份等分组求和，无需底表）",
                        variable=self.mode_var, value=2, command=self.on_mode_change).grid(row=1, column=0, sticky=tk.W,
                                                                                           pady=2)
        row += 1

        # ========== 动态区域（根据模式显示不同内容） ==========
        self.dynamic_frame = ttk.Frame(main_frame)
        self.dynamic_frame.grid(row=row, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        self.dynamic_frame.columnconfigure(0, weight=1)
        row += 1

        # ----- 模式1 容器 -----
        self.mode1_container = ttk.Frame(self.dynamic_frame)
        self.mode1_container.grid(row=0, column=0, sticky=(tk.W, tk.E))
        self.mode1_container.columnconfigure(0, weight=1)

        # 底表
        base_frame = ttk.LabelFrame(self.mode1_container, text="【底表】第一次汇总表 / 模板表（单文件，支持多Sheet）",
                                    padding="10")
        base_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        base_frame.columnconfigure(0, weight=1)

        self.base_path_var = tk.StringVar(value="（未选择）")
        self.base_entry = ttk.Entry(base_frame, textvariable=self.base_path_var, state="readonly")
        self.base_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(base_frame, text="选择底表...", command=self.select_base_file).grid(row=0, column=1)

        ttk.Label(base_frame,
                  text="说明：底表决定最终汇总表有哪些地区、哪些字段、哪些Sheet。收表按Sheet名匹配，只覆盖对应Sheet的更新数据。",
                  foreground="gray", wraplength=700).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))

        # 收表
        update_frame = ttk.LabelFrame(self.mode1_container, text="【收表】各地区更新表（可多选文件或文件夹，支持多Sheet）",
                                      padding="10")
        update_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        update_frame.columnconfigure(0, weight=1)

        source_btn_frame = ttk.Frame(update_frame)
        source_btn_frame.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=5)
        ttk.Button(source_btn_frame, text="添加Excel文件...", command=self.add_update_files).grid(row=0, column=0,
                                                                                                  padx=(0, 10))
        ttk.Button(source_btn_frame, text="添加文件夹...", command=self.add_update_folder).grid(row=0, column=1,
                                                                                                padx=(0, 10))
        ttk.Button(source_btn_frame, text="清空所有", command=self.clear_update_files).grid(row=0, column=2)

        list_frame = ttk.Frame(update_frame)
        list_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.update_listbox = tk.Listbox(list_frame, height=6, selectmode=tk.EXTENDED, font=('Consolas', 9))
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.update_listbox.yview)
        self.update_listbox.configure(yscrollcommand=scrollbar.set)
        self.update_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        del_btn_frame = ttk.Frame(update_frame)
        del_btn_frame.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=2)
        ttk.Button(del_btn_frame, text="删除选中", command=self.delete_selected_updates).grid(row=0, column=0)

        ttk.Label(update_frame,
                  text="说明：收表按地区匹配底表，有值的字段覆盖底表，空字段保留底表原值。Sheet名必须与底表一致。",
                  foreground="gray", wraplength=700).grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))

        # ----- 模式2 容器 -----
        self.mode2_container = ttk.Frame(self.dynamic_frame)
        self.mode2_container.grid(row=0, column=0, sticky=(tk.W, tk.E))
        self.mode2_container.columnconfigure(0, weight=1)

        # 数据表
        data_frame = ttk.LabelFrame(self.mode2_container, text="【数据表】要汇总的原始表（可多选文件或文件夹）",
                                    padding="10")
        data_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        data_frame.columnconfigure(0, weight=1)

        data_btn_frame = ttk.Frame(data_frame)
        data_btn_frame.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=5)
        ttk.Button(data_btn_frame, text="添加Excel文件...", command=self.add_mode2_files).grid(row=0, column=0,
                                                                                               padx=(0, 10))
        ttk.Button(data_btn_frame, text="添加文件夹...", command=self.add_mode2_folder).grid(row=0, column=1,
                                                                                             padx=(0, 10))
        ttk.Button(data_btn_frame, text="清空所有", command=self.clear_mode2_files).grid(row=0, column=2)

        data_list_frame = ttk.Frame(data_frame)
        data_list_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        data_list_frame.columnconfigure(0, weight=1)
        data_list_frame.rowconfigure(0, weight=1)

        self.mode2_listbox = tk.Listbox(data_list_frame, height=6, selectmode=tk.EXTENDED, font=('Consolas', 9))
        data_scrollbar = ttk.Scrollbar(data_list_frame, orient=tk.VERTICAL, command=self.mode2_listbox.yview)
        self.mode2_listbox.configure(yscrollcommand=data_scrollbar.set)
        self.mode2_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        data_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        data_del_frame = ttk.Frame(data_frame)
        data_del_frame.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=2)
        ttk.Button(data_del_frame, text="删除选中", command=self.delete_selected_mode2).grid(row=0, column=0)

        ttk.Label(data_frame, text="说明：模式2不需要底表，直接把所有数据表合并后按A字段分组汇总。",
                  foreground="gray", wraplength=700).grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))

        # 模式2设置
        a_field_frame = ttk.LabelFrame(self.mode2_container, text="【模式2设置】分组与汇总方式", padding="10")
        a_field_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)

        ttk.Label(a_field_frame, text="A字段名称（如：任务批次、月份、类别等）：").grid(row=0, column=0, sticky=tk.W)
        self.a_field_var = tk.StringVar(value="任务批次")
        ttk.Entry(a_field_frame, textvariable=self.a_field_var, width=25).grid(row=0, column=1, sticky=tk.W,
                                                                               padx=(5, 0))

        ttk.Label(a_field_frame, text="数字列汇总方式：").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        self.agg_var = tk.StringVar(value="sum")
        ttk.Combobox(a_field_frame, textvariable=self.agg_var, values=["sum", "mean", "count"],
                     state="readonly", width=15).grid(row=1, column=1, sticky=tk.W, padx=(5, 0), pady=(10, 0))

        # ========== 共用设置 ==========
        row += 1
        common_frame = ttk.LabelFrame(main_frame, text="【第2步】共用设置", padding="10")
        common_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=5)
        common_frame.columnconfigure(1, weight=1)

        cr = 0
        ttk.Label(common_frame, text="输出文件路径：").grid(row=cr, column=0, sticky=tk.W, pady=5)
        self.output_path_var = tk.StringVar()
        ttk.Entry(common_frame, textvariable=self.output_path_var).grid(row=cr, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(common_frame, text="浏览...", command=self.browse_output).grid(row=cr, column=2)
        cr += 1

        ttk.Label(common_frame, text="表头所在行（从1开始计数）：").grid(row=cr, column=0, sticky=tk.W, pady=5)
        self.header_row_var = tk.StringVar(value="1")
        ttk.Spinbox(common_frame, from_=1, to=20, textvariable=self.header_row_var, width=8).grid(row=cr, column=1,
                                                                                                  sticky=tk.W, padx=5)
        ttk.Label(common_frame, text="（底表和数据表的表头必须在同一行）", foreground="gray").grid(row=cr, column=2,
                                                                                                 sticky=tk.W)
        cr += 1

        ttk.Label(common_frame, text="代表地区的字段名称：").grid(row=cr, column=0, sticky=tk.W, pady=5)
        self.region_col_var = tk.StringVar(value="地区")
        ttk.Entry(common_frame, textvariable=self.region_col_var, width=30).grid(row=cr, column=1, sticky=tk.W, padx=5)
        cr += 1

        # ========== 操作按钮 ==========
        row += 1
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=row, column=0, pady=20)
        ttk.Button(btn_frame, text="[开始] 开始汇总", command=self.start_merge, width=20).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="[日志] 查看日志", command=self.show_log, width=20).grid(row=0, column=1, padx=5)
        row += 1

        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 0))

        self.log_messages = []

        # 初始化显示
        self.on_mode_change()

    # ========== 模式切换 ==========
    def on_mode_change(self):
        if self.mode_var.get() == 1:
            self.mode1_container.grid()
            self.mode2_container.grid_remove()
        else:
            self.mode1_container.grid_remove()
            self.mode2_container.grid()
        self.update_status()

    # ========== 模式1：底表操作 ==========
    def select_base_file(self):
        file = filedialog.askopenfilename(
            title="选择底表（第一次汇总表或模板表）",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file:
            self.base_file = file
            self.base_path_var.set(file)
            self.log(f"已选择底表: {Path(file).name}")

    # ========== 模式1：收表操作 ==========
    def add_update_files(self):
        files = filedialog.askopenfilenames(
            title="选择各地区更新表",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        for f in files:
            if f not in self.update_files:
                self.update_files.append(f)
                self.update_listbox.insert(tk.END, f)
        self.update_status()

    def add_update_folder(self):
        folder = filedialog.askdirectory(title="选择包含更新表的文件夹")
        if not folder:
            return
        found = 0
        for ext in ['*.xlsx', '*.xls']:
            for f in Path(folder).glob(ext):
                f_str = str(f)
                if f_str not in self.update_files:
                    self.update_files.append(f_str)
                    self.update_listbox.insert(tk.END, f_str)
                    found += 1
        self.log(f"从文件夹添加了 {found} 个更新表")
        self.update_status()

    def delete_selected_updates(self):
        selected = self.update_listbox.curselection()
        for idx in reversed(selected):
            self.update_listbox.delete(idx)
            del self.update_files[idx]
        self.update_status()

    def clear_update_files(self):
        self.update_listbox.delete(0, tk.END)
        self.update_files.clear()
        self.update_status()

    # ========== 模式2：数据表操作 ==========
    def add_mode2_files(self):
        files = filedialog.askopenfilenames(
            title="选择要汇总的数据表",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        for f in files:
            if f not in self.mode2_files:
                self.mode2_files.append(f)
                self.mode2_listbox.insert(tk.END, f)
        self.update_status()

    def add_mode2_folder(self):
        folder = filedialog.askdirectory(title="选择包含数据表的文件夹")
        if not folder:
            return
        found = 0
        for ext in ['*.xlsx', '*.xls']:
            for f in Path(folder).glob(ext):
                f_str = str(f)
                if f_str not in self.mode2_files:
                    self.mode2_files.append(f_str)
                    self.mode2_listbox.insert(tk.END, f_str)
                    found += 1
        self.log(f"从文件夹添加了 {found} 个数据表")
        self.update_status()

    def delete_selected_mode2(self):
        selected = self.mode2_listbox.curselection()
        for idx in reversed(selected):
            self.mode2_listbox.delete(idx)
            del self.mode2_files[idx]
        self.update_status()

    def clear_mode2_files(self):
        self.mode2_listbox.delete(0, tk.END)
        self.mode2_files.clear()
        self.update_status()

    def update_status(self):
        if self.mode_var.get() == 1:
            base_status = f"底表: {Path(self.base_file).name if self.base_file else '未选择'}"
            update_status = f"收表: {len(self.update_files)} 个文件"
            self.status_var.set(f"模式1 | {base_status} | {update_status}")
        else:
            data_status = f"数据表: {len(self.mode2_files)} 个文件"
            self.status_var.set(f"模式2 | {data_status}")

    def browse_output(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if path:
            self.output_path_var.set(path)

    def log(self, msg):
        self.log_messages.append(msg)
        self.status_var.set(msg)
        self.root.update_idletasks()
        print(msg)

    def read_excel_with_header(self, filepath, header_row):
        pd_header = header_row - 1
        try:
            df = pd.read_excel(filepath, header=pd_header, engine='openpyxl')
        except Exception as e:
            try:
                df = pd.read_excel(filepath, header=pd_header, engine='xlrd')
            except:
                raise e
        return df

    def read_excel_all_sheets(self, filepath, header_row):
        """读取Excel所有Sheet，返回 {sheet_name: df}"""
        pd_header = header_row - 1
        try:
            sheets = pd.read_excel(filepath, sheet_name=None, header=pd_header, engine='openpyxl')
        except Exception as e:
            try:
                sheets = pd.read_excel(filepath, sheet_name=None, header=pd_header, engine='xlrd')
            except:
                raise e
        return sheets

    def clean_dataframe(self, df):
        df = df.dropna(how='all').reset_index(drop=True)
        df = df.dropna(axis=1, how='all')
        return df

    # ========== 【新增】文件占用检测与自动备份 ==========
    def check_file_locked(self, filepath):
        """检测文件是否被其他程序（如Excel）占用"""
        if not os.path.exists(filepath):
            return False
        try:
            # 尝试以读写模式打开，Windows/Linux/Mac通用
            with open(filepath, 'r+b') as f:
                pass
            return False
        except PermissionError:
            return True
        except OSError:
            return True

    def backup_existing_file(self, filepath):
        """如果输出文件已存在且未被占用，自动备份"""
        if not os.path.exists(filepath):
            return
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        p = Path(filepath)
        backup_name = f"{p.stem}_备份_{timestamp}{p.suffix}"
        backup_path = p.parent / backup_name
        os.rename(filepath, backup_path)
        self.log(f"[自动备份] 原文件已备份至: {backup_path}")

    def open_output_folder(self, output_path):
        """自动打开输出文件所在文件夹"""
        folder = os.path.dirname(os.path.abspath(output_path))
        try:
            syst = platform.system()
            if syst == "Windows":
                os.startfile(folder)
            elif syst == "Darwin":
                subprocess.call(["open", folder])
            else:
                subprocess.call(["xdg-open", folder])
            self.log(f"[系统] 已打开输出文件夹: {folder}")
        except Exception as e:
            self.log(f"[提示] 无法自动打开文件夹: {str(e)}")

    def prepare_output_file(self, output_path):
        """
        写入前的统一准备：
        1. 检测文件是否被占用
        2. 如果存在且未占用，自动备份旧文件
        被占用时抛出 PermissionError，由外层捕获弹窗
        """
        if self.check_file_locked(output_path):
            raise PermissionError(
                f"输出文件已被占用，请先关闭 Excel 或其他程序：\n{output_path}\n\n"
                f"关闭后重新点击【开始汇总】即可。"
            )
        self.backup_existing_file(output_path)

    # ========== 模式1：多Sheet汇总 ==========
    def mode1_merge(self, base_path, update_paths, output_path, header_row, region_col):
        self.log("=" * 40)
        self.log("【模式1】底表 + 收表更新（多Sheet汇总）")

        # 【新增】写入前检测占用并备份
        self.prepare_output_file(output_path)

        # 1. 读取底表所有Sheet
        self.log("正在读取底表...")
        try:
            base_sheets_raw = self.read_excel_all_sheets(base_path, header_row)
        except Exception as e:
            raise ValueError(f"读取底表失败: {str(e)}")

        base_sheets = {}
        for sheet_name, df in base_sheets_raw.items():
            df = self.clean_dataframe(df)
            if region_col not in df.columns:
                self.log(f"[跳过] 底表 sheet '{sheet_name}' 未找到地区字段 '{region_col}'")
                continue
            df = df[df[region_col].notna()].copy()
            if len(df) == 0:
                self.log(f"[跳过] 底表 sheet '{sheet_name}' 无有效数据行")
                continue
            base_sheets[sheet_name] = df
            self.log(f"[底表] sheet '{sheet_name}'：{len(df)} 行，{len(df.columns)} 列")

        if not base_sheets:
            raise ValueError("底表所有Sheet均无有效数据，或缺少地区字段")

        # 2. 读取所有收表的所有Sheet
        update_data = {}
        for fp in update_paths:
            fname = Path(fp).name
            try:
                sheets_raw = self.read_excel_all_sheets(fp, header_row)
                update_data[fname] = {}
                for sheet_name, df in sheets_raw.items():
                    df = self.clean_dataframe(df)
                    if region_col not in df.columns:
                        continue
                    df = df[df[region_col].notna()].copy()
                    if len(df) == 0:
                        continue
                    update_data[fname][sheet_name] = df
                sheet_count = len(update_data[fname])
                self.log(f"[成功] 收表 {fname}：{sheet_count} 个有效Sheet")
            except Exception as e:
                self.log(f"[失败] 读取 {fname} 失败: {str(e)}")

        # 3. 逐Sheet汇总
        result_sheets = {}
        all_change_records = []
        all_missing_records = []

        for sheet_name, base_df in base_sheets.items():
            self.log(f"--- 正在汇总 sheet '{sheet_name}' ---")

            base_columns = list(base_df.columns)
            all_columns = list(base_df.columns)
            all_regions = base_df[region_col].dropna().astype(str).str.strip().tolist()
            seen = set()
            ordered_regions = []
            for r in all_regions:
                if r not in seen:
                    seen.add(r)
                    ordered_regions.append(r)

            sheet_updates = {}
            sheet_col_missing = []

            for fname, sheets in update_data.items():
                if sheet_name in sheets:
                    df = sheets[sheet_name]
                    missing_cols = set(base_columns) - set(df.columns) - {region_col}
                    if missing_cols:
                        missing_str = ", ".join(sorted(missing_cols))
                        self.log(f"[警告] {fname} [{sheet_name}] 缺少列: {missing_str}")
                        sheet_col_missing.append(f"{fname}：{missing_str}")

                    for col in df.columns:
                        if col not in all_columns:
                            all_columns.append(col)
                            base_df[col] = np.nan

                    sheet_updates[fname] = df
                else:
                    self.log(f"[提示] {fname} 无 sheet '{sheet_name}'")

            result_data = []
            sheet_changes = []
            missing_regions = []

            for region in ordered_regions:
                base_row_match = base_df[base_df[region_col].astype(str).str.strip() == region]
                if len(base_row_match) == 0:
                    continue

                base_row_original = base_row_match.iloc[0].to_dict()
                row_data = base_row_match.iloc[0].to_dict()
                region_updated = False

                for fname, df in sheet_updates.items():
                    mask = df[region_col].astype(str).str.strip() == region
                    matched = df[mask]
                    if len(matched) > 0:
                        region_updated = True
                        source_row = matched.iloc[0]
                        for col in all_columns:
                            if col == region_col:
                                continue
                            if col in source_row.index and pd.notna(source_row[col]):
                                row_data[col] = source_row[col]

                if not region_updated:
                    missing_regions.append(region)

                for col in all_columns:
                    if col == region_col:
                        continue
                    old_val = base_row_original.get(col)
                    final_val = row_data.get(col)
                    old_empty = pd.isna(old_val) or str(old_val).strip() == ''
                    new_empty = pd.isna(final_val) or str(final_val).strip() == ''

                    if not new_empty and (old_empty or str(old_val) != str(final_val)):
                        sheet_changes.append({
                            'Sheet': sheet_name,
                            '地区': region,
                            '字段': col,
                            '底表原值': old_val if not old_empty else '(空)',
                            '更新后值': final_val
                        })

                result_data.append(row_data)

            result_df = pd.DataFrame(result_data)
            for col in all_columns:
                if col not in result_df.columns:
                    result_df[col] = np.nan
            result_df = result_df[all_columns]
            result_sheets[sheet_name] = result_df

            all_change_records.extend(sheet_changes)
            if missing_regions:
                all_missing_records.append({
                    'type': '地区缺失',
                    'sheet': sheet_name,
                    'regions': "、".join(missing_regions)
                })
            if sheet_col_missing:
                all_missing_records.append({
                    'type': '列缺失',
                    'sheet': sheet_name,
                    'details': "\n".join(sheet_col_missing)
                })

            self.log(f"[完成] sheet '{sheet_name}'：{len(result_df)} 行，变更 {len(sheet_changes)} 处")

        # 4. 写入输出文件
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet_name, df in result_sheets.items():
                safe_name = sheet_name[:31]
                df.to_excel(writer, sheet_name=safe_name, index=False)

            if all_change_records:
                change_df = pd.DataFrame(all_change_records)
                change_df.to_excel(writer, sheet_name='变更记录', index=False)
                self.log(f"[变更记录] 共 {len(change_df)} 条变更，已写入输出文件")
            else:
                self.log("[变更记录] 本次汇总无任何数据变更")

        self.log(f"[完成] 汇总完成！已保存至: {output_path}")
        self.log(f"   共 {len(result_sheets)} 个Sheet")

        # 5. 完整性检查弹窗
        if all_missing_records:
            lines = ["【收表完整性检查】\n"]
            for rec in all_missing_records:
                if rec['type'] == '地区缺失':
                    lines.append(f"• Sheet '{rec['sheet']}' 未收到以下地区数据：{rec['regions']}")
                else:
                    lines.append(f"• Sheet '{rec['sheet']}' 部分收表缺少列：")
                    for d in rec['details'].split("\n"):
                        lines.append(f"    - {d}")
            lines.append("\n上述地区/列已保留底表原值，请核对后补报。")
            messagebox.showwarning("收表完整性检查", "\n".join(lines))

        # 6. 自动打开文件夹
        self.open_output_folder(output_path)
        return result_sheets

    # ========== 模式2：按A字段汇总 ==========
    def mode2_merge(self, file_paths, output_path, header_row, region_col, a_field, agg_func='sum'):
        self.log("正在读取输入文件...")

        # 【新增】写入前检测占用并备份
        self.prepare_output_file(output_path)

        all_raw_data = []
        all_columns = set()

        for fp in file_paths:
            fname = Path(fp).name
            try:
                df = self.read_excel_with_header(fp, header_row)
                df = self.clean_dataframe(df)
                if a_field not in df.columns:
                    self.log(f"[警告] 跳过 {fname}：未找到A字段 '{a_field}'")
                    continue
                if region_col in df.columns:
                    df[region_col] = df[region_col].ffill()
                df['_数据来源'] = fname
                if region_col in df.columns:
                    df['_地区'] = df[region_col].fillna('未指定')
                else:
                    df['_地区'] = fname
                all_raw_data.append(df)
                all_columns.update(df.columns)
                self.log(f"[成功] 读取 {fname}：{len(df)} 行")
            except Exception as e:
                self.log(f"[失败] 读取 {fname} 失败: {str(e)}")
        if not all_raw_data:
            raise ValueError("没有成功读取任何有效数据文件")

        final_columns = []
        for df in all_raw_data:
            for col in df.columns:
                if col not in final_columns:
                    final_columns.append(col)
        aligned_data = []
        for df in all_raw_data:
            for col in final_columns:
                if col not in df.columns:
                    df[col] = np.nan
            aligned_data.append(df[final_columns])

        raw_combined = pd.concat(aligned_data, ignore_index=True)
        self.log(f"原始数据共 {len(raw_combined)} 行")
        numeric_cols = []
        for col in raw_combined.columns:
            if col in [a_field, '_数据来源', '_地区', region_col]:
                continue
            try:
                converted = pd.to_numeric(raw_combined[col], errors='coerce')
                if converted.notna().sum() > 0:
                    numeric_cols.append(col)
            except:
                pass
        self.log(f"识别到 {len(numeric_cols)} 个数字列: {numeric_cols}")
        if agg_func == 'sum':
            agg_dict = {col: 'sum' for col in numeric_cols}
        elif agg_func == 'mean':
            agg_dict = {col: 'mean' for col in numeric_cols}
        elif agg_func == 'count':
            agg_dict = {col: 'count' for col in numeric_cols}
        else:
            agg_dict = {col: 'sum' for col in numeric_cols}
        summary_df = raw_combined[raw_combined[a_field].notna()].copy()
        if len(summary_df) == 0:
            raise ValueError(f"没有有效数据包含A字段 '{a_field}'")
        grouped = summary_df.groupby(a_field, sort=False).agg(agg_dict).reset_index()
        count_df = summary_df.groupby(a_field, sort=False).size().reset_index(name='数据行数')
        grouped = grouped.merge(count_df, on=a_field, how='left')
        ordered_cols = [a_field, '数据行数'] + numeric_cols
        for col in grouped.columns:
            if col not in ordered_cols:
                ordered_cols.append(col)
        grouped = grouped[ordered_cols]
        self.log(f"汇总结果共 {len(grouped)} 个{a_field}")

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            grouped.to_excel(writer, sheet_name='汇总数据', index=False)
            raw_output = raw_combined.drop(columns=['_数据来源', '_地区'], errors='ignore')
            raw_output.to_excel(writer, sheet_name='原始数据', index=False)

        self.log(f"[完成] 汇总完成！已保存至: {output_path}")
        self.log(f"   Sheet1 '汇总数据': {len(grouped)} 行")
        self.log(f"   Sheet2 '原始数据': {len(raw_combined)} 行")

        # 模式2也自动打开文件夹
        self.open_output_folder(output_path)
        return grouped, raw_combined

    def start_merge(self):
        output_path = self.output_path_var.get().strip()
        if not output_path:
            messagebox.showerror("错误", "请选择输出文件路径")
            return
        try:
            header_row = int(self.header_row_var.get())
            if header_row < 1:
                raise ValueError()
        except:
            messagebox.showerror("错误", "表头行必须是大于等于1的整数")
            return
        region_col = self.region_col_var.get().strip()
        if not region_col:
            messagebox.showerror("错误", "请填写地区字段名称")
            return

        mode = self.mode_var.get()
        try:
            self.log("=" * 40)
            self.log("开始汇总...")

            if mode == 1:
                if not self.base_file:
                    messagebox.showerror("错误", "请先选择底表文件")
                    return
                if not self.update_files:
                    messagebox.showerror("错误", "请至少添加一个收表文件")
                    return
                self.mode1_merge(self.base_file, self.update_files, output_path, header_row, region_col)
            else:
                if not self.mode2_files:
                    messagebox.showerror("错误", "模式2需要至少添加一个数据文件")
                    return
                a_field = self.a_field_var.get().strip()
                if not a_field:
                    messagebox.showerror("错误", "模式2需要填写A字段名称")
                    return
                agg_func = self.agg_var.get()
                self.mode2_merge(self.mode2_files, output_path, header_row, region_col, a_field, agg_func)

            messagebox.showinfo("完成", "汇总完成！\n" + self.status_var.get())
        except PermissionError as e:
            # 【新增】友好提示文件被占用
            self.log(f"[错误] {str(e)}")
            messagebox.showerror("文件被占用", str(e))
        except Exception as e:
            self.log(f"[错误] {str(e)}")
            import traceback
            print(traceback.format_exc())
            messagebox.showerror("汇总失败", str(e))

    def show_log(self):
        log_window = tk.Toplevel(self.root)
        log_window.title("运行日志")
        log_window.geometry("600x400")
        text = scrolledtext.ScrolledText(log_window, wrap=tk.WORD, font=('Consolas', 10))
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        for msg in self.log_messages:
            text.insert(tk.END, msg + "\n")
        text.see(tk.END)


def main():
    try:
        root = tk.Tk()
        app = TableMergeApp(root)
        root.mainloop()
    except Exception as e:
        print('[致命错误]', str(e))
        import traceback
        print(traceback.format_exc())
        sys.exit(1)


if __name__ == "__main__":
    main()