#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PyInstaller 构建脚本
用于在 GitHub Actions 中打包 tables_merge.py 为 exe
"""
import PyInstaller.__main__
import os
import sys

def build():
    # 获取项目根目录
    root = os.path.dirname(os.path.abspath(__file__))
    
    # 主脚本路径
    main_script = os.path.join(root, "tables_merge.py")
    
    # 图标路径（如果存在）
    icon_path = os.path.join(root, "tables_merge.png")
    icon_arg = f"--icon={icon_path}" if os.path.exists(icon_path) else ""
    
    # 构建参数
    args = [
        main_script,
        "--onefile",           # 打包为单个 exe 文件
        "--windowed",          # 不显示控制台窗口（GUI程序）
        "--name=表格汇总工具",  # 输出文件名
        "--clean",             # 清理临时文件
        "--noconfirm",         # 不询问确认
        # 隐藏导入（tkinter 相关）
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.filedialog",
        "--hidden-import=tkinter.messagebox",
        "--hidden-import=tkinter.scrolledtext",
        # pandas/numpy 相关隐藏导入
        "--hidden-import=pandas",
        "--hidden-import=numpy",
        "--hidden-import=openpyxl",
        "--hidden-import=xlrd",
        # 收集所有数据文件
        "--collect-all=pandas",
        "--collect-all=openpyxl",
    ]
    
    # 如果有图标，添加图标参数
    if icon_arg:
        args.append(icon_arg)
    
    # 添加资源文件（如果有图标）
    if os.path.exists(icon_path):
        args.append(f"--add-data={icon_path};.")  # Windows 用 ; 分隔
    
    print(f"构建参数: {args}")
    PyInstaller.__main__.run(args)
    print("构建完成！")

if __name__ == "__main__":
    build()
