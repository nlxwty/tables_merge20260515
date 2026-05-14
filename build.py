#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PyInstaller build script for tables_merge.py
"""
import PyInstaller.__main__
import os
import sys
import io

# Fix Windows console encoding issue in GitHub Actions
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

def build():
    root = os.path.dirname(os.path.abspath(__file__))
    main_script = os.path.join(root, "tables_merge.py")
    icon_path = os.path.join(root, "tables_merge.png")
    
    # Build arguments
    args = [
        main_script,
        "--onefile",
        "--windowed",
        "--name=tables_merge_tool",
        "--clean",
        "--noconfirm",
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.filedialog",
        "--hidden-import=tkinter.messagebox",
        "--hidden-import=tkinter.scrolledtext",
        "--hidden-import=pandas",
        "--hidden-import=numpy",
        "--hidden-import=openpyxl",
        "--hidden-import=xlrd",
        "--collect-all=pandas",
        "--collect-all=openpyxl",
    ]
    
    # Add icon if exists
    if os.path.exists(icon_path):
        args.append(f"--icon={icon_path}")
        args.append(f"--add-data={icon_path};.")
    
    # Use ASCII-only print to avoid encoding issues
    print("Build args:", args)
    PyInstaller.__main__.run(args)
    print("Build completed!")

if __name__ == "__main__":
    build()
