#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
编译脚本 - 将Python程序打包为exe文件
"""

import os
import subprocess
import sys

def build_exe():
    """编译exe文件"""
    print("开始编译Poedit自动翻译...")
    
    # PyInstaller命令
    cmd = [
        "python", "-m", "PyInstaller",
        "--onefile",  # 打包成单个exe文件
        "--windowed",  # 不显示控制台窗口
        "--name=Poedit自动翻译",  # 指定exe文件名
        "--icon=icon.ico",  # 图标文件（如果存在）
        "--add-data=requirements.txt;.",  # 包含依赖文件
        "poedit_auto_translator.py"
    ]
    
    # 如果没有图标文件，移除图标参数
    if not os.path.exists("icon.ico"):
        cmd = [c for c in cmd if not c.startswith("--icon")]
    
    try:
        # 执行编译命令
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("编译成功！")
        print(f"输出: {result.stdout}")
        
        # 检查生成的文件
        exe_path = os.path.join("dist", "Poedit自动翻译.exe")
        if os.path.exists(exe_path):
            print(f"可执行文件已生成: {exe_path}")
            print(f"文件大小: {os.path.getsize(exe_path) / 1024 / 1024:.2f} MB")
        else:
            print("警告: 未找到生成的exe文件")
            
    except subprocess.CalledProcessError as e:
        print(f"编译失败: {e}")
        print(f"错误输出: {e.stderr}")
        return False
    except FileNotFoundError:
        print("错误: 未找到PyInstaller，请先安装:")
        print("pip install pyinstaller")
        return False
    
    return True

def install_pyinstaller():
    """安装PyInstaller"""
    print("正在安装PyInstaller...")
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
        print("PyInstaller安装成功！")
        return True
    except subprocess.CalledProcessError as e:
        print(f"PyInstaller安装失败: {e}")
        return False

if __name__ == "__main__":
    print("Poedit自动翻译 - 编译脚本")
    print("=" * 40)
    
    # 检查是否安装了PyInstaller
    try:
        subprocess.run(["python", "-m", "PyInstaller", "--version"], check=True, capture_output=True)
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("未检测到PyInstaller，正在安装...")
        if not install_pyinstaller():
            sys.exit(1)
    
    # 开始编译
    if build_exe():
        print("\n编译完成！可执行文件位于 dist/Poedit自动翻译.exe")
    else:
        print("\n编译失败，请检查错误信息")
        sys.exit(1)