#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FY4B 云图 IDM 自动导入脚本
功能：
1. 检查待导入文件中的日期时间是否已过期
2. 调用 IDM 导入下载任务
3. 使用 Excel COM 接口更新文件并获取计算后的值
"""

import os
import re
import sys
import subprocess
from datetime import datetime, timedelta
from pathlib import Path

# 配置
BASE_DIR = Path(r"E:\云图\FY4B\待导入日期")
IDM_PATH = Path(r"D:\Program Files (x86)\Internet Download Manager\IDMan.exe")
EXCEL_FILE = BASE_DIR / "链接生成-有效期24小时.xlsx"

def log(msg):
    """带时间戳的日志"""
    # 移除 emoji 避免编码问题
    msg = msg.replace('❌', '[FAIL]').replace('✅', '[OK]').replace('📄', '[FILE]')
    msg = msg.replace('📅', '[DATE]').replace('🕐', '[TIME]').replace('📦', '[IDM]')
    msg = msg.replace('📊', '[EXCEL]').replace('📝', '[COPY]').replace('📋', '[LINKS]')
    msg = msg.replace('⏳', '[WAIT]').replace('📂', '[OPEN]')
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

def extract_datetime_from_filename(filename):
    """从文件名提取日期时间
    格式：截止：2026年04月28日0907.txt
    """
    match = re.search(r'截止：(\d{4})年(\d{2})月(\d{2})日(\d{2})(\d{2})\.txt', filename)
    if match:
        year, month, day, hour, minute = match.groups()
        return datetime(int(year), int(month), int(day), int(hour), int(minute))
    return None

def find_pending_import_file():
    """查找待导入的txt文件"""
    for f in BASE_DIR.glob("截止：*.txt"):
        return f
    return None

def update_excel_and_file(txt_file):
    """使用 Excel COM 接口更新 Excel 和文本文件"""
    try:
        import win32com.client
    except ImportError:
        log("[FAIL] 需要安装 pywin32: pip install pywin32")
        return False, None
    
    excel = None
    wb = None
    try:
        # 启动 Excel
        log("[EXCEL] 启动 Excel...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # 打开工作簿
        log(f"[OPEN] 打开文件: {EXCEL_FILE}")
        wb = excel.Workbooks.Open(str(EXCEL_FILE))
        ws = wb.Worksheets(1)
        
        # 复制 B162 到 B2（只保留值）
        b162_value = ws.Range("B162").Value
        ws.Range("B2").Value = b162_value
        log(f"[COPY] B162 -> B2: {b162_value}")
        
        # 强制重新计算
        excel.Calculate()
        
        # 收集 E2:E161 的值
        e_values = []
        for row in range(2, 162):
            cell_value = ws.Range(f"E{row}").Value
            if cell_value:
                e_values.append(str(cell_value))
        
        log(f"[LINKS] 收集到 {len(e_values)} 个链接")
        
        # 获取 B1 的新值（用于更新文件名）
        b1_value = ws.Range("B1").Value
        log(f"[DATE] 新的 B1 值: {b1_value}")
        
        # 保存 Excel
        wb.Save()
        log("[OK] Excel 已保存")
        
        # 更新文本文件内容
        with open(txt_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(e_values))
        log(f"[OK] 文本文件已更新: {txt_file}")
        
        # 更新文件名（根据 B1 的新值）
        new_filename = None
        if b1_value:
            # B1 是 datetime 类型
            if isinstance(b1_value, datetime):
                new_dt = b1_value
            else:
                # Excel 返回的是 float (OLE Automation Date)
                new_dt = datetime(1899, 12, 30) + timedelta(days=float(b1_value))
            
            new_filename = f"截止：{new_dt.strftime('%Y年%m月%d日%H%M')}.txt"
            new_txt_path = BASE_DIR / new_filename
            
            if txt_file != new_txt_path:
                txt_file.rename(new_txt_path)
                log(f"[OK] 文件已重命名: {new_filename}")
        
        return True, new_filename
        
    except Exception as e:
        log(f"[FAIL] 更新 Excel/文本文件失败: {e}")
        import traceback
        traceback.print_exc()
        return False, None
    finally:
        # 关闭 Excel
        if wb:
            wb.Close(SaveChanges=False)
        if excel:
            excel.Quit()

def check_and_import():
    """主流程"""
    log("=" * 50)
    log("FY4B IDM 自动导入任务开始")
    log("=" * 50)
    
    # 1. 查找待导入文件
    txt_file = find_pending_import_file()
    if not txt_file:
        log("[FAIL] 未找到待导入文件")
        return False
    
    log(f"[FILE] 找到文件: {txt_file.name}")
    
    # 2. 提取文件名中的日期时间
    file_datetime = extract_datetime_from_filename(txt_file.name)
    if not file_datetime:
        log(f"[FAIL] 无法从文件名提取日期时间: {txt_file.name}")
        return False
    
    log(f"[DATE] 文件日期时间: {file_datetime.strftime('%Y-%m-%d %H:%M')}")
    
    # 3. 检查当前时间是否大于文件时间
    now = datetime.now()
    log(f"[TIME] 当前时间: {now.strftime('%Y-%m-%d %H:%M')}")
    
    if now <= file_datetime:
        log("[WAIT] 当前时间未超过文件时间，无需导入")
        return True
    
    log("[OK] 当前时间已超过文件时间，开始导入流程")
    
    # 4. 调用 IDM 导入
    log(f"[IDM] 调用 IDM 导入: {txt_file}")
    try:
        # IDM 命令行参数：
        # /s - 静默模式
        # /n - 不显示确认对话框
        # /import - 从文本文件导入
        result = subprocess.run(
            [str(IDM_PATH), "/s", "/import", str(txt_file)],
            capture_output=True,
            text=True,
            timeout=60
        )
        log(f"[OK] IDM 导入命令已执行 (返回码: {result.returncode})")
    except subprocess.TimeoutExpired:
        log("[FAIL] IDM 导入超时")
        return False
    except Exception as e:
        log(f"[FAIL] IDM 导入失败: {e}")
        return False
    
    # 5. 更新 Excel 和文本文件
    log("[EXCEL] 开始更新 Excel 和文本文件...")
    success, new_filename = update_excel_and_file(txt_file)
    
    if not success:
        return False
    
    log("=" * 50)
    log("[OK] FY4B IDM 自动导入任务完成")
    log("=" * 50)
    return True

if __name__ == "__main__":
    success = check_and_import()
    sys.exit(0 if success else 1)
