#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FY4B 云图 IDM 自动导入脚本

功能：
1. 检查待下载文件中的日期时间是否已过期
2. 调用 IDM 导入下载任务
3. 使用 Excel COM 接口更新文件并获取计算后的值

所有路径、软件位置、参数均从 fy4b_config.json 读取。
"""

import json
import re
import sys
import subprocess
from datetime import datetime, timedelta
from pathlib import Path


# ===================== 加载配置 =====================
_CFG_PATH = Path(__file__).parent / "fy4b_config.json"

with open(_CFG_PATH, "r", encoding="utf-8") as _f:
    C = json.load(_f)

# 路径
TXT_DIR    = Path(C["路径"]["txt文件目录"])
EXCEL_FILE = Path(C["路径"]["Excel文件"])
IDM_PATH   = Path(C["路径"]["IDM程序"])

# Excel 结构
_XL_SHEET     = C["Excel"]["工作表索引"]
_XL_B162      = C["Excel"]["B162单元格"]
_XL_B2        = C["Excel"]["B2单元格"]
_XL_B1        = C["Excel"]["B1单元格"]
_XL_LINK_COL  = C["Excel"]["链接列"]
_XL_ROW_START = C["Excel"]["链接起始行"]
_XL_ROW_END   = C["Excel"]["链接结束行"]

# txt 文件
TXT_PREFIX  = C["txt文件"]["前缀"]
TXT_DATEFMT = C["txt文件"]["日期格式"]


# ===================== 辅助函数 =====================
def log(msg):
    """带时间戳的日志（自动替换 emoji 避免编码问题）"""
    for emoji, tag in [
        ('❌', '[FAIL]'), ('✅', '[OK]'), ('📄', '[FILE]'),
        ('📅', '[DATE]'), ('🕐', '[TIME]'), ('📦', '[IDM]'),
        ('📊', '[EXCEL]'), ('📝', '[COPY]'), ('📋', '[LINKS]'),
        ('⏳', '[WAIT]'), ('📂', '[OPEN]'),
    ]:
        msg = msg.replace(emoji, tag)
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


def ole_to_datetime(value):
    """Excel OLE 日期 / datetime → Python datetime"""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, (int, float)):
        return datetime(1899, 12, 30) + timedelta(days=float(value))
    return None


def extract_datetime_from_filename(filename):
    """从文件名提取日期时间，如 截止T：2026年04月28日0907.txt"""
    pat = rf"{re.escape(TXT_PREFIX)}(\d{{4}})年(\d{{2}})月(\d{{2}})日(\d{{2}})(\d{{2}})\.txt"
    match = re.search(pat, filename)
    if match:
        y, m, d, h, mi = match.groups()
        return datetime(int(y), int(m), int(d), int(h), int(mi))
    return None


def find_pending_import_file():
    """查找待下载的 txt 文件"""
    for f in TXT_DIR.glob(f"{TXT_PREFIX}*.txt"):
        return f
    return None


# ===================== Excel 操作 =====================
def update_excel_and_file(txt_file):
    """使用 Excel COM 接口更新 Excel 和文本文件，返回 (成功?, 新文件名?)"""
    try:
        import win32com.client
    except ImportError:
        log("[FAIL] 需要安装 pywin32: pip install pywin32")
        return False, None

    excel = wb = None
    try:
        log("[EXCEL] 启动 Excel...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        log(f"[OPEN] 打开文件: {EXCEL_FILE}")
        wb = excel.Workbooks.Open(str(EXCEL_FILE))
        ws = wb.Worksheets(_XL_SHEET)

        # B162 → B2
        b162_value = ws.Range(_XL_B162).Value
        ws.Range(_XL_B2).Value = b162_value
        log(f"[COPY] {_XL_B162} → {_XL_B2}: {b162_value}")

        excel.Calculate()

        # 收集链接
        e_values = []
        for row in range(_XL_ROW_START, _XL_ROW_END + 1):
            cell_value = ws.Range(f"{_XL_LINK_COL}{row}").Value
            if cell_value:
                e_values.append(str(cell_value))
        log(f"[LINKS] 收集到 {len(e_values)} 个链接")

        # 读取新的 B1
        b1_value = ws.Range(_XL_B1).Value
        log(f"[DATE] 新的 B1 值: {b1_value}")

        wb.Save()
        log("[OK] Excel 已保存")

        # 更新文本文件内容
        with open(txt_file, "w", encoding="utf-8") as f:
            f.write("\n".join(e_values))
        log(f"[OK] 文本文件已更新: {txt_file}")

        # 更新文件名
        new_filename = None
        if b1_value:
            new_dt = ole_to_datetime(b1_value)
            if new_dt:
                new_filename = f"{TXT_PREFIX}{new_dt.strftime(TXT_DATEFMT)}.txt"
                new_txt_path = TXT_DIR / new_filename
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
        if wb:
            wb.Close(SaveChanges=False)
        if excel:
            excel.Quit()


# ===================== 主流程 =====================
def check_and_import():
    """主流程"""
    log("=" * 50)
    log("FY4B IDM 自动导入任务开始")
    log(f"[CFG] 配置文件: {_CFG_PATH}")
    log("=" * 50)

    # 1. 查找待下载文件
    txt_file = find_pending_import_file()
    if not txt_file:
        log("[FAIL] 未找到待下载文件")
        return False
    log(f"[FILE] 找到文件: {txt_file.name}")

    # 2. 提取文件名中的日期时间
    file_datetime = extract_datetime_from_filename(txt_file.name)
    if not file_datetime:
        log(f"[FAIL] 无法从文件名提取日期时间: {txt_file.name}")
        return False
    log(f"[DATE] 文件日期时间: {file_datetime.strftime('%Y-%m-%d %H:%M')}")

    # 3. 检查是否过期
    now = datetime.now()
    log(f"[TIME] 当前时间: {now.strftime('%Y-%m-%d %H:%M')}")

    if now <= file_datetime:
        log("[WAIT] 当前时间未超过文件时间，无需导入")
        return True

    log("[OK] 当前时间已超过文件时间，开始导入流程")

    # 4. 调用 IDM 导入
    log(f"[IDM] 调用 IDM 导入: {txt_file}")
    try:
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
    sys.exit(0 if check_and_import() else 1)