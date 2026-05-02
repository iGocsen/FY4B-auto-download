#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FY4B Excel/txt 文件检查器

功能：
1. 检查 Excel 文件是否存在
2. 检查 txt 文件是否存在
3. 可被其他脚本 import 调用

所有路径从 fy4b_config.json 读取。
"""

import json
import re
from datetime import datetime
from pathlib import Path


# ===================== 加载配置 =====================
_CFG_PATH = Path(__file__).parent / "fy4b_config.json"

with open(_CFG_PATH, "r", encoding="utf-8") as _f:
    C = json.load(_f)

_SCRIPT_DIR = Path(__file__).parent


def _resolve_path(raw_path, base_dir=None):
    """智能路径解析：绝对路径直接使用，相对路径以 base_dir 为基准解析，~ 自动展开。"""
    if base_dir is None:
        base_dir = _SCRIPT_DIR
    # 接受 str 或 Path 对象
    if isinstance(raw_path, Path):
        p_raw = raw_path
    else:
        p_raw = Path(str(raw_path).strip())
    # 展开 ~
    p_raw = p_raw.expanduser()
    if p_raw.is_absolute():
        return p_raw.resolve()
    return (base_dir / p_raw).resolve()


# 路径（所有路径都走智能解析）
EXCEL_FILE = _resolve_path(C["路径"]["Excel文件"])
TXT_DIR = _resolve_path(C["路径"]["txt文件目录"])
TXT_PREFIX = C["txt文件"]["前缀"]


# ===================== 检查函数 =====================
def parse_txt_datetime(filename: str) -> datetime | None:
    """从截止T：YYYY年MM月DD日hhmm.txt 提取 datetime"""
    pat = rf"{re.escape(TXT_PREFIX)}(\d{{4}})年(\d{{2}})月(\d{{2}})日(\d{{2}})(\d{{2}})\.txt$"
    match = re.search(pat, filename)
    if match:
        y, mo, d, h, mi = match.groups()
        return datetime(int(y), int(mo), int(d), int(h), int(mi))
    return None


def find_txt_file() -> tuple[Path | None, datetime | None]:
    """查找截止T.txt 文件，返回 (Path, datetime) 或 (None, None)"""
    if not TXT_DIR.exists():
        return None, None
    
    for f in TXT_DIR.glob(f"{TXT_PREFIX}*.txt"):
        dt = parse_txt_datetime(f.name)
        if dt:
            return f, dt
    return None, None


def check_files_exist() -> tuple[bool, bool, Path | None, datetime | None]:
    """
    检查 Excel 和 txt 文件是否存在。
    
    Returns:
        (excel_exists, txt_exists, txt_path, txt_datetime)
        - excel_exists: Excel 文件是否存在
        - txt_exists: txt 文件是否存在
        - txt_path: txt 文件路径（不存在则为 None）
        - txt_datetime: txt 文件名中的时间（不存在则为 None）
    """
    excel_exists = EXCEL_FILE.exists()
    txt_path, txt_dt = find_txt_file()
    txt_exists = txt_path is not None
    
    return excel_exists, txt_exists, txt_path, txt_dt


# ===================== 主流程 =====================
def main():
    """命令行入口：检查并打印结果"""
    print("=" * 50)
    print("FY4B 文件检查器")
    print(f"[CFG] 配置文件: {_CFG_PATH}")
    print("=" * 50)
    
    excel_ok, txt_ok, txt_path, txt_dt = check_files_exist()
    
    print(f"[EXCEL] {EXCEL_FILE}")
    print(f"  存在: {'是' if excel_ok else '否'}")
    
    print(f"[TXT] 目录: {TXT_DIR}")
    print(f"  存在: {'是' if txt_ok else '否'}")
    if txt_path:
        print(f"  文件: {txt_path.name}")
        print(f"  时间: {txt_dt.strftime('%Y-%m-%d %H:%M')}")
    
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main())