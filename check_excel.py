#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
调试用：检查 Excel 文件中 B 列的值（仅读取，不修改）。

所有路径从 fy4b_config.json 读取。
"""

import json
from pathlib import Path

# 加载配置
_CFG_PATH = Path(__file__).parent / "fy4b_config.json"
with open(_CFG_PATH, "r", encoding="utf-8") as _f:
    C = json.load(_f)

EXCEL_FILE = Path(C["路径"]["Excel文件"])
_XL_ROW_END = C["Excel"]["链接结束行"] + 1  # 检查到 162 行

import openpyxl

wb = openpyxl.load_workbook(str(EXCEL_FILE), data_only=True)
ws = wb.active

for row in range(1, _XL_ROW_END + 1):
    b_val = ws[f"B{row}"].value
    if b_val is not None and not str(b_val).startswith("="):
        print(f"Row {row}: B={b_val}")