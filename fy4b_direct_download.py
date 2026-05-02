#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FY4B 云图自动下载脚本

流程:
1. 查找 截止T：{B1}.txt 文件，提取过期时间
2. 当前时间 > B1（已过期）:
   a. 读取旧 txt → 下载到 E:\云图\FY4B_40H\
   b. 更新 Excel (B162→B2)
   c. 将 FY4B_40H\ 所有图片移动至 E:\云图\FY4B\
   d. 移动完成后，生成新的截止T txt
3. 当前时间 <= B1（未过期）:
   a. 读取现有 txt → 下载到 E:\云图\FY4B_40H\
   b. （不更新 Excel，不移动，不生成新 txt）
"""

import os
import re
import sys
import time
import shutil
from pathlib import Path
from datetime import datetime, timedelta

# 配置
EXCEL_FILE   = Path(r"E:\云图\FY4B\待导入日期\链接生成-有效期24小时.xlsx")
TXT_DIR      = Path(r"E:\云图\FY4B\待导入日期")
STAGING_DIR  = Path(r"E:\云图\FY4B_40H")
TARGET_DIR   = Path(r"E:\云图\FY4B")
TXT_PREFIX   = "截止T："


def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


def ole_to_datetime(value):
    """将 Excel OLE 日期或 datetime 转为 Python datetime"""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, (int, float)):
        return datetime(1899, 12, 30) + timedelta(days=float(value))
    return None


def format_b1_filename(dt):
    """将 datetime 格式化为文件名: 截止T：2026年04月28日0907.txt"""
    return f"{TXT_PREFIX}{dt.strftime('%Y年%m月%d日%H%M')}.txt"


def parse_txt_datetime(filename):
    """从截止T：YYYY年MM月DD日HHMM.txt 提取 datetime"""
    match = re.search(r'截止T：(\d{4})年(\d{2})月(\d{2})日(\d{2})(\d{2})\.txt$', filename)
    if match:
        y, m, d, h, mi = match.groups()
        return datetime(int(y), int(m), int(d), int(h), int(mi))
    return None


def find_txt_file():
    """查找截止T.txt文件，返回 (Path, datetime) 或 (None, None)"""
    for f in TXT_DIR.glob(f"{TXT_PREFIX}*.txt"):
        dt = parse_txt_datetime(f.name)
        if dt:
            return f, dt
    return None, None


def read_txt_links(txt_path):
    """从txt文件读取链接列表"""
    with open(txt_path, 'r', encoding='utf-8') as f:
        return [line.strip() for line in f if line.strip()]


def move_staging_to_target():
    """将 STAGING_DIR 中所有 .jpg/.png 文件移动到 TARGET_DIR"""
    if not STAGING_DIR.exists():
        return 0
    moved = 0
    for f in STAGING_DIR.iterdir():
        if f.suffix.lower() in ('.jpg', '.jpeg', '.png', '.gif'):
            dest = TARGET_DIR / f.name
            try:
                shutil.move(str(f), str(dest))
                moved += 1
            except Exception as e:
                log(f"[MOVE FAIL] {f.name}: {e}")
    if moved:
        log(f"[MOVE] {moved} 个文件已移动到 {TARGET_DIR}")
    return moved


def update_excel_and_generate_txt():
    """
    更新 Excel (B162→B2, 重算)，生成新的待下载txt。
    返回新 txt 的 Path，失败返回 None。
    """
    log("[EXCEL] 启动 Excel 更新流程...")
    import win32com.client
    excel = wb = ws = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(EXCEL_FILE))
        ws = wb.Worksheets(1)
        excel.Calculate()

        # B162 → B2
        b162_val = ws.Range("B162").Value
        if b162_val is None:
            log("[FAIL] B162 为空，无法更新")
            return None

        ws.Range("B2").Value = b162_val
        excel.Calculate()

        # 读取新的 B1 和链接
        new_b1 = ole_to_datetime(ws.Range("B1").Value)
        links = []
        for row in range(2, 162):
            v = ws.Range(f"E{row}").Value
            if v:
                links.append(str(v))

        if not new_b1 or not links:
            log("[FAIL] 更新后 B1 或链接为空")
            return None

        wb.Save()
        log(f"[EXCEL] B162->B2 完成, 新过期时间: {new_b1.strftime('%Y-%m-%d %H:%M')}")

        # 写入新 txt (不删除旧文件，由调用方控制)
        new_txt_name = format_b1_filename(new_b1)
        new_txt_path = TXT_DIR / new_txt_name
        with open(new_txt_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(links))
        log(f"[OK] 新文件: {new_txt_name} ({len(links)} 个链接)")

        return new_txt_path

    except Exception as e:
        log(f"[FAIL] Excel 更新失败: {e}")
        import traceback
        traceback.print_exc()
        return None
    finally:
        if wb:
            wb.Close(SaveChanges=False)
        if excel:
            excel.Quit()


def download_links(links):
    """批量下载到 STAGING_DIR，跳过已存在于 STAGING_DIR 或 TARGET_DIR 的文件"""
    success = skip = fail = 0
    total = len(links)

    log(f"[DOWNLOAD] 共 {total} 个文件, 目标: {STAGING_DIR}")

    STAGING_DIR.mkdir(parents=True, exist_ok=True)

    for i, url in enumerate(links, 1):
        filename = url.split('/')[-1]
        staging_path = STAGING_DIR / filename
        target_path  = TARGET_DIR / filename

        if staging_path.exists() or target_path.exists():
            skip += 1
            continue

        try:
            import urllib.request
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=30) as resp:
                with open(staging_path, 'wb') as f:
                    f.write(resp.read())
            success += 1
            # 每下载 20 个打印一次进度
            if success % 20 == 0 or i == total:
                log(f"[PROGRESS] {i}/{total} | 新下载: {success} | 跳过: {skip} | 失败: {fail}")
        except Exception as e:
            fail += 1
            log(f"[FAIL] {filename}: {e}")

        # 每 10 个请求暂停一下
        if i % 10 == 0:
            time.sleep(0.5)

    log(f"[DONE] 新下载: {success} | 跳过: {skip} | 失败: {fail}")
    return success, skip, fail


def main():
    log("=" * 50)
    log("FY4B 云图自动下载")
    log("=" * 50)

    txt_path, b1_dt = find_txt_file()
    now = datetime.now()
    need_update = False

    # 1. 判断是否需要更新 Excel
    if b1_dt:
        log(f"[CHECK] 过期时间: {b1_dt.strftime('%Y-%m-%d %H:%M')}")
        log(f"[CHECK] 当前时间: {now.strftime('%Y-%m-%d %H:%M')}")
        if now > b1_dt:
            need_update = True
            log("[INFO] 已过期，先下载当前批次，再更新 Excel")
        else:
            log("[INFO] 未过期，直接下载")
    else:
        need_update = True
        log("[CHECK] 未找到待下载txt文件，需要更新 Excel")

    # 2. 读取 txt 并下载到临时目录（无论是否过期都要先下载）
    if txt_path and txt_path.exists():
        links = read_txt_links(txt_path)
        log(f"[DOWNLOAD] 当前批次: {txt_path.name} ({len(links)} 个链接)")
    elif need_update:
        # 无旧 txt 但需要更新时，先更新 Excel 获取新链接
        links = None
    else:
        log("[FAIL] 无待下载文件且无需更新，退出")
        return False

    if links is not None:
        ok, skip, fail = download_links(links)
        log(f"[DONE] 当前批次下载完成 | 新下载: {ok} | 跳过: {skip} | 失败: {fail}")

    # 3. 过期 → 更新 Excel → 移动临时目录文件到目标 → 生成新 txt
    if need_update:
        old_txt_path = txt_path

        # 3a. 更新 Excel，生成新 txt
        new_txt_path = update_excel_and_generate_txt()
        if not new_txt_path:
            log("[FAIL] Excel 更新失败，退出")
            return False

        # 3b. 移动临时目录所有图片到目标目录
        moved = move_staging_to_target()
        log(f"[MOVE] 本次共移动 {moved} 个文件到 {TARGET_DIR}")

        # 3c. 删除旧 txt
        if old_txt_path and old_txt_path.exists():
            old_txt_path.unlink()
            log(f"[DELETE] 旧文件: {old_txt_path.name}")

        # 3d. 读取新 txt（下一批次链接），下载到临时目录（等下次更新再移动）
        new_links = read_txt_links(new_txt_path)
        log(f"[DOWNLOAD] 下一批次: {new_txt_path.name} ({len(new_links)} 个链接)")
        ok2, skip2, fail2 = download_links(new_links)
        log(f"[DONE] 下一批次预下载 | 新下载: {ok2} | 跳过: {skip2} | 失败: {fail2}")

    log("=" * 50)
    log("[OK] 任务完成")
    log("=" * 50)
    return True


if __name__ == "__main__":
    sys.exit(0 if main() else 1)
