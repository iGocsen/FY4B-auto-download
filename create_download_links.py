#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FY4B 云图下载链接生成器

功能：
1. 根据 url_template.md 模板生成 160 个下载链接
2. 保存到截止T.txt 文件

所有参数从 fy4b_config.json 读取。
"""

import json
import re
from datetime import datetime, timedelta
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
TXT_DIR = _resolve_path(C["路径"]["txt文件目录"])
# TEMPLATE_FILE = _resolve_path(C["路径"].get("模板文件", "url_template.md"))
TEMPLATE_FILE = _resolve_path(C["路径"]["模板文件"])

TXT_PREFIX       = C["txt文件"]["前缀"]
TXT_DATEFMT      = C["txt文件"]["日期格式"]
LINK_COUNT       = C["txt文件"]["链接数量"]
OFFSET_HOURS     = C["txt文件"]["时间偏移小时"]
TOTAL_HOURS      = C["txt文件"]["总时长小时"]
INTERVAL_MINUTES = C["txt文件"]["链接间隔分钟"]
EXPIRY_OFFSET_MINUTES = C["txt文件"]["截止时间偏移分钟"]


# ===================== 模板解析 =====================
def parse_template() -> str:
    """
    从 url_template.md 解析链接模板。
    
    Returns:
        模板字符串，如：
        https://image.nmc.cn/product/YYYY/MM/DD/WXBL/SEVP_NSMC_WXBL_FY4B_ETCC_ACHN_LNO_PY_YYYYMMDDhhmm00000.JPG
    """
    if not TEMPLATE_FILE.exists():
        raise FileNotFoundError(f"模板文件不存在: {TEMPLATE_FILE}")
    
    with open(TEMPLATE_FILE, "r", encoding="utf-8") as f:
        content = f.read()
    
    # 提取 ``` 代码块中的模板，或直接提取 https:// 链接
    code_block_match = re.search(r"```\s*\n?(https?://[^\s`]+)\n?```", content, re.DOTALL)
    if code_block_match:
        return code_block_match.group(1).strip()
    
    # 尝试直接提取链接
    link_match = re.search(r"https?://[^\s`]+", content)
    if link_match:
        return link_match.group(0).strip()
    
    raise ValueError(f"模板文件中未找到有效链接: {TEMPLATE_FILE}")


def format_link(template: str, dt: datetime) -> str:
    """
    将模板中的占位符替换为实际时间值。
    
    Args:
        template: 链接模板
        dt: 时间
    
    Returns:
        完整链接
    """
    result = template
    result = result.replace("YYYY", dt.strftime("%Y"))
    result = result.replace("MM", dt.strftime("%m"))
    result = result.replace("DD", dt.strftime("%d"))
    result = result.replace("hh", dt.strftime("%H"))
    result = result.replace("mm", dt.strftime("%M"))
    return result


# ===================== 链接生成 =====================
def generate_links(start_time: datetime = None) -> list[str]:
    """
    生成下载链接列表。
    
    Args:
        start_time: 起始时间，默认为 (当前时间 - 36h) 的整点
    
    Returns:
        链接列表（160 个）
    """
    template = parse_template()
    
    if start_time is None:
        now = datetime.now()
        start_time = (now - timedelta(hours=OFFSET_HOURS)).replace(minute=0, second=0, microsecond=0)
    
    links = []
    for i in range(LINK_COUNT):
        dt = start_time + timedelta(minutes=i * INTERVAL_MINUTES)
        link = format_link(template, dt)
        links.append(link)
    
    return links


def generate_txt_filename(links: list[str]) -> str:
    """
    根据最后一个链接的时间生成 txt 文件名。
    
    最后一个链接时间 + 22 分钟 → 截止时间
    
    Args:
        links: 链接列表
    
    Returns:
        文件名，如 "截止T：2026年05月02日2007.txt"
    """
    # 从最后一个链接提取时间
    last_link = links[-1]
    # 提取 YYYYMMDDhhmm
    match = re.search(r"(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})00000\.JPG$", last_link, re.IGNORECASE)
    if not match:
        # 备用：从链接中解析时间
        raise ValueError(f"无法从链接中提取时间: {last_link}")
    
    y, mo, d, h, mi = match.groups()
    last_dt = datetime(int(y), int(mo), int(d), int(h), int(mi))
    
    # 加上截止时间偏移
    expiry_dt = last_dt + timedelta(minutes=EXPIRY_OFFSET_MINUTES)
    
    return f"{TXT_PREFIX}{expiry_dt.strftime(TXT_DATEFMT)}.txt"


def generate_and_save_txt(start_time: datetime = None) -> Path:
    """
    生成链接并保存到 txt 文件。
    
    Args:
        start_time: 起始时间，默认为 (当前时间 - 36h) 的整点
    
    Returns:
        生成的 txt 文件路径
    """
    links = generate_links(start_time)
    filename = generate_txt_filename(links)
    txt_path = TXT_DIR / filename
    
    TXT_DIR.mkdir(parents=True, exist_ok=True)
    
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(links))
    
    return txt_path


# ===================== 主流程 =====================
def main():
    """命令行入口：生成链接并保存"""
    print("=" * 50)
    print("FY4B 云图链接生成器")
    print(f"[CFG] 配置文件: {_CFG_PATH}")
    print("=" * 50)
    
    try:
        template = parse_template()
        print(f"[TEMPLATE] {template}")
        
        txt_path = generate_and_save_txt()
        print(f"[OK] 生成文件: {txt_path.name}")
        print(f"[INFO] 包含 {LINK_COUNT} 个链接")
        
        # 显示时间范围
        links = generate_links()
        first_link = links[0]
        last_link = links[-1]
        print(f"[TIME] 时间范围: {first_link[-20:-4]} → {last_link[-20:-4]}")
        
        return 0
    except Exception as e:
        print(f"[FAIL] {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    import sys
    sys.exit(main())