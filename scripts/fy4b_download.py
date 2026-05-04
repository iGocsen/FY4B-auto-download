# -*- coding: utf-8 -*-
"""
FY4B 云图下载 - 配置文件驱动入口
所有关键配置从 config.json 读取
"""

import json, subprocess, sys, re, pathlib, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# ============ 配置读取 ============
_SCRIPT_DIR = pathlib.Path(__file__).parent
_config_path = _SCRIPT_DIR / "config.json"
_config = json.loads(_config_path.read_text(encoding="utf-8"))

DOWNLOADER   = pathlib.Path(_config["downloader_script"])
_DOWNLOADER_CWD = DOWNLOADER.parent

def parse_output(raw: str):
    """从脚本输出中解析下载统计四元组 (新下载, 已存在, 未来, 失败)"""
    pattern = r"\[DONE\]\s*新下载:\s*(\d+)\s*\|\s*已存在:\s*(\d+)\s*\|\s*未来:\s*(\d+)\s*\|\s*失败:\s*(\d+)"
    m = re.search(pattern, raw)
    if m:
        return int(m[1]), int(m[2]), int(m[3]), int(m[4])
    return None

def build_message(new, skip_exist, skip_future, fail):
    tmpl = _config["push_message_template"]
    return tmpl.format(new=new, skip_exist=skip_exist, skip_future=skip_future, fail=fail)

def main():
    # 1. 执行核心下载脚本（设置 cwd 确保内部 import 正常）
    result = subprocess.run(
        [sys.executable, str(DOWNLOADER)],
        capture_output=True, text=True,
        encoding="gbk", errors="replace",
        cwd=str(_DOWNLOADER_CWD)
    )
    out = result.stdout + result.stderr

    # 调试：将原始输出写入文件（方便排查）
    debug_file = _SCRIPT_DIR / "last_output.txt"
    debug_file.write_text(f"=== STDOUT ===\n{result.stdout}\n=== STDERR ===\n{result.stderr}", encoding="utf-8")

    # 2. 解析结果
    parsed = parse_output(out)
    if parsed is None:
        # 输出末尾500字供调试
        print(f"⚠️ 解析失败，原始输出末尾500字：\n{out[-500:]}")
        return
    new, skip_exist, skip_future, fail = parsed

    # 3. 输出 Markdown 报告（窗口可见）
    msg = build_message(new, skip_exist, skip_future, fail)
    print(msg)

if __name__ == "__main__":
    main()