# -*- coding: utf-8 -*-
"""
FY4B 云图下载 - 配置文件驱动入口
所有关键配置从 skill-config.json 读取
"""

import json, subprocess, sys, re, pathlib, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# ============ 配置读取 ============
_SCRIPT_DIR = pathlib.Path(__file__).parent
_config_path = _SCRIPT_DIR / "skill-config.json"
_config = json.loads(_config_path.read_text(encoding="utf-8"))

# DOWNLOADER = _SCRIPT_DIR / _config["paths"]["downloader_script"]
DOWNLOADER   = pathlib.Path(_config["paths"]["downloader_script"])
_DOWNLOADER_CWD = DOWNLOADER.parent

# def safe_print(text: str):
#     """兼容 Windows 控制台：emoji 替换为标签，UTF-8 bytes 写入 stdout buffer"""
#     for _emo, _tag in [
#         ("📡", "[FY4B]"), ("✅", "[OK]"), ("⏭️", "[SKIP]"),
#         ("🔮", "[FUTURE]"), ("❌", "[FAIL]"), ("⚠️", "[WARN]"), ("[!]", "[!]"),
#     ]:
#         text = text.replace(_emo, _tag)
#     sys.stdout.buffer.write(text.encode("utf-8") + b"\n")


# # GBK 字节码（从子进程 raw bytes 提取验证）
# # 新下载 = d0c2cf c2d4d8
# # 已存在 = d2d1b4e6 d4da
# # 未来   = ceb4c0b4
# # 失败   = caa7b0dc
# _GBK_PAT_FULL = rb"\[DONE\]\s*\xd0\xc2\xcf\xc2\xd4\xd8:\s*(\d+)\s*\|\s*\xd2\xd1\xb4\xe6\xd4\xda:\s*(\d+)\s*\|\s*\xce\xb4\xc0\xb4:\s*(\d+)\s*\|\s*\xca\xa7\xb0\xdc:\s*(\d+)"
# _GBK_PAT_SIMPLE = rb"\[DONE\]\s*\xd0\xc2\xcf\xc2\xd4\xd8:\s*(\d+)"


# def parse_output_gbk(raw: str) -> tuple | None:
#     """从子进程 GBK 解码字符串中解析下载统计四元组。
#     关键：子进程用 sys.stdout.write() + 默认 UTF-8 编码打印，
#     但我们用 GBK 解码（因为 fy4b_direct_download.py 的 print() 实际
#     受 PYTHONIOENCODING/Windows 控制台编码影响，最终是 GBK 字节）。
#     所以 raw 字符串虽然是乱码，但 encode('gbk') 后可以得到原始字节。
#     """
#     b = raw.encode("gbk", errors="replace")
#     matches = list(re.finditer(_GBK_PAT_FULL, b))
#     if matches:
#         m = matches[-1]
#         return int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
#     m2 = re.search(_GBK_PAT_SIMPLE, b)
#     if m2:
#         return int(m2.group(1)), 0, 0, 0
#     return None


def parse_output(raw: str):
    """从脚本输出中解析下载统计四元组 (新下载, 已存在, 未来, 失败)"""
    pattern = r"\[DONE\]\s*新下载:\s*(\d+)\s*\|\s*已存在:\s*(\d+)\s*\|\s*未来:\s*(\d+)\s*\|\s*失败:\s*(\d+)"
    m = re.search(pattern, raw)
    if m:
        return int(m[1]), int(m[2]), int(m[3]), int(m[4])
    return None

def build_message(new, skip_exist, skip_future, fail, exec_time):
    tmpl = _config["push_message_template"]
    return tmpl.format(new=new, skip_exist=skip_exist, skip_future=skip_future, fail=fail, exec_time=exec_time)

def main():
    # 1. 执行核心下载脚本（设置 cwd 确保内部 import 正常）
    result = subprocess.run(
        [sys.executable, str(DOWNLOADER)],
        capture_output=True, text=True,
        encoding="gbk", errors="replace",
        # cwd=str(_SCRIPT_DIR)
        cwd=str(_DOWNLOADER_CWD)
    )
    out = result.stdout + result.stderr

    # 调试：将原始输出写入文件（方便排查）
    debug_file = _SCRIPT_DIR / "last_output.txt"
    try:
        # debug_file.write_text(out, encoding="utf-8")
        debug_file.write_text(f"=== STDOUT ===\n{result.stdout}\n=== STDERR ===\n{result.stderr}", encoding="utf-8")
    except Exception:
        pass

    # 解析结果（GBK 字节串正则）
    # parsed = parse_output_gbk(out)
    parsed = parse_output(out)
    if parsed is None:
        # 输出末尾500字供调试
        # safe_print(f"[!] 解析失败，原始输出末尾500字：\n{out[-500:]}")
        print(f"⚠️ 解析失败，原始输出末尾500字：\n{out[-500:]}")
        return
    new, skip_exist, skip_future, fail = parsed

    # 输出 Markdown 报告（窗口可见）
    exec_time = subprocess.run(
        ['powershell', '-Command', '(Get-Date).ToString("yyyy-MM-dd HH:mm")'],
        capture_output=True, text=True, encoding='utf-8'
    ).stdout.strip()
    # safe_print("[FY4B] FY4B 云图下载报告")
    # safe_print(f"✅ 新下载: **{new}**")
    # safe_print(f"⏭️ 已存在: **{skip_exist}**")
    # safe_print(f"🔮 未就绪: **{skip_future}**")
    # safe_print(f"❌ 失败: **{fail}**")
    # safe_print(f"⏰ 执行时间：{exec_time}")
    msg = build_message(new, skip_exist, skip_future, fail, exec_time)
    print(msg)

if __name__ == "__main__":
    main()