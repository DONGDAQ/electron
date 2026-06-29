"""独立脚本：用修复后的逻辑统计TK结算数据"""
import subprocess, json, sys, time
from pathlib import Path
from datetime import date

# ---- 找到 lark-cli ----
def _find_cli():
    for name in ("lark-cli", "lark-cli.cmd"):
        p = subprocess.shutil.which(name)
        if p:
            return p
    return None

LARK = _find_cli()
if not LARK:
    print("lark-cli 未找到，请先安装 npm install -g @larksuite/cli")
    sys.exit(1)

SHEET_URL = "https://my.feishu.cn/sheets/R8QWsav9Nh77V3tWdTQchuCfnMc"
current_month = date.today().month

def run_lark(*args, max_retry=3):
    for i in range(max_retry):
        r = subprocess.run([LARK] + list(args), capture_output=True, text=True, timeout=30)
        if r.returncode == 0 and r.stdout.strip():
            return r.stdout.strip()
        time.sleep(2)
    return r.stdout.strip()

print("正在读取 TK 表 N/O 列数据...")
# 使用 cells-get 命令（参考 auto_fill_tk.py 的逻辑）
# 先获取表 sheet_id
info_raw = run_lark("sheets", "info", SHEET_URL)
print(f"info: {info_raw[:200]}")

# 直接用正确的 lark-cli 命令格式
# 根据代码，正确的命令是: lark-cli sheets cells-get <sheet_url> --range <range>
raw = run_lark("sheets", "cells-get", SHEET_URL, "--range", f"N1:O300")
print(f"前200字符: {raw[:200]}")

if not raw or raw.startswith("{"):
    print("可能需要不同的命令格式，尝试 sheet-id 方式...")
    # 解析 sheet info 获取 sheet_id
    try:
        info = json.loads(info_raw)
        sheet_id = info.get("data", {}).get("sheet_id", "")
        print(f"sheet_id: {sheet_id}")
    except:
        print("无法解析 sheet info")
        sys.exit(1)
else:
    print("读取成功，开始解析...")
    # 解析逻辑同 auto_fill_tk.py
    pass
