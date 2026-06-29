"""TK 填表诊断 v6 - 修复 sparse 数据解析 bug"""
import json
import subprocess
import sys
import os
import shutil
from pathlib import Path

SHEET_URL = "https://my.feishu.cn/sheets/R8QWsav9Nh77V3tWdTQchuCfnMc"
SHEET_ID = "d73cd4"

_LARK_CLI_CMD = shutil.which("lark-cli") or shutil.which("lark-cli.cmd")
_WB = Path(os.path.expanduser("~/.workbuddy"))
_LARK_CLI_DIR = _WB / "binaries" / "node" / "cli-connector-packages"
_RUN_JS = _LARK_CLI_DIR / "node_modules" / "@larksuite" / "cli" / "scripts" / "run.js"
_NODE_EXE = None
_node_versions = sorted(
    (_WB / "binaries" / "node" / "versions").glob("*/node.exe"), reverse=True
)
if _node_versions:
    _NODE_EXE = str(_node_versions[0])

def _cli(*args, timeout=120):
    if _LARK_CLI_CMD:
        cmd = [_LARK_CLI_CMD, "--format", "json"] + list(args)
    elif _RUN_JS.exists() and _NODE_EXE:
        cmd = [_NODE_EXE, str(_RUN_JS), "--format", "json"] + list(args)
    else:
        raise FileNotFoundError("lark-cli not found")
    result = subprocess.run(
        cmd, capture_output=True, text=True, encoding="utf-8", timeout=timeout,
        env={**os.environ, "LARK_CLI_NO_PROXY": "1"},
    )
    if not result.stdout:
        raise RuntimeError(f"lark-cli 无输出: {result.stderr[:300]}")
    return result

def read_sheet_fixed():
    """
    正确解析 lark-cli cells-get 的返回。
    cells-get 返回格式：
      col_indices: 所有请求的列字母列表，如 ['A','B',...,'O']
      cells: 二维数组，每行长度 == len(col_indices)，
            空单元格用 {} 表示（NOT sparse！）
    但如果某些行尾部的列全部为空，lark-cli 可能省略那些 {}，
    导致 row_cells 长度 < len(col_indices)。
    正确做法：遍历 row_cells，用 col_indices[ci] 获取列字母。
    """
    print("获取表格维度...")
    result = _cli("sheets", "+workbook-info", "--url", SHEET_URL)
    info = json.loads(result.stdout)
    sheets = info["data"]["sheets"]
    target = next(s for s in sheets if s["sheet_id"] == SHEET_ID)
    row_count = target["row_count"]
    print(f"  表格共有 {row_count} 行\n")

    all_rows = []
    chunk_size = 100
    for start_row in range(1, row_count + 1, chunk_size):
        end_row = min(start_row + chunk_size - 1, row_count)
        range_str = f"A{start_row}:O{end_row}"
        print(f"  读取 {range_str}...")

        result = _cli(
            "sheets", "+cells-get",
            "--url", SHEET_URL,
            "--sheet-id", SHEET_ID,
            "--range", range_str,
            timeout=120,
        )
        data = json.loads(result.stdout)
        if not data.get("ok"):
            raise RuntimeError(f"读取失败: {data}")

        for rng in data["data"]["ranges"]:
            row_indices = rng["row_indices"]
            col_indices = rng["col_indices"]   # e.g. ['A','B',...,'O']
            cells = rng["cells"]

            for ri, row_cells in enumerate(cells):
                row_num = row_indices[ri]
                row_data = {"row_num": row_num, "values": {}, "attachments": {}}

                # 正确做法：遍历 row_cells，用 col_indices[ci] 获取列字母
                # 如果 row_cells 长度 < len(col_indices)，说明尾部空列被省略了
                for ci, cell in enumerate(row_cells):
                    if ci >= len(col_indices):
                        break   # 防御性：不应该发生
                    col = col_indices[ci]
                    if not cell:
                        row_data["values"][col] = ""
                        continue
                    row_data["values"][col] = cell.get("value", "")

                    # 提取附件信息
                    rt = cell.get("rich_text")
                    if rt:
                        for elem in rt:
                            if elem.get("type") == "attachment":
                                row_data["attachments"][col] = {
                                    "token": elem.get("attachment_token", ""),
                                    "name": elem.get("text", ""),
                                }
                                break

                # 对于没有被 row_cells 覆盖的列（尾部省略的空列），设为 ""
                for ci2 in range(len(row_cells), len(col_indices)):
                    col = col_indices[ci2]
                    row_data["values"].setdefault(col, "")

                all_rows.append(row_data)
    return all_rows

# ---- 主逻辑 ----
print("=" * 60)
print("TK 填表诊断工具 v6 (修复 sparse 数据解析)")
print("=" * 60)

all_rows = read_sheet_fixed()
print(f"\n总共读取 {len(all_rows)} 行")
print(f"首行: {all_rows[0]['row_num']}, 末行: {all_rows[-1]['row_num']}\n")

# 检查 195-205 行
print("===== 195-205 行解析结果 =====")
for r in all_rows:
    rn = r["row_num"]
    if 195 <= rn <= 205:
        n_val = r["values"].get("N", "(missing)")
        o_val = r["values"].get("O", "(missing)")
        n_attach = r["attachments"].get("N")
        print(f"  Row {rn}: N={repr(n_val)[:45]}, O={repr(o_val)[:20]}, N_attach={n_attach is not None}")

# 模拟 find_candidates (blank_stop=200)
print("\n===== 模拟 find_candidates (blank_stop=200) =====")
candidates = []
blank_streak = 0
for r in all_rows:
    rn = r["row_num"]
    n_val = str(r["values"].get("N", "") or "").strip()
    o_val = r["values"].get("O", "")
    n_attach = r["attachments"].get("N")

    # O 已填 → 跳过
    o_filled = False
    if o_val is not None:
        if isinstance(o_val, str) and not o_val.strip():
            o_filled = False
        elif isinstance(o_val, (int, float)):
            o_filled = True
        else:
            o_filled = bool(o_val)

    if o_filled:
        blank_streak = 0
        continue

    # N 空 → 累计空白
    n_has = bool(n_attach or n_val)
    if not n_has:
        blank_streak += 1
        if blank_streak >= 200:
            print(f"  [STOP] 行 {rn}: 连续空白达到 200 行，停止扫描")
            break
        continue

    blank_streak = 0
    candidates.append({"row": rn, "n_attach": n_attach, "n_text": n_val})
    if 190 <= rn <= 210:
        print(f"  ✓ Candidate: Row {rn}: N有数据, O空, attach={n_attach is not None}")

print(f"\n共找到 {len(candidates)} 个候选行")
if candidates:
    print("\n===== 所有候选行 =====")
    for c in candidates[:30]:
        name = c["n_attach"]["name"][:35] if c["n_attach"] else c["n_text"][:35]
        print(f"  Row {c['row']}: {name}")
