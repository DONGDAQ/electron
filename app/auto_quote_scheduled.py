"""定时自动报价：每天自动检查项目的新需求并生成报价单"""
import io
import json
import sys
import traceback
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).parent
LOG_DIR = Path(r"D:\baojia\electron\outputs") / "logs" / "auto_quote"


def _save_log(project_key: str, output: str, status: str):
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = LOG_DIR / f"{project_key}_{ts}.log"
    log_file.write_text(output, encoding="utf-8")
    index_file = LOG_DIR / "_index.json"
    if index_file.exists():
        try:
            index = json.loads(index_file.read_text(encoding="utf-8"))
        except Exception:
            index = []
    else:
        index = []
    index.append({
        "project": project_key,
        "timestamp": datetime.now().isoformat(),
        "status": status,
        "log_file": str(log_file),
        "summary": output.strip().split("\n")[-1] if output.strip() else "",
    })
    if len(index) > 500:
        index = index[-500:]
    index_file.write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")


def _run_captured(label: str, fn):
    old = sys.stdout
    sys.stdout = buf = io.StringIO()
    try:
        fn()
        out = buf.getvalue()
        _save_log(label, out, "success")
    except Exception as e:
        out = buf.getvalue()
        err = out + "\n" + traceback.format_exc()
        _save_log(label, err, "error")
        print(f"{label} 执行失败: {e}", file=old)
    finally:
        sys.stdout = old


if __name__ == "__main__":
    from quote_system.auto_quote import run as run_huanta
    from quote_system.auto_quote_zhan_shuang import run as run_zhan_shuang
    from quote_system.auto_quote_zhan_shuang_feishu import run as run_zhan_shuang_feishu
    from quote_system.auto_quote_4399 import run_scheduled as run_4399

    # 完美世界三个项目（幻塔、异环游戏内、异环发行）
    for project_key in ["huanta", "yihuan_nei", "yihuan_faxing"]:
        _run_captured(project_key, lambda pk=project_key: run_huanta(pk))

    # 战双邮件
    _run_captured("zhan_shuang", run_zhan_shuang)

    # 战双飞书（双月汇总报价单）
    _run_captured("zhan_shuang_feishu", run_zhan_shuang_feishu)

    # 4399 在线表填表
    _run_captured("4399", run_4399)

    # 归档已结算记录
    try:
        sys.path.insert(0, str(ROOT))
        from settlement.settlement_tracker import archive_settled
        n = archive_settled(months=3)
        if n:
            print(f"归档了 {n} 条已结算记录")
    except Exception as e:
        print(f"归档失败: {e}")
