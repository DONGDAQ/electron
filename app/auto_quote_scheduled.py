"""定时自动报价：每天自动检查项目的新需求并生成报价单"""
import io
import os
import sys
import traceback
from pathlib import Path

ROOT = Path(__file__).parent
_OUTPUTS_DIR = ROOT.parent / "outputs"
if not _OUTPUTS_DIR.exists():
    _OUTPUTS_DIR = Path(r"D:\baojia\electron\outputs")  # 兼容旧路径
LOG_DIR = _OUTPUTS_DIR / "logs" / "auto_quote"
FILL_LOG_DIR = _OUTPUTS_DIR / "logs"

# ---- 防止并发执行 ----
LOCK_FILE = LOG_DIR / ".auto_quote.lock"

def _acquire_lock():
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    if LOCK_FILE.exists():
        try:
            old_pid = int(LOCK_FILE.read_text().strip())
            import ctypes
            h = ctypes.windll.kernel32.OpenProcess(0x0400, False, old_pid)
            if h:
                ctypes.windll.kernel32.CloseHandle(h)
                print(f"[LOCK] 已有实例运行中 (PID={old_pid})，退出")
                return False
        except (ValueError, OSError):
            pass
    LOCK_FILE.write_text(str(os.getpid()))
    return True

def _release_lock():
    try:
        LOCK_FILE.unlink(missing_ok=True)
    except OSError:
        pass

if not _acquire_lock():
    sys.exit(0)

from quote_system.utils import save_auto_quote_log, write_fill_meta


def _run_captured(label: str, fn):
    old = sys.stdout
    sys.stdout = buf = io.StringIO()
    try:
        fn()
        out = buf.getvalue()
        save_auto_quote_log(LOG_DIR, label, out, "success", "auto")
    except Exception as e:
        out = buf.getvalue()
        err = out + "\n" + traceback.format_exc()
        save_auto_quote_log(LOG_DIR, label, err, "error", "auto")
        print(f"{label} 执行失败: {e}", file=old)
    finally:
        sys.stdout = old


if __name__ == "__main__":
    from quote_system.auto_quote import run as run_huanta
    from quote_system.auto_quote_zhan_shuang import run as run_zhan_shuang
    from quote_system.auto_quote_zhan_shuang_feishu import run as run_zhan_shuang_feishu
    from quote_system.auto_quote_bang2 import run as run_bang2
    from quote_system.auto_quote_bang2_shequ import run as run_bang2_shequ
    from quote_system.auto_quote_4399 import run_scheduled as run_4399
    from quote_system.auto_quote_bilibili import run as run_bilibili

    # 完美世界三个项目（幻塔、异环游戏内、异环发行）
    for project_key in ["huanta", "yihuan_nei", "yihuan_faxing"]:
        _run_captured(project_key, lambda pk=project_key: run_huanta(pk))

    # 战双邮件 + 战双飞书（库洛2个项目）
    _run_captured("zhan_shuang", run_zhan_shuang)
    _run_captured("zhan_shuang_feishu", run_zhan_shuang_feishu)

    # BANG2 + BANG2社区
    _run_captured("bang2", run_bang2)
    _run_captured("bang2_shequ", run_bang2_shequ)

    # 马娘 + HBR
    for project_key in ["maniang", "hbr"]:
        _run_captured(project_key, lambda pk=project_key: run_bilibili(pk))

    # TK
    try:
        from quote_system.auto_fill_tk import run_scheduled as run_tk
        write_fill_meta(FILL_LOG_DIR, "tk", "auto")
        _run_captured("tk", run_tk)
    except Exception as e:
        print(f"TK自动报价跳过: {e}")

    # 4399 在线表填表
    write_fill_meta(FILL_LOG_DIR, "4399", "auto")
    _run_captured("4399", run_4399)

    # 同步 TK 飞书表数据到本地缓存
    try:
        from quote_system.auto_fill_tk import sync_tk_data
        sync_tk_data()
    except Exception as e:
        print(f"TK 数据同步失败: {e}")

    # 同步报告缓存
    try:
        from quote_system.report_cache import sync_report_cache
        sync_report_cache()
    except Exception as e:
        print(f"报告缓存同步失败: {e}")

    _release_lock()
