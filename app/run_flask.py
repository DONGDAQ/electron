import sys
import os
from pathlib import Path

if getattr(sys, 'frozen', False):
    base = Path(sys._MEIPASS)
else:
    base = Path(__file__).parent

sys.path.insert(0, str(base))
os.chdir(str(base))

from quote_system.web_app import app

# 启动时后台刷新缓存（先同步TK，再同步报告）
import threading
def _startup_refresh():
    try:
        from quote_system.auto_fill_tk import sync_tk_data
        sync_tk_data()
    except Exception:
        pass
    try:
        from quote_system.report_cache import sync_report_cache
        sync_report_cache()
    except Exception:
        pass
threading.Thread(target=_startup_refresh, daemon=True).start()

app.run(host="127.0.0.1", port=5000, debug=False)
