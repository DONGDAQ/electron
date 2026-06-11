"""后台调度守护进程：不依赖 Windows 任务计划程序，免疫 360 清理"""
import json
import os
import sys
import subprocess
import time
import urllib.request
from datetime import datetime
from pathlib import Path
from logging import basicConfig, getLogger, INFO

ROOT = Path(__file__).parent
LOG_DIR = Path(r"D:\baojia\electron\outputs") / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

basicConfig(
    filename=str(LOG_DIR / "daemon.log"),
    level=INFO,
    format="%(asctime)s %(message)s",
)
LOG = getLogger(__name__)

PYTHON = sys.executable
SCRIPT = str(ROOT / "auto_quote_scheduled.py")


def _notify_dingtalk(title: str, text: str):
    url = os.getenv("DINGTALK_WEBHOOK", "")
    if not url:
        return
    body = json.dumps({
        "msgtype": "markdown",
        "markdown": {"title": title, "text": f"### {title}\n\n{text}"},
    }).encode()
    try:
        req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"})
        urllib.request.urlopen(req, timeout=10)
    except Exception:
        pass


def _next_9am():
    now = datetime.now()
    target = now.replace(hour=9, minute=0, second=0, microsecond=0)
    if now >= target:
        from datetime import timedelta
        target += timedelta(days=1)
    while target.weekday() >= 5:
        target += timedelta(days=1)
    return target


def _ensure_smb():
    share = r"\\192.168.110.111\【管理者专用】"
    try:
        subprocess.run(
            ["net", "use", share, "/user:dong_daqian", "dq46460055"],
            capture_output=True, text=True, timeout=10,
        )
    except Exception:
        pass


def _run():
    LOG.info("执行自动报价...")
    _ensure_smb()
    try:
        r = subprocess.run(
            [PYTHON, SCRIPT],
            cwd=str(ROOT),
            capture_output=True,
            text=True,
            timeout=600,
            encoding="utf-8",
            errors="replace",
        )
        if r.returncode == 0:
            LOG.info(f"完成: {r.stdout.strip()[-300:]}")
        else:
            err = r.stderr.strip()[-500:]
            LOG.error(f"失败 (exit={r.returncode}): {err}")
            _notify_dingtalk("自动报价失败", f"退出码: {r.returncode}\n\n```\n{err}\n```")
    except Exception as e:
        LOG.error(f"异常: {e}")
        _notify_dingtalk("自动报价异常", str(e))


def _archive_old_records():
    try:
        sys.path.insert(0, str(ROOT))
        from settlement.settlement_tracker import archive_settled
        n = archive_settled(months=3)
        if n:
            LOG.info(f"归档了 {n} 条已结算记录")
    except Exception as e:
        LOG.error(f"归档失败: {e}")


def main():
    LOG.info("守护进程启动")
    last_run_date = None

    while True:
        try:
            now = datetime.now()

            if now.weekday() < 5 and last_run_date != now.date():
                if now.hour == 9 and now.minute < 5:
                    _run()
                    _archive_old_records()
                    last_run_date = now.date()
                elif now.hour >= 10 and now.hour < 12:
                    LOG.info(f"补执行（错过9点窗口，当前 {now.strftime('%H:%M')}）")
                    _run()
                    _archive_old_records()
                    last_run_date = now.date()

            time.sleep(30)
        except KeyboardInterrupt:
            break
        except Exception as e:
            LOG.error(f"守护进程异常: {e}")
            time.sleep(60)

    LOG.info("守护进程停止")


if __name__ == "__main__":
    main()
