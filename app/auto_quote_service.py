"""Windows 服务包装器：将 auto_quote_daemon 注册为系统服务，免疫 360 清理"""
import os
import sys
import subprocess
import time
import logging
from pathlib import Path

import win32serviceutil
import win32service
import win32event
import servicemanager

ROOT = Path(__file__).parent
LOG_DIR = Path(r"D:\baojia\electron\outputs\logs")
LOG_DIR.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    filename=str(LOG_DIR / "service.log"),
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
LOG = logging.getLogger("AutoQuoteService")


class AutoQuoteService(win32serviceutil.ServiceFramework):
    _svc_name_ = "AutoQuoteDaemon"
    _svc_display_name_ = "翻译报价自动执行服务"
    _svc_description_ = "工作日每天9点自动执行翻译报价（完美世界/战双/4399），无需用户登录"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.process = None

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)
        if self.process and self.process.poll() is None:
            self.process.terminate()
            LOG.info("守护进程已停止")

    def SvcDoRun(self):
        LOG.info("服务启动")
        try:
            python_exe = r"C:\Users\admin\AppData\Local\Python\pythoncore-3.14-64\python.exe"
            daemon_script = str(ROOT / "auto_quote_daemon.py")
            self.process = subprocess.Popen(
                [python_exe, daemon_script],
                cwd=str(ROOT),
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            LOG.info(f"守护进程已启动 (PID={self.process.pid})")
            # 等待停止信号或子进程退出
            while True:
                rc = win32event.WaitForSingleObject(self.stop_event, 5000)
                if rc == win32event.WAIT_OBJECT_0:
                    break
                if self.process.poll() is not None:
                    LOG.warning(f"守护进程意外退出 (code={self.process.returncode})，5秒后重启")
                    time.sleep(5)
                    self.process = subprocess.Popen(
                        [python_exe, daemon_script],
                        cwd=str(ROOT),
                        creationflags=subprocess.CREATE_NO_WINDOW,
                    )
                    LOG.info(f"守护进程已重启 (PID={self.process.pid})")
        except Exception as e:
            LOG.error(f"服务异常: {e}")
        finally:
            if self.process and self.process.poll() is None:
                self.process.terminate()
            LOG.info("服务停止")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "install":
        # 安装服务
        win32serviceutil.HandleCommandLine(AutoQuoteService)
    elif len(sys.argv) > 1:
        win32serviceutil.HandleCommandLine(AutoQuoteService)
    else:
        # 直接运行（调试用）
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(AutoQuoteService)
        servicemanager.StartServiceCtrlDispatcher()
