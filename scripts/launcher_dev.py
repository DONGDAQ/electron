"""开发版启动器：双击运行，无黑窗口，直接打开报价工具（源码模式，无需打包）"""
import os
import subprocess
import sys

ROOT = r"D:\baojia\electron"


def kill_stale() -> None:
    """杀掉残留的 Flask 进程（开发模式的 python.exe + 打包版 flask_server.exe）"""
    try:
        ps = subprocess.run(
            [
                "powershell", "-NoProfile", "-Command",
                "Get-CimInstance Win32_Process | Where-Object { $_.Name -like 'python*.exe' "
                "-and $_.CommandLine -like '*quote_system.web_app*' } "
                "| ForEach-Object { Stop-Process -Id $_.ProcessId -Force }",
            ],
            capture_output=True, timeout=30,
        )
    except Exception:
        pass
    try:
        subprocess.run(["taskkill", "/F", "/IM", "flask_server.exe"],
                       capture_output=True, timeout=15)
    except Exception:
        pass


def main() -> None:
    # 清掉会破坏 Electron 的环境变量（某些 IDE 会注入）
    for k in ("ELECTRON_RUN_AS_NODE", "NODE_OPTIONS"):
        os.environ.pop(k, None)

    kill_stale()

    electron = os.path.join(ROOT, "node_modules", "electron", "dist", "electron.exe")
    subprocess.Popen([electron, ROOT], cwd=ROOT, close_fds=True)


if __name__ == "__main__":
    main()
