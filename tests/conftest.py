import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

import pytest


@pytest.fixture(autouse=True)
def _reset_fill_state():
    """每个测试前后重置 TK/4399 填表的全局状态，防止测试间互相污染。

    之前 test_web_app.py 里的填表/结算测试会真实启动后台线程和子进程，
    全局状态（TK_FILL_RUNNING / FILL_4399_PROCESS 等）会泄漏到后续测试，
    导致下一次真实运行时出现意外的 409 或文件移动。
    """
    from quote_system import web_app
    web_app.TK_FILL_RUNNING = False
    web_app.TK_FILL_THREAD = None
    web_app.FILL_4399_PROCESS = None
    yield
    web_app.TK_FILL_RUNNING = False
    web_app.TK_FILL_THREAD = None
    web_app.FILL_4399_PROCESS = None
