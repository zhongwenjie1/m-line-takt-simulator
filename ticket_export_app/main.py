# -*- coding: utf-8 -*-
"""
组合票程序独立入口
兼容两种启动方式：
1) 直接运行 main.py
2) 在项目目录中运行 python main.py / python3 main.py
"""

import os
import sys
import datetime, tempfile, atexit

# —— 关键：当以脚本直接运行时，补齐包路径 —— #
if __name__ == "__main__" and (__package__ is None or __package__ == ""):
    # 当前文件所在目录：.../ticket_export_app
    _pkg_dir = os.path.dirname(os.path.abspath(__file__))
    _parent = os.path.dirname(_pkg_dir)
    if _parent not in sys.path:
        sys.path.insert(0, _parent)
    # 让后续导入以包名开头
    #__package__ = "checker_ui"

from PySide6.QtWidgets import QApplication

# 统一使用包导入，避免“顶层 ui”导致的相对导入失败
from ui.export_ticket_window import ExportTicketWindow

# === Logging: redirect stdout/stderr to file when no attached terminal ===
def _setup_logging():
    # mac: ~/Library/Logs/ticket_export_app ; Windows: %LOCALAPPDATA%\ticket_export_app\logs
    if sys.platform.startswith("darwin"):
        log_dir = os.path.expanduser("~/Library/Logs/ticket_export_app")
    elif os.name == "nt":
        log_dir = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "ticket_export_app", "logs")
    else:
        log_dir = os.path.join(tempfile.gettempdir(), "ticket_export_app_logs")
    os.makedirs(log_dir, exist_ok=True)

    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    log_path = os.path.join(log_dir, f"run-{stamp}.log")

    # only redirect when not launched from a terminal
    if not sys.stdout or not getattr(sys.stdout, "isatty", lambda: False)():
        f = open(log_path, "a", buffering=1, encoding="utf-8", errors="ignore")
        sys.stdout = f
        sys.stderr = f
        print(f"[INFO] Log file: {log_path}")
        atexit.register(lambda: f.close())

# initialize logging before QApplication is created
_setup_logging()
# === End Logging setup ===


def main():
    app = QApplication(sys.argv)
    win = ExportTicketWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
