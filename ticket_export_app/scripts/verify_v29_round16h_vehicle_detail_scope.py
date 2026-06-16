#!/usr/bin/env python3
"""Verify round-16H vehicle detail wording and scope boundaries."""

from __future__ import annotations

import os
import sys
from pathlib import Path


os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

APP_DIR = Path(__file__).resolve().parents[1]
ROOT = APP_DIR.parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from PySide6.QtWidgets import QApplication  # noqa: E402
from ui.export_ticket_window import ExportTicketWindow  # noqa: E402


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def main() -> None:
    app = QApplication.instance() or QApplication([])
    window = ExportTicketWindow()
    assert window.btn_vehicle_log.text() == "车辆查询"
    assert window.btn_export_analysis_report.text() == "导出分析报告"
    assert "车辆查询缓存" in window.txt_schedule_debug.placeholderText()

    ui_source = _read(APP_DIR / "ui" / "export_ticket_window.py")
    assert "车辆明细 / 调试日志" not in ui_source
    assert "查看车辆日志" not in ui_source
    assert "排程运行日志" not in ui_source
    assert "tabs.addTab(cause_table" not in ui_source
    assert "type_filter = QComboBox" not in ui_source
    assert "station_filter = QComboBox" not in ui_source
    assert "wait_filter = QComboBox" not in ui_source
    assert "导出分析报告" in ui_source
    assert "车辆查询仅用于按车号快速锁定当前分析结果" in ui_source
    assert '"车辆", "车型", "投车时间", "下线时间", "实际等待"' in ui_source

    usage = _read(APP_DIR / "docs" / "M-Line混流节拍仿真系统_使用说明_v2.9.md")
    assert "车辆明细 / 调试日志" not in usage
    assert "车辆日志" not in usage
    assert "车辆查询" in usage
    assert "仅支持按车号快速锁定" in usage
    assert "完整批次复算、完整工程链、等待真因和计算口径说明请使用“导出分析报告”" in usage

    matrix = _read(APP_DIR / "docs" / "第16轮_信息归属矩阵.md")
    assert "车辆查询界面和当前查询CSV继续用于日常查车" in matrix
    assert "完整Excel用于批次复算和留档" in matrix
    assert "不再保留“等待真因”第二页" in matrix

    print({
        "vehicle_button": window.btn_vehicle_log.text(),
        "report_button": window.btn_export_analysis_report.text(),
        "old_debug_terms_removed_from_ui": True,
        "usage_doc_updated": True,
        "assertions": "passed",
    })
    window.close()
    app.processEvents()


if __name__ == "__main__":
    main()
