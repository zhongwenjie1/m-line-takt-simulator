#!/usr/bin/env python3
"""Verify that schedule inputs invalidate cached analysis results."""

from __future__ import annotations

import os
import sys
from pathlib import Path


os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from PySide6.QtWidgets import QApplication  # noqa: E402
from ui.export_ticket_window import ExportTicketWindow  # noqa: E402


def _assert_invalidated(window: ExportTicketWindow) -> None:
    assert window.last_schedule_rows == []
    assert window.last_analysis is None
    assert window.last_max_finish == 0.0
    assert window._last_model_result_summary is None
    assert window._analysis_result_stale is True
    assert window.sim_time == 0.0
    assert not window.sim_timer.isActive()
    assert not window.btn_export_analysis_report.isEnabled()
    assert not window.btn_sim_play.isEnabled()
    assert not window.btn_sim_pause.isEnabled()
    assert not window.btn_sim_reset.isEnabled()
    assert not window.btn_vehicle_log.isEnabled()
    assert "请重新点击" in window.lbl_vehicle_summary.text()


def main() -> int:
    app = QApplication.instance() or QApplication([])
    window = ExportTicketWindow()
    window.fill_sample()
    window.spn_target_takt.setValue(118)
    window.do_analyze()

    original_finish = window.last_max_finish
    original_wait = window._last_model_result_summary["total_actual_wait"]
    assert original_finish == 1082.0
    assert original_wait == 82.0
    assert window.btn_vehicle_log.isEnabled()

    window.sim_timer.start()
    window.spn_target_takt.setValue(58)
    _assert_invalidated(window)

    window.do_analyze()
    refreshed_finish = window.last_max_finish
    refreshed_summary = dict(window._last_model_result_summary or {})
    assert refreshed_finish == 940.0
    assert refreshed_summary.get("target_takt") == 58.0
    assert refreshed_summary.get("overall_takt") == 146.0
    assert refreshed_summary.get("total_actual_wait") == 510.0
    assert refreshed_summary.get("total_excess_wait") == 98.0
    assert window._analysis_result_stale is False
    assert window.btn_export_analysis_report.isEnabled()
    assert window.btn_sim_play.isEnabled()
    assert window.btn_vehicle_log.isEnabled()

    window.tbl.item(0, 5).setText("101")
    _assert_invalidated(window)

    result = {
        "old_finish": original_finish,
        "old_actual_wait": original_wait,
        "refreshed_finish": refreshed_finish,
        "refreshed_overall_takt": refreshed_summary.get("overall_takt"),
        "refreshed_actual_wait": refreshed_summary.get("total_actual_wait"),
        "refreshed_excess_wait": refreshed_summary.get("total_excess_wait"),
        "table_edit_invalidated": True,
        "assertions": "passed",
    }
    print(result)
    window.close()
    app.processEvents()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
