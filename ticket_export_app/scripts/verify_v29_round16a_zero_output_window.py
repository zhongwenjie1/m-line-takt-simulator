#!/usr/bin/env python3
"""Verify that a ratio analysis window with no outputs remains an empty scope."""

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


def main() -> int:
    app = QApplication.instance() or QApplication([])
    window = ExportTicketWindow()
    window.fill_sample()
    window.cmb_launch_mode.setCurrentText("按比例投车")
    window.spn_a_cars.setValue(4)
    window.spn_b_cars.setValue(2)
    window.spn_c_cars.setValue(0)
    window.spn_total_cars.setValue(1)
    window.spn_target_takt.setValue(58)
    window.do_analyze()

    summary = dict(window._last_model_result_summary or {})
    core_summary = dict((window.last_analysis or {}).get("summary", {}) or {})
    explanation = window._build_model_result_explanation_text(summary)

    assert core_summary.get("actual_output_count_in_window") == 0
    assert summary.get("output_count") == 0
    assert summary.get("qualified_count") == 0
    assert summary.get("qualified_rate") is None
    assert summary.get("overall_takt") is None
    assert summary.get("first_out") is None
    assert summary.get("last_out") is None
    assert summary.get("last_output_car_no") is None
    assert "当前分析时间内暂无车辆完成下线" in explanation
    assert "第52台车辆" not in explanation

    result = {
        "analysis_time_seconds": summary.get("analysis_time_seconds"),
        "core_output_count": core_summary.get("actual_output_count_in_window"),
        "ui_output_count": summary.get("output_count"),
        "qualified_count": summary.get("qualified_count"),
        "overall_takt": summary.get("overall_takt"),
        "generated_simulation_vehicles": len(
            window._build_realtime_model_result(0.0).get("all_vehicles", [])
        ),
        "assertions": "passed",
    }
    print(result)
    window.close()
    app.processEvents()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
