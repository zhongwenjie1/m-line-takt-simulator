#!/usr/bin/env python3
"""Verify quantity/ratio result scope wording uses one authoritative source."""

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
from utils.result_scope_text import build_result_scope_text  # noqa: E402


def _verify_pure_wording() -> None:
    quantity = build_result_scope_text(
        is_ratio_mode=False,
        output_count=6,
        analysis_time_seconds=3600,
        last_output_car_no=6,
        last_output_car_out=1082.0,
    )
    quantity_text = " ".join(str(value) for value in quantity.values())
    assert quantity["mode"] == "quantity"
    assert quantity["analysis_time_minutes"] == 0.0
    assert quantity["analysis_time_seconds"] == 0.0
    assert quantity["title"] == "模型结果（目标批次终值：6台）"
    assert quantity["note"] == "统计范围：按目标批次共6台统计"
    assert "Car#6" in quantity["vehicle_current"]
    assert "1082s" in quantity["vehicle_current"]
    assert "分析时间" not in quantity_text
    assert "分钟" not in quantity_text

    ratio = build_result_scope_text(
        is_ratio_mode=True,
        output_count=1024,
        analysis_time_seconds=69000,
        last_output_car_no=1024,
        last_output_car_out=68996.5,
    )
    assert ratio["mode"] == "ratio"
    assert ratio["analysis_time_minutes"] == 1150.0
    assert ratio["analysis_time_seconds"] == 69000.0
    assert ratio["title"] == "模型结果（分析窗口终值：1150分钟）"
    assert ratio["note"] == "统计范围：按分析时间1150分钟（69000s）统计"
    assert "Car#1024" in ratio["vehicle_current"]
    assert "68996.5s ≤ 69000s" in ratio["vehicle_current"]

    empty_ratio = build_result_scope_text(
        is_ratio_mode=True,
        output_count=0,
        analysis_time_seconds=60,
    )
    assert "当前分析时间内" in empty_ratio["vehicle_current"]
    assert "窗口1分钟，60s" in empty_ratio["vehicle_current"]
    assert "Car#" not in empty_ratio["vehicle_current"]


def _verify_ui_integration() -> dict:
    app = QApplication.instance() or QApplication([])
    window = ExportTicketWindow()
    window.fill_sample()
    window.spn_total_cars.setValue(60)
    window.spn_target_takt.setValue(118)
    window.do_analyze()

    summary = dict(window._last_model_result_summary or {})
    scope = dict(summary.get("result_scope_text", {}) or {})
    explanation = window._build_model_result_explanation_text(summary)
    summary_html = window.lbl_vehicle_summary.text()

    assert summary.get("is_ratio_mode") is False
    assert summary.get("analysis_time_minutes") == 0.0
    assert summary.get("analysis_time_seconds") == 0.0
    assert scope.get("mode") == "quantity"
    assert scope.get("title") in summary_html
    assert scope.get("note") in summary_html
    assert scope.get("vehicle_definition") in explanation
    assert scope.get("vehicle_rule") in explanation
    assert scope.get("vehicle_current") in explanation
    assert "目标批次" in explanation
    assert "分析时间60分钟" not in explanation
    assert "分析时间3600" not in explanation

    quantity_result = {
        "mode": scope.get("mode"),
        "title": scope.get("title"),
        "note": scope.get("note"),
        "output_count": summary.get("output_count"),
        "analysis_time_seconds": summary.get("analysis_time_seconds"),
    }

    window.cmb_launch_mode.setCurrentText("按比例投车")
    window.spn_a_cars.setValue(4)
    window.spn_b_cars.setValue(2)
    window.spn_c_cars.setValue(0)
    window.spn_total_cars.setValue(1)
    window.spn_target_takt.setValue(58)
    window.do_analyze()

    ratio_summary = dict(window._last_model_result_summary or {})
    ratio_scope = dict(ratio_summary.get("result_scope_text", {}) or {})
    ratio_explanation = window._build_model_result_explanation_text(ratio_summary)
    ratio_summary_html = window.lbl_vehicle_summary.text()

    assert ratio_summary.get("is_ratio_mode") is True
    assert ratio_summary.get("analysis_time_minutes") == 1.0
    assert ratio_summary.get("analysis_time_seconds") == 60.0
    assert ratio_scope.get("mode") == "ratio"
    assert ratio_scope.get("title") in ratio_summary_html
    assert ratio_scope.get("note") in ratio_summary_html
    assert ratio_scope.get("vehicle_definition") in ratio_explanation
    assert ratio_scope.get("vehicle_rule") in ratio_explanation
    assert ratio_scope.get("vehicle_current") in ratio_explanation
    assert "当前分析时间内暂无车辆完成下线" in ratio_explanation
    assert "目标批次终值" not in ratio_summary_html

    result = {
        "quantity": quantity_result,
        "ratio": {
            "mode": ratio_scope.get("mode"),
            "title": ratio_scope.get("title"),
            "note": ratio_scope.get("note"),
            "output_count": ratio_summary.get("output_count"),
            "analysis_time_seconds": ratio_summary.get("analysis_time_seconds"),
        },
        "assertions": "passed",
    }
    window.close()
    app.processEvents()
    return result


def main() -> int:
    _verify_pure_wording()
    result = _verify_ui_integration()
    print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
