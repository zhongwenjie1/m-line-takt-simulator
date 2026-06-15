#!/usr/bin/env python3
"""Verify the round-16G four-card result summary and compact layout."""

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


VIEWPORTS = ((1080, 700), (1366, 768), (1440, 900))


def _assert_four_card_summary(window: ExportTicketWindow) -> None:
    html = window.lbl_vehicle_summary.text()
    report_payload = window._build_analysis_report_payload()

    assert html.count("<td width='25%'") == 4
    assert "下线车辆" in html
    assert "批次完成时刻" in html
    assert "整体节拍" in html
    assert "累计节拍外等待" in html
    assert "累计实际等待" in html
    assert "达标车辆" not in html
    assert "达标率" not in html
    assert "等待真因：" not in html
    assert "风险提示：" not in html

    assert report_payload.get("completion_label") == "批次总完成时刻"
    assert report_payload.get("completion_time") == 1082.0
    assert not hasattr(window, "btn_model_result_explanation")
    assert window.btn_export_analysis_report.text() == "导出分析报告"


def _capture_viewports(window: ExportTicketWindow) -> list[dict]:
    app = QApplication.instance()
    captures = []
    window.multi_tabs.setCurrentWidget(window.page_multi_result_scroll)
    for width, height in VIEWPORTS:
        window.resize(width, height)
        window.show()
        app.processEvents()
        window._position_analysis_report_button()
        app.processEvents()

        label = window.lbl_vehicle_summary
        button = window.btn_export_analysis_report
        assert label.minimumHeight() == 160
        assert label.maximumHeight() == 205
        assert button.x() >= 0
        assert button.y() >= 0
        assert button.x() + button.width() <= label.width()
        assert button.y() + button.height() <= label.height()
        assert button.y() + button.height() <= 32

        output = Path("/tmp") / f"v29_round16g_{width}x{height}.png"
        label.grab().save(str(output))
        captures.append(
            {
                "viewport": f"{width}x{height}",
                "summary_size": [label.width(), label.height()],
                "button_geometry": [button.x(), button.y(), button.width(), button.height()],
                "vertical_scroll_max": window.page_multi_result_scroll.verticalScrollBar().maximum(),
                "screenshot": str(output),
            }
        )
    return captures


def _assert_ratio_labels(window: ExportTicketWindow) -> dict:
    window.cmb_launch_mode.setCurrentText("按比例投车")
    window.spn_a_cars.setValue(4)
    window.spn_b_cars.setValue(2)
    window.spn_c_cars.setValue(0)
    window.spn_total_cars.setValue(1)
    window.spn_target_takt.setValue(58)
    window.do_analyze()

    html = window.lbl_vehicle_summary.text()
    report_payload = window._build_analysis_report_payload()
    assert html.count("<td width='25%'") == 4
    assert "窗口内下线" in html
    assert "窗口末台下线" in html
    assert report_payload.get("completion_label") == "窗口末台下线"
    assert "批次完成时刻" not in html
    assert "达标车辆" not in html
    assert "达标率" not in html
    return {
        "output_count": window._last_model_result_summary.get("output_count"),
        "completion_time": window._last_model_result_summary.get("completion_time"),
    }


def main() -> int:
    app = QApplication.instance() or QApplication([])
    window = ExportTicketWindow()
    window.fill_sample()
    window.spn_target_takt.setValue(118)
    window.do_analyze()

    _assert_four_card_summary(window)
    captures = _capture_viewports(window)
    ratio = _assert_ratio_labels(window)

    print(
        {
            "quantity_cards": 4,
            "quantity_completion_time": 1082.0,
            "ratio": ratio,
            "viewports": captures,
            "assertions": "passed",
        }
    )
    window.close()
    app.processEvents()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
