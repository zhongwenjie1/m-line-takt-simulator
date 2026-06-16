#!/usr/bin/env python3
"""Verify shared-equipment animation is rendered as one slot without touching scheduling."""

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


def _inject_two_station_case(window: ExportTicketWindow) -> None:
    window.current_defs = [
        {
            "seq": 1,
            "group": "引取",
            "display": "引取",
            "device_count": 1,
            "line_scope": "双线共用",
            "run_mode": "双线单设备",
        },
        {
            "seq": 2,
            "group": "集约",
            "display": "集约",
            "device_count": 2,
            "line_scope": "双线",
            "run_mode": "双线双设备",
        },
    ]
    window.last_schedule_rows = [
        {
            "car": 1,
            "car_type": "A",
            "step_seq": 1,
            "step_display": "引取",
            "group": "引取",
            "line_no": "1号线",
            "line_scope": "双线共用",
            "resource_key": "引取::双线共用",
            "start": 0.0,
            "dur": 56.0,
            "svc_finish": 56.0,
            "depart": 56.0,
        },
        {
            "car": 1,
            "car_type": "A",
            "step_seq": 2,
            "step_display": "集约",
            "group": "集约",
            "line_no": "1号线",
            "line_scope": "双线",
            "resource_key": "集约::1号线",
            "start": 56.0,
            "dur": 100.0,
            "svc_finish": 156.0,
            "depart": 156.0,
        },
        {
            "car": 2,
            "car_type": "A",
            "step_seq": 1,
            "step_display": "引取",
            "group": "引取",
            "line_no": "2号线",
            "line_scope": "双线共用",
            "resource_key": "引取::双线共用",
            "start": 56.0,
            "dur": 56.0,
            "svc_finish": 112.0,
            "depart": 112.0,
        },
        {
            "car": 2,
            "car_type": "A",
            "step_seq": 2,
            "step_display": "集约",
            "group": "集约",
            "line_no": "2号线",
            "line_scope": "双线",
            "resource_key": "集约::2号线",
            "start": 112.0,
            "dur": 100.0,
            "svc_finish": 212.0,
            "depart": 212.0,
        },
    ]


def _scene_texts(window: ExportTicketWindow) -> list[str]:
    return [
        item.toPlainText()
        for item in window.sim_scene.items()
        if hasattr(item, "toPlainText")
    ]


def main() -> int:
    app = QApplication.instance() or QApplication([])
    window = ExportTicketWindow()
    _inject_two_station_case(window)
    window.sim_time = 60.0
    window._draw_sim_scene()
    texts = _scene_texts(window)
    assert "共用设备" in texts
    assert any("→" in text for text in texts), texts
    assert sum(1 for text in texts if text == "空闲") >= 3
    assert sum(1 for text in texts if text == "不适用") == 0
    assert window.last_schedule_rows
    shared_rows = [
        row for row in window.last_schedule_rows
        if str(row.get("line_scope", "")) == "双线共用"
    ]
    assert shared_rows
    assert len({row.get("resource_key") for row in shared_rows}) == 1
    print({
        "shared_badge": "共用设备" in texts,
        "flow_direction_label": any("→" in text for text in texts),
        "shared_resource_keys": sorted({row.get("resource_key") for row in shared_rows}),
        "assertions": "passed",
    })
    window.close()
    app.processEvents()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
