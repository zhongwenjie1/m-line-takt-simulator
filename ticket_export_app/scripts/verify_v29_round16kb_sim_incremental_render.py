#!/usr/bin/env python3
"""Verify round-16K-B simulation incremental rendering boundaries."""

from __future__ import annotations

import importlib.util
import os
import sys
from pathlib import Path


os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from PySide6.QtWidgets import QApplication  # noqa: E402
from core import tickets  # noqa: E402
from ui.export_ticket_window import ExportTicketWindow  # noqa: E402


def _load_baseline():
    path = APP_DIR / "scripts" / "verify_v29_real_data_baseline.py"
    spec = importlib.util.spec_from_file_location("v29_baseline", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载基线脚本：{path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main() -> None:
    baseline = _load_baseline()
    rows, max_finish = tickets.schedule(
        baseline.station_defs(),
        1148,
        vehicle_counts={"A": 276, "B": 306, "C": 566},
        sequence_mode="alternate",
        max_consecutive=5,
        launch_takt=58.0,
    )
    analysis = tickets.analyze_schedule(rows, max_finish, 58.0)

    app = QApplication.instance() or QApplication([])
    window = ExportTicketWindow()
    window.current_defs = baseline.station_defs()
    window.last_schedule_rows = rows
    window.last_analysis = analysis
    window.last_max_finish = max_finish
    window.sim_time = 0.0

    assert not window.btn_export_frozen_sequence.isVisible()
    assert not window.btn_import_frozen_sequence.isVisible()

    window._draw_sim_scene()
    initial_signature = window._sim_static_signature
    initial_scene_items = len(window.sim_scene.items())
    assert initial_signature is not None
    assert initial_scene_items > 0
    static_item_count = initial_scene_items - len(window._sim_dynamic_items)

    dynamic_counts = []
    scene_counts = []
    for current in (0.0, 1340.0, 1830.0, 3600.0, 76152.5):
        window.sim_time = current
        window._update_sim_view()
        window._draw_sim_scene()
        assert window._sim_static_signature == initial_signature
        dynamic_counts.append(len(window._sim_dynamic_items))
        scene_counts.append(len(window.sim_scene.items()))
        assert len(window.sim_scene.items()) == static_item_count + len(window._sim_dynamic_items)

    assert any(count > 0 for count in dynamic_counts)
    assert window._sim_event_cache is not None
    assert window._sim_car_rows_cache is not None

    window.close()
    app.processEvents()
    print({
        "rows": len(rows),
        "static_signature_stable": True,
        "dynamic_counts": dynamic_counts,
        "scene_count_range": [min(scene_counts), max(scene_counts)],
        "freeze_io_buttons_hidden": True,
        "assertions": "passed",
    })


if __name__ == "__main__":
    main()
