#!/usr/bin/env python3
"""Verify round-12 one-vehicle-per-row log export with real data."""

from __future__ import annotations

import csv
import importlib.util
import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from PySide6.QtWidgets import QApplication, QFileDialog  # noqa: E402
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
        baseline.station_defs(), 1148,
        vehicle_counts={"A": 276, "B": 306, "C": 566},
        sequence_mode="alternate", max_consecutive=5, launch_takt=58.0,
    )
    analysis = tickets.analyze_schedule(rows, max_finish, 58.0)

    app = QApplication.instance() or QApplication([])
    window = ExportTicketWindow()
    window.last_schedule_rows = rows
    window.last_analysis = analysis
    window.last_max_finish = max_finish
    window.sim_time = max_finish
    window.spn_target_takt.setValue(58.0)

    columns, log_rows = window._build_vehicle_log_rows()
    car_11 = window._filter_vehicle_log_rows(log_rows, "Car#11")
    car_12 = window._filter_vehicle_log_rows(log_rows, "12")
    last_car = window._filter_vehicle_log_rows(log_rows, "1148")
    actual_wait = window._filter_vehicle_log_rows(log_rows, wait_filter="有实际等待")
    excess_wait = window._filter_vehicle_log_rows(log_rows, wait_filter="有节拍外等待")

    assert len(log_rows) == 17_220
    assert len(car_11) == 15 and {row[0] for row in car_11} == {"11"}
    assert len(car_12) == 15 and {row[0] for row in car_12} == {"12"}
    assert len(last_car) == 15 and {row[0] for row in last_car} == {"1148"}
    assert len(actual_wait) == 6_736
    assert len(excess_wait) == 255
    assert {row[3] for row in excess_wait} == {"L2++", "电检1", "空气悬挂+快充"}

    car_11_ids = {row[0] for row in car_11}
    car_11_source_rows = [row for row in rows if str(row.get("car", "")) in car_11_ids]
    car_11_text = window._build_schedule_debug_log(car_11_source_rows, limit=9999)
    assert "Car#11" in car_11_text
    assert "Car#111 " not in car_11_text
    assert len([line for line in car_11_text.splitlines() if line.startswith("Car#")]) == 1
    assert "SEGMENTS" in car_11_text
    assert "开:" in car_11_text and "加:" in car_11_text
    assert "等前:" in car_11_text and "等后:" in car_11_text

    selected_car_ids = {row[0] for row in car_11}
    selected_schedule_rows = [row for row in rows if str(row.get("car", "")) in selected_car_ids]
    compact_columns, compact_rows = window._build_compact_vehicle_log_csv_rows(selected_schedule_rows)
    all_compact_columns, all_compact_rows = window._build_compact_vehicle_log_csv_rows(rows)
    assert compact_columns == ["CAR", "TYPE", "IN(s)", "OUT(s)", "WAIT(s)", "FLOW(s)", "能力判断", "SEGMENTS"]
    assert all_compact_columns == compact_columns and len(all_compact_rows) == 1_148
    assert len(compact_rows) == 1 and compact_rows[0][0] == "Car#11"
    assert compact_rows[0][7].count("ST") == 15

    export_path = Path(tempfile.gettempdir()) / "mline_round12_vehicle_log.csv"
    original_dialog = QFileDialog.getSaveFileName
    QFileDialog.getSaveFileName = lambda *args, **kwargs: (str(export_path), "CSV (*.csv)")
    try:
        window._export_vehicle_log_csv(compact_columns, compact_rows, window)
    finally:
        QFileDialog.getSaveFileName = original_dialog
    with export_path.open("r", encoding="utf-8-sig", newline="") as csv_file:
        exported = list(csv.reader(csv_file))
    export_path.unlink(missing_ok=True)
    assert exported[0] == compact_columns and exported[1:] == compact_rows

    print({
        "all_segment_rows": len(log_rows), "car_11_rows": len(car_11),
        "car_12_rows": len(car_12), "last_car_rows": len(last_car),
        "actual_wait_rows": len(actual_wait), "excess_wait_rows": len(excess_wait),
        "excess_wait_stations": sorted({row[3] for row in excess_wait}),
        "csv_matches_filter": exported[1:] == compact_rows,
        "csv_vehicle_rows": len(compact_rows),
        "all_csv_vehicle_rows": len(all_compact_rows),
        "compact_log_vehicle_lines": len([line for line in car_11_text.splitlines() if line.startswith("Car#")]),
    })
    window.close()


if __name__ == "__main__":
    main()
