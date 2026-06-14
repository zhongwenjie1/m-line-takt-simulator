#!/usr/bin/env python3
"""Verify round-11 target-batch and simulation-buffer boundaries."""

from __future__ import annotations

import importlib.util
import math
import sys
from pathlib import Path


APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from core import tickets  # noqa: E402
from core.analysis import apply_time_window_analysis  # noqa: E402
from core.input_parser import parse_multi_project_inputs  # noqa: E402


def _load_baseline():
    path = APP_DIR / "scripts" / "verify_v29_real_data_baseline.py"
    spec = importlib.util.spec_from_file_location("v29_baseline", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载基线脚本：{path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _raw_station_rows(definitions):
    return [
        {
            "seq": item["seq"],
            "display": item["display"],
            "device_count": item["device_count"],
            "line_scope": item["line_scope"],
            "group": item["group"],
            "duration_a": item["duration_a"],
            "duration_b": item["duration_b"],
            "duration_c": item["duration_c"],
        }
        for item in definitions
    ]


def main() -> None:
    baseline = _load_baseline()
    parsed = parse_multi_project_inputs({
        "project": "v2.9真实比例场景",
        "cars_a": 2,
        "cars_b": 3,
        "cars_c": 5,
        "analysis_minutes": 1150,
        "target_takt": 58,
        "is_ratio_mode": True,
        "max_consecutive": 5,
        "station_rows": _raw_station_rows(baseline.station_defs()),
    })

    assert parsed["theoretical_launch_count"] == 1190
    assert parsed["simulation_buffer_count"] == 50
    assert parsed["simulation_vehicle_count"] == 1240
    assert parsed["cars"] == 1240

    rows, max_finish = tickets.schedule(
        parsed["defs"],
        parsed["cars"],
        parsed["vehicle_counts"],
        parsed["sequence_mode"],
        parsed["max_consecutive"],
        parsed["ratio_pattern"],
        launch_takt=parsed["target_takt"],
    )
    analysis = tickets.analyze_schedule(rows, max_finish, parsed["target_takt"])
    analysis = apply_time_window_analysis(
        analysis,
        rows,
        parsed["target_takt"],
        parsed["analysis_time_seconds"],
        parsed["theoretical_launch_count"],
        parsed["simulation_buffer_count"],
    )
    summary = analysis["summary"]

    assert summary["theoretical_launch_count"] == 1190
    assert summary["simulation_buffer_count"] == 50
    assert summary["simulation_vehicle_count"] == 1240
    assert math.isclose(summary["target_batch_time"], 69823.0)
    assert math.isclose(summary["actual_finish_time"], 80010.5)
    assert math.isclose(summary["target_batch_finish_delta"], 10187.5)
    assert math.isclose(summary["batch_overrun_time"], 10187.5)
    assert math.isclose(summary["batch_early_time"], 0.0)
    assert summary["batch_overrun_result"] == "延后完成"
    assert math.isclose(max_finish, 83308.0)
    assert summary["actual_finish_time"] < max_finish
    assert summary["actual_output_count_in_window"] == 1024

    early_rows = [
        {"car": 1, "step_seq": 1, "step_display": "ST1", "start": 0.0, "dur": 100.0, "svc_finish": 100.0, "depart": 100.0},
        {"car": 2, "step_seq": 1, "step_display": "ST1", "start": 58.0, "dur": 40.0, "svc_finish": 98.0, "depart": 98.0},
    ]
    early_analysis = apply_time_window_analysis(
        {"summary": {}}, early_rows, 58.0, 116.0, 2, 0
    )
    early_summary = early_analysis["summary"]
    assert math.isclose(early_summary["target_batch_time"], 158.0)
    assert math.isclose(early_summary["actual_finish_time"], 100.0)
    assert math.isclose(early_summary["target_batch_finish_delta"], -58.0)
    assert math.isclose(early_summary["batch_overrun_time"], 0.0)
    assert math.isclose(early_summary["batch_early_time"], 58.0)
    assert early_summary["batch_overrun_result"] == "提前完成"

    print({
        "theoretical_launch_count": summary["theoretical_launch_count"],
        "simulation_buffer_count": summary["simulation_buffer_count"],
        "target_batch_time": summary["target_batch_time"],
        "target_batch_actual_finish": summary["actual_finish_time"],
        "target_batch_finish_delta": summary["target_batch_finish_delta"],
        "batch_overrun_time": summary["batch_overrun_time"],
        "batch_early_time": summary["batch_early_time"],
        "window_output_count": summary["actual_output_count_in_window"],
        "early_finish_case": {
            "planned_finish": early_summary["target_batch_time"],
            "actual_finish": early_summary["actual_finish_time"],
            "finish_delta": early_summary["target_batch_finish_delta"],
            "result": early_summary["batch_overrun_result"],
        },
    })


if __name__ == "__main__":
    main()
