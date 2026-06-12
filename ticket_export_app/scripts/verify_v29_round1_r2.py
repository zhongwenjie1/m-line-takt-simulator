#!/usr/bin/env python3
"""Verify the v2.9 round-1 physical station-slot capacity change."""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path
from typing import Any


APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from core import tickets  # noqa: E402


def _load_baseline_module():
    path = APP_DIR / "scripts" / "verify_v29_real_data_baseline.py"
    spec = importlib.util.spec_from_file_location("v29_baseline", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载基线脚本：{path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


BASELINE = _load_baseline_module()


def _chain_checks(rows: list[dict[str, Any]]) -> dict[str, int]:
    overlap = 0
    gap = 0
    finish_errors = 0
    wait_errors = 0
    for segments in BASELINE.group_rows(rows).values():
        for segment in segments:
            if abs(float(segment["start"]) + float(segment["dur"]) - float(segment["svc_finish"])) > 1e-9:
                finish_errors += 1
            if abs(float(segment["depart"]) - float(segment["svc_finish"]) - float(segment["block_wait"])) > 1e-9:
                wait_errors += 1
        for previous, current in zip(segments, segments[1:]):
            delta = float(current["start"]) - float(previous["depart"])
            if delta < -1e-9:
                overlap += 1
            elif delta > 1e-9:
                gap += 1
    return {
        "overlap": overlap,
        "gap": gap,
        "finish_errors": finish_errors,
        "wait_errors": wait_errors,
    }


def _minimal_backpressure() -> dict[str, Any]:
    definitions = [
        {
            "seq": 1,
            "display": "A工位",
            "group": "A工位",
            "duration_a": 30,
            "duration_b": 30,
            "duration_c": 30,
            "capacity": 1,
            "device_count": 1,
            "line_scope": "1号线",
            "run_mode": "单线单设备",
        },
        {
            "seq": 2,
            "display": "零工时点",
            "group": "零工时点",
            "duration_a": 0,
            "duration_b": 0,
            "duration_c": 0,
            "capacity": 1,
            "device_count": 1,
            "line_scope": "1号线",
            "run_mode": "单线单设备",
        },
        {
            "seq": 3,
            "display": "B瓶颈",
            "group": "B瓶颈",
            "duration_a": 100,
            "duration_b": 100,
            "duration_c": 100,
            "capacity": 1,
            "device_count": 1,
            "line_scope": "1号线",
            "run_mode": "单线单设备",
        },
    ]
    rows, _ = tickets.schedule(
        definitions,
        4,
        vehicle_counts={"A": 4, "B": 0, "C": 0},
        sequence_mode="grouped",
        launch_takt=0,
    )
    grouped = BASELINE.group_rows(rows)
    car4_first_start = float(grouped[4][0]["start"])
    if car4_first_start != 130.0:
        raise AssertionError(f"最小回堵场景Car#4开始时间应为130，实际为{car4_first_start}")
    overlaps = BASELINE.station_slot_overlap_summary(rows)
    if overlaps["pair_count"] != 0:
        raise AssertionError(f"最小回堵场景仍有同槽重叠：{overlaps}")
    return {
        "car4_first_station_start": car4_first_start,
        "station_slot_overlaps": overlaps,
        "chain_checks": _chain_checks(rows),
    }


def _run_real(name: str) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    definitions = BASELINE.station_defs()
    if name == "ratio":
        rows, _ = tickets.schedule(
            definitions,
            1240,
            vehicle_counts={"A": 248, "B": 372, "C": 620},
            sequence_mode="ratio",
            ratio_pattern={"A": 2, "B": 3, "C": 5},
            launch_takt=BASELINE.TARGET_TAKT,
        )
    else:
        rows, _ = tickets.schedule(
            definitions,
            1148,
            vehicle_counts={"A": 276, "B": 306, "C": 566},
            sequence_mode="alternate",
            max_consecutive=5,
            launch_takt=BASELINE.TARGET_TAKT,
        )
    return rows, BASELINE.summarize(rows)


def _scenario_result(name: str, old: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    rows, summary = _run_real(name)
    overlaps = summary["station_slot_overlaps"]
    occupancy = summary["station_slot_occupancy"]
    chain = _chain_checks(rows)
    if overlaps["pair_count"] != 0 or occupancy["maximum"] != 1:
        raise AssertionError(f"{name}场景物理车位仍重叠：{overlaps}, {occupancy}")
    if summary["processing_capacity_violations"] != 0:
        raise AssertionError(f"{name}场景正加工能力超限")
    if any(chain.values()):
        raise AssertionError(f"{name}场景车辆时间链错误：{chain}")

    fields = (
        "within_window",
        "qualified",
        "first_out",
        "last_out",
        "overall_takt",
        "raw_block_wait_in_window",
        "full_last_out",
        "full_raw_block_wait",
    )
    result = {field: summary[field] for field in fields}
    result["delta_from_round0"] = {field: summary[field] - old[field] for field in fields}
    result["processing_capacity_violations"] = summary["processing_capacity_violations"]
    result["station_slot_occupancy"] = occupancy
    result["station_slot_overlaps"] = overlaps
    result["chain_checks"] = chain
    result["segment_count"] = len(rows)
    result["zero_duration_segment_count"] = sum(abs(float(row["dur"])) <= 1e-9 for row in rows)
    return result, rows


def _screenshot_chain(rows: list[dict[str, Any]]) -> dict[str, Any]:
    grouped = BASELINE.group_rows(rows)
    result: dict[str, Any] = {}
    for car in (12, 14, 16):
        selected = []
        for segment in grouped[car]:
            if segment["step_display"] in ("ADAS", "ARHUD+慢充"):
                selected.append(
                    {
                        "station": segment["step_display"],
                        "line": segment["line_no"],
                        "start": segment["start"],
                        "svc_finish": segment["svc_finish"],
                        "depart": segment["depart"],
                        "duration": segment["dur"],
                        "block_wait": segment["block_wait"],
                    }
                )
        result[str(car)] = selected
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline", type=Path, required=True)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    old = json.loads(args.baseline.read_text(encoding="utf-8"))
    ratio, ratio_rows = _scenario_result("ratio", old["ratio"])
    quantity, _ = _scenario_result("quantity", old["quantity"])
    result = {
        "round": "v2.9-round1-r2",
        "minimal_backpressure": _minimal_backpressure(),
        "ratio": ratio,
        "quantity": quantity,
        "screenshot_chain_after_fix": _screenshot_chain(ratio_rows),
        "assertions": "passed",
    }
    output = json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(output, encoding="utf-8")
    sys.stdout.write(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
