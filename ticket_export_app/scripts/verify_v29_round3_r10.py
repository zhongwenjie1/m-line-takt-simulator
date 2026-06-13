#!/usr/bin/env python3
"""Verify R10 arrival-first service with the v2.9 real-data models."""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from collections import Counter
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
EXPECTED_QUANTITY_SEQUENCE_HASH = (
    "dac6086f3b896255eeb77cb82679b0334e49ab972a5e9ed6aa0cd4e6d347f08f"
)


def _run_real(name: str) -> tuple[list[dict[str, Any]], list[str]]:
    definitions = BASELINE.station_defs()
    if name == "quantity":
        counts = {"A": 276, "B": 306, "C": 566}
        sequence = tickets.build_vehicle_sequence(
            sum(counts.values()),
            counts,
            sequence_mode="alternate",
            max_consecutive=5,
        )
        rows, _ = tickets.schedule(
            definitions,
            len(sequence),
            vehicle_counts=counts,
            sequence_mode="alternate",
            max_consecutive=5,
            launch_takt=BASELINE.TARGET_TAKT,
            vehicle_sequence=sequence,
        )
    else:
        counts = {"A": 248, "B": 372, "C": 620}
        sequence = tickets.build_vehicle_sequence(
            sum(counts.values()),
            counts,
            sequence_mode="ratio",
            ratio_pattern={"A": 2, "B": 3, "C": 5},
        )
        rows, _ = tickets.schedule(
            definitions,
            len(sequence),
            vehicle_counts=counts,
            sequence_mode="ratio",
            ratio_pattern={"A": 2, "B": 3, "C": 5},
            launch_takt=BASELINE.TARGET_TAKT,
            vehicle_sequence=sequence,
        )
    return rows, sequence


def _arrival_time(segments: list[dict[str, Any]], index: int) -> float:
    if index == 0:
        return float(segments[index].get("theory_launch_time", 0.0))
    return float(segments[index - 1]["svc_finish"])


def _service_checks(rows: list[dict[str, Any]]) -> dict[str, Any]:
    grouped = BASELINE.group_rows(rows)
    definitions = BASELINE.station_defs()
    first_station = sorted(
        (float(segments[0]["start"]), car)
        for car, segments in grouped.items()
    )
    launch_overtakes = sum(
        next_car < previous_car
        for (_, previous_car), (_, next_car) in zip(first_station, first_station[1:])
    )

    shared: dict[str, Any] = {}
    total_fcfs_violations = 0
    for index, definition in enumerate(definitions):
        if definition["line_scope"] not in ("双线共用", "1号线", "2号线"):
            continue
        services = []
        for car, segments in grouped.items():
            segment = segments[index]
            services.append(
                {
                    "car": car,
                    "arrival": _arrival_time(segments, index),
                    "start": float(segment["start"]),
                    "line": str(segment["line_no"]),
                }
            )
        services.sort(key=lambda item: (item["start"], item["car"]))
        violations = sum(
            (current["arrival"], current["car"])
            < (previous["arrival"], previous["car"])
            for previous, current in zip(services, services[1:])
        )
        total_fcfs_violations += violations

        cumulative = Counter()
        maximum_line_difference = 0
        maximum_same_line_run = 0
        current_run = 0
        previous_line = ""
        waits = {"1号线": [], "2号线": []}
        for service in services:
            line = service["line"]
            cumulative[line] += 1
            maximum_line_difference = max(
                maximum_line_difference,
                abs(cumulative["1号线"] - cumulative["2号线"]),
            )
            current_run = current_run + 1 if line == previous_line else 1
            previous_line = line
            maximum_same_line_run = max(maximum_same_line_run, current_run)
            if line in waits:
                waits[line].append(service["start"] - service["arrival"])

        shared[definition["display"]] = {
            "fcfs_violations": violations,
            "line_counts": dict(Counter(item["line"] for item in services)),
            "maximum_cumulative_line_difference": maximum_line_difference,
            "maximum_same_line_run": maximum_same_line_run,
            "average_wait_by_line": {
                line: sum(values) / len(values) if values else 0.0
                for line, values in waits.items()
            },
            "maximum_wait_by_line": {
                line: max(values, default=0.0)
                for line, values in waits.items()
            },
        }

    return {
        "launch_overtakes": launch_overtakes,
        "fcfs_violations": total_fcfs_violations,
        "shared_or_single_stations": shared,
    }


def _minimal_arrival_first() -> dict[str, Any]:
    definitions = [
        {
            "seq": 1,
            "display": "长工位",
            "group": "长工位",
            "duration_a": 100,
            "duration_b": -1,
            "duration_c": -1,
            "capacity": 2,
            "device_count": 2,
            "line_scope": "双线",
            "run_mode": "双线双设备",
        },
        {
            "seq": 2,
            "display": "末端共用岗位",
            "group": "末端共用岗位",
            "duration_a": 10,
            "duration_b": 10,
            "duration_c": 10,
            "capacity": 1,
            "device_count": 1,
            "line_scope": "双线共用",
            "run_mode": "双线单设备",
        },
    ]
    rows, _ = tickets.schedule(
        definitions,
        2,
        vehicle_counts={"A": 1, "B": 1, "C": 0},
        launch_takt=0,
        vehicle_sequence=["A", "B"],
    )
    common = [row for row in rows if row["step_display"] == "末端共用岗位"]
    common.sort(key=lambda row: (float(row["start"]), int(row["car"])))
    order = [int(row["car"]) for row in common]
    if order != [2, 1]:
        raise AssertionError(f"后到车辆仍先服务：{order}")
    return {
        "service_order": order,
        "car2_start": float(common[0]["start"]),
        "car1_start": float(common[1]["start"]),
    }


def _scenario(name: str) -> dict[str, Any]:
    rows, sequence = _run_real(name)
    summary = BASELINE.summarize(rows)
    checks = _service_checks(rows)
    overlaps = summary["station_slot_overlaps"]
    if checks["launch_overtakes"] != 0:
        raise AssertionError(f"{name}首工位发生越车")
    if checks["fcfs_violations"] != 0:
        raise AssertionError(f"{name}共用/单设备未按实际到达顺序服务")
    if overlaps["pair_count"] != 0 or summary["processing_capacity_violations"] != 0:
        raise AssertionError(f"{name}发生资源或物理车位重叠")

    return {
        "generated": summary["generated"],
        "within_69000_second_auxiliary_window": summary["within_window"],
        "first_out": summary["first_out"],
        "last_out": summary["full_last_out"],
        "overall_takt": (
            (summary["full_last_out"] - summary["full_first_out"])
            / (summary["generated"] - 1)
        ),
        "full_raw_block_wait": summary["full_raw_block_wait"],
        "sequence_first_30": "".join(sequence[:30]),
        "sequence_hash": tickets.vehicle_sequence_hash(sequence),
        "service_checks": checks,
        "processing_capacity_violations": summary["processing_capacity_violations"],
        "station_slot_overlaps": overlaps,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    quantity = _scenario("quantity")
    if quantity["sequence_hash"] != EXPECTED_QUANTITY_SEQUENCE_HASH:
        raise AssertionError("R9冻结排列哈希发生变化")
    result = {
        "round": "v2.9-round3-r10",
        "minimal_arrival_first": _minimal_arrival_first(),
        "quantity": quantity,
        "ratio": _scenario("ratio"),
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
