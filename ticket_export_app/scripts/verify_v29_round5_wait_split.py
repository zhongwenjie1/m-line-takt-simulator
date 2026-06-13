#!/usr/bin/env python3
"""Verify actual waiting and holding-capacity bottleneck waiting."""

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

from core import analysis, tickets  # noqa: E402


def _load_baseline_module():
    path = APP_DIR / "scripts" / "verify_v29_real_data_baseline.py"
    spec = importlib.util.spec_from_file_location("v29_baseline", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载基线脚本：{path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


BASELINE = _load_baseline_module()
TARGET = 58.0
EXPECTED_CAPACITY_OVER_TAKT = {
    "quantity": {"total": 19489.0, "count": 1438},
    "ratio": {"total": 21514.0, "count": 1612},
}


def _row(
    car: int,
    seq: int,
    station: str,
    duration: float,
    capacity: int,
    start: float,
    block_wait: float = 0.0,
    launch_wait: float = 0.0,
) -> dict[str, Any]:
    finish = start + duration
    return {
        "car": car,
        "step_seq": seq,
        "step_display": station,
        "dur": duration,
        "capacity": capacity,
        "device_count": capacity,
        "start": start,
        "svc_finish": finish,
        "depart": finish + block_wait,
        "block_wait": block_wait,
        "launch_wait": launch_wait,
    }


def _minimal_cases() -> dict[str, Any]:
    dual_rows = [
        _row(1, 1, "等待工程", 100, 2, 0, block_wait=120),
        _row(1, 2, "超节拍工程", 143.5, 2, 220),
    ]
    single_rows = [
        _row(1, 1, "等待工程", 50, 1, 0, block_wait=70),
        _row(1, 2, "超节拍工程", 116, 1, 120),
    ]
    zero_node_rows = [
        _row(1, 1, "ADAS", 107, 2, 0, block_wait=143.5),
        _row(1, 2, "零工时节点", 0, 1, 250.5),
        _row(1, 3, "ARHUD+慢充", 143.5, 2, 250.5),
    ]

    def summary(rows: list[dict[str, Any]]) -> dict[str, Any]:
        result = analysis.analyze_schedule_v2(
            rows, max(float(row["depart"]) for row in rows), TARGET
        )["summary"]
        return {
            "raw": result["total_raw_flow_wait"],
            "bottleneck": result["total_bottleneck_wait"],
            "bottleneck_stations": result["bottleneck_wait_by_station"],
            "downstream_stations": result["bottleneck_wait_by_downstream_station"],
            "over_takt_processes": result["capacity_over_takt_processes"],
        }

    dual = summary(dual_rows)
    single = summary(single_rows)
    zero_node = summary(zero_node_rows)
    if dual["raw"] != 120 or dual["bottleneck"] != 4:
        raise AssertionError(f"双设备等待工程可接纳116秒口径错误：{dual}")
    if single["raw"] != 70 or single["bottleneck"] != 12:
        raise AssertionError(f"单设备等待工程可接纳58秒口径错误：{single}")
    if zero_node["bottleneck_stations"] != [
        {"station": "ADAS", "wait_time": 27.5}
    ] or zero_node["downstream_stations"] != [
        {"station": "零工时节点", "wait_time": 27.5}
    ] or zero_node["over_takt_processes"][0]["station"] != "ARHUD+慢充":
        raise AssertionError(f"0工时节点后的瓶颈归因错误：{zero_node}")
    return {"dual_device_116": dual, "single_device_116": single, "zero_node": zero_node}


def _run_real(name: str) -> list[dict[str, Any]]:
    definitions = BASELINE.station_defs()
    if name == "quantity":
        counts = {"A": 276, "B": 306, "C": 566}
        sequence = tickets.build_vehicle_sequence(
            sum(counts.values()), counts, sequence_mode="alternate", max_consecutive=5
        )
        rows, _ = tickets.schedule(
            definitions,
            len(sequence),
            vehicle_counts=counts,
            sequence_mode="alternate",
            max_consecutive=5,
            vehicle_sequence=sequence,
            launch_takt=TARGET,
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
            vehicle_sequence=sequence,
            launch_takt=TARGET,
        )
    return rows


def _scenario(name: str) -> dict[str, Any]:
    rows = _run_real(name)
    before = [
        (row["car"], row["step_seq"], row["start"], row["svc_finish"], row["depart"])
        for row in rows
    ]
    result = analysis.analyze_schedule_v2(
        rows, max(float(row["depart"]) for row in rows), TARGET
    )["summary"]
    after = [
        (row["car"], row["step_seq"], row["start"], row["svc_finish"], row["depart"])
        for row in rows
    ]
    for car_rows in BASELINE.group_rows(rows).values():
        for current_row, next_row in zip(car_rows, car_rows[1:]):
            if abs(float(current_row["depart"]) - float(next_row["start"])) > 1e-9:
                raise AssertionError(
                    f"{name}下游接收工程时间链不连续："
                    f"Car#{current_row['car']} {current_row['step_display']} -> "
                    f"{next_row['step_display']}"
                )
    direct_raw = sum(max(0.0, float(row.get("block_wait", 0.0))) for row in rows)
    if before != after:
        raise AssertionError(f"{name}分析修改了排程时间链")
    if abs(result["total_raw_flow_wait"] - direct_raw) > 1e-9:
        raise AssertionError(f"{name}原始等待不能从过程记录复算")
    if result["total_bottleneck_wait"] > result["total_raw_flow_wait"] + 1e-9:
        raise AssertionError(f"{name}瓶颈等待超过原始等待")
    event_total = 0.0
    for event in result["bottleneck_wait_events"]:
        expected = max(
            0.0,
            float(event["actual_wait"]) - float(event["holding_limit"]),
        )
        if abs(float(event["bottleneck_wait"]) - expected) > 1e-9:
            raise AssertionError(f"{name}瓶颈等待不能逐条复算：{event}")
        event_total += float(event["bottleneck_wait"])
    if abs(event_total - result["total_bottleneck_wait"]) > 1e-9:
        raise AssertionError(f"{name}瓶颈等待事件合计不一致")
    over_takt_count = 0
    over_takt_total = 0.0
    for item in result["capacity_over_takt_processes"]:
        if float(item["duration"]) <= float(item["capacity_limit"]) + 1e-9:
            raise AssertionError(f"{name}普通工程被误标为超节拍工程：{item}")
        over_takt_count += int(item["count"])
        over_takt_total += float(item["total_over_takt"])
    expected_over = EXPECTED_CAPACITY_OVER_TAKT[name]
    if over_takt_count != expected_over["count"] or abs(
        over_takt_total - expected_over["total"]
    ) > 1e-9:
        raise AssertionError(
            f"{name}超节拍工程汇总不一致：{over_takt_count}, {over_takt_total}"
        )
    return {
        "generated": len(BASELINE.group_rows(rows)),
        "schedule_time_chain_unchanged": before == after,
        "total_raw_flow_wait": result["total_raw_flow_wait"],
        "raw_launch_wait": result["raw_launch_wait"],
        "raw_post_process_wait": result["raw_post_process_wait"],
        "total_bottleneck_wait": result["total_bottleneck_wait"],
        "bottleneck_launch_wait": result["bottleneck_launch_wait"],
        "bottleneck_post_process_wait": result["bottleneck_post_process_wait"],
        "bottleneck_wait_event_count": result["bottleneck_wait_event_count"],
        "bottleneck_wait_by_station": result["bottleneck_wait_by_station"],
        "bottleneck_wait_by_downstream_station": result[
            "bottleneck_wait_by_downstream_station"
        ],
        "capacity_over_takt_processes": result["capacity_over_takt_processes"],
        "capacity_over_takt_total_time": result["capacity_over_takt_total_time"],
        "existing_total_wait": result["total_wait"],
        "existing_displayed_blocking_time": result["total_blocking_time"],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    result = {
        "round": "v2.9-round5-wait-split",
        "minimal_cases": _minimal_cases(),
        "quantity": _scenario("quantity"),
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
