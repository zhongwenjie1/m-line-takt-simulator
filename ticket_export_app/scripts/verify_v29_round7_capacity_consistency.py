#!/usr/bin/env python3
"""Verify Round 7 vehicle capacity decisions are identical across consumers."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path


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


def _run(name: str):
    if name == "quantity":
        counts = {"A": 276, "B": 306, "C": 566}
        mode = "alternate"
        ratio = None
    else:
        counts = {"A": 248, "B": 372, "C": 620}
        mode = "ratio"
        ratio = {"A": 2, "B": 3, "C": 5}
    sequence = tickets.build_vehicle_sequence(
        sum(counts.values()), counts, sequence_mode=mode,
        max_consecutive=5, ratio_pattern=ratio,
    )
    rows, max_finish = tickets.schedule(
        BASELINE.station_defs(), len(sequence), vehicle_counts=counts,
        sequence_mode=mode, max_consecutive=5, ratio_pattern=ratio,
        vehicle_sequence=sequence, launch_takt=TARGET,
    )
    before = [
        (row["car"], row["step_seq"], row["start"], row["svc_finish"], row["depart"])
        for row in rows
    ]
    result = analysis.analyze_schedule_v2(rows, max_finish, TARGET)
    after = [
        (row["car"], row["step_seq"], row["start"], row["svc_finish"], row["depart"])
        for row in rows
    ]
    if before != after:
        raise AssertionError(f"{name}：能力判断修改了排程时间链")

    capacity_results = result["car_capacity_results"]
    if len(capacity_results) != len(sequence):
        raise AssertionError(f"{name}：逐车能力结果数量不完整")
    by_car = {int(item["car"]): item for item in capacity_results}

    for car, car_rows in BASELINE.group_rows(rows).items():
        expected_stations = []
        for row in car_rows:
            duration = float(row.get("dur", 0.0) or 0.0)
            if duration <= 0:
                continue
            capacity = max(1, int(row.get("capacity", 1) or 1))
            if duration / capacity > TARGET + 1e-9:
                expected_stations.append(str(row.get("step_display", "")))
        item = by_car[car]
        actual_stations = [value["station"] for value in item["over_capacity_stations"]]
        if actual_stations != expected_stations:
            raise AssertionError(f"{name} Car#{car}：超目标工程不一致")
        if bool(item["meets_capacity_target"]) != (not expected_stations):
            raise AssertionError(f"{name} Car#{car}：能力判断不一致")

    full_meets = sum(bool(item["meets_capacity_target"]) for item in capacity_results)
    if name == "quantity" and full_meets != 276:
        raise AssertionError(f"数量模式能力满足车辆应为276，实际{full_meets}")
    if any(
        any(float(row.get("dur", 0.0) or 0.0) == 0 for row in BASELINE.group_rows(rows)[int(item["car"])])
        and any(value["duration"] <= 0 for value in item["over_capacity_stations"])
        for item in capacity_results
    ):
        raise AssertionError(f"{name}：0工时节点错误参加能力判断")
    return {
        "vehicles": len(capacity_results),
        "capacity_meets": full_meets,
        "capacity_exceeds": len(capacity_results) - full_meets,
        "schedule_time_chain_unchanged": before == after,
        "first_20": [
            {
                "car": item["car"],
                "type": item["car_type"],
                "status": item["capacity_status"],
                "over_stations": item["over_capacity_station_text"],
            }
            for item in capacity_results[:20]
        ],
    }


def main() -> None:
    import json
    print(json.dumps({name: _run(name) for name in ("quantity", "ratio")}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
