#!/usr/bin/env python3
"""Verify Round 6 wait-cause chains against the two real-data scenarios."""

from __future__ import annotations

import importlib.util
import json
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
EXPECTED = {
    "quantity": {
        "actual_wait": 239975.5,
        "excess_wait": 5430.0,
        "top_chain": ("电检1", "电检2", "小跑道前车占用", 78, 3384.0),
    },
    "ratio": {
        "actual_wait": 238387.0,
        "excess_wait": 3768.5,
        "top_chain": ("电检1", "电检2", "小跑道前车占用", 62, 1672.5),
    },
}


def _schedule(name: str):
    if name == "quantity":
        counts = {"A": 276, "B": 306, "C": 566}
        mode = "alternate"
        ratio = None
    else:
        counts = {"A": 248, "B": 372, "C": 620}
        mode = "ratio"
        ratio = {"A": 2, "B": 3, "C": 5}
    sequence = tickets.build_vehicle_sequence(
        sum(counts.values()),
        counts,
        sequence_mode=mode,
        max_consecutive=5,
        ratio_pattern=ratio,
    )
    return tickets.schedule(
        BASELINE.station_defs(),
        len(sequence),
        vehicle_counts=counts,
        sequence_mode=mode,
        max_consecutive=5,
        ratio_pattern=ratio,
        vehicle_sequence=sequence,
        launch_takt=TARGET,
    )


def _verify(name: str) -> dict:
    rows, max_finish = _schedule(name)
    before = [
        (row["car"], row["step_seq"], row["start"], row["svc_finish"], row["depart"])
        for row in rows
    ]
    summary = analysis.analyze_schedule_v2(rows, max_finish, TARGET)["summary"]
    after = [
        (row["car"], row["step_seq"], row["start"], row["svc_finish"], row["depart"])
        for row in rows
    ]
    expected = EXPECTED[name]
    if before != after:
        raise AssertionError(f"{name}：原因链分析修改了排程时间链")
    if abs(summary["total_actual_wait"] - expected["actual_wait"]) > 1e-9:
        raise AssertionError(f"{name}：累计实际等待变化")
    if abs(summary["total_bottleneck_wait"] - expected["excess_wait"]) > 1e-9:
        raise AssertionError(f"{name}：累计节拍外等待变化")
    if abs(summary["wait_cause_chain_total_time"] - expected["excess_wait"]) > 1e-9:
        raise AssertionError(f"{name}：等待真因工时未覆盖全部节拍外等待")
    if summary["wait_cause_chain_incomplete_time"] != 0:
        raise AssertionError(f"{name}：存在未解析原因链")
    if summary["wait_cause_chain_coverage_rate"] != 1.0:
        raise AssertionError(f"{name}：原因链覆盖率不是100%")
    if any("signature" in cause_slice for row in rows for cause_slice in row.get("wait_cause_slices", [])):
        raise AssertionError(f"{name}：内部原因签名泄漏到最终 rows")

    top = summary["wait_cause_chain_summary"][0]
    actual_top = (
        top["waiting_station"],
        top["direct_blocking_station"],
        top["terminal_cause"],
        top["event_count"],
        top["wait_time"],
    )
    if actual_top != expected["top_chain"]:
        raise AssertionError(f"{name}：首要原因链变化：{actual_top}")
    return {
        "max_finish": max_finish,
        "actual_wait": summary["total_actual_wait"],
        "excess_wait": summary["total_bottleneck_wait"],
        "cause_chain_total": summary["wait_cause_chain_total_time"],
        "coverage_rate": summary["wait_cause_chain_coverage_rate"],
        "incomplete_time": summary["wait_cause_chain_incomplete_time"],
        "cause_chains": summary["wait_cause_chain_summary"],
        "schedule_time_chain_unchanged": before == after,
    }


def main() -> None:
    result = {name: _verify(name) for name in ("quantity", "ratio")}
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
