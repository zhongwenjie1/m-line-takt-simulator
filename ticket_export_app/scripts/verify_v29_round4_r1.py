#!/usr/bin/env python3
"""Verify R1 blocking-root attribution without changing schedule times."""

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
EXPECTED = {
    "quantity": {"moved_events": 911, "moved_time": 32687.0},
    "ratio": {"moved_events": 862, "moved_time": 33132.0},
}


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
    return rows


def _row_time_snapshot(rows: list[dict[str, Any]]) -> list[tuple[Any, ...]]:
    return [
        (
            row["car"],
            row["step_seq"],
            row["line_no"],
            row["start"],
            row["svc_finish"],
            row["depart"],
            row["block_wait"],
            row.get("launch_wait", 0.0),
        )
        for row in rows
    ]


def _scenario(name: str) -> dict[str, Any]:
    rows = _run_real(name)
    before = _row_time_snapshot(rows)
    max_finish = max(float(row["depart"]) for row in rows)
    result = analysis.analyze_schedule_v2(rows, max_finish, BASELINE.TARGET_TAKT)
    after = _row_time_snapshot(rows)
    if before != after:
        raise AssertionError(f"{name}分析过程修改了排程时间链")

    summary = result["summary"]
    moved = [
        item
        for item in summary["blocking_root_attributions"]
        if item["skipped_zero_duration_node"]
    ]
    moved_events = sum(int(item["event_count"]) for item in moved)
    moved_time = sum(float(item["blocking_time"]) for item in moved)
    expected = EXPECTED[name]
    if moved_events != expected["moved_events"] or abs(moved_time - expected["moved_time"]) > 1e-9:
        raise AssertionError(
            f"{name}跨0工时节点归因不符：{moved_events}条/{moved_time}s"
        )

    return {
        "generated": len(BASELINE.group_rows(rows)),
        "schedule_time_chain_unchanged": before == after,
        "attributed_blocking_time": summary["attributed_blocking_time"],
        "attributed_launch_wait_time": summary["attributed_launch_wait_time"],
        "attributed_post_process_wait_time": summary["attributed_post_process_wait_time"],
        "displayed_blocking_time": summary["total_blocking_time"],
        "moved_events": moved_events,
        "moved_time": moved_time,
        "moved_attributions": moved,
    }


def _minimal_zero_node_trace() -> dict[str, Any]:
    rows = [
        {
            "car": 1,
            "step_seq": 1,
            "step_display": "气囊安装+绝缘",
            "dur": 58,
            "start": 0,
            "svc_finish": 58,
            "depart": 78,
            "block_wait": 20,
            "launch_wait": 0,
            "capacity": 1,
        },
        {
            "car": 1,
            "step_seq": 2,
            "step_display": "L2++",
            "dur": 0,
            "start": 78,
            "svc_finish": 78,
            "depart": 78,
            "block_wait": 0,
            "launch_wait": 0,
            "capacity": 1,
        },
        {
            "car": 1,
            "step_seq": 3,
            "step_display": "小跑道",
            "dur": 50,
            "start": 78,
            "svc_finish": 128,
            "depart": 128,
            "block_wait": 0,
            "launch_wait": 0,
            "capacity": 1,
        },
    ]
    summary = analysis.analyze_schedule_v2(rows, 128, 58)["summary"]
    attribution = summary["blocking_root_attributions"][0]
    if attribution["root_station"] != "小跑道":
        raise AssertionError(f"0工时节点仍被当作根因：{attribution}")
    if attribution["immediate_next_station"] != "L2++":
        raise AssertionError(f"最小场景紧邻节点记录错误：{attribution}")
    return attribution


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    result = {
        "round": "v2.9-round4-r1",
        "minimal_zero_node_trace": _minimal_zero_node_trace(),
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
