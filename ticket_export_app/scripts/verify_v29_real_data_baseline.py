#!/usr/bin/env python3
"""Historical investigation snapshot for the original v2.9 real-data results.

This module remains available because later verification scripts reuse its
station definitions and inspection helpers. Its embedded expected values are
from an earlier scheduler stage and are no longer the current pass/fail gate.

Use verify_v29_round16c_authoritative_baseline.py as the authoritative current
regression entry point.
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any


APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from core import tickets  # noqa: E402


TARGET_TAKT = 58.0
ANALYSIS_SECONDS = 69_000.0

STATION_NAMES = [
    "引取",
    "集约",
    "空气悬挂+快充",
    "四轮定位上、下",
    "ADAS",
    "ARHUD+慢充",
    "转毂",
    "制动",
    "EG漏液/底盘漏液",
    "全景/空调检查",
    "气囊安装+绝缘",
    "L2++",
    "电检1",
    "电检2",
    "小跑道",
]

DURATIONS = {
    "A": [55, 106, 0, 115, 97, 0, 116, 108, 115, 98, 0, 0, 0, 0, 51],
    "B": [56, 111, 108.5, 120.5, 107, 51.5, 114.5, 114, 112.5, 0, 116, 88, 84.5, 97.5, 50],
    "C": [56, 111, 108.5, 120.5, 107, 143.5, 114.5, 114, 112.5, 0, 116, 88, 84.5, 97.5, 50],
}

EXPECTED = {
    "ratio": {
        "generated": 1240,
        "within_window": 1046,
        "qualified": 210,
        "first_out": 861.0,
        "last_out": 68987.5,
        "last_scope_car": 1046,
        "overall_takt": 65.19282296650718,
        "raw_block_wait_in_window": 243807.0,
        "full_last_out": 81543.5,
        "full_raw_block_wait": 285937.0,
        "out_gap_over_target_count": 469,
    },
    "quantity": {
        "generated": 1148,
        "within_window": 1036,
        "qualified": 239,
        "first_out": 1423.5,
        "last_out": 68952.5,
        "last_scope_car": 1036,
        "overall_takt": 65.24541062801933,
        "raw_block_wait_in_window": 235263.5,
        "full_last_out": 75885.5,
        "full_raw_block_wait": 260263.5,
        "out_gap_over_target_count": 570,
        "first_a_position": 186,
        "max_consecutive": 5,
    },
}

CAR_LINE = re.compile(
    r"^Car#(\d+)\s*([ABC])\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+(\S+)\s+(.*)$"
)
SEGMENT_LINE = re.compile(
    r"ST(\d+)\s+(.+?)\(开:([\d.]+)s 加:([\d.]+)s 等前:([\d.]+)s 等后:([\d.]+)s\)"
)


def station_defs() -> list[dict[str, Any]]:
    result = []
    for index, name in enumerate(STATION_NAMES):
        shared = index in (0, 14)
        result.append(
            {
                "seq": index + 1,
                "display": name,
                "group": name,
                "duration_a": DURATIONS["A"][index],
                "duration_b": DURATIONS["B"][index],
                "duration_c": DURATIONS["C"][index],
                "device_count": 1 if shared else 2,
                "capacity": 1 if shared else 2,
                "line_scope": "双线共用" if shared else "双线",
                "run_mode": "双线单设备" if shared else "双线双设备",
            }
        )
    return result


def group_rows(rows: list[dict[str, Any]]) -> dict[int, list[dict[str, Any]]]:
    grouped: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        grouped[int(row["car"])].append(row)
    for car_rows in grouped.values():
        car_rows.sort(key=lambda item: (int(item.get("step_seq", 0)), float(item.get("start", 0.0))))
    return dict(grouped)


def vehicle_records(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    records = []
    for car, segments in sorted(group_rows(rows).items()):
        records.append(
            {
                "car": car,
                "type": str(segments[0]["car_type"]),
                "in": float(segments[0]["start"]),
                "out": float(segments[-1]["depart"]),
                "flow": float(segments[-1]["depart"]) - float(segments[0]["start"]),
                "wait": sum(
                    float(segment.get("launch_wait", 0.0)) + float(segment.get("block_wait", 0.0))
                    for segment in segments
                ),
                "segments": segments,
            }
        )
    return records


def raw_block_by_station(records: list[dict[str, Any]], cutoff: float | None) -> dict[str, float]:
    totals: dict[str, float] = defaultdict(float)
    for record in records:
        segments = record["segments"]
        for index, segment in enumerate(segments):
            wait = max(0.0, float(segment.get("block_wait", 0.0)))
            if wait <= 0:
                continue
            finish = float(segment["svc_finish"])
            depart = float(segment["depart"])
            if cutoff is None:
                occurred = wait
            elif cutoff <= finish:
                occurred = 0.0
            elif cutoff < depart:
                occurred = min(wait, cutoff - finish)
            else:
                occurred = wait
            if occurred <= 0:
                continue
            station = segments[index + 1]["step_display"] if index + 1 < len(segments) else segment["step_display"]
            totals[str(station)] += occurred
    return dict(sorted(totals.items(), key=lambda item: (-item[1], item[0])))


def is_qualified(record: dict[str, Any]) -> bool:
    for segment in record["segments"]:
        duration = float(segment.get("dur", 0.0))
        if duration <= 0:
            continue
        capacity = max(1, int(segment.get("capacity", 1) or 1))
        if duration / capacity > TARGET_TAKT + 1e-9:
            return False
    return True


def max_run(sequence: list[str]) -> int:
    maximum = 0
    current = 0
    previous = None
    for vehicle_type in sequence:
        current = current + 1 if vehicle_type == previous else 1
        previous = vehicle_type
        maximum = max(maximum, current)
    return maximum


def overlap_counts(rows: list[dict[str, Any]]) -> dict[str, int]:
    grouped: dict[tuple[str, str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        if float(row.get("dur", 0.0)) <= 0:
            continue
        key = (
            str(row.get("step_display", "")),
            str(row.get("resource_key", "")),
            str(row.get("line_no", "")),
        )
        grouped[key].append(row)

    processing = 0
    occupancy = 0
    overtakes = 0
    for values in grouped.values():
        values.sort(key=lambda item: (float(item["start"]), int(item["car"])))
        for previous, current in zip(values, values[1:]):
            if float(current["start"]) < float(previous["svc_finish"]) - 1e-9:
                processing += 1
            if float(current["start"]) < float(previous["depart"]) - 1e-9:
                occupancy += 1
            if int(current["car"]) < int(previous["car"]):
                overtakes += 1
    return {
        "processing_overlap": processing,
        "occupancy_overlap": occupancy,
        "same_resource_overtake": overtakes,
    }


def maximum_zero_node_occupancy(rows: list[dict[str, Any]]) -> dict[str, Any]:
    grouped: dict[tuple[str, str], list[tuple[float, int, int]]] = defaultdict(list)
    for row in rows:
        if abs(float(row.get("dur", 0.0))) > 1e-9:
            continue
        start = float(row.get("start", 0.0))
        depart = float(row.get("depart", start))
        if depart <= start + 1e-9:
            continue
        key = (str(row.get("step_display", "")), str(row.get("line_no", "")))
        car = int(row.get("car", 0))
        grouped[key].append((start, 1, car))
        grouped[key].append((depart, -1, car))

    maximum = 0
    maximum_key = ("", "")
    maximum_time = None
    maximum_cars: list[int] = []
    for key, events in grouped.items():
        # Departures are processed before arrivals at the same timestamp.
        events.sort(key=lambda item: (item[0], item[1], item[2]))
        active: set[int] = set()
        for event_time, delta, car in events:
            if delta < 0:
                active.discard(car)
            else:
                active.add(car)
            if len(active) > maximum:
                maximum = len(active)
                maximum_key = key
                maximum_time = event_time
                maximum_cars = sorted(active)
    return {
        "maximum": maximum,
        "station": maximum_key[0],
        "line": maximum_key[1],
        "time": maximum_time,
        "cars": maximum_cars,
    }


def maximum_station_slot_occupancy(rows: list[dict[str, Any]]) -> dict[str, Any]:
    """Count every vehicle physically present in the same station/line slot."""
    grouped: dict[tuple[str, str], list[tuple[float, int, int, str]]] = defaultdict(list)
    for row in rows:
        start = float(row.get("start", 0.0))
        depart = float(row.get("depart", start))
        if depart <= start + 1e-9:
            continue
        key = (str(row.get("step_display", "")), str(row.get("line_no", "")))
        car = int(row.get("car", 0))
        state = "zero_wait" if abs(float(row.get("dur", 0.0))) <= 1e-9 else "processing_or_blocked"
        grouped[key].append((start, 1, car, state))
        grouped[key].append((depart, -1, car, state))

    maximum = 0
    evidence: dict[str, Any] = {"maximum": 0, "station": "", "line": "", "time": None, "cars": []}
    for key, events in grouped.items():
        events.sort(key=lambda item: (item[0], item[1], item[2]))
        active: dict[int, str] = {}
        for event_time, delta, car, state in events:
            if delta < 0:
                active.pop(car, None)
            else:
                active[car] = state
            if len(active) > maximum:
                maximum = len(active)
                evidence = {
                    "maximum": maximum,
                    "station": key[0],
                    "line": key[1],
                    "time": event_time,
                    "cars": [
                        {"car": active_car, "state": active[active_car]}
                        for active_car in sorted(active)
                    ],
                }
    return evidence


def station_slot_overlap_summary(rows: list[dict[str, Any]]) -> dict[str, Any]:
    """Summarize physical same-station/same-line overlaps, including zero-duration waits."""
    grouped: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        start = float(row.get("start", 0.0))
        depart = float(row.get("depart", start))
        if depart <= start + 1e-9:
            continue
        if str(row.get("line_scope", "") or "") == "双线共用":
            key = (str(row.get("step_display", "")), "共享")
        else:
            key = (str(row.get("step_display", "")), str(row.get("line_no", "")))
        grouped[key].append(row)

    pair_counts: Counter[str] = Counter()
    affected_processing_cars: set[int] = set()
    maximum_overlap: dict[str, Any] = {"seconds": 0.0, "station": "", "line": "", "cars": []}
    for (station, line), values in grouped.items():
        values.sort(key=lambda item: (float(item["start"]), int(item["car"])))
        active: list[dict[str, Any]] = []
        for current in values:
            current_start = float(current["start"])
            active = [item for item in active if float(item["depart"]) > current_start + 1e-9]
            for previous in active:
                previous_zero = abs(float(previous.get("dur", 0.0))) <= 1e-9
                current_zero = abs(float(current.get("dur", 0.0))) <= 1e-9
                overlap_type = (
                    ("zero" if previous_zero else "process")
                    + "->"
                    + ("zero" if current_zero else "process")
                )
                pair_counts[overlap_type] += 1
                if previous_zero and not current_zero:
                    affected_processing_cars.add(int(current["car"]))
                overlap = min(float(previous["depart"]), float(current["depart"])) - current_start
                if overlap > float(maximum_overlap["seconds"]):
                    maximum_overlap = {
                        "seconds": overlap,
                        "station": station,
                        "line": line,
                        "cars": [int(previous["car"]), int(current["car"])],
                    }
            active.append(current)

    return {
        "pair_count": sum(pair_counts.values()),
        "pair_types": dict(sorted(pair_counts.items())),
        "processing_cars_entering_occupied_zero_node": len(affected_processing_cars),
        "maximum_overlap": maximum_overlap,
    }


def processing_capacity_violations(rows: list[dict[str, Any]]) -> int:
    """Count moments where active processing exceeds a resource's configured capacity."""
    grouped: dict[str, list[tuple[float, int, int]]] = defaultdict(list)
    capacities: dict[str, int] = {}
    for row in rows:
        if float(row.get("dur", 0.0)) <= 0:
            continue
        key = str(row.get("resource_key", "") or "")
        capacities[key] = max(1, int(row.get("capacity", 1) or 1))
        grouped[key].append((float(row["start"]), 1, int(row["car"])))
        grouped[key].append((float(row["svc_finish"]), -1, int(row["car"])))

    violations = 0
    for key, events in grouped.items():
        events.sort(key=lambda item: (item[0], item[1], item[2]))
        active = 0
        for _, delta, _ in events:
            active += delta
            if active > capacities[key]:
                violations += 1
    return violations


def summarize(rows: list[dict[str, Any]]) -> dict[str, Any]:
    records = vehicle_records(rows)
    window_records = sorted(
        (record for record in records if record["out"] <= ANALYSIS_SECONDS + 1e-9),
        key=lambda record: (record["out"], record["car"]),
    )
    ordered = sorted(records, key=lambda record: (record["out"], record["car"]))
    gaps = [current["out"] - previous["out"] for previous, current in zip(ordered, ordered[1:])]
    window_block = raw_block_by_station(records, ANALYSIS_SECONDS)
    full_block = raw_block_by_station(records, None)
    sequence = [record["type"] for record in records]
    last = ordered[-1]
    first = ordered[0]
    return {
        "generated": len(records),
        "vehicle_counts": dict(sorted(Counter(sequence).items())),
        "sequence_first_30": "".join(sequence[:30]),
        "first_a_position": sequence.index("A") + 1 if "A" in sequence else None,
        "max_consecutive": max_run(sequence),
        "within_window": len(window_records),
        "qualified": sum(1 for record in window_records if is_qualified(record)),
        "first_out": window_records[0]["out"],
        "last_out": window_records[-1]["out"],
        "last_scope_car": window_records[-1]["car"],
        "overall_takt": (window_records[-1]["out"] - window_records[0]["out"]) / (len(window_records) - 1),
        "raw_block_wait_in_window": sum(window_block.values()),
        "raw_block_stations_in_window": window_block,
        "full_first_out": first["out"],
        "full_last_out": last["out"],
        "full_last_car": last["car"],
        "full_raw_block_wait": sum(full_block.values()),
        "out_gap_sum": sum(gaps),
        "out_gap_over_target_count": sum(gap > TARGET_TAKT + 1e-9 for gap in gaps),
        "out_gap_positive_excess": sum(max(0.0, gap - TARGET_TAKT) for gap in gaps),
        "out_gap_negative_compensation": sum(min(0.0, gap - TARGET_TAKT) for gap in gaps),
        "out_gap_net_delta": sum(gap - TARGET_TAKT for gap in gaps),
        "final_vehicle": {
            "car": last["car"],
            "type": last["type"],
            "actual_launch": last["in"],
            "flow": last["flow"],
            "out": last["out"],
            "launch_plus_flow": last["in"] + last["flow"],
            "first_out_plus_gap_sum": first["out"] + sum(gaps),
        },
        "resource_checks": overlap_counts(rows),
        "processing_capacity_violations": processing_capacity_violations(rows),
        "zero_node_occupancy": maximum_zero_node_occupancy(rows),
        "station_slot_occupancy": maximum_station_slot_occupancy(rows),
        "station_slot_overlaps": station_slot_overlap_summary(rows),
        "bottleneck_wait_status": "not_implemented_in_current_baseline",
    }


def parse_evidence(path: Path) -> dict[str, list[dict[str, Any]]]:
    text = path.read_text(encoding="utf-8")
    ratio_text, quantity_text = text.split("按数量投车", 1)

    def parse_section(section: str) -> list[dict[str, Any]]:
        records = []
        for raw_line in section.splitlines():
            match = CAR_LINE.match(raw_line.strip())
            if not match:
                continue
            records.append(
                {
                    "car": int(match.group(1)),
                    "type": match.group(2),
                    "in": float(match.group(3)),
                    "out": float(match.group(4)),
                    "wait": float(match.group(5)),
                    "flow": float(match.group(6)),
                    "segments": [
                        {
                            "seq": int(segment.group(1)),
                            "station": segment.group(2),
                            "start": float(segment.group(3)),
                            "dur": float(segment.group(4)),
                            "launch_wait": float(segment.group(5)),
                            "block_wait": float(segment.group(6)),
                        }
                        for segment in SEGMENT_LINE.finditer(match.group(8))
                    ],
                }
            )
        return records

    return {"ratio": parse_section(ratio_text), "quantity": parse_section(quantity_text)}


def compare_evidence(evidence: list[dict[str, Any]], rows: list[dict[str, Any]]) -> dict[str, Any]:
    actual = {record["car"]: record for record in vehicle_records(rows)}
    differences = []
    for expected in evidence:
        current = actual.get(expected["car"])
        if current is None:
            differences.append({"car": expected["car"], "reason": "missing"})
            continue
        for field in ("type", "in", "out", "wait", "flow"):
            expected_value = expected[field]
            current_value = current[field]
            if isinstance(expected_value, float):
                equal = math.isclose(expected_value, float(current_value), abs_tol=1e-9)
            else:
                equal = expected_value == current_value
            if not equal:
                differences.append(
                    {"car": expected["car"], "field": field, "expected": expected_value, "actual": current_value}
                )
                break
        else:
            expected_segments = expected.get("segments", [])
            current_segments = current["segments"]
            if len(expected_segments) != len(current_segments):
                differences.append(
                    {
                        "car": expected["car"],
                        "field": "segment_count",
                        "expected": len(expected_segments),
                        "actual": len(current_segments),
                    }
                )
                continue
            for expected_segment, current_segment in zip(expected_segments, current_segments):
                segment_fields = (
                    ("seq", expected_segment["seq"], int(current_segment["step_seq"])),
                    ("station", expected_segment["station"], str(current_segment["step_display"])),
                    ("start", expected_segment["start"], float(current_segment["start"])),
                    ("dur", expected_segment["dur"], float(current_segment["dur"])),
                    (
                        "launch_wait",
                        expected_segment["launch_wait"],
                        float(current_segment.get("launch_wait", 0.0)),
                    ),
                    (
                        "block_wait",
                        expected_segment["block_wait"],
                        float(current_segment.get("block_wait", 0.0)),
                    ),
                )
                mismatch = next(
                    (
                        (field, expected_value, actual_value)
                        for field, expected_value, actual_value in segment_fields
                        if expected_value != actual_value
                    ),
                    None,
                )
                if mismatch:
                    field, expected_value, actual_value = mismatch
                    differences.append(
                        {
                            "car": expected["car"],
                            "segment": expected_segment["seq"],
                            "field": field,
                            "expected": expected_value,
                            "actual": actual_value,
                        }
                    )
                    break
    return {"evidence_records": len(evidence), "differences": len(differences), "examples": differences[:10]}


def assert_expected(name: str, summary: dict[str, Any]) -> None:
    errors = []
    for field, expected in EXPECTED[name].items():
        actual = summary.get(field)
        if isinstance(expected, float):
            valid = actual is not None and math.isclose(float(actual), expected, rel_tol=0.0, abs_tol=1e-9)
        else:
            valid = actual == expected
        if not valid:
            errors.append(f"{name}.{field}: expected {expected!r}, got {actual!r}")
    if errors:
        raise AssertionError("Baseline mismatch:\n" + "\n".join(errors))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--evidence", type=Path, help="Optional 调查资料.txt path for row-by-row verification")
    parser.add_argument("--output", type=Path, help="Optional JSON output path; stdout is always written")
    args = parser.parse_args()

    definitions = station_defs()
    ratio_rows, _ = tickets.schedule(
        definitions,
        1240,
        vehicle_counts={"A": 248, "B": 372, "C": 620},
        sequence_mode="ratio",
        ratio_pattern={"A": 2, "B": 3, "C": 5},
        launch_takt=TARGET_TAKT,
    )
    quantity_rows, _ = tickets.schedule(
        definitions,
        1148,
        vehicle_counts={"A": 276, "B": 306, "C": 566},
        sequence_mode="alternate",
        max_consecutive=5,
        launch_takt=TARGET_TAKT,
    )

    ratio_summary = summarize(ratio_rows)
    quantity_summary = summarize(quantity_rows)
    assert_expected("ratio", ratio_summary)
    assert_expected("quantity", quantity_summary)

    result: dict[str, Any] = {
        "baseline": "v2.9-before-improvement-rounds",
        "target_takt": TARGET_TAKT,
        "analysis_seconds": ANALYSIS_SECONDS,
        "ratio": ratio_summary,
        "quantity": quantity_summary,
        "baseline_assertions": "passed",
    }

    if args.evidence:
        evidence = parse_evidence(args.evidence)
        result["evidence_verification"] = {
            "ratio": compare_evidence(evidence["ratio"], ratio_rows),
            "quantity": compare_evidence(evidence["quantity"], quantity_rows),
        }
        if result["evidence_verification"]["ratio"]["differences"]:
            raise AssertionError("Ratio evidence differs from current scheduler output")
        if result["evidence_verification"]["quantity"]["differences"]:
            raise AssertionError("Quantity evidence differs from current scheduler output")

    output = json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(output, encoding="utf-8")
    sys.stdout.write(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
