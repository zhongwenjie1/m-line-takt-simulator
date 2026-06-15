#!/usr/bin/env python3
"""Verify round-14 schedule path isolation without changing any result data."""

from __future__ import annotations

import hashlib
import importlib.util
import json
import sys
from pathlib import Path

APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from core import tickets  # noqa: E402


EXPECTED = {
    "quantity": "04a96cbbcae32c159df6fd31774bc8e7e91f8d6a21b93281b942713c182b8af0",
    "ratio": "34447d9ada7294c783398bab126b5b22d9dd1e895c6a8064b7093464d6ba5ba7",
    "legacy_zone_gate": "38b814e12bd62f33d8f46611864e92a99f6ea738ed6ab80d0bc1f24ab6c2ff75",
}


def _load_baseline():
    path = APP_DIR / "scripts" / "verify_v29_real_data_baseline.py"
    spec = importlib.util.spec_from_file_location("v29_baseline", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _normalize(value):
    if isinstance(value, float):
        return round(value, 9)
    if isinstance(value, dict):
        return {key: _normalize(item) for key, item in sorted(value.items())}
    if isinstance(value, (list, tuple)):
        return [_normalize(item) for item in value]
    return value


def _digest(value) -> str:
    raw = json.dumps(
        _normalize(value), ensure_ascii=False, sort_keys=True, separators=(",", ":")
    ).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def _run_real_case(defs, sequence):
    rows, max_finish = tickets.schedule(
        defs, len(sequence), launch_takt=58.0, vehicle_sequence=sequence
    )
    analysis = tickets.analyze_schedule(rows, max_finish, 58.0)
    return {"rows": rows, "max_finish": max_finish, "analysis": analysis}


def _legacy_zone_defs():
    return [
        {
            "seq": 1, "display": "入口", "group": "入口",
            "duration_a": 10, "duration_b": 12, "duration_c": 8,
            "device_count": 1, "capacity": 1,
            "line_scope": "双线共用", "run_mode": "双线单设备",
            "gate_zone_id": "Z1", "gate_buffer": 1,
        },
        {
            "seq": 2, "display": "区域作业", "group": "区域作业",
            "duration_a": 20, "duration_b": 25, "duration_c": 15,
            "device_count": 1, "capacity": 1,
            "line_scope": "双线共用", "run_mode": "双线单设备",
            "zone_id": "Z1", "zone_capacity": 1,
        },
        {
            "seq": 3, "display": "出口", "group": "出口",
            "duration_a": 9, "duration_b": 11, "duration_c": 7,
            "device_count": 1, "capacity": 1,
            "line_scope": "双线共用", "run_mode": "双线单设备",
        },
    ]


def main():
    baseline = _load_baseline()
    defs = baseline.station_defs()
    quantity_sequence = tickets.build_vehicle_sequence(
        1148, {"A": 276, "B": 306, "C": 566}, "alternate", 5
    )
    ratio_sequence = tickets.build_vehicle_sequence(
        1240, {"A": 2, "B": 3, "C": 5}, "ratio", 10,
        {"A": 2, "B": 3, "C": 5},
    )

    quantity = _run_real_case(defs, quantity_sequence)
    ratio = _run_real_case(defs, ratio_sequence)
    zone_result = tickets.schedule(
        _legacy_zone_defs(), 12,
        vehicle_counts={"A": 4, "B": 4, "C": 4},
        sequence_mode="alternate", max_consecutive=2, launch_takt=8.0,
    )

    actual = {
        "quantity": _digest(quantity),
        "ratio": _digest(ratio),
        "legacy_zone_gate": _digest(zone_result),
    }
    assert actual == EXPECTED, {"expected": EXPECTED, "actual": actual}
    print({
        "hashes_match_pre_isolation": True,
        "quantity_rows": len(quantity["rows"]),
        "ratio_rows": len(ratio["rows"]),
        "legacy_zone_gate_rows": len(zone_result[0]),
        "hashes": actual,
    })


if __name__ == "__main__":
    main()
