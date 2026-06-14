#!/usr/bin/env python3
"""Verify round-9 recent-takt wording against the real-data model."""

from __future__ import annotations

import importlib.util
import math
import sys
from pathlib import Path


APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from core import tickets  # noqa: E402


def _load_baseline():
    path = APP_DIR / "scripts" / "verify_v29_real_data_baseline.py"
    spec = importlib.util.spec_from_file_location("v29_baseline", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载基线脚本：{path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _recent_takt(records, current_time):
    finished = sorted(
        (item for item in records if item["out"] <= current_time + 1e-9),
        key=lambda item: (item["out"], item["car"]),
    )
    recent = finished[-6:]
    gaps = [recent[index]["out"] - recent[index - 1]["out"] for index in range(1, len(recent))]
    value = sum(gaps) / len(gaps) if gaps else None
    return finished, recent, gaps, value


def main() -> None:
    baseline = _load_baseline()
    rows, _ = tickets.schedule(
        baseline.station_defs(),
        1240,
        vehicle_counts={"A": 248, "B": 372, "C": 620},
        sequence_mode="ratio",
        ratio_pattern={"A": 2, "B": 3, "C": 5},
        launch_takt=baseline.TARGET_TAKT,
    )
    records = baseline.vehicle_records(rows)

    finished_1340, recent_1340, gaps_1340, takt_1340 = _recent_takt(records, 1340.0)
    finished_1830, recent_1830, gaps_1830, takt_1830 = _recent_takt(records, 1830.0)
    if len(finished_1340) != 2 or gaps_1340 != [58.0] or not math.isclose(takt_1340 or 0.0, 58.0):
        raise AssertionError("1340秒近期节拍复算不一致")
    expected_1830 = [528.5, 58.0, 62.5, 150.0, 62.5]
    if gaps_1830 != expected_1830 or not math.isclose(takt_1830 or 0.0, 172.3):
        raise AssertionError("1830秒近期节拍复算不一致")

    window = sorted(
        (item for item in records if item["out"] <= 69000.0 + 1e-9),
        key=lambda item: (item["out"], item["car"]),
    )
    overall = (window[-1]["out"] - window[0]["out"]) / (len(window) - 1)
    print({
        "1340s": {"finished": len(finished_1340), "cars": [item["car"] for item in recent_1340], "gaps": gaps_1340, "recent_takt": takt_1340},
        "1830s": {"finished": len(finished_1830), "cars": [item["car"] for item in recent_1830], "gaps": gaps_1830, "recent_takt": takt_1830},
        "analysis_window": {"finished": len(window), "overall_takt": overall},
    })


if __name__ == "__main__":
    main()
