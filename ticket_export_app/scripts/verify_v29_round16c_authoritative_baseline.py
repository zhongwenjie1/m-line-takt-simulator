#!/usr/bin/env python3
"""Round-16C authoritative regression baseline + SHA-256 lock.

本脚本汇总分散在 Round-1/3/5/10/11/16A 等已审定结果的当前权威指标，
并对每组真实场景的排程 rows 计算 SHA-256 锁，作为后续轮回归的统一通过条件。

旧 verify_v29_real_data_baseline.py 自本轮起不再作为通过条件，仅保留为
历史调查快照与公共工具函数库（station_defs、DURATIONS、group_rows 等）。

通过条件：
- 两组真实场景的所有权威指标全部对得上；
- 排程 rows 的 SHA-256 与本文件锁定值一致。

权威值来源（每项均可追溯到已审定文档章节）：
- 数量 A276/B306/C566 → 第16轮审计 §2 与第3轮R10验证记录
- 比例 2:3:5、1150 分钟 → 第1轮R2 与第11轮目标批次锁定
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
import sys
from pathlib import Path
from typing import Any


APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from core import tickets  # noqa: E402
from core.analysis import apply_time_window_analysis  # noqa: E402
from core.input_parser import parse_multi_project_inputs  # noqa: E402


def _load_legacy_baseline():
    """Reuse station definitions and helpers from the legacy baseline module."""
    path = APP_DIR / "scripts" / "verify_v29_real_data_baseline.py"
    spec = importlib.util.spec_from_file_location("v29_baseline", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载基线脚本：{path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


LEGACY = _load_legacy_baseline()


def _load_round3_checks():
    """Reuse the established FCFS and launch-order checks from round 3."""
    path = APP_DIR / "scripts" / "verify_v29_round3_r10.py"
    spec = importlib.util.spec_from_file_location("v29_round3", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载R10验证脚本：{path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


ROUND3 = _load_round3_checks()


# ---------------------------------------------------------------------------
# 权威指标（按已审定文档章节登记）
# ---------------------------------------------------------------------------

AUTHORITATIVE_QUANTITY = {
    # 来源：第16轮审计报告 §2、第3轮R10 真实数据复核
    "generated": 1148,
    "finished_full_batch": 1148,
    "first_out": 1423.5,
    "last_out_full_batch": 76152.5,
    "overall_takt_full_batch": 65.15170008718395,  # (76152.5 - 1423.5) / 1147
    "total_actual_wait_full_batch": 239975.5,
    "total_bottleneck_wait_full_batch": 5430.0,
    # 物理约束（R1 / R2 / R10 综合结果）
    "same_resource_processing_overlap": 0,
    "station_slot_overlap": 0,
    "shared_resource_arrival_order_violation": 0,
    "first_station_launch_skip": 0,
}

AUTHORITATIVE_RATIO = {
    # 来源：R10后当前rows、第11轮目标批次锁定、第16A轮零下线窗口修复
    "generated": 1240,
    "theoretical_launch_count": 1190,
    "simulation_buffer_count": 50,
    "within_window": 1024,
    "first_out_in_window": 861.0,
    "last_out_in_window": 68996.5,
    "overall_takt_in_window": 66.60361681329424,  # (68996.5 - 861.0) / 1023
    "target_batch_time": 69823.0,
    "target_batch_actual_finish_in_window": 80010.5,
    "max_finish_full_batch": 83308.0,
    "total_actual_wait_at_cutoff": 199015.5,
    "total_bottleneck_wait_at_cutoff": 3126.0,
    "same_resource_processing_overlap": 0,
    "station_slot_overlap": 0,
    "shared_resource_arrival_order_violation": 0,
    "first_station_launch_skip": 0,
}

ROWS_SHA256_LOCK = {
    "quantity": "1e586079387868fa5fc2bb0770a65672cf2080985b20b04b3938129ede60f2fb",
    "ratio": "4e8b24f9fd609f4c36481bff96d3ad0a4844b7227efde24180d4164a069010d8",
}

# 规范化指标JSON哈希在权威值复核后填入，后续轮必须同时通过rows锁和指标锁。
SUMMARY_SHA256_LOCK = {
    "quantity": "55c788272c900139d246963353dc34e42d4f318d8f78da54d37f06f3091ca2f8",
    "ratio": "c9dcd80156728f7d653d863bf25a0621c91fff226b7c823e7bef31c10d49b197",
}


# ---------------------------------------------------------------------------
# 场景构建
# ---------------------------------------------------------------------------


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


def _run_quantity_scenario():
    parsed = parse_multi_project_inputs({
        "project": "v2.9真实数量场景",
        "cars_a": 276,
        "cars_b": 306,
        "cars_c": 566,
        "target_takt": 58,
        "is_ratio_mode": False,
        "sequence_mode_index": 1,
        "max_consecutive": 5,
        "station_rows": _raw_station_rows(LEGACY.station_defs()),
    })
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
    return parsed, rows, max_finish, analysis


def _run_ratio_scenario():
    parsed = parse_multi_project_inputs({
        "project": "v2.9真实比例场景",
        "cars_a": 2,
        "cars_b": 3,
        "cars_c": 5,
        "analysis_minutes": 1150,
        "target_takt": 58,
        "is_ratio_mode": True,
        "max_consecutive": 5,
        "station_rows": _raw_station_rows(LEGACY.station_defs()),
    })
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
    return parsed, rows, max_finish, analysis


# ---------------------------------------------------------------------------
# 汇总指标 (从 rows / analysis 提取)
# ---------------------------------------------------------------------------


def _summary_quantity(rows, max_finish, analysis):
    grouped = LEGACY.group_rows(rows)
    out_times = sorted(float(g[-1]["depart"]) for g in grouped.values())
    first_out = out_times[0]
    last_out = out_times[-1]
    total_actual_wait = sum(
        float(seg.get("block_wait", 0.0)) for car_rows in grouped.values() for seg in car_rows
    )
    service_checks = ROUND3._service_checks(rows)
    overlap_summary = LEGACY.station_slot_overlap_summary(rows)
    return {
        "generated": len(grouped),
        "finished_full_batch": len(out_times),
        "first_out": first_out,
        "last_out_full_batch": last_out,
        "overall_takt_full_batch": (last_out - first_out) / (len(out_times) - 1),
        "total_actual_wait_full_batch": total_actual_wait,
        "total_bottleneck_wait_full_batch": float(
            analysis["summary"].get("total_bottleneck_wait", 0.0)
        ),
        "same_resource_processing_overlap": LEGACY.processing_capacity_violations(rows),
        "station_slot_overlap": int(overlap_summary["pair_count"]),
        "shared_resource_arrival_order_violation": int(service_checks["fcfs_violations"]),
        "first_station_launch_skip": int(service_checks["launch_overtakes"]),
    }


def _summary_ratio(rows, max_finish, analysis):
    summary = analysis["summary"]
    grouped = LEGACY.group_rows(rows)
    target_takt = 58.0
    analysis_seconds = 69_000.0
    within = [
        float(g[-1]["depart"])
        for g in grouped.values()
        if float(g[-1]["depart"]) <= analysis_seconds + 1e-9
    ]
    within.sort()
    first_in_window = within[0] if within else float("nan")
    last_in_window = within[-1] if within else float("nan")
    overall_in_window = (
        (last_in_window - first_in_window) / (len(within) - 1) if len(within) > 1 else float("nan")
    )
    # 截止时点口径：累加截至69000s已经真实发生的等待，
    # 包含当时仍在等待中的已发生部分，与主界面终值口径一致。
    total_actual_wait_at_cutoff = 0.0
    total_bottleneck_wait_at_cutoff = 0.0
    for segment in rows:
        wait_start = float(segment.get("svc_finish", 0.0) or 0.0)
        wait_end = float(segment.get("depart", wait_start) or wait_start)
        occurred_end = min(analysis_seconds, wait_end)
        actual_wait = max(0.0, occurred_end - wait_start)
        total_actual_wait_at_cutoff += actual_wait

        try:
            capacity = max(1, int(float(
                segment.get("capacity", segment.get("device_count", 1)) or 1
            )))
        except Exception:
            capacity = 1
        excess_start = wait_start + capacity * target_takt
        total_bottleneck_wait_at_cutoff += max(0.0, occurred_end - excess_start)

    service_checks = ROUND3._service_checks(rows)
    overlap_summary = LEGACY.station_slot_overlap_summary(rows)
    return {
        "generated": len(grouped),
        "theoretical_launch_count": int(summary["theoretical_launch_count"]),
        "simulation_buffer_count": int(summary["simulation_buffer_count"]),
        "within_window": len(within),
        "first_out_in_window": first_in_window,
        "last_out_in_window": last_in_window,
        "overall_takt_in_window": overall_in_window,
        "target_batch_time": float(summary["target_batch_time"]),
        "target_batch_actual_finish_in_window": float(summary["actual_finish_time"]),
        "max_finish_full_batch": float(max_finish),
        "total_actual_wait_at_cutoff": total_actual_wait_at_cutoff,
        "total_bottleneck_wait_at_cutoff": total_bottleneck_wait_at_cutoff,
        "same_resource_processing_overlap": LEGACY.processing_capacity_violations(rows),
        "station_slot_overlap": int(overlap_summary["pair_count"]),
        "shared_resource_arrival_order_violation": int(service_checks["fcfs_violations"]),
        "first_station_launch_skip": int(service_checks["launch_overtakes"]),
    }


# ---------------------------------------------------------------------------
# rows SHA-256 锁
# ---------------------------------------------------------------------------


_ROW_LOCK_FIELDS = (
    "car",
    "car_type",
    "step_seq",
    "step_display",
    "start",
    "dur",
    "svc_finish",
    "depart",
    "launch_wait",
    "block_wait",
    "line",
)


def _rows_sha256(rows):
    digest = hashlib.sha256()
    for row in rows:
        record = []
        for field in _ROW_LOCK_FIELDS:
            value = row.get(field, "")
            if isinstance(value, float):
                # 业务时间均为 0.5 秒整数倍；用 1 位小数即可稳定
                record.append(f"{value:.4f}")
            else:
                record.append(str(value))
        digest.update("|".join(record).encode("utf-8"))
        digest.update(b"\n")
    return digest.hexdigest()


def _summary_sha256(summary):
    payload = json.dumps(
        summary,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        allow_nan=False,
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


# ---------------------------------------------------------------------------
# 断言
# ---------------------------------------------------------------------------


def _assert_authoritative(name, actual, expected):
    diffs = []
    for key, exp in expected.items():
        act = actual.get(key)
        if isinstance(exp, float):
            ok = act is not None and math.isclose(float(act), exp, rel_tol=0.0, abs_tol=1e-6)
        else:
            ok = act == exp
        if not ok:
            diffs.append({"key": key, "expected": exp, "actual": act})
    if diffs:
        raise AssertionError(f"{name} 权威基线断言失败：{json.dumps(diffs, ensure_ascii=False)}")


def _assert_rows_lock(name, actual_hash):
    expected = ROWS_SHA256_LOCK.get(name, "")
    if not expected:
        return False
    if actual_hash != expected:
        raise AssertionError(
            f"{name} rows SHA-256 漂移：expected={expected} actual={actual_hash}"
        )
    return True


def _assert_summary_lock(name, actual_hash):
    expected = SUMMARY_SHA256_LOCK.get(name, "")
    if not expected:
        return False
    if actual_hash != expected:
        raise AssertionError(
            f"{name} summary SHA-256 漂移：expected={expected} actual={actual_hash}"
        )
    return True


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--write-report",
        action="store_true",
        help="将权威指标与 SHA-256 锁结果写入 reports/v2.9_round16c_authoritative_baseline.json",
    )
    args = parser.parse_args()

    _, q_rows, q_max, q_analysis = _run_quantity_scenario()
    q_summary = _summary_quantity(q_rows, q_max, q_analysis)
    q_hash = _rows_sha256(q_rows)
    q_summary_hash = _summary_sha256(q_summary)

    _, r_rows, r_max, r_analysis = _run_ratio_scenario()
    r_summary = _summary_ratio(r_rows, r_max, r_analysis)
    r_hash = _rows_sha256(r_rows)
    r_summary_hash = _summary_sha256(r_summary)

    _assert_authoritative("数量场景", q_summary, AUTHORITATIVE_QUANTITY)
    _assert_authoritative("比例场景", r_summary, AUTHORITATIVE_RATIO)
    q_lock_passed = _assert_rows_lock("quantity", q_hash)
    r_lock_passed = _assert_rows_lock("ratio", r_hash)
    q_summary_lock_passed = _assert_summary_lock("quantity", q_summary_hash)
    r_summary_lock_passed = _assert_summary_lock("ratio", r_summary_hash)

    report = {
        "baseline_version": "v2.9-round16c",
        "waiting_scope": {
            "authoritative": "截止69000s时已经真实发生的等待，包含当时仍在等待中的已发生部分",
            "not_authoritative_finished_segments_only": 198904.5,
            "not_authoritative_finished_vehicles_only": 197230.5,
        },
        "quantity": {
            "summary": q_summary,
            "rows_sha256": q_hash,
            "rows_lock_passed": q_lock_passed,
            "summary_sha256": q_summary_hash,
            "summary_lock_passed": q_summary_lock_passed,
        },
        "ratio": {
            "summary": r_summary,
            "rows_sha256": r_hash,
            "rows_lock_passed": r_lock_passed,
            "summary_sha256": r_summary_hash,
            "summary_lock_passed": r_summary_lock_passed,
        },
    }
    print(json.dumps(report, ensure_ascii=False, indent=2))

    if args.write_report:
        out = APP_DIR / "reports" / "v2.9_round16c_authoritative_baseline.json"
        out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
