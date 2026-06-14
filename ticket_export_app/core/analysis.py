# -*- coding: utf-8 -*-
"""
排程结果分析模块。

职责边界：
- 只负责分析 schedule rows，不负责排程、不负责导出、不负责 UI。
- v2-5 开始承接等待、节拍、溢出等统计逻辑。
- tickets.py 后续只保留兼容入口，避免继续膨胀。
"""

from __future__ import annotations

import math
from collections import defaultdict
from typing import Any, Dict, Iterable, List


def _to_float(value: Any, default: float = 0.0) -> float:
    """安全转 float。"""
    try:
        if value is None or value == "":
            return default
        return float(value)
    except Exception:
        return default


def _to_int(value: Any, default: int = 0) -> int:
    """安全转 int。"""
    try:
        if value is None or value == "":
            return default
        return int(float(value))
    except Exception:
        return default


def _car_key(row: Dict[str, Any]) -> int:
    return _to_int(row.get("car", 0), 0)


def _row_start(row: Dict[str, Any]) -> float:
    return _to_float(row.get("start", row.get("begin", 0.0)), 0.0)


def _row_depart(row: Dict[str, Any]) -> float:
    return _to_float(row.get("depart", row.get("end", row.get("svc_finish", 0.0))), 0.0)


def _row_wait(row: Dict[str, Any]) -> float:
    """当前排程行等待时间。"""
    block_wait = _to_float(row.get("block_wait", 0.0), 0.0)
    launch_wait = _to_float(row.get("launch_wait", 0.0), 0.0)
    return max(0.0, block_wait) + max(0.0, launch_wait)


def _row_block_wait(row: Dict[str, Any]) -> float:
    """当前工序完成后，因为下一工序/资源未释放产生的等待。"""
    return max(0.0, _to_float(row.get("block_wait", 0.0), 0.0))


def _row_launch_wait(row: Dict[str, Any]) -> float:
    """车辆进入当前工序前，因为当前工序/资源未释放产生的等待。"""
    return max(0.0, _to_float(row.get("launch_wait", 0.0), 0.0))


def _row_duration(row: Dict[str, Any]) -> float:
    """当前排程行加工时间。"""
    return max(0.0, _to_float(row.get("dur", row.get("duration", 0.0)), 0.0))


def _row_station(row: Dict[str, Any]) -> str:
    """兼容读取岗位/工程名称。"""
    return str(row.get("step_display", row.get("station", row.get("group", ""))) or "")


def _row_line_scope(row: Dict[str, Any]) -> str:
    """兼容读取所属线别/资源范围。"""
    return str(row.get("line_scope", row.get("line", "")) or "")


def _row_capacity(row: Dict[str, Any]) -> int:
    """兼容读取设备数量/资源容量。"""
    return max(1, _to_int(row.get("capacity", row.get("device_count", 1)), 1))


def _row_wait_split(row: Dict[str, Any], target_takt: float) -> Dict[str, float]:
    """
    将等待时间拆分为：
    - total_wait：原始等待，包含所有等待累计，仅作参考。
    - absorbed_wait：节拍内可吸收等待，即工时低于目标节拍时，可被节拍余量消化的部分。
    - overflow_wait：节拍外溢出等待，即超过节拍余量后真正造成节拍损失的等待。
    """
    wait = _row_wait(row)
    dur = _row_duration(row)

    if target_takt <= 0:
        return {
            "total_wait": wait,
            "absorbed_wait": 0.0,
            "overflow_wait": 0.0,
            "process_over_takt": 0.0,
        }

    absorb_capacity = max(0.0, target_takt - dur)
    absorbed_wait = min(wait, absorb_capacity)
    overflow_wait = max(0.0, wait - absorbed_wait)
    process_over_takt = max(0.0, dur - target_takt)

    return {
        "total_wait": wait,
        "absorbed_wait": absorbed_wait,
        "overflow_wait": overflow_wait,
        "process_over_takt": process_over_takt,
    }


def _group_rows_by_car(rows: Iterable[Dict[str, Any]]) -> Dict[int, List[Dict[str, Any]]]:
    by_car: Dict[int, List[Dict[str, Any]]] = defaultdict(list)
    for row in rows:
        car = _car_key(row)
        if car <= 0:
            continue
        by_car[car].append(row)

    for car_rows in by_car.values():
        car_rows.sort(key=lambda item: (_row_start(item), _row_depart(item), _to_int(item.get("step_seq", 0))))
    return dict(by_car)


def _compute_car_finish_times(rows: Iterable[Dict[str, Any]]) -> Dict[int, float]:
    """每台车最终完成时间。"""
    finish_by_car: Dict[int, float] = {}
    for row in rows:
        car = _car_key(row)
        if car <= 0:
            continue
        finish_by_car[car] = max(finish_by_car.get(car, 0.0), _row_depart(row))
    return finish_by_car


def _planned_fill_car_key(row: Dict[str, Any]) -> int:
    for key in ("car", "car_no", "car_index", "idx"):
        car = _to_int(row.get(key, 0), 0)
        if car > 0:
            return car
    return 0


def _planned_fill_row_duration(row: Dict[str, Any]) -> float:
    duration = _to_float(row.get("dur", row.get("duration", 0.0)), 0.0)
    if duration > 0:
        return duration

    start = _to_float(row.get("start", row.get("start_time", row.get("begin", 0.0))), 0.0)
    for key in ("svc_finish", "finish", "end"):
        finish = _to_float(row.get(key, 0.0), 0.0)
        if finish > start:
            return max(0.0, finish - start)
    return 0.0


def _compute_planned_fill_time_from_rows(rows: Iterable[Dict[str, Any]]) -> float:
    """首台车经过所有有效工位的标准加工时间合计，不包含等待。"""
    rows = list(rows or [])
    first_car = 0
    for row in rows:
        car = _planned_fill_car_key(row)
        if car <= 0:
            continue
        first_car = car if first_car <= 0 else min(first_car, car)

    if first_car <= 0:
        return 0.0

    first_car_rows = [
        row for row in rows
        if _planned_fill_car_key(row) == first_car
    ]
    first_car_rows.sort(
        key=lambda item: (
            _to_int(item.get("step_seq", item.get("seq", 0)), 0),
            _to_float(item.get("start", item.get("start_time", item.get("begin", 0.0))), 0.0),
        )
    )

    return sum(_planned_fill_row_duration(row) for row in first_car_rows)


def _compute_wait_summary(rows: List[Dict[str, Any]], target_takt: float) -> Dict[str, Any]:
    """
    等待统计。

    total_wait 仍保留为原始等待累计，便于追溯。
    overflow_wait_time 才是 v2-5 后续重点关注的“节拍外等待”。
    """
    splits = [_row_wait_split(row, target_takt) for row in rows]
    total_wait = sum(item["total_wait"] for item in splits)
    absorbed_wait_time = sum(item["absorbed_wait"] for item in splits)
    overflow_wait_time = sum(item["overflow_wait"] for item in splits)
    process_over_takt_time = sum(item["process_over_takt"] for item in splits)

    finish_by_car = _compute_car_finish_times(rows)
    car_count = len(finish_by_car)
    avg_wait = total_wait / car_count if car_count else 0.0
    avg_overflow_wait = overflow_wait_time / car_count if car_count else 0.0
    max_row_wait = max((item["total_wait"] for item in splits), default=0.0)
    max_row_overflow_wait = max((item["overflow_wait"] for item in splits), default=0.0)

    total_wait_equivalent_cars = total_wait / target_takt if target_takt > 0 else 0.0
    overflow_wait_equivalent_cars = overflow_wait_time / target_takt if target_takt > 0 else 0.0

    return {
        "total_wait": total_wait,
        "avg_wait": avg_wait,
        "average_wait": avg_wait,
        "max_wait": max_row_wait,
        "car_count": car_count,
        "absorbed_wait_time": absorbed_wait_time,
        "overflow_wait_time": overflow_wait_time,
        "avg_overflow_wait": avg_overflow_wait,
        "max_overflow_wait": max_row_overflow_wait,
        "process_over_takt_time": process_over_takt_time,
        "total_wait_equivalent_cars": total_wait_equivalent_cars,
        "overflow_wait_equivalent_cars": overflow_wait_equivalent_cars,
    }


def _compute_station_summary(rows: List[Dict[str, Any]], target_takt: float) -> List[Dict[str, Any]]:
    """按岗位汇总节拍/等待。"""
    grouped: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for row in rows:
        station = _row_station(row)
        grouped[station].append(row)

    result: List[Dict[str, Any]] = []
    for station, station_rows in grouped.items():
        durations = [_row_duration(row) for row in station_rows]
        wait_splits = [_row_wait_split(row, target_takt) for row in station_rows]
        waits = [item["total_wait"] for item in wait_splits]
        overflow_waits = [item["overflow_wait"] for item in wait_splits]
        avg_dur = sum(durations) / len(durations) if durations else 0.0
        max_dur = max(durations, default=0.0)
        total_wait = sum(waits)
        overflow_wait_time = sum(overflow_waits)
        avg_overflow_wait = overflow_wait_time / len(station_rows) if station_rows else 0.0
        over_takt_count = sum(1 for dur in durations if target_takt > 0 and dur > target_takt)

        result.append({
            "station": station,
            "step_display": station,
            "count": len(station_rows),
            "avg_duration": avg_dur,
            "max_duration": max_dur,
            "total_wait": total_wait,
            "overflow_wait_time": overflow_wait_time,
            "blocking_time": overflow_wait_time,
            "avg_overflow_wait": avg_overflow_wait,
            "over_takt_count": over_takt_count,
            "status": "NG" if over_takt_count > 0 or overflow_wait_time > 0 else "OK",
        })

    result.sort(key=lambda item: item.get("station", ""))
    return result


def _compute_car_type_summary(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """按车型汇总。"""
    by_type: Dict[str, set[int]] = defaultdict(set)
    finish_by_type: Dict[str, List[float]] = defaultdict(list)
    finish_by_car = _compute_car_finish_times(rows)

    for row in rows:
        car = _car_key(row)
        if car <= 0:
            continue
        car_type = str(row.get("car_type", row.get("duration_source", "A")) or "A").upper()
        by_type[car_type].add(car)

    for car, finish_time in finish_by_car.items():
        car_rows = [row for row in rows if _car_key(row) == car]
        if not car_rows:
            continue
        car_type = str(car_rows[0].get("car_type", car_rows[0].get("duration_source", "A")) or "A").upper()
        finish_by_type[car_type].append(finish_time)

    result: List[Dict[str, Any]] = []
    for car_type in sorted(by_type.keys()):
        finishes = finish_by_type.get(car_type, [])
        result.append({
            "car_type": car_type,
            "count": len(by_type[car_type]),
            "last_finish": max(finishes, default=0.0),
        })
    return result


def compute_car_capacity_results(
    rows: Iterable[Dict[str, Any]], target_takt: float
) -> List[Dict[str, Any]]:
    """按唯一能力口径判断每台车，并保留超过目标的工程明细。"""
    target = max(0.0, _to_float(target_takt, 0.0))
    results: List[Dict[str, Any]] = []
    for car, car_rows in sorted(_group_rows_by_car(rows).items()):
        over_stations: List[Dict[str, Any]] = []
        car_type = str(car_rows[0].get("car_type", "") or "") if car_rows else ""
        for row in car_rows:
            duration = _row_duration(row)
            if duration <= 0:
                continue
            capacity = _row_capacity(row)
            capacity_limit = capacity * target
            capacity_takt = duration / capacity
            if target > 0 and capacity_takt > target + 1e-9:
                over_stations.append({
                    "station": _row_station(row),
                    "duration": duration,
                    "effective_capacity": capacity,
                    "capacity_limit": capacity_limit,
                    "capacity_takt": capacity_takt,
                    "over_time": duration - capacity_limit,
                })
        results.append({
            "car": car,
            "car_type": car_type,
            "meets_capacity_target": target > 0 and not over_stations,
            "capacity_status": "能力满足" if target > 0 and not over_stations else (
                "未设定目标" if target <= 0 else "能力超目标"
            ),
            "over_capacity_stations": over_stations,
            "over_capacity_station_text": "、".join(
                item["station"] for item in over_stations
            ) or "无",
        })
    return results


# === Station profile helpers for blocking root-cause analysis ===

def _build_station_profiles(rows: List[Dict[str, Any]], target_takt: float) -> Dict[str, Dict[str, Any]]:
    """
    汇总每个工程的基础能力特征，用于过滤“误归因”的阻塞工程。

    当前业务口径：
    - 真正显示为阻塞工程，至少应存在该工程自身工时超目标节拍的风险。
    - 仅因为车辆在进入该工程前发生瞬时等待，不代表该工程就是根因瓶颈。
    """
    profiles: Dict[str, Dict[str, Any]] = {}
    for row in rows:
        station = _row_station(row)
        if not station:
            continue

        profile = profiles.setdefault(
            station,
            {
                "station": station,
                "max_duration": 0.0,
                "min_capacity": 999999,
                "line_scopes": set(),
                "resource_keys": set(),
                "has_process_over_takt": False,
            },
        )

        dur = _row_duration(row)
        profile["max_duration"] = max(float(profile["max_duration"]), dur)
        profile["min_capacity"] = min(int(profile["min_capacity"]), _row_capacity(row))
        profile["line_scopes"].add(_row_line_scope(row))
        profile["resource_keys"].add(str(row.get("resource_key", "") or ""))

        if target_takt > 0 and dur > target_takt:
            profile["has_process_over_takt"] = True

    return profiles


def _is_plausible_blocking_root(profile: Dict[str, Any], target_takt: float) -> bool:
    """
    判断一个被归因的工程是否应该显示为“阻塞工程”。

    先采用保守规则：必须该工程自身存在工时超目标节拍。
    这样可以避免把“节拍内、双线、双设备”的接收工程误显示为根因。
    """
    if not profile:
        return False
    if target_takt <= 0:
        return False
    return bool(profile.get("has_process_over_takt", False))


def _next_processing_row(
    car_rows: List[Dict[str, Any]], current_index: int
) -> Dict[str, Any] | None:
    """Find the next row where this vehicle actually requires processing."""
    for next_row in car_rows[current_index + 1:]:
        if _row_duration(next_row) <= 0:
            continue
        if _row_station(next_row):
            return next_row
    return None


def _next_route_row(
    car_rows: List[Dict[str, Any]], current_index: int
) -> Dict[str, Any] | None:
    """Return the vehicle's immediate next route node, including zero-duration nodes."""
    next_index = current_index + 1
    return car_rows[next_index] if next_index < len(car_rows) else None


def _next_processing_station(car_rows: List[Dict[str, Any]], current_index: int) -> str:
    """Find the next station where this vehicle actually requires processing."""
    next_row = _next_processing_row(car_rows, current_index)
    if next_row is not None:
        return _row_station(next_row)
    return _row_station(car_rows[current_index]) if current_index < len(car_rows) else ""


def _compute_wait_classification(
    rows: List[Dict[str, Any]], target_takt: float
) -> Dict[str, Any]:
    """Split in-line actual waiting from waiting beyond station holding capacity.

    对外显示术语定版（2026-06-13，UI/报表第6轮按此呈现；内部英文键名保持不变以兼容）：
      - total_bottleneck_wait / bottleneck_wait      -> 「节拍外等待」= max(0, 实际等待 − 可接纳上限)
      - holding_limit                                 -> 「可接纳上限」= 有效设备数 × 目标节拍
      - actual_wait / total_actual_wait               -> 「实际等待 / 累计实际等待」
      - downstream_station /
        bottleneck_wait_by_downstream_station         -> 「下一工程」（仅路线，非根因，也不代表它超节拍）
    不对外使用“瓶颈阻塞 / 根因工程 / 阻塞根因”等措辞；最终远端根因需后续“根因事件记录”轮次。
    """
    actual_by_station: Dict[str, float] = defaultdict(float)
    bottleneck_by_station: Dict[str, float] = defaultdict(float)
    bottleneck_by_downstream_station: Dict[str, float] = defaultdict(float)
    total_launch_wait = 0.0
    total_actual_wait = 0.0
    total_bottleneck_wait = 0.0
    bottleneck_events: List[Dict[str, Any]] = []

    for car_rows in _group_rows_by_car(rows).values():
        for index, row in enumerate(car_rows):
            waiting_station = _row_station(row)
            total_launch_wait += _row_launch_wait(row)
            block_wait = _row_block_wait(row)
            if block_wait <= 0:
                continue

            total_actual_wait += block_wait
            if waiting_station:
                actual_by_station[waiting_station] += block_wait

            holding_limit = max(0.0, target_takt) * _row_capacity(row)
            bottleneck_wait = max(0.0, block_wait - holding_limit)
            if bottleneck_wait <= 0:
                continue

            downstream_row = _next_route_row(car_rows, index)
            downstream_station = (
                _row_station(downstream_row) if downstream_row is not None else ""
            )
            total_bottleneck_wait += bottleneck_wait
            if waiting_station:
                bottleneck_by_station[waiting_station] += bottleneck_wait
            if downstream_station:
                bottleneck_by_downstream_station[downstream_station] += bottleneck_wait
            bottleneck_events.append(
                {
                    "car": _car_key(row),
                    "car_type": str(row.get("car_type", "") or ""),
                    "waiting_station": waiting_station,
                    "downstream_station": downstream_station,
                    "actual_wait": block_wait,
                    "holding_capacity": _row_capacity(row),
                    "holding_limit": holding_limit,
                    "bottleneck_wait": bottleneck_wait,
                }
            )

    def station_items(values: Dict[str, float]) -> List[Dict[str, Any]]:
        return [
            {"station": station, "wait_time": wait_time}
            for station, wait_time in sorted(
                values.items(), key=lambda item: (-item[1], item[0])
            )
            if wait_time > 0
        ]

    return {
        "total_actual_wait": total_actual_wait,
        "actual_wait_by_station": station_items(actual_by_station),
        "total_raw_flow_wait": total_actual_wait,
        "raw_launch_wait": total_launch_wait,
        "raw_post_process_wait": total_actual_wait,
        "raw_flow_wait_by_station": station_items(actual_by_station),
        "total_bottleneck_wait": total_bottleneck_wait,
        "bottleneck_launch_wait": 0.0,
        "bottleneck_post_process_wait": total_bottleneck_wait,
        "bottleneck_wait_by_station": station_items(bottleneck_by_station),
        "bottleneck_wait_by_downstream_station": station_items(
            bottleneck_by_downstream_station
        ),
        "bottleneck_wait_event_count": len(bottleneck_events),
        "bottleneck_wait_events": bottleneck_events,
    }


def _compute_capacity_over_takt_summary(
    rows: List[Dict[str, Any]], target_takt: float
) -> Dict[str, Any]:
    """Summarize processing time beyond each row's capacity-based takt limit."""
    grouped: Dict[tuple[str, str], Dict[str, Any]] = {}
    if target_takt <= 0:
        return {
            "capacity_over_takt_processes": [],
            "capacity_over_takt_total_time": 0.0,
            "capacity_over_takt_station_text": "无",
        }

    for row in rows:
        station = _row_station(row)
        if not station:
            continue
        duration = _row_duration(row)
        capacity = _row_capacity(row)
        capacity_limit = capacity * target_takt
        over_takt_time = max(0.0, duration - capacity_limit)
        if over_takt_time <= 0:
            continue
        car_type = str(
            row.get("car_type", row.get("duration_source", "")) or ""
        ).upper()
        key = (station, car_type)
        item = grouped.setdefault(
            key,
            {
                "station": station,
                "car_type": car_type,
                "count": 0,
                "duration": duration,
                "effective_capacity": capacity,
                "capacity_limit": capacity_limit,
                "single_over_takt": over_takt_time,
                "total_over_takt": 0.0,
            },
        )
        item["count"] += 1
        item["duration"] = max(float(item["duration"]), duration)
        item["single_over_takt"] = max(
            float(item["single_over_takt"]), over_takt_time
        )
        item["total_over_takt"] += over_takt_time

    processes = sorted(
        grouped.values(),
        key=lambda item: (-float(item["total_over_takt"]), item["station"], item["car_type"]),
    )
    total_time = sum(float(item["total_over_takt"]) for item in processes)
    station_text = "、".join(
        f"{item['station']}{item['car_type']}超{item['single_over_takt']:.1f}s"
        for item in processes
    ) or "无"
    return {
        "capacity_over_takt_processes": processes,
        "capacity_over_takt_total_time": total_time,
        "capacity_over_takt_station_text": station_text,
    }


def _compute_wait_cause_chain_summary(
    rows: List[Dict[str, Any]], target_takt: float
) -> Dict[str, Any]:
    """Aggregate scheduler-recorded cause slices for waiting beyond holding limits."""
    grouped: Dict[tuple[str, str, str, str], Dict[str, Any]] = {}
    details: List[Dict[str, Any]] = []
    explained_time = 0.0
    incomplete_time = 0.0

    for row in rows:
        actual_wait = _row_block_wait(row)
        if actual_wait <= 0:
            continue
        waiting_station = _row_station(row)
        holding_limit = max(0.0, target_takt) * _row_capacity(row)
        excess_start = _to_float(row.get("svc_finish"), 0.0) + holding_limit
        excess_end = _to_float(row.get("depart"), excess_start)
        if excess_end <= excess_start:
            continue

        event_keys: set[tuple[str, str, str, str]] = set()
        covered_time = 0.0
        for cause_slice in row.get("wait_cause_slices", []) or []:
            slice_start = max(excess_start, _to_float(cause_slice.get("start"), excess_start))
            slice_end = min(excess_end, _to_float(cause_slice.get("end"), excess_end))
            duration = max(0.0, slice_end - slice_start)
            if duration <= 0:
                continue
            covered_time += duration

            chain = list(cause_slice.get("chain", []) or [])
            direct = chain[0] if chain else {}
            direct_station = str(direct.get("blocked_station", "") or "未知工程")
            direct_type = str(direct.get("block_type", "") or "unresolved")
            terminal_type = str(cause_slice.get("terminal_type", "") or "unresolved")
            terminal_station = str(cause_slice.get("terminal_station", "") or "未知工程")
            terminal_resource = str(cause_slice.get("terminal_resource", "") or "")
            terminal_car_type = str(cause_slice.get("terminal_car_type", "") or "")
            complete = bool(cause_slice.get("chain_complete", False))

            if not complete:
                terminal_label = "原因链不完整"
                incomplete_time += duration
            elif terminal_type == "over_takt_processing":
                type_suffix = f" {terminal_car_type}" if terminal_car_type else ""
                terminal_label = f"{terminal_station}{type_suffix}超节拍加工占用"
                explained_time += duration
            else:
                resource_label = terminal_station or terminal_resource or "未知资源"
                terminal_label = f"{resource_label}前车占用"
                explained_time += duration

            key = (waiting_station, direct_station, terminal_label, direct_type)
            item = grouped.setdefault(
                key,
                {
                    "waiting_station": waiting_station,
                    "direct_blocking_station": direct_station,
                    "terminal_cause": terminal_label,
                    "direct_block_type": direct_type,
                    "event_count": 0,
                    "wait_time": 0.0,
                    "chain_complete": complete,
                },
            )
            item["wait_time"] += duration
            event_keys.add(key)
            details.append(
                {
                    "car": _car_key(row),
                    "car_type": str(row.get("car_type", "") or ""),
                    "waiting_station": waiting_station,
                    "wait_start": slice_start,
                    "wait_end": slice_end,
                    "wait_time": duration,
                    "actual_wait": actual_wait,
                    "holding_limit": holding_limit,
                    "direct_blocker_car": direct.get("blocker_car"),
                    "direct_blocking_station": direct_station,
                    "direct_blocking_resource": str(direct.get("blocked_resource", "") or ""),
                    "direct_block_type": direct_type,
                    "terminal_car": cause_slice.get("terminal_car"),
                    "terminal_station": terminal_station,
                    "terminal_resource": terminal_resource,
                    "terminal_type": terminal_type,
                    "terminal_cause": terminal_label,
                    "chain_complete": complete,
                    "chain": chain,
                }
            )
        for key in event_keys:
            grouped[key]["event_count"] += 1

        expected_time = excess_end - excess_start
        if covered_time + 1e-9 < expected_time:
            missing_time = expected_time - covered_time
            incomplete_time += missing_time
            missing_key = (waiting_station, "未知工程", "原因链不完整", "unresolved")
            missing_item = grouped.setdefault(
                missing_key,
                {
                    "waiting_station": waiting_station,
                    "direct_blocking_station": "未知工程",
                    "terminal_cause": "原因链不完整",
                    "direct_block_type": "unresolved",
                    "event_count": 0,
                    "wait_time": 0.0,
                    "chain_complete": False,
                },
            )
            missing_item["event_count"] += 1
            missing_item["wait_time"] += missing_time
            details.append(
                {
                    "car": _car_key(row),
                    "car_type": str(row.get("car_type", "") or ""),
                    "waiting_station": waiting_station,
                    "wait_start": excess_start + covered_time,
                    "wait_end": excess_end,
                    "wait_time": missing_time,
                    "actual_wait": actual_wait,
                    "holding_limit": holding_limit,
                    "direct_blocker_car": None,
                    "direct_blocking_station": "未知工程",
                    "direct_blocking_resource": "",
                    "direct_block_type": "unresolved",
                    "terminal_car": None,
                    "terminal_station": "",
                    "terminal_resource": "",
                    "terminal_type": "unresolved",
                    "terminal_cause": "原因链不完整",
                    "chain_complete": False,
                    "chain": [],
                }
            )

    summary = sorted(
        grouped.values(),
        key=lambda item: (-float(item["wait_time"]), -int(item["event_count"]), item["waiting_station"], item["terminal_cause"]),
    )
    total_time = explained_time + incomplete_time
    coverage_rate = explained_time / total_time if total_time > 0 else 1.0
    return {
        "wait_cause_chain_summary": summary,
        "wait_cause_chain_details": details,
        "wait_cause_chain_total_time": total_time,
        "wait_cause_chain_explained_time": explained_time,
        "wait_cause_chain_incomplete_time": incomplete_time,
        "wait_cause_chain_coverage_rate": coverage_rate,
    }


def _compute_blocking_summary(rows: List[Dict[str, Any]], target_takt: float) -> Dict[str, Any]:
    """
    v2-5 主业务口径：阻塞根因分析。

    口径说明：
    - launch_wait：车辆进入当前工序前等待，根因归属当前工序。
    - block_wait：车辆完成当前工序后等待下游接收，根因沿路径追到下一个
      工时大于0的真实作业工位；中间0工时节点只表示车辆停留位置。
    - 只统计超过目标节拍余量后仍无法被吸收的等待，即节拍外阻塞工时。
    """
    blocking_by_station: Dict[str, float] = defaultdict(float)
    attribution_by_pair: Dict[tuple[str, str], Dict[str, Any]] = {}
    attributed_launch_wait_time = 0.0
    attributed_post_process_wait_time = 0.0
    station_profiles = _build_station_profiles(rows, target_takt)
    by_car = _group_rows_by_car(rows)

    for car_rows in by_car.values():
        for idx, row in enumerate(car_rows):
            dur = _row_duration(row)
            absorb_capacity = max(0.0, target_takt - dur) if target_takt > 0 else 0.0

            launch_wait = _row_launch_wait(row)
            block_wait = _row_block_wait(row)

            remaining_absorb = absorb_capacity

            absorbed_launch = min(launch_wait, remaining_absorb)
            launch_overflow = max(0.0, launch_wait - absorbed_launch)
            remaining_absorb = max(0.0, remaining_absorb - absorbed_launch)

            absorbed_block = min(block_wait, remaining_absorb)
            block_overflow = max(0.0, block_wait - absorbed_block)

            current_station = _row_station(row)
            if launch_overflow > 0 and current_station:
                blocking_by_station[current_station] += launch_overflow
                attributed_launch_wait_time += launch_overflow

            if block_overflow > 0:
                immediate_station = (
                    _row_station(car_rows[idx + 1])
                    if idx + 1 < len(car_rows)
                    else current_station
                )
                root_station = _next_processing_station(car_rows, idx)
                if root_station:
                    blocking_by_station[root_station] += block_overflow
                    attributed_post_process_wait_time += block_overflow
                    pair = (current_station, root_station)
                    item = attribution_by_pair.setdefault(
                        pair,
                        {
                            "source_station": current_station,
                            "immediate_next_station": immediate_station,
                            "root_station": root_station,
                            "event_count": 0,
                            "blocking_time": 0.0,
                            "skipped_zero_duration_node": immediate_station != root_station,
                        },
                    )
                    item["event_count"] += 1
                    item["blocking_time"] += block_overflow

    blocking_stations: List[Dict[str, Any]] = []
    for station, blocking_time in sorted(blocking_by_station.items()):
        if blocking_time <= 0:
            continue

        profile = station_profiles.get(station, {})
        if not _is_plausible_blocking_root(profile, target_takt):
            continue

        equivalent_cars = blocking_time / target_takt if target_takt > 0 else 0.0
        blocking_stations.append({
            "station": station,
            "blocking_time": blocking_time,
            "overflow_vehicle_count": equivalent_cars,
            "max_duration": profile.get("max_duration", 0.0),
        })

    total_blocking_time = sum(item["blocking_time"] for item in blocking_stations)
    attributed_blocking_time = sum(blocking_by_station.values())
    overflow_vehicle_count = total_blocking_time / target_takt if target_takt > 0 else 0.0
    blocking_root_attributions = sorted(
        attribution_by_pair.values(),
        key=lambda item: (-float(item["blocking_time"]), item["root_station"], item["source_station"]),
    )
    attributed_blocking_stations = [
        {"station": station, "blocking_time": blocking_time}
        for station, blocking_time in sorted(
            blocking_by_station.items(),
            key=lambda item: (-item[1], item[0]),
        )
        if blocking_time > 0
    ]

    if blocking_stations:
        blocking_station_text = "、".join(item["station"] for item in blocking_stations)
        blocking_station_detail_text = "、".join(
            f"{item['station']} {item['blocking_time']:.1f}s"
            for item in blocking_stations
        )
        blocking_result = "有阻塞"
    else:
        blocking_station_text = "无"
        blocking_station_detail_text = "无"
        blocking_result = "无阻塞"

    return {
        "blocking_result": blocking_result,
        "blocking_station_count": len(blocking_stations),
        "blocking_stations": blocking_stations,
        "blocking_station_text": blocking_station_text,
        "blocking_station_detail_text": blocking_station_detail_text,
        "total_blocking_time": total_blocking_time,
        "overflow_vehicle_count": overflow_vehicle_count,
        "blocking_time": total_blocking_time,
        "attributed_blocking_time": attributed_blocking_time,
        "attributed_launch_wait_time": attributed_launch_wait_time,
        "attributed_post_process_wait_time": attributed_post_process_wait_time,
        "attributed_blocking_stations": attributed_blocking_stations,
        "blocking_root_attributions": blocking_root_attributions,
        "overflow_wait_time": total_blocking_time,
        "overflow_wait_equivalent_cars": overflow_vehicle_count,
        "wait_equivalent_cars": overflow_vehicle_count,
    }


# === Batch overrun & process root-cause analysis ===

def _compute_batch_overrun_summary(
    car_count: int,
    max_finish: float,
    target_takt: float,
) -> Dict[str, Any]:
    """
    批次节拍溢出分析。

    口径说明：
    - 该指标不等同于等待/阻塞。
    - 即使没有 launch_wait / block_wait，只要实际完成时间超过“总台数 × 目标节拍”，
      也说明当前模型存在批次完成时间溢出。
    """
    target_batch_time = car_count * target_takt if target_takt > 0 else 0.0
    actual_finish_time = max(0.0, max_finish)
    batch_overrun_time = max(0.0, actual_finish_time - target_batch_time)
    batch_overrun_cars = batch_overrun_time / target_takt if target_takt > 0 else 0.0

    return {
        "target_batch_time": target_batch_time,
        "actual_finish_time": actual_finish_time,
        "batch_overrun_time": batch_overrun_time,
        "batch_overrun_cars": batch_overrun_cars,
        "batch_overrun_result": "有溢出" if batch_overrun_time > 0 else "无溢出",
    }


def _compute_process_over_takt_roots(
    rows: List[Dict[str, Any]],
    target_takt: float,
) -> Dict[str, Any]:
    """
    工时型根因分析。

    口径说明：
    - 当没有等待阻塞，但批次总完成时间仍超出目标节拍时，通常是工序自身工时超节拍导致。
    - 按“工程 + 车型”统计单次工时超出量和累计超出量。
    """
    if target_takt <= 0:
        return {
            "process_over_takt_roots": [],
            "process_over_takt_root_text": "无",
            "process_over_takt_total_time": 0.0,
        }

    grouped: Dict[tuple[str, str], Dict[str, Any]] = {}
    for row in rows:
        station = _row_station(row)
        if not station:
            continue
        car_type = str(row.get("car_type", row.get("duration_source", "")) or "").upper()
        dur = _row_duration(row)
        over = max(0.0, dur - target_takt)
        if over <= 0:
            continue

        key = (station, car_type)
        item = grouped.setdefault(
            key,
            {
                "station": station,
                "car_type": car_type,
                "count": 0,
                "max_duration": 0.0,
                "single_over_takt": over,
                "total_over_takt": 0.0,
            },
        )
        item["count"] += 1
        item["max_duration"] = max(float(item["max_duration"]), dur)
        item["single_over_takt"] = max(float(item["single_over_takt"]), over)
        item["total_over_takt"] += over

    roots = sorted(
        grouped.values(),
        key=lambda item: float(item.get("total_over_takt", 0.0)),
        reverse=True,
    )
    total_time = sum(float(item.get("total_over_takt", 0.0)) for item in roots)

    if roots:
        root_text = "、".join(
            f"{item['station']}{item['car_type']}超{item['single_over_takt']:.1f}s"
            for item in roots[:3]
        )
    else:
        root_text = "无"

    return {
        "process_over_takt_roots": roots,
        "process_over_takt_root_text": root_text,
        "process_over_takt_total_time": total_time,
    }


def _compute_overflow_summary(
    rows: List[Dict[str, Any]],
    max_finish: float,
    target_takt: float,
    total_wait: float,
    overflow_wait_time: float,
) -> Dict[str, Any]:
    """
    v2-5 溢出分析。

    当前第一版不新增 UI 输入项，默认以实际总完成时间 max_finish 作为观察时间窗口。
    注意：wait_equivalent_cars 从本版开始代表“节拍外溢出等待折算台数”，不再使用原始总等待折算。
    """
    finish_by_car = _compute_car_finish_times(rows)
    time_window_seconds = max(0.0, _to_float(max_finish, 0.0))

    if target_takt > 0 and time_window_seconds > 0:
        theory_output_count = int(time_window_seconds // target_takt)
    else:
        theory_output_count = 0

    actual_output_count = sum(1 for finish_time in finish_by_car.values() if finish_time <= time_window_seconds)
    overflow_count = max(0, theory_output_count - actual_output_count)

    total_wait_equivalent_cars = total_wait / target_takt if target_takt > 0 else 0.0
    overflow_wait_equivalent_cars = overflow_wait_time / target_takt if target_takt > 0 else 0.0

    return {
        "time_window_seconds": time_window_seconds,
        "theory_output_count": theory_output_count,
        "actual_output_count": actual_output_count,
        "overflow_count": overflow_count,
        "total_wait_equivalent_cars": total_wait_equivalent_cars,
        "overflow_wait_equivalent_cars": overflow_wait_equivalent_cars,
        "wait_equivalent_cars": overflow_wait_equivalent_cars,
    }


def analyze_schedule_v2(
    rows: List[Dict[str, Any]],
    max_finish: float,
    target_takt: float | None = None,
) -> Dict[str, Any]:
    """
    v2 分析入口。

    返回结构尽量兼容 tickets.py 原有 UI 调用，同时新增 v2-5 溢出分析字段。
    """
    rows = list(rows or [])
    target = _to_float(target_takt, 0.0)
    max_finish_value = _to_float(max_finish, 0.0)

    wait_summary = _compute_wait_summary(rows, target)
    wait_classification = _compute_wait_classification(rows, target)
    capacity_over_takt_summary = _compute_capacity_over_takt_summary(rows, target)
    wait_cause_chain_summary = _compute_wait_cause_chain_summary(rows, target)
    station_summary = _compute_station_summary(rows, target)
    car_type_summary = _compute_car_type_summary(rows)
    car_capacity_results = compute_car_capacity_results(rows, target)

    blocking_summary = _compute_blocking_summary(rows, target)
    batch_overrun_summary = _compute_batch_overrun_summary(
        wait_summary["car_count"],
        max_finish_value,
        target,
    )
    process_root_summary = _compute_process_over_takt_roots(rows, target)

    overflow_summary = _compute_overflow_summary(
        rows,
        max_finish_value,
        target,
        wait_summary["total_wait"],
        blocking_summary["total_blocking_time"],
    )

    summary = {
        "max_finish": max_finish_value,
        "target_takt": target,
        **wait_summary,
        **wait_classification,
        **capacity_over_takt_summary,
        **wait_cause_chain_summary,
        **overflow_summary,
        **blocking_summary,
        **batch_overrun_summary,
        **process_root_summary,
    }

    return {
        "summary": summary,
        "stations": station_summary,
        "station_summary": station_summary,
        "car_types": car_type_summary,
        "car_type_summary": car_type_summary,
        "car_capacity_results": car_capacity_results,
        "rows": rows,
    }


def apply_time_window_analysis(
    analysis,
    rows,
    target_takt,
    analysis_time_seconds,
    theoretical_launch_count,
):
    """Apply v2-6D line takt time-window analysis without touching scheduling."""
    if not isinstance(analysis, dict):
        return analysis

    try:
        analysis_time_seconds = float(analysis_time_seconds or 0.0)
        theoretical_launch_count = int(theoretical_launch_count or 0)
        target_takt = float(target_takt or 0.0)
    except Exception:
        return analysis

    if not analysis_time_seconds or not theoretical_launch_count or target_takt <= 0:
        return analysis

    summary = analysis.setdefault("summary", {})

    car_finish_times = {}
    station_names = []
    seen_stations = set()

    for row in rows or []:
        station = str(row.get("step_display", row.get("station", row.get("group", ""))) or "")
        if station and station not in seen_stations:
            station_names.append(station)
            seen_stations.add(station)

        try:
            car = int(row.get("car", 0) or 0)
        except Exception:
            car = 0

        if car <= 0:
            continue

        finish = 0.0
        for key in ("depart", "end", "svc_finish"):
            try:
                finish = max(finish, float(row.get(key, 0.0) or 0.0))
            except Exception:
                pass

        car_finish_times[car] = max(car_finish_times.get(car, 0.0), finish)

    finish_times = sorted(car_finish_times.values())
    actual_output_count = sum(
        1 for finish in finish_times
        if finish <= analysis_time_seconds + 1e-9
    )

    station_count = max(1, len(station_names))
    planned_fill_time = _compute_planned_fill_time_from_rows(rows or [])
    if planned_fill_time <= 0:
        planned_fill_time = station_count * target_takt
    line_lead_time = planned_fill_time
    if analysis_time_seconds + 1e-9 < line_lead_time:
        planned_output_count = 0
    else:
        planned_output_count = math.floor(
            (analysis_time_seconds - line_lead_time) / target_takt
        ) + 1

    display_actual_output_count = (
        min(actual_output_count, planned_output_count)
        if planned_output_count > 0
        else actual_output_count
    )

    planned_n_finish_time = None
    actual_n_finish_time = None
    finish_delta = None
    actual_line_takt = None

    if planned_output_count <= 0:
        achievement_rate = 0.0
        final_result = "未判定"
    else:
        achievement_rate = min(actual_output_count / planned_output_count, 1.0)
        planned_n_finish_time = line_lead_time + (planned_output_count - 1) * target_takt

        if len(finish_times) >= planned_output_count:
            actual_n_finish_time = finish_times[planned_output_count - 1]
            finish_delta = actual_n_finish_time - planned_n_finish_time
            if planned_output_count > 1:
                actual_line_takt = target_takt + finish_delta / (planned_output_count - 1)
            else:
                actual_line_takt = target_takt

        if actual_output_count < planned_output_count:
            final_result = "NG"
        elif actual_n_finish_time is None:
            final_result = "NG"
        elif actual_n_finish_time > planned_n_finish_time + 1e-9:
            final_result = "NG"
        else:
            final_result = "OK"

    summary.update({
        "analysis_time_seconds": analysis_time_seconds,
        "analysis_time_minutes": analysis_time_seconds / 60.0,
        "theoretical_launch_count": theoretical_launch_count,
        "station_count": station_count,
        "line_lead_time": line_lead_time,
        "planned_fill_time": planned_fill_time,
        "planned_output_count_in_window": planned_output_count,
        "actual_output_count_in_window": actual_output_count,
        "display_actual_output_count_in_window": display_actual_output_count,
        "actual_equivalent_count_in_window": display_actual_output_count,
        "achievement_rate": achievement_rate,
        "planned_n_finish_time": planned_n_finish_time,
        "actual_n_finish_time": actual_n_finish_time,
        "finish_delta": finish_delta,
        "actual_line_takt_in_window": actual_line_takt,
        "actual_production_takt_in_window": actual_line_takt,
        "time_window_result": final_result,
    })

    summary.pop("ok_output_count_in_window", None)
    summary.pop("output_gap_count", None)
    summary.pop("output_gap_time", None)
    summary.pop("actual_output_takt_in_window", None)
    summary.pop("completed_step_count_in_window", None)
    summary.pop("planned_step_count_in_window", None)
    summary.pop("planned_equivalent_count_raw_in_window", None)
    summary.pop("actual_equivalent_count_raw_in_window", None)
    summary.pop("entered_step_count_in_window", None)
    summary.pop("actual_takt_in_window", None)

    return analysis
