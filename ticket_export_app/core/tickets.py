# -*- coding: utf-8 -*-
"""
组合票排程 + 导出
当前阶段说明：
- 多工程组合票已切换为『运行方式 + 岗位/设备 + A/B/C 工时』录入模型。
- v2-1 支持固定投车节拍；v2-2 支持车型工时为 0 时保留流转节点；v2-3B 建立资源 key；v2-4A 建立自动线别分配基础；v2-4B 支持线别连续性与强制线别；v2-4C 支持未来强制线别预判。
- 当前排程/导出逻辑会按车辆类型使用 A/B/C 对应工时；若 B/C 未填写则继承 A 工时。
- 运行方式会自动推导资源能力：
    * 单线单设备 -> capacity = 1
    * 双线单设备 -> capacity = 1
    * 双线双设备 -> capacity = 2
- Zone / gate_zone / gate_buffer 的旧逻辑仍保留兼容，便于后续继续升级复杂等待规则。
"""

from __future__ import annotations
import math
import heapq
import hashlib
from collections import defaultdict
from typing import List, Dict, Any, Tuple, Optional
import pandas as pd
from core.analysis import analyze_schedule_v2


# ---------------- Excel 引擎选择 ---------------- #
def _choose_engine():
    try:
        import xlsxwriter  # noqa: F401
        return "xlsxwriter"
    except Exception:
        try:
            import openpyxl  # noqa: F401
            return "openpyxl"
        except Exception:
            return None


# ---------------- 解析步骤与 Zone ---------------- #
def _normalize_defs(step_defs: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], Dict[str, Dict[str, Any]], Dict[str, int]]:
    """
    返回 (steps, zones, gate_buffers)

    steps: 每步 {
        seq, display, group, duration, duration_a, duration_b, duration_c,
        run_mode, capacity, device_count, line_scope, resource_key, zone_id, gate_zone_id
    }
    zones: {zid: {"capacity": int, "first_seq": int, "last_seq": int}}
    gate_buffers: {zid: gate_buffer_int}  # 若未显式提供，默认=2
    """
    steps: List[Dict[str, Any]] = []
    gate_buffers: Dict[str, int] = {}

    for d in step_defs:
        display = str(d.get("display", "")).strip()
        group = str(d.get("group", "")).strip() or display
        run_mode = str(d.get("run_mode", "单线单设备") or "单线单设备").strip()
        line_scope = str(d.get("line_scope", "") or "").strip()
        if not line_scope:
            if run_mode == "双线双设备":
                line_scope = "双线"
            elif run_mode == "双线单设备":
                line_scope = "双线共用"
            else:
                line_scope = "1号线"

        # 兼容新旧两种工时来源：
        # 1) 新结构：duration_a / duration_b / duration_c
        # 2) 旧结构：durations 列表（第一项作为 A 工时，B/C 缺省时继承 A）
        durations = list(d.get("durations", []))
        duration_a_raw = d.get("duration_a", None)
        duration_b_raw = d.get("duration_b", None)
        duration_c_raw = d.get("duration_c", None)

        if duration_a_raw is None:
            if durations:
                duration_a_raw = durations[0]
            else:
                duration_a_raw = None

        if not display or duration_a_raw is None:
            continue

        try:
            duration_a = float(duration_a_raw)
        except Exception:
            continue

        try:
            duration_b = float(duration_b_raw) if duration_b_raw not in (None, "") else duration_a
        except Exception:
            duration_b = duration_a

        try:
            duration_c = float(duration_c_raw) if duration_c_raw not in (None, "") else duration_a
        except Exception:
            duration_c = duration_a

        # capacity 允许显式传入；若未传，则由运行方式自动推导
        cap_raw = d.get("capacity", None)
        if cap_raw is None:
            capacity = 2 if run_mode == "双线双设备" else 1
        else:
            try:
                capacity = max(1, int(float(cap_raw)))
            except Exception:
                capacity = 2 if run_mode == "双线双设备" else 1

        device_count_raw = d.get("device_count", None)
        if device_count_raw is None:
            device_count = 2 if run_mode == "双线双设备" else 1
        else:
            try:
                device_count = max(1, int(float(device_count_raw)))
            except Exception:
                device_count = 2 if run_mode == "双线双设备" else 1
        if device_count not in (1, 2):
            device_count = 2 if device_count >= 2 else 1

        # v2-3B：资源 key。
        # resource_key 用于表达“同一个物理设备/资源”的占用关系。
        # 当前阶段只建立 key 并按 key 共享资源堆；自动择线留到 v2-4。
        if line_scope == "1号线":
            resource_key = f"{group}::1号线"
        elif line_scope == "2号线":
            resource_key = f"{group}::2号线"
        elif line_scope == "双线共用":
            resource_key = f"{group}::双线共用"
        else:
            resource_key = f"{group}::双线"

        zone_id = str(d.get("zone_id", "") or "").strip()
        gate_zone_id = str(d.get("gate_zone_id", "") or "").strip()

        # 聚合 gate_buffer
        if gate_zone_id:
            gb = d.get("gate_buffer", None)
            if gb is None:
                gb = 2
            try:
                gb = max(1, int(float(gb)))
            except Exception:
                gb = 2
            if gate_zone_id in gate_buffers:
                gate_buffers[gate_zone_id] = max(gate_buffers[gate_zone_id], gb)
            else:
                gate_buffers[gate_zone_id] = gb

        steps.append({
            "seq": int(d.get("seq", len(steps) + 1)),
            "display": display,
            "group": group,
            "duration": duration_a,
            "duration_a": duration_a,
            "duration_b": duration_b,
            "duration_c": duration_c,
            "run_mode": run_mode,
            "capacity": capacity,
            "device_count": device_count,
            "line_scope": line_scope,
            "resource_key": resource_key,
            "zone_id": zone_id,
            "gate_zone_id": gate_zone_id,
        })

    steps.sort(key=lambda x: x["seq"])
    if not steps:
        raise ValueError("没有有效的步骤定义")

    zones: Dict[str, Dict[str, Any]] = {}
    for s in steps:
        zid = s.get("zone_id", "")
        if not zid:
            continue
        z = zones.setdefault(zid, {"capacity": 1, "first_seq": s["seq"], "last_seq": s["seq"]})
        z["first_seq"] = min(z["first_seq"], s["seq"])
        z["last_seq"] = max(z["last_seq"], s["seq"])

    for d in step_defs:
        zid = str(d.get("zone_id", "") or "").strip()
        if not zid or zid not in zones:
            continue
        zcap = d.get("zone_capacity", None)
        if zcap is not None:
            try:
                zones[zid]["capacity"] = max(int(zones[zid]["capacity"]), int(zcap))
            except Exception:
                pass

    return steps, zones, gate_buffers


QUANTITY_SEQUENCE_RULE_VERSION = "quantity-balanced-run-tau1-v1"
QUANTITY_SEQUENCE_LAG_TOLERANCE = 1.0


def vehicle_sequence_hash(vehicle_sequence: List[str]) -> str:
    """Return the stable SHA-256 used to identify a frozen vehicle sequence."""
    return hashlib.sha256("".join(vehicle_sequence).encode("utf-8")).hexdigest()


def build_vehicle_sequence(cars: int,
                           vehicle_counts: Dict[str, int] | None = None,
                           sequence_mode: str = "grouped",
                           max_consecutive: int = 10,
                           ratio_pattern: Optional[Dict[str, int]] = None) -> List[str]:
    """Build the deterministic vehicle sequence used by schedule and sequence freeze."""
    vehicle_seq: List[str] = []
    if vehicle_counts:
        counts: Dict[str, int] = {}
        for vehicle_type in ("A", "B", "C"):
            try:
                counts[vehicle_type] = max(0, int(vehicle_counts.get(vehicle_type, 0) or 0))
            except Exception:
                counts[vehicle_type] = 0

        max_consecutive = max(1, int(max_consecutive or 1))

        if sequence_mode == "alternate":
            # R9: let the current type run toward the configured limit, but insert a
            # lagging type as soon as its cumulative deficit reaches one vehicle.
            target_counts = dict(counts)
            remaining = dict(counts)
            placed = {vehicle_type: 0 for vehicle_type in ("A", "B", "C")}
            total_cars = sum(target_counts.values())
            order = {"A": 0, "B": 1, "C": 2}
            last_type = ""
            run_len = 0

            for position in range(1, total_cars + 1):
                deficits = {
                    vehicle_type: (
                        target_counts[vehicle_type] / total_cars * position
                        - placed[vehicle_type]
                    )
                    for vehicle_type in ("A", "B", "C")
                }
                can_continue = (
                    bool(last_type)
                    and remaining[last_type] > 0
                    and run_len < max_consecutive
                )
                another_type_is_lagging = any(
                    vehicle_type != last_type
                    and remaining[vehicle_type] > 0
                    and deficits[vehicle_type] >= QUANTITY_SEQUENCE_LAG_TOLERANCE - 1e-12
                    for vehicle_type in ("A", "B", "C")
                )

                if can_continue and not another_type_is_lagging:
                    pick = last_type
                else:
                    other_type_remains = bool(last_type) and any(
                        remaining[vehicle_type] > 0 and vehicle_type != last_type
                        for vehicle_type in ("A", "B", "C")
                    )
                    candidates = [
                        vehicle_type
                        for vehicle_type in ("A", "B", "C")
                        if remaining[vehicle_type] > 0
                        and not (
                            vehicle_type == last_type
                            and run_len >= max_consecutive
                            and other_type_remains
                        )
                    ]
                    pick = max(
                        candidates,
                        key=lambda vehicle_type: (
                            round(deficits[vehicle_type], 12),
                            remaining[vehicle_type],
                            -order[vehicle_type],
                        ),
                    )

                vehicle_seq.append(pick)
                placed[pick] += 1
                remaining[pick] -= 1
                if pick == last_type:
                    run_len += 1
                else:
                    last_type = pick
                    run_len = 1
        elif sequence_mode == "ratio":
            # Ratio mode remains the existing fixed-block behavior in R9.
            pattern = ratio_pattern or {}
            a = max(0, int(pattern.get("A", 0) or 0))
            b = max(0, int(pattern.get("B", 0) or 0))
            c = max(0, int(pattern.get("C", 0) or 0))
            block: List[str] = (["A"] * a) + (["B"] * b) + (["C"] * c)
            if not block:
                raise ValueError("按比例运行模式下，比例块不能为空")
            total_cars = max(0, int(cars))
            while len(vehicle_seq) < total_cars:
                need = total_cars - len(vehicle_seq)
                vehicle_seq.extend(block[:need])
        else:
            for vehicle_type in ("A", "B", "C"):
                vehicle_seq.extend([vehicle_type] * counts[vehicle_type])

    if not vehicle_seq:
        vehicle_seq = ["A" for _ in range(max(0, int(cars)))]
    return vehicle_seq


def _schedule_arrival_fcfs(steps: List[Dict[str, Any]],
                           vehicle_seq: List[str],
                           launch_takt: float) -> Tuple[List[Dict[str, Any]], float]:
    """Event-based blocking flow scheduler with arrival-first resource service."""

    def pick_duration(step: Dict[str, Any], car_type: str) -> float:
        source = str(car_type or "A").strip().upper()
        if source == "B":
            return float(step.get("duration_b", step["duration"]))
        if source == "C":
            return float(step.get("duration_c", step["duration"]))
        return float(step.get("duration_a", step["duration"]))

    routes = {
        car: [index for index, step in enumerate(steps) if pick_duration(step, car_type) >= 0]
        for car, car_type in enumerate(vehicle_seq, start=1)
    }

    def upcoming_forced_line(start_index: int, car_type: str) -> str:
        for future_step in steps[start_index:]:
            if pick_duration(future_step, car_type) <= 0:
                continue
            scope = str(future_step.get("line_scope", "") or "")
            if scope in ("1号线", "2号线"):
                return scope
        return ""

    resource_tokens: Dict[str, List[str]] = {}
    for step in steps:
        key = str(step.get("resource_key", "") or "")
        if key in resource_tokens:
            continue
        scope = str(step.get("line_scope", "") or "")
        capacity = max(1, int(step.get("capacity", 1) or 1))
        if scope == "双线":
            resource_tokens[key] = (["1号线", "2号线"] + [
                f"{index + 3}号线" for index in range(max(0, capacity - 2))
            ])[:capacity]
        else:
            resource_tokens[key] = [f"资源{index + 1}" for index in range(capacity)]

    resource_owner = {
        (key, token): None
        for key, tokens in resource_tokens.items()
        for token in tokens
    }
    station_slot_owner: Dict[Tuple[int, str], Optional[int]] = {}
    line_assignment_counts: Dict[Tuple[int, str], int] = defaultdict(int)
    waiting_by_step: Dict[int, List[Tuple[float, int]]] = {
        index: [] for index in range(len(steps))
    }
    events: List[Tuple[float, int, int]] = []  # time, kind(0=finish/1=launch), car
    states: Dict[int, Dict[str, Any]] = {}
    for car, car_type in enumerate(vehicle_seq, start=1):
        launch_time = (car - 1) * launch_takt if launch_takt > 0 else 0.0
        states[car] = {
            "car_type": car_type,
            "route_pos": 0,
            "current_row": None,
            "current_step": None,
            "current_line": "",
            "resource_token": None,
            "done": False,
        }
        heapq.heappush(events, (launch_time, 1, car))

    def candidate_lines(step_index: int, current_line: str, car_type: str) -> List[str]:
        scope = str(steps[step_index].get("line_scope", "") or "")
        if scope in ("1号线", "2号线"):
            return [scope]
        if current_line in ("1号线", "2号线"):
            return [current_line]
        forced = upcoming_forced_line(step_index + 1, car_type)
        return [forced] if forced else ["1号线", "2号线"]

    def free_resource_token(step: Dict[str, Any], line_no: str) -> Optional[Tuple[str, str]]:
        key = str(step.get("resource_key", "") or "")
        if str(step.get("line_scope", "") or "") == "双线":
            token = (key, line_no)
            return token if resource_owner.get(token) is None else None
        for name in resource_tokens.get(key, []):
            token = (key, name)
            if resource_owner.get(token) is None:
                return token
        return None

    def downstream_line_score(car: int, step_index: int, line_no: str, now: float):
        """Prefer the line whose next physical slot can receive the car first."""
        state = states[car]
        route = routes[car]
        next_route_pos = state["route_pos"] + 1
        if next_route_pos >= len(route):
            return (0.0, 0, line_assignment_counts[(step_index, line_no)], line_no)

        next_step_index = route[next_route_pos]
        owner = station_slot_owner.get((next_step_index, line_no))
        if owner is None:
            slot_delay = 0.0
        else:
            owner_row = states[owner].get("current_row")
            finish = float(owner_row["svc_finish"]) if owner_row is not None else now
            # A vehicle that has finished but is still in the slot is blocked by
            # its downstream station, so that line is not currently releasable.
            slot_delay = finish - now if finish > now else float("inf")

        queued = sum(
            1
            for _, waiting_car in waiting_by_step[next_step_index]
            if states[waiting_car].get("current_line") == line_no
        )
        return (
            slot_delay,
            queued,
            line_assignment_counts[(step_index, line_no)],
            line_no,
        )

    def available_assignment(car: int, step_index: int, now: float):
        state = states[car]
        step = steps[step_index]
        duration = pick_duration(step, state["car_type"])
        candidates = candidate_lines(step_index, state["current_line"], state["car_type"])
        if len(candidates) > 1:
            candidates.sort(key=lambda line: downstream_line_score(car, step_index, line, now))
        for line_no in candidates:
            if station_slot_owner.get((step_index, line_no)) is not None:
                continue
            token = None if duration <= 0 else free_resource_token(step, line_no)
            if duration > 0 and token is None:
                continue
            return line_no, token, duration
        return None

    def release_current(car: int, depart: float) -> None:
        state = states[car]
        row = state["current_row"]
        step_index = state["current_step"]
        if row is None or step_index is None:
            return
        row["depart"] = depart
        row["block_wait"] = max(0.0, depart - float(row["svc_finish"]))
        station_slot_owner[(step_index, row["line_no"])] = None
        if state["resource_token"] is not None:
            resource_owner[state["resource_token"]] = None
        state["current_row"] = None
        state["current_step"] = None
        state["resource_token"] = None

    def enqueue_next(car: int, ready_time: float) -> None:
        state = states[car]
        route = routes[car]
        if state["route_pos"] >= len(route):
            release_current(car, ready_time)
            state["done"] = True
            return
        heapq.heappush(waiting_by_step[route[state["route_pos"]]], (ready_time, car))

    def choose_waiting(step_index: int, now: float):
        queue = waiting_by_step[step_index]
        if not queue:
            return None
        scope = str(steps[step_index].get("line_scope", "") or "")
        # Shared/single resources are strict FCFS. Dual-line stations may start the
        # earliest eligible car on either independent line.
        candidate_indexes = [0] if scope != "双线" else range(len(queue))
        best = None
        for index in candidate_indexes:
            ready_time, car = queue[index]
            assignment = available_assignment(car, step_index, now)
            if assignment is None:
                if scope != "双线":
                    return None
                continue
            item = (ready_time, car, index, assignment)
            if best is None or item[:2] < best[:2]:
                best = item
        return best

    def waiting_step_for_car(car: int) -> Optional[int]:
        state = states[car]
        route = routes[car]
        route_pos = int(state.get("route_pos", 0))
        if route_pos >= len(route):
            return None
        step_index = route[route_pos]
        if any(waiting_car == car for _, waiting_car in waiting_by_step[step_index]):
            return step_index
        return None

    def direct_wait_blocker(car: int, step_index: int, now: float) -> Dict[str, Any]:
        """Describe the concrete object preventing a waiting car from advancing."""
        step = steps[step_index]
        state = states[car]
        queue = waiting_by_step[step_index]
        scope = str(step.get("line_scope", "") or "")
        if scope != "双线" and queue:
            ready_time, first_car = min(queue)
            if ready_time <= now + epsilon and first_car != car:
                return {
                    "blocker_car": first_car,
                    "block_type": "queue_order",
                    "blocked_station": str(step.get("display", "") or ""),
                    "blocked_resource": str(step.get("resource_key", "") or ""),
                }

        duration = pick_duration(step, state["car_type"])
        candidates = candidate_lines(step_index, state.get("current_line", ""), state["car_type"])
        blockers: List[Dict[str, Any]] = []
        for line_no in candidates:
            slot_owner = station_slot_owner.get((step_index, line_no))
            if slot_owner is not None and slot_owner != car:
                blockers.append({
                    "blocker_car": slot_owner,
                    "block_type": "station_slot",
                    "blocked_station": str(step.get("display", "") or ""),
                    "blocked_resource": f"{step.get('display', '')}::{line_no}车位",
                })
                continue
            if duration <= 0:
                continue
            key = str(step.get("resource_key", "") or "")
            tokens = (
                [(key, line_no)]
                if scope == "双线"
                else [(key, token_name) for token_name in resource_tokens.get(key, [])]
            )
            owners = sorted({resource_owner.get(token) for token in tokens if resource_owner.get(token) is not None})
            for owner in owners:
                if owner == car:
                    continue
                blockers.append({
                    "blocker_car": owner,
                    "block_type": "processing_resource",
                    "blocked_station": str(step.get("display", "") or ""),
                    "blocked_resource": key,
                })

        if not blockers:
            return {
                "blocker_car": None,
                "block_type": "unresolved",
                "blocked_station": str(step.get("display", "") or ""),
                "blocked_resource": str(step.get("resource_key", "") or ""),
            }
        blockers.sort(key=lambda item: (int(item["blocker_car"]), item["block_type"], item["blocked_resource"]))
        return blockers[0]

    def wait_cause_chain(car: int, step_index: int, now: float) -> Dict[str, Any]:
        """Follow direct blockers until the active terminal occupation is reached."""
        chain: List[Dict[str, Any]] = []
        visited = {car}
        current_car = car
        current_step = step_index
        while True:
            direct = direct_wait_blocker(current_car, current_step, now)
            blocker_car = direct.get("blocker_car")
            chain.append({
                **direct,
                "waiting_car": current_car,
            })
            if blocker_car is None:
                return {
                    "chain": chain,
                    "chain_complete": False,
                    "terminal_type": "unresolved",
                    "terminal_car": None,
                    "terminal_station": direct.get("blocked_station", ""),
                    "terminal_resource": direct.get("blocked_resource", ""),
                }
            blocker_car = int(blocker_car)
            if blocker_car in visited:
                return {
                    "chain": chain,
                    "chain_complete": False,
                    "terminal_type": "cycle",
                    "terminal_car": blocker_car,
                    "terminal_station": direct.get("blocked_station", ""),
                    "terminal_resource": direct.get("blocked_resource", ""),
                }
            visited.add(blocker_car)
            blocker_state = states[blocker_car]
            blocker_row = blocker_state.get("current_row")
            blocker_wait_step = waiting_step_for_car(blocker_car)
            if blocker_row is None:
                return {
                    "chain": chain,
                    "chain_complete": False,
                    "terminal_type": "unresolved",
                    "terminal_car": blocker_car,
                    "terminal_station": direct.get("blocked_station", ""),
                    "terminal_resource": direct.get("blocked_resource", ""),
                }
            if blocker_wait_step is not None and float(blocker_row.get("svc_finish", 0.0)) <= now + epsilon:
                current_car = blocker_car
                current_step = blocker_wait_step
                continue

            duration = float(blocker_row.get("dur", 0.0) or 0.0)
            capacity = max(1, int(blocker_row.get("capacity", 1) or 1))
            capacity_limit = capacity * launch_takt if launch_takt > 0 else 0.0
            over_takt = launch_takt > 0 and duration > capacity_limit + epsilon
            terminal_type = "over_takt_processing" if over_takt else "resource_occupation"
            return {
                "chain": chain,
                "chain_complete": True,
                "terminal_type": terminal_type,
                "terminal_car": blocker_car,
                "terminal_car_type": blocker_state.get("car_type", ""),
                "terminal_station": str(blocker_row.get("step_display", "") or ""),
                "terminal_resource": str(blocker_row.get("resource_key", "") or direct.get("blocked_resource", "")),
                "terminal_start": float(blocker_row.get("start", now) or now),
                "terminal_finish": float(blocker_row.get("svc_finish", now) or now),
                "terminal_duration": duration,
                "terminal_capacity_limit": capacity_limit,
            }

    def append_wait_cause_slice(row: Dict[str, Any], start: float, end: float, cause: Dict[str, Any]) -> None:
        if end <= start + epsilon:
            return
        slices = row.setdefault("wait_cause_slices", [])
        signature = (
            tuple((item.get("waiting_car"), item.get("blocker_car"), item.get("block_type"), item.get("blocked_station"), item.get("blocked_resource")) for item in cause.get("chain", [])),
            cause.get("chain_complete"), cause.get("terminal_type"), cause.get("terminal_car"),
            cause.get("terminal_station"), cause.get("terminal_resource"),
        )
        if slices and slices[-1].get("signature") == signature and abs(float(slices[-1]["end"]) - start) <= epsilon:
            slices[-1]["end"] = end
            slices[-1]["duration"] = end - float(slices[-1]["start"])
            return
        slices.append({
            "start": start,
            "end": end,
            "duration": end - start,
            **cause,
            "signature": signature,
        })

    def record_wait_cause_slices(start: float, end: float) -> None:
        """Record one immutable cause slice for every vehicle blocked in this interval."""
        for car, state in states.items():
            row = state.get("current_row")
            if row is None or float(row.get("svc_finish", 0.0)) > start + epsilon:
                continue
            step_index = waiting_step_for_car(car)
            if step_index is None:
                continue
            cause = wait_cause_chain(car, step_index, start)
            append_wait_cause_slice(row, start, end, cause)

    rows: List[Dict[str, Any]] = []
    current_time = 0.0
    max_time = 0.0
    completed = 0
    epsilon = 1e-9

    while completed < len(states):
        if events and not any(waiting_by_step.values()):
            current_time = max(current_time, events[0][0])

        while events and events[0][0] <= current_time + epsilon:
            event_time, kind, car = heapq.heappop(events)
            state = states[car]
            if kind == 1:
                enqueue_next(car, event_time)
            else:
                before = state["done"]
                enqueue_next(car, event_time)
                if state["done"] and not before:
                    completed += 1
                    max_time = max(max_time, event_time)

        made_progress = True
        while made_progress:
            made_progress = False
            for step_index in range(len(steps)):
                selected = choose_waiting(step_index, current_time)
                if selected is None:
                    continue
                ready_time, car, queue_index, assignment = selected
                if ready_time > current_time + epsilon:
                    continue
                queue = waiting_by_step[step_index]
                queue.pop(queue_index)
                heapq.heapify(queue)
                line_no, token, duration = assignment
                state = states[car]
                if state["current_row"] is not None:
                    release_current(car, current_time)
                step = steps[step_index]
                station_slot_owner[(step_index, line_no)] = car
                line_assignment_counts[(step_index, line_no)] += 1
                if token is not None:
                    resource_owner[token] = car
                theory_launch = (car - 1) * launch_takt if launch_takt > 0 else 0.0
                row = {
                    "car": car,
                    "step_seq": step["seq"],
                    "step_display": step["display"],
                    "group": step["group"],
                    "run_mode": step.get("run_mode", "单线单设备"),
                    "capacity": int(step.get("capacity", 1) or 1),
                    "device_count": int(step.get("device_count", step.get("capacity", 1)) or 1),
                    "line_scope": str(step.get("line_scope", "") or ""),
                    "resource_key": str(step.get("resource_key", "") or ""),
                    "line_no": line_no,
                    "car_type": state["car_type"],
                    "duration_source": str(state["car_type"] or "A").strip().upper(),
                    "theory_launch_time": theory_launch,
                    "launch_takt": launch_takt,
                    "launch_wait": max(0.0, current_time - theory_launch) if state["route_pos"] == 0 else 0.0,
                    "dur": duration,
                    "start": current_time,
                    "svc_finish": current_time + duration,
                    "depart": current_time + duration,
                    "block_wait": 0.0,
                    "wait_cause_slices": [],
                }
                rows.append(row)
                state["current_row"] = row
                state["current_step"] = step_index
                state["resource_token"] = token
                state["current_line"] = line_no
                state["route_pos"] += 1
                heapq.heappush(events, (current_time + duration, 0, car))
                made_progress = True

        if completed >= len(states):
            break
        if events:
            next_time = events[0][0]
            if next_time <= current_time + epsilon:
                continue
            record_wait_cause_slices(current_time, next_time)
            current_time = next_time
            continue
        blocked = [car for car, state in states.items() if not state["done"]]
        raise RuntimeError(f"排程发生资源死锁，未完成车辆：{blocked[:10]}")

    for row in rows:
        for cause_slice in row.get("wait_cause_slices", []) or []:
            cause_slice.pop("signature", None)
    rows.sort(key=lambda row: (int(row["car"]), int(row["step_seq"])))
    return rows, max_time


# ---------------- 排程入口与历史兼容路径 ---------------- #
def schedule(step_defs: List[Dict[str, Any]],
             cars: int,
             vehicle_counts: Dict[str, int] | None = None,
             sequence_mode: str = "grouped",
             max_consecutive: int = 10,
             ratio_pattern: Optional[Dict[str, int]] = None,
             launch_takt: Optional[float] = None,
             vehicle_sequence: Optional[List[str]] = None) -> Tuple[List[Dict[str, Any]], float]:
    """正式排程入口。

    当前M-Line UI生成的岗位定义不含Zone/gate字段，因此直接进入
    按实际到达顺序服务的事件排程。只有历史外部输入显式传入
    zone_id / gate_zone_id 时，才进入保留的兼容路径。
    """
    steps, zones, gate_buffers = _normalize_defs(step_defs)
    if zones or gate_buffers:
        return _schedule_legacy_zone_gate(
            step_defs=step_defs,
            cars=cars,
            vehicle_counts=vehicle_counts,
            sequence_mode=sequence_mode,
            max_consecutive=max_consecutive,
            ratio_pattern=ratio_pattern,
            launch_takt=launch_takt,
            vehicle_sequence=vehicle_sequence,
        )

    if vehicle_sequence is not None:
        vehicle_seq = [str(vehicle_type).strip().upper() for vehicle_type in vehicle_sequence]
        if not vehicle_seq or any(vehicle_type not in ("A", "B", "C") for vehicle_type in vehicle_seq):
            raise ValueError("冻结排列包含无效车型，必须重新冻结。")
        if vehicle_counts and sequence_mode != "ratio":
            expected = {
                vehicle_type: max(0, int(vehicle_counts.get(vehicle_type, 0) or 0))
                for vehicle_type in ("A", "B", "C")
            }
            actual = {vehicle_type: vehicle_seq.count(vehicle_type) for vehicle_type in ("A", "B", "C")}
            if actual != expected:
                raise ValueError("冻结排列与当前A/B/C数量不一致，必须重新冻结。")
    else:
        vehicle_seq = build_vehicle_sequence(
            cars,
            vehicle_counts,
            sequence_mode,
            max_consecutive,
            ratio_pattern,
        )

    try:
        launch_takt_value = float(launch_takt or 0.0)
    except Exception:
        launch_takt_value = 0.0
    if launch_takt_value < 0:
        launch_takt_value = 0.0

    return _schedule_arrival_fcfs(steps, vehicle_seq, launch_takt_value)


def _schedule_legacy_zone_gate(step_defs: List[Dict[str, Any]],
                               cars: int,
                               vehicle_counts: Dict[str, int] | None = None,
                               sequence_mode: str = "grouped",
                               max_consecutive: int = 10,
                               ratio_pattern: Optional[Dict[str, int]] = None,
                               launch_takt: Optional[float] = None,
                               vehicle_sequence: Optional[List[str]] = None) -> Tuple[List[Dict[str, Any]], float]:
    """
    历史Zone/gate兼容排程。当前M-Line UI不会生成这些字段。

    本函数仅为兼容旧的外部输入保留，不属于v2.9正式业务路径。
    修改正式排程时不应同时改动此兼容实现。

    返回：
      rows: 每车-每步记录：
        {car, step_seq, step_display, group, dur, start, svc_finish, depart, block_wait}
      max_time: 全局最后 depart

    v2-1：固定投车节拍
      launch_takt > 0 时，第 n 台车理论投车时间 = (n - 1) × launch_takt。
      首岗位实际开始时间 = max(理论投车时间, 首岗位资源释放时间)。
      launch_takt 未传或 <= 0 时，保持旧排程逻辑。

    v2-2：车型跳过岗位
      当前车型在当前岗位的工时 < 0 时，表示该车型跳过该岗位。
      工时 = 0 时仍然经过该岗位，不占用加工设备，但占用当前工位的物理流转车位。
      仅负工时岗位不占用资源、不产生等待、不写入 rows。

    v2-3B：设备数量与所属线别资源 key
      当前阶段根据 device_count / line_scope / group 生成 resource_key。
      资源堆从“按步骤”升级为“按资源 key”。
      暂不做自动择线，只为后续 v2-4 做基础。

    v2-4A：自动线别分配基础
      资源堆元素从单纯 ready_time 升级为 (ready_time, line_no)。
      1号线 / 2号线岗位固定线别；双线岗位按最早释放资源自动分配线别。
      当前阶段只记录 line_no，不引入复杂缓冲区与真实线体拓扑。

    v2-4B：线别连续性与强制线别
      1号线 / 2号线岗位为强制线别，即使该线较晚释放，也必须等待该线资源。
      双线岗位若车辆已有当前线别，则优先保持同一线别连续流转。
      双线共用岗位占用共享资源，但车辆实际线别尽量沿用进入前线别。

    v2-4C：未来强制线别预判
      如果车辆后续存在需要进入 1号线 / 2号线的有效岗位，前置双线岗位会提前优先选择该线别。
      只有工时 > 0 的后续岗位才构成强制线别约束。
    """
    steps, zones, gate_buffers = _normalize_defs(step_defs)
    m = len(steps)

    if vehicle_sequence is not None:
        vehicle_seq = [str(vehicle_type).strip().upper() for vehicle_type in vehicle_sequence]
        if not vehicle_seq or any(vehicle_type not in ("A", "B", "C") for vehicle_type in vehicle_seq):
            raise ValueError("冻结排列包含无效车型，必须重新冻结。")
        if vehicle_counts and sequence_mode != "ratio":
            expected = {
                vehicle_type: max(0, int(vehicle_counts.get(vehicle_type, 0) or 0))
                for vehicle_type in ("A", "B", "C")
            }
            actual = {vehicle_type: vehicle_seq.count(vehicle_type) for vehicle_type in ("A", "B", "C")}
            if actual != expected:
                raise ValueError("冻结排列与当前A/B/C数量不一致，必须重新冻结。")
    else:
        vehicle_seq = build_vehicle_sequence(
            cars,
            vehicle_counts,
            sequence_mode,
            max_consecutive,
            ratio_pattern,
        )

    cars = len(vehicle_seq)
    try:
        launch_takt_value = float(launch_takt or 0.0)
    except Exception:
        launch_takt_value = 0.0
    if launch_takt_value < 0:
        launch_takt_value = 0.0

    if not zones and not gate_buffers:
        raise ValueError("历史Zone/gate兼容路径缺少zone_id或gate_zone_id。")
 
    def _initial_resource_slots(st: Dict[str, Any]) -> List[Tuple[float, str]]:
        """根据所属线别生成初始资源槽：(释放时间, 线别)。"""
        line_scope = str(st.get("line_scope", "") or "")
        cap = max(1, int(st.get("capacity", 1) or 1))

        if line_scope == "1号线":
            slots = [(0.0, "1号线") for _ in range(cap)]
        elif line_scope == "2号线":
            slots = [(0.0, "2号线") for _ in range(cap)]
        elif line_scope == "双线":
            slots = [(0.0, "1号线"), (0.0, "2号线")]
            # 理论上双线只支持设备数量=2；若后续扩容，保持补齐逻辑。
            while len(slots) < cap:
                slots.append((0.0, f"{len(slots) + 1}号线"))
        elif line_scope == "双线共用":
            slots = [(0.0, "双线共用") for _ in range(cap)]
        else:
            slots = [(0.0, "1号线") for _ in range(cap)]

        return slots[:cap]

    def _upcoming_forced_line(all_steps: List[Dict[str, Any]], start_index: int, car_type: str) -> str:
        """
        v2-4C：未来强制线别预判。
        如果当前车辆后续还有必须进入 1号线 / 2号线的岗位，
        则前置双线岗位应提前优先选择该线别。
        只有工时 > 0 的岗位才构成强制约束。
        """
        for future_st in all_steps[start_index:]:
            try:
                future_dur = float(_pick_duration(future_st, car_type) or 0.0)
            except Exception:
                future_dur = 0.0

            if future_dur <= 0:
                continue

            future_scope = str(future_st.get("line_scope", "") or "")
            if future_scope in ("1号线", "2号线"):
                return future_scope

        return ""

    def _forced_or_preferred_line(st: Dict[str, Any], current_line_no: str, upcoming_forced_line: str = "") -> str:
        """
        根据岗位所属线别、车辆当前线别、后续强制线别，返回必须/优先使用的线别。
        空字符串表示取最早释放资源。
        """
        line_scope = str(st.get("line_scope", "") or "")

        if line_scope == "1号线":
            return "1号线"
        if line_scope == "2号线":
            return "2号线"

        if line_scope == "双线":
            if current_line_no in ("1号线", "2号线"):
                return current_line_no
            if upcoming_forced_line in ("1号线", "2号线"):
                return upcoming_forced_line

        return ""

    def _pop_resource_slot(
        heap: List[Tuple[float, str]],
        st: Dict[str, Any],
        current_line_no: str,
        upcoming_forced_line: str = "",
        station_ready_by_line: Optional[Dict[str, float]] = None,
    ) -> Tuple[float, str]:
        """按强制线别/线别连续性/未来强制线别预判取出资源槽；无约束时取最早释放。"""
        preferred_line = _forced_or_preferred_line(st, current_line_no, upcoming_forced_line)
        if preferred_line:
            best_idx = -1
            best_item: Tuple[float, str] | None = None
            for idx, item in enumerate(heap):
                ready_time, line_no = item
                if line_no != preferred_line:
                    continue
                if best_item is None or ready_time < best_item[0]:
                    best_idx = idx
                    best_item = item
            if best_idx >= 0 and best_item is not None:
                heap.pop(best_idx)
                heapq.heapify(heap)
                return best_item
        if station_ready_by_line:
            best_idx = min(
                range(len(heap)),
                key=lambda idx: (
                    max(heap[idx][0], station_ready_by_line.get(heap[idx][1], 0.0)),
                    heap[idx][0],
                    heap[idx][1],
                ),
            )
            best_item = heap.pop(best_idx)
            heapq.heapify(heap)
            return best_item
        return heapq.heappop(heap)

    def _peek_resource_slot(
        heap: List[Tuple[float, str]],
        st: Dict[str, Any],
        current_line_no: str,
        upcoming_forced_line: str = "",
        station_ready_by_line: Optional[Dict[str, float]] = None,
    ) -> Tuple[float, str]:
        """按下一岗位线别约束预估资源释放时间和资源线别，不修改资源堆。"""
        if not heap:
            return 0.0, current_line_no or "1号线"
        preferred_line = _forced_or_preferred_line(st, current_line_no, upcoming_forced_line)
        if preferred_line:
            matched = [ready_time for ready_time, line_no in heap if line_no == preferred_line]
            if matched:
                return min(matched), preferred_line
        if station_ready_by_line:
            return min(
                heap,
                key=lambda item: (
                    max(item[0], station_ready_by_line.get(item[1], 0.0)),
                    item[0],
                    item[1],
                ),
            )
        return min(heap)

    def _actual_line_no(st: Dict[str, Any], resource_line_no: str, current_line_no: str, upcoming_forced_line: str = "") -> str:
        """用于 rows 记录的车辆实际线别。双线共用资源尽量沿用进入前线别。"""
        line_scope = str(st.get("line_scope", "") or "")
        if line_scope == "双线共用":
            if current_line_no in ("1号线", "2号线"):
                return current_line_no
            if upcoming_forced_line in ("1号线", "2号线"):
                return upcoming_forced_line
        return resource_line_no
    # v2-4A：按 resource_key 建立资源释放堆，堆元素为 (ready_time, line_no)。
    # resource_key 相同的步骤共享同一套资源。
    resource_heaps: Dict[str, List[Tuple[float, str]]] = {}
    for st in steps:
        key = str(st.get("resource_key", "") or "")
        if not key:
            key = f"{st.get('group', '')}::{st.get('line_scope', '')}"
            st["resource_key"] = key

        cap = max(1, int(st.get("capacity", 1) or 1))
        if key not in resource_heaps:
            heap = _initial_resource_slots(st)
            heapq.heapify(heap)
            resource_heaps[key] = heap
        else:
            # 同一 resource_key 如果后续出现更大容量，则扩容到最大容量。
            # 正常情况下同一 resource_key 的 capacity 应保持一致。
            heap = resource_heaps[key]
            existing_lines = {line for _, line in heap}
            for _, line in _initial_resource_slots(st):
                if len(heap) >= cap:
                    break
                if line not in existing_lines:
                    heapq.heappush(heap, (0.0, line))
                    existing_lines.add(line)
            while len(heap) < cap:
                heapq.heappush(heap, (0.0, f"{len(heap) + 1}号线"))

    # 每个工位、每条实际流转线只有一个物理车位。
    # 正工时车辆同时占加工资源和物理车位；0工时车辆只占物理车位。
    station_slot_ready: Dict[Tuple[int, str], float] = {}

    def _slot_ready_time(step_index: int, line_no: str) -> float:
        return station_slot_ready.get((step_index, line_no or "1号线"), 0.0)

    def _set_slot_ready_time(step_index: int, line_no: str, ready_time: float) -> None:
        station_slot_ready[(step_index, line_no or "1号线")] = ready_time

    # Zone 名额堆：zid -> [free_time, ...]（长度=capacity）
    zone_heaps: Dict[str, List[float]] = {}
    for zid, zinfo in zones.items():
        cap = int(zinfo.get("capacity", 1)) or 1
        zone_heaps[zid] = [0.0 for _ in range(cap)]
        heapq.heapify(zone_heaps[zid])

    # 闸门缓冲：对每个 gate_zone 维护“尚未进入该 zone 的车辆的预计进入时刻”最小堆
    # pre_heap[z] 中的元素是“已经通过闸门但尚未进入 zone 的车辆的 ‘zone 入口开始时间’ ”
    pre_heap: Dict[str, List[float]] = {}

    def is_zone_entry(idx: int) -> bool:
        s = steps[idx]
        zid = s.get("zone_id", "")
        if not zid:
            return False
        return s["seq"] == zones[zid]["first_seq"]

    def is_zone_exit(idx: int) -> bool:
        s = steps[idx]
        zid = s.get("zone_id", "")
        if not zid:
            return False
        return s["seq"] == zones[zid]["last_seq"]

    rows: List[Dict[str, Any]] = []
    max_time = 0.0

    def _pick_duration(st: Dict[str, Any], car_type: str) -> float:
        car_type = str(car_type or "A").strip().upper()
        if car_type == "B":
            return float(st.get("duration_b", st["duration"]))
        if car_type == "C":
            return float(st.get("duration_c", st["duration"]))
        return float(st.get("duration_a", st["duration"]))

    def _find_next_effective_step(
        all_steps: List[Dict[str, Any]],
        current_index: int,
        car_type: str,
    ) -> Tuple[Optional[int], Optional[Dict[str, Any]]]:
        for next_index in range(current_index + 1, len(all_steps)):
            next_step = all_steps[next_index]
            try:
                next_duration = float(_pick_duration(next_step, car_type) or 0.0)
            except Exception:
                next_duration = 0.0
            if next_duration > 0:
                return next_index, next_step
        return None, None

    def _find_next_active_step(
        all_steps: List[Dict[str, Any]],
        current_index: int,
        car_type: str,
    ) -> Tuple[Optional[int], Optional[Dict[str, Any]], float]:
        """
        返回后续第一个“未跳过”的节点：
        - duration < 0：跳过
        - duration = 0：pass-through 节点
        - duration > 0：真实作业节点
        """
        for next_index in range(current_index + 1, len(all_steps)):
            next_step = all_steps[next_index]
            try:
                next_duration = float(_pick_duration(next_step, car_type) or 0.0)
            except Exception:
                next_duration = 0.0
            if next_duration >= 0:
                return next_index, next_step, next_duration
        return None, None, -1.0

    def _pass_through_line(step_index: int, current_line_no: str, car_type: str) -> str:
        """0工时节点不强制线别；优先保持车辆当前线别。"""
        if current_line_no in ("1号线", "2号线"):
            return current_line_no

        next_idx, next_step = _find_next_effective_step(steps, step_index, car_type)
        if next_idx is not None and next_step is not None:
            next_key = str(next_step.get("resource_key", "") or "")
            next_heap = resource_heaps[next_key]
            next_forced_line = _upcoming_forced_line(steps, next_idx + 1, car_type)
            _, resource_line = _peek_resource_slot(
                next_heap,
                next_step,
                current_line_no,
                next_forced_line,
                {
                    line_no: _slot_ready_time(next_idx, _actual_line_no(
                        next_step,
                        line_no,
                        current_line_no,
                        next_forced_line,
                    ))
                    for _, line_no in next_heap
                },
            )
            return _actual_line_no(next_step, resource_line, current_line_no, next_forced_line)

        return "1号线"

    def _next_step_ready(
        current_index: int,
        current_line_no: str,
        car_type: str,
    ) -> float:
        """返回紧邻下一个流转节点可接收当前车辆的最早时间。"""
        next_idx, next_step, next_duration = _find_next_active_step(steps, current_index, car_type)
        if next_idx is None or next_step is None:
            return 0.0

        if next_duration <= 0:
            next_line = _pass_through_line(next_idx, current_line_no, car_type)
            return _slot_ready_time(next_idx, next_line)

        next_key = str(next_step.get("resource_key", "") or "")
        next_heap = resource_heaps[next_key]
        next_forced_line = _upcoming_forced_line(steps, next_idx + 1, car_type)
        resource_ready, resource_line = _peek_resource_slot(
            next_heap,
            next_step,
            current_line_no,
            next_forced_line,
            {
                line_no: _slot_ready_time(next_idx, _actual_line_no(
                    next_step,
                    line_no,
                    current_line_no,
                    next_forced_line,
                ))
                for _, line_no in next_heap
            },
        )
        next_line = _actual_line_no(
            next_step,
            resource_line,
            current_line_no,
            next_forced_line,
        )
        next_ready = max(resource_ready, _slot_ready_time(next_idx, next_line))

        if is_zone_entry(next_idx):
            nzid = steps[next_idx]["zone_id"]
            nheap = zone_heaps[nzid]
            next_ready = max(next_ready, nheap[0] if nheap else 0.0)

        return next_ready

    for car, car_type in enumerate(vehicle_seq, start=1):
        theory_launch_time = (car - 1) * launch_takt_value if launch_takt_value > 0 else 0.0
        prev_depart = theory_launch_time
        current_line_no = ""
        # 记录该车是否经过某个 gate_zone（用于之后把它的“进入 zone 的时刻”加入 pre_heap）
        car_gate_zones: set[str] = set()

        for j, st in enumerate(steps):
            cur_duration = _pick_duration(st, car_type)
            try:
                cur_duration = float(cur_duration or 0.0)
            except Exception:
                cur_duration = 0.0

            # v2-2：车型跳过岗位逻辑
            # A/B/C 对应工时 < 0 时，表示该车型不需要该岗位，直接跳过。
            # 工时 = 0 表示该车型仍经过岗位，不占用加工设备，但占用物理流转车位。
            if cur_duration < 0:
                continue

            is_zero_duration = abs(cur_duration) <= 1e-9

            # ---- 计算本步开始时间：所有车辆受物理车位约束；正工时车辆另受加工资源约束 ----
            cur_key = str(st.get("resource_key", "") or "")
            cur_heap = resource_heaps[cur_key]
            cur_ready = prev_depart
            resource_line_no = current_line_no
            selected_line_no = current_line_no

            if is_zero_duration:
                selected_line_no = _pass_through_line(j, current_line_no, car_type)
                cur_ready = max(cur_ready, _slot_ready_time(j, selected_line_no))
            else:
                upcoming_forced_line = _upcoming_forced_line(steps, j + 1, car_type)
                cur_ready, resource_line_no = _pop_resource_slot(
                    cur_heap,
                    st,
                    current_line_no,
                    upcoming_forced_line,
                    {
                        line_no: _slot_ready_time(j, _actual_line_no(
                            st,
                            line_no,
                            current_line_no,
                            upcoming_forced_line,
                        ))
                        for _, line_no in cur_heap
                    },
                )
                selected_line_no = _actual_line_no(
                    st,
                    resource_line_no,
                    current_line_no,
                    upcoming_forced_line,
                )
                cur_ready = max(cur_ready, _slot_ready_time(j, selected_line_no))

            start = max(cur_ready, prev_depart)

            # ---- 闸门缓冲约束（在 start 阶段判断）：允许“闸门→区域入口”链路上最多 gate_buffer 辆 ----
            gz = st.get("gate_zone_id", "")
            if gz:
                car_gate_zones.add(gz)
                gb = max(1, int(gate_buffers.get(gz, 2)))
                heap = pre_heap.setdefault(gz, [])

                while heap and heap[0] <= start:
                    heapq.heappop(heap)

                while len(heap) >= gb:
                    start = max(start, heap[0])
                    while heap and heap[0] <= start:
                        heapq.heappop(heap)

            svc_finish = start + cur_duration

            # ---- depart 受紧邻下一个流转节点的物理车位/加工资源/zone容量约束 ----
            # 车辆只有在下一个节点可接收时才离开当前节点，等待逐级向上游回传。
            next_ready = _next_step_ready(j, selected_line_no, car_type)
            depart = max(svc_finish, next_ready)

            block_wait = max(0.0, depart - svc_finish)

            rows.append({
                "car": car,
                "step_seq": st["seq"],
                "step_display": st["display"],
                "group": st["group"],
                "run_mode": st.get("run_mode", "单线单设备"),
                "capacity": int(st.get("capacity", 1) or 1),
                "device_count": int(st.get("device_count", st.get("capacity", 1)) or 1),
                "line_scope": str(st.get("line_scope", "") or ""),
                "resource_key": str(st.get("resource_key", "") or ""),
                "line_no": selected_line_no,
                "car_type": car_type,
                "duration_source": str(car_type or "A").strip().upper(),
                "theory_launch_time": theory_launch_time,
                "launch_takt": launch_takt_value,
                "launch_wait": max(0.0, start - theory_launch_time) if j == 0 else 0.0,
                "dur": cur_duration,
                "start": start,
                "svc_finish": svc_finish,
                "depart": depart,
                "block_wait": block_wait,
            })

            # ---- Zone 名额占用/释放 ----
            # 进入 Zone：仅在“Zone 入口步骤”发生，占用一个名额
            if is_zone_entry(j):
                zid = st["zone_id"]
                # 如果该车之前通过过指向该 zid 的某个闸门，则把它“进入 zone 的时刻（=本步 start=上一步 depart）”加入 pre_heap
                if zid in car_gate_zones:
                    heap = pre_heap.setdefault(zid, [])
                    heapq.heappush(heap, start)  # 之后其他车的“闸门开始”会受这个时间点约束

                heap = zone_heaps[zid]
                if heap:
                    heapq.heappop(heap)  # 占用一个 zone 名额

            # 离开 Zone：仅在“Zone 最后一步”释放一个名额（名额释放时刻=本步 depart）
            if is_zone_exit(j):
                zid = st["zone_id"]
                heap = zone_heaps[zid]
                heapq.heappush(heap, depart)

            # ---- 更新状态，进入下一步 ----
            if not is_zero_duration:
                heapq.heappush(cur_heap, (depart, resource_line_no))
            _set_slot_ready_time(j, selected_line_no, depart)
            prev_depart = depart
            if not is_zero_duration and selected_line_no:
                current_line_no = selected_line_no
            max_time = max(max_time, depart)

    return rows, max_time


# ---------------- 等待统计 ---------------- #
def _build_car_slices(rows: List[Dict[str, Any]]) -> Dict[int, List[Dict[str, Any]]]:
    by_car: Dict[int, List[Dict[str, Any]]] = {}
    for r in rows:
        by_car.setdefault(r["car"], []).append(r)
    for k in by_car:
        by_car[k].sort(key=lambda x: x["step_seq"])
    return by_car


def _compute_entry_wait(by_car: Dict[int, List[Dict[str, Any]]]) -> Dict[int, float]:
    """
    入站等待：
    - v2-1 固定投车节拍模式下：第一步 start - 理论投车时间（<0 计 0）。
      这样按节拍等待投车的时间不会被误算为等待损失。
    - 旧模式下：保持原逻辑，车 i 第一步 start - 车 i-1 第一步 depart（<0 计 0）。
    """
    wait_map: Dict[int, float] = {}
    prev_first_depart = 0.0
    for car in sorted(by_car.keys()):
        steps = by_car[car]
        if not steps:
            continue
        first = steps[0]
        first_start = float(first.get("start", 0.0))
        launch_takt_value = float(first.get("launch_takt", 0.0) or 0.0)
        if launch_takt_value > 0:
            theory_launch_time = float(first.get("theory_launch_time", 0.0) or 0.0)
            wait_map[car] = max(0.0, first_start - theory_launch_time)
        else:
            wait_map[car] = max(0.0, first_start - prev_first_depart)
        prev_first_depart = float(first.get("depart", 0.0))
    return wait_map


def _compute_total_wait(by_car: Dict[int, List[Dict[str, Any]]]) -> Dict[int, float]:
    """总等待 = 入站等待 + Σ block_wait（所有步）"""
    entry_map = _compute_entry_wait(by_car)
    total_map: Dict[int, float] = {}
    for car, steps in by_car.items():
        inter = sum(max(0.0, s.get("block_wait", 0.0)) for s in steps)
        total_map[car] = float(entry_map.get(car, 0.0) + inter)
    return total_map


# ---------------- 结果分析 ---------------- #
def analyze_schedule(rows: List[Dict[str, Any]], max_finish: float, target_takt: Optional[float] = None) -> Dict[str, Any]:
    """
    分析兼容入口。

    v2-5B：tickets.py 不再承载完整分析逻辑，实际分析交给 core.analysis.analyze_schedule_v2。
    本函数只负责把 v2 分析结果补齐为旧 UI/导出仍可能读取的字段名，避免一次性改动 UI。
    """
    result = analyze_schedule_v2(rows, max_finish, target_takt)

    summary = dict(result.get("summary", {}) or {})
    station_summary = list(result.get("station_summary", result.get("stations", [])) or [])
    car_type_summary = list(result.get("car_type_summary", result.get("car_types", [])) or [])

    # ---- 兼容旧 summary 字段 ----
    if "total_cars" not in summary:
        summary["total_cars"] = int(summary.get("car_count", 0) or 0)
    if "avg_wait" not in summary:
        summary["avg_wait"] = float(summary.get("average_wait", 0.0) or 0.0)
    if "average_wait" not in summary:
        summary["average_wait"] = float(summary.get("avg_wait", 0.0) or 0.0)
    if "bottleneck_station" not in summary:
        summary["bottleneck_station"] = ""

    target_takt_value = float(summary.get("target_takt", target_takt or 0.0) or 0.0)
    ng_stations = [s for s in station_summary if str(s.get("status", s.get("takt_result", ""))) == "NG"]
    if "takt_result" not in summary:
        summary["takt_result"] = "NG" if target_takt_value > 0 and ng_stations else ("OK" if target_takt_value > 0 else "未设定")
    if "over_takt_station_count" not in summary:
        summary["over_takt_station_count"] = len(ng_stations)

    # ---- 兼容旧 station_stats 字段 ----
    station_stats: List[Dict[str, Any]] = []
    for item in station_summary:
        stat = dict(item)
        if "avg_process" not in stat:
            stat["avg_process"] = float(stat.get("avg_duration", 0.0) or 0.0)
        if "total_process" not in stat:
            stat["total_process"] = stat["avg_process"] * int(stat.get("count", 0) or 0)
        if "total_block_wait" not in stat:
            stat["total_block_wait"] = float(stat.get("total_wait", 0.0) or 0.0)
        if "avg_block_wait" not in stat:
            count = int(stat.get("count", 0) or 0)
            stat["avg_block_wait"] = stat["total_block_wait"] / count if count else 0.0
        if "takt_result" not in stat:
            stat["takt_result"] = str(stat.get("status", "未设定") or "未设定")
        if "over_takt_types" not in stat:
            stat["over_takt_types"] = "—"
        station_stats.append(stat)

    # ---- 兼容旧 type_stats 字段 ----
    type_stats: List[Dict[str, Any]] = []
    for item in car_type_summary:
        stat = dict(item)
        stat.setdefault("avg_wait", 0.0)
        stat.setdefault("avg_through_time", 0.0)
        stat.setdefault("avg_process", 0.0)
        type_stats.append(stat)

    # ---- 兼容旧 car_stats 字段 ----
    # 当前 v2 分析模块暂不返回完整 car_stats；这里保留空列表，后续如 UI 需要再迁移到 analysis.py。
    car_stats = list(result.get("car_stats", []) or [])

    result["summary"] = summary
    result["station_summary"] = station_summary
    result["stations"] = station_summary
    result["car_type_summary"] = car_type_summary
    result["car_types"] = car_type_summary
    result["station_stats"] = station_stats
    result["type_stats"] = type_stats
    result["car_stats"] = car_stats
    return result


# ---------------- 导出入口 ---------------- #
def schedule_and_export(defs: List[Dict[str, Any]],
                        cars: int,
                        grid_step: float,
                        wait_policy: str,   # "before"/"after" 仅影响是否绘入站等待条
                        project: str,
                        dst_path: str,
                        vehicle_counts: Dict[str, int] | None = None,
                        sequence_mode: str = "grouped",
                        max_consecutive: int = 10,
                        ratio_pattern: Optional[Dict[str, int]] = None,
                        target_takt: Optional[float] = None,
                        vehicle_sequence: Optional[List[str]] = None) -> Dict[str, Any]:
    grid_step = 1.0 if (not isinstance(grid_step, (int, float)) or grid_step <= 0) else float(grid_step)
    rows, max_finish = schedule(
        defs,
        cars,
        vehicle_counts,
        sequence_mode,
        max_consecutive,
        ratio_pattern,
        launch_takt=target_takt,
        vehicle_sequence=vehicle_sequence,
    )
    analysis = analyze_schedule(rows, max_finish, target_takt)
    # ---- 收集用户自定义颜色 (display -> hex)
    step_color_map = {d.get("display"): d.get("color") for d in defs if d.get("color")}
    engine = _choose_engine()
    if engine is None:
        raise RuntimeError("未找到可用的 Excel 引擎，请安装 xlsxwriter 或 openpyxl")

    if engine == "xlsxwriter":
        _export_with_xlsxwriter(rows, max_finish, grid_step, wait_policy, project, dst_path, step_color_map)
    else:
        _export_with_openpyxl(rows, max_finish, grid_step, wait_policy, project, dst_path)
    return analysis


# ---------------- 样式与工具 ---------------- #
def _palette():
    group_colors = [
        "#4CAF50", "#2196F3", "#9C27B0", "#FF9800", "#009688",
        "#795548", "#3F51B5", "#E91E63", "#00BCD4", "#8BC34A",
    ]
    wait_color = "#FFC107"
    return group_colors, wait_color


def _fmt_num(x: float) -> str:
    if abs(x - round(x)) < 1e-9:
        return str(int(round(x)))
    return f"{x:.1f}"


# ---------------- xlsxwriter 彩色导出 ---------------- #
def _export_with_xlsxwriter(rows: List[Dict[str, Any]], max_finish: float,
                             grid_step: float, wait_policy: str,
                             project: str, dst_path: str,
                             step_color_map: Dict[str, str]) -> None:
    import xlsxwriter  # type: ignore

    by_car = _build_car_slices(rows)
    entry_wait = _compute_entry_wait(by_car)
    total_wait = _compute_total_wait(by_car)

    n_cols_grid = max(1, int(math.ceil(max_finish / grid_step)))

    with pd.ExcelWriter(dst_path, engine="xlsxwriter") as writer:
        wb = writer.book
        ws = wb.add_worksheet("作业组合票")
        writer.sheets["作业组合票"] = ws

        fmt_header = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "bg_color": "#EEEEEE", "border": 1})
        fmt_text   = wb.add_format({"align": "center", "valign": "vcenter", "border": 1})
        fmt_left   = wb.add_format({"align": "left", "valign": "vcenter", "border": 1})
        fmt_wait   = wb.add_format({"align": "left", "valign": "vcenter", "border": 1, "bg_color": "#FFF9C4"})
        fmt_bar_wait = wb.add_format({"bg_color": "#FFE082", "border": 0})
        fmt_car    = wb.add_format({"align": "center", "valign": "vcenter", "border": 1, "bg_color": "#F5F5F5"})

        group_colors, _ = _palette()
        group_fmt_cache: Dict[str, Any] = {}
        def bar_fmt(group: str, display: str):
            # 若当前步骤有自定义颜色，优先使用
            custom_hex = step_color_map.get(display)
            if custom_hex:
                if custom_hex not in group_fmt_cache:
                    group_fmt_cache[custom_hex] = wb.add_format({"bg_color": custom_hex, "border": 0})
                return group_fmt_cache[custom_hex]
            # 否则按 group 调色盘
            if group not in group_fmt_cache:
                idx = (hash(group) >> 1) % len(group_colors)
                group_fmt_cache[group] = wb.add_format({"bg_color": group_colors[idx], "border": 0})
            return group_fmt_cache[group]

        ws.set_column(0, 0, 36)
        ws.set_column(1, 1, 8)
        ws.set_column(2, 2, 8)
        ws.set_column(3, 3, 18)
        ws.set_column(4, 4, 10)
        ws.set_column(5, 5 + n_cols_grid - 1, 2.8)

        ws.write(0, 0, f"连续投入{project}等待时间", fmt_header)
        ws.write(0, 1, "车号", fmt_header)
        ws.write(0, 2, "车型", fmt_header)
        ws.write(0, 3, "项目", fmt_header)
        ws.write(0, 4, "时间", fmt_header)
        for i in range(n_cols_grid):
            ws.write(0, 5 + i, f"{grid_step:.1f}", fmt_header)
        ws.freeze_panes(1, 0)

        row_cursor = 1
        for car in sorted(by_car.keys()):
            steps = by_car[car]
            if not steps:
                continue

            ewait = float(entry_wait.get(car, 0.0))
            twait = float(total_wait.get(car, 0.0))
            car_type = steps[0].get("car_type", "")
            ws.write(row_cursor, 0, f"入站等待{_fmt_num(ewait)}秒；总等待{_fmt_num(twait)}秒", fmt_wait if ewait > 0 else fmt_left)
            ws.write(row_cursor, 1, car, fmt_car)
            ws.write(row_cursor, 2, car_type, fmt_car)
            ws.write(row_cursor, 3, "", fmt_left)
            ws.write(row_cursor, 4, _fmt_num(ewait) if ewait > 0 else "", fmt_text if ewait > 0 else fmt_left)

            if ewait > 0 and wait_policy == "before":
                first_start = steps[0]["start"]
                c0 = 5
                c1 = 5 + int(math.ceil(first_start / grid_step)) - 1
                c1 = max(c1, c0 - 1)
                for c in range(c0, c1 + 1):
                    ws.write(row_cursor, c, "", fmt_bar_wait)
            row_cursor += 1

            for idx, s in enumerate(steps):
                # 服务条
                ws.write(row_cursor, 0, "", fmt_left)
                ws.write(row_cursor, 1, "", fmt_text)
                ws.write(row_cursor, 2, s.get("car_type", ""), fmt_text)
                ws.write(row_cursor, 3, s["step_display"], fmt_left)
                ws.write(row_cursor, 4, _fmt_num(s["dur"]), fmt_text)
                c_start = 5 + int(math.floor(s["start"] / grid_step))
                c_end_svc = 5 + int(math.ceil(s["svc_finish"] / grid_step)) - 1
                c_end_svc = max(c_end_svc, c_start)
                bf = bar_fmt(s["group"], s["step_display"])
                for c in range(c_start, c_end_svc + 1):
                    ws.write(row_cursor, c, "", bf)
                row_cursor += 1

                # 等待条（svc_finish → depart）
                if s["block_wait"] > 1e-9 and idx < len(steps) - 1:
                    wait_val = s["block_wait"]
                    next_name = steps[idx + 1]["step_display"]
                    ws.write(row_cursor, 0, f"等待{_fmt_num(wait_val)}秒（{s['step_display']} → {next_name}）", fmt_wait)
                    ws.write(row_cursor, 1, "", fmt_text)
                    ws.write(row_cursor, 2, s.get("car_type", ""), fmt_text)
                    ws.write(row_cursor, 3, "", fmt_wait)
                    ws.write(row_cursor, 4, _fmt_num(wait_val), fmt_text)
                    c_w0 = 5 + int(math.floor(s["svc_finish"] / grid_step))
                    c_w1 = 5 + int(math.ceil(s["depart"] / grid_step)) - 1
                    c_w1 = max(c_w1, c_w0)
                    for c in range(c_w0, c_w1 + 1):
                        ws.write(row_cursor, c, "", fmt_bar_wait)
                    row_cursor += 1

            row_cursor += 1  # 车与车之间空一行


# ---------------- openpyxl 回退导出（文字） ---------------- #
def _export_with_openpyxl(rows: List[Dict[str, Any]], max_finish: float,
                           grid_step: float, wait_policy: str,
                           project: str, dst_path: str) -> None:
    by_car = _build_car_slices(rows)
    entry_wait = _compute_entry_wait(by_car)
    total_wait = _compute_total_wait(by_car)

    out_rows = []
    for car in sorted(by_car.keys()):
        steps = by_car[car]
        if not steps:
            continue
        ewait = float(entry_wait.get(car, 0.0))
        twait = float(total_wait.get(car, 0.0))
        car_type = steps[0].get("car_type", "")
        out_rows.append({
            "车号": car,
            "车型": car_type,
            "项目": "(入站等待/总等待)",
            "时间": ewait,
            "说明": f"入站等待{_fmt_num(ewait)}秒；总等待{_fmt_num(twait)}秒"
        })
        for idx, s in enumerate(steps):
            out_rows.append({"车号": car, "车型": s.get("car_type", ""), "项目": s["step_display"], "时间": s["dur"], "说明": ""})
            if s["block_wait"] > 1e-9 and idx < len(steps) - 1:
                out_rows.append({
                    "车号": car,
                    "车型": s.get("car_type", ""),
                    "项目": f"(等待：{s['step_display']}→{steps[idx+1]['step_display']})",
                    "时间": s["block_wait"],
                    "说明": f"等待{_fmt_num(s['block_wait'])}秒"
                })
        out_rows.append({"车号": "", "车型": "", "项目": "", "时间": "", "说明": ""})

    df = pd.DataFrame(out_rows, columns=["车号", "车型", "项目", "时间", "说明"])
    # TODO: 未应用自定义颜色（step_color_map）到 openpyxl 导出
    with pd.ExcelWriter(dst_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="作业组合票")
        try:
            ws = writer.sheets["作业组合票"]
            ws.freeze_panes = "A2"
        except Exception:
            pass
