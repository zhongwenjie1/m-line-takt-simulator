# -*- coding: utf-8 -*-
"""
组合票排程 + 导出
当前阶段说明：
- 多工程组合票已切换为『运行方式 + 岗位/设备 + A/B/C 工时』录入模型。
- v2-1 支持固定投车节拍；v2-2 支持车型工时为 0 时跳过该岗位；v2-3B 建立资源 key；v2-4A 建立自动线别分配基础；v2-4B 支持线别连续性与强制线别；v2-4C 支持未来强制线别预判。
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


# ---------------- 调度（含 Zone + gate_buffer） ---------------- #
def schedule(step_defs: List[Dict[str, Any]],
             cars: int,
             vehicle_counts: Dict[str, int] | None = None,
             sequence_mode: str = "grouped",
             max_consecutive: int = 10,
             ratio_pattern: Optional[Dict[str, int]] = None,
             launch_takt: Optional[float] = None) -> Tuple[List[Dict[str, Any]], float]:
    """
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
      工时 = 0 时仍然经过该岗位，但不占用加工时间，也不占用当前工位资源。
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

    # 车型序列
    vehicle_seq: List[str] = []
    if vehicle_counts:
        counts: Dict[str, int] = {}
        for vt in ("A", "B", "C"):
            try:
                counts[vt] = max(0, int(vehicle_counts.get(vt, 0) or 0))
            except Exception:
                counts[vt] = 0

        max_consecutive = max(1, int(max_consecutive or 1))

        if sequence_mode == "alternate":
            # 交替混流 + 最大连续台数约束
            # 规则：尽量避免同车型连续超过 max_consecutive；若无可选车型则放宽约束继续排。
            last_type = ""
            run_len = 0
            while any(v > 0 for v in counts.values()):
                candidates = []
                for vt in ("A", "B", "C"):
                    if counts[vt] <= 0:
                        continue
                    if vt == last_type and run_len >= max_consecutive:
                        continue
                    candidates.append(vt)

                if not candidates:
                    # 所有可用车型都被连续台数限制挡住，则放宽约束，优先取剩余最多的车型
                    candidates = [vt for vt in ("A", "B", "C") if counts[vt] > 0]
                    candidates.sort(key=lambda x: (-counts[x], x))
                    pick = candidates[0]
                else:
                    # 优先取剩余最多的车型；同数时按 A/B/C 顺序
                    candidates.sort(key=lambda x: (-counts[x], x))
                    pick = candidates[0]

                vehicle_seq.append(pick)
                counts[pick] -= 1
                if pick == last_type:
                    run_len += 1
                else:
                    last_type = pick
                    run_len = 1
        elif sequence_mode == "ratio":
            # 按比例运行：固定比例块循环，例如 6:4 -> AAAAAABBBB，然后重复展开到总台数。
            pattern = ratio_pattern or {}
            a = max(0, int(pattern.get("A", 0) or 0))
            b = max(0, int(pattern.get("B", 0) or 0))
            c = max(0, int(pattern.get("C", 0) or 0))

            block: List[str] = (["A"] * a) + (["B"] * b) + (["C"] * c)
            if not block:
                raise ValueError("按比例运行模式下，比例块不能为空")

            vehicle_seq = []
            total_cars = max(0, int(cars))
            while len(vehicle_seq) < total_cars:
                need = total_cars - len(vehicle_seq)
                vehicle_seq.extend(block[:need])
        else:
            # 顺排：A -> B -> C
            # 这里明确忽略 max_consecutive，后续如果做方式B界面联动再在 UI 上灰掉提示
            for vt in ("A", "B", "C"):
                vehicle_seq.extend([vt] * counts[vt])

    if not vehicle_seq:
        vehicle_seq = ["A" for _ in range(max(0, int(cars)))]

    cars = len(vehicle_seq)
    try:
        launch_takt_value = float(launch_takt or 0.0)
    except Exception:
        launch_takt_value = 0.0
    if launch_takt_value < 0:
        launch_takt_value = 0.0
 
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

    def _pop_resource_slot(heap: List[Tuple[float, str]], st: Dict[str, Any], current_line_no: str, upcoming_forced_line: str = "") -> Tuple[float, str]:
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
        return heapq.heappop(heap)

    def _peek_resource_ready(heap: List[Tuple[float, str]], st: Dict[str, Any], current_line_no: str, upcoming_forced_line: str = "") -> float:
        """按下一岗位强制线别/线别连续性/未来强制线别预估可接收时间。"""
        if not heap:
            return 0.0
        preferred_line = _forced_or_preferred_line(st, current_line_no, upcoming_forced_line)
        if preferred_line:
            matched = [ready_time for ready_time, line_no in heap if line_no == preferred_line]
            if matched:
                return min(matched)
        return heap[0][0]

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
            # 工时 = 0 表示该车型仍经过岗位，但不占用加工时间，也不占用资源。
            if cur_duration < 0:
                continue

            is_zero_duration = abs(cur_duration) <= 1e-9

            # ---- 计算本步开始时间：0 工时只继承车辆流转时间；>0 工时仍受当前工位资源约束 ----
            cur_key = str(st.get("resource_key", "") or "")
            cur_heap = resource_heaps[cur_key]
            cur_ready = prev_depart
            resource_line_no = current_line_no
            selected_line_no = current_line_no

            if not is_zero_duration:
                upcoming_forced_line = _upcoming_forced_line(steps, j + 1, car_type)
                cur_ready, resource_line_no = _pop_resource_slot(
                    cur_heap,
                    st,
                    current_line_no,
                    upcoming_forced_line,
                )
                selected_line_no = _actual_line_no(
                    st,
                    resource_line_no,
                    current_line_no,
                    upcoming_forced_line,
                )

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

            # ---- depart 受“下步可接收（服务器释放 + zone 容量）”约束 ----
            # 0 工时 segment 只是流转节点，不占当前资源，也不在该节点等待下步资源。
            if is_zero_duration:
                next_idx, next_st = _find_next_effective_step(steps, j, car_type)
                if next_st is not None and next_idx is not None:
                    next_key = str(next_st.get("resource_key", "") or "")
                    next_heap = resource_heaps[next_key]
                    next_upcoming_forced_line = _upcoming_forced_line(steps, next_idx + 1, car_type)
                    next_ready = _peek_resource_ready(
                        next_heap,
                        next_st,
                        selected_line_no,
                        next_upcoming_forced_line,
                    )

                    if is_zone_entry(next_idx):
                        nzid = steps[next_idx]["zone_id"]
                        nheap = zone_heaps[nzid]
                        next_ready = max(next_ready, nheap[0] if nheap else 0.0)

                    depart = max(svc_finish, next_ready)
                else:
                    depart = svc_finish
            else:
                next_active_idx, next_active_st, next_active_duration = _find_next_active_step(steps, j, car_type)
                if next_active_st is not None and next_active_idx is not None and next_active_duration <= 0:
                    # 当前步后面紧跟 0 工时 pass-through 节点。
                    # 车辆可以先从当前作业工位释放到 pass-through 节点，
                    # 下游真正不可接收时，应把等待挂在该 0 工时节点，而不是回传到当前工位。
                    depart = svc_finish
                else:
                    next_idx, next_st = _find_next_effective_step(steps, j, car_type)
                    if next_st is not None and next_idx is not None:
                        next_key = str(next_st.get("resource_key", "") or "")
                        next_heap = resource_heaps[next_key]
                        next_upcoming_forced_line = _upcoming_forced_line(steps, next_idx + 1, car_type)
                        next_ready = _peek_resource_ready(
                            next_heap,
                            next_st,
                            selected_line_no,
                            next_upcoming_forced_line,
                        )

                        if is_zone_entry(next_idx):
                            nzid = steps[next_idx]["zone_id"]
                            nheap = zone_heaps[nzid]
                            next_ready = max(next_ready, nheap[0] if nheap else 0.0)

                        depart = max(svc_finish, next_ready)
                    else:
                        depart = svc_finish

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
                        target_takt: Optional[float] = None) -> Dict[str, Any]:
    grid_step = 1.0 if (not isinstance(grid_step, (int, float)) or grid_step <= 0) else float(grid_step)
    rows, max_finish = schedule(
        defs,
        cars,
        vehicle_counts,
        sequence_mode,
        max_consecutive,
        ratio_pattern,
        launch_takt=target_takt,
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
