# -*- coding: utf-8 -*-
"""
Input parsing helpers for ticket_export_app.

当前阶段只提供纯输入解析工具函数。
暂不接入 ui/export_ticket_window.py。
"""

import math


def normalize_text(value):
    if value is None:
        return ""
    return str(value).strip()


def is_blank(value):
    return normalize_text(value) == ""


def parse_float(value, default=0.0):
    if is_blank(value):
        return default
    try:
        return float(value)
    except Exception:
        return default


def parse_int(value, default=0):
    if is_blank(value):
        return default
    try:
        return int(float(value))
    except Exception:
        return default


def parse_required_float(value, field_name="数值"):
    try:
        if is_blank(value):
            raise ValueError
        return float(value)
    except Exception:
        raise ValueError(f"{field_name} 不是有效数字：{value}")


def parse_non_negative_float(value, field_name="数值"):
    result = parse_required_float(value, field_name)
    if result < 0:
        raise ValueError(f"{field_name} 不能小于 0：{value}")
    return result


def parse_positive_float(value, field_name="数值"):
    result = parse_required_float(value, field_name)
    if result <= 0:
        raise ValueError(f"{field_name} 必须大于 0：{value}")
    return result


def validate_vehicle_values(values, mode_name="数量"):
    result = {}
    for key in ("A", "B", "C"):
        value = parse_int((values or {}).get(key, 0), 0)
        if value < 0:
            raise ValueError(f"A/B/C{mode_name}不能小于 0。")
        result[key] = value

    if sum(result.values()) <= 0:
        raise ValueError(f"A/B/C{mode_name}合计必须大于 0。")

    return result


def validate_target_takt(target_takt, required=False):
    value = parse_required_float(target_takt, "目标节拍")
    if required and value <= 0:
        raise ValueError("目标节拍必须大于 0。")
    if not required and value < 0:
        raise ValueError("目标节拍不能小于 0。")
    return value


def validate_analysis_minutes(minutes):
    value = parse_int(minutes, 0)
    if value <= 0:
        raise ValueError("按比例投车模式下，请填写分析时间，且分析时间必须大于 0 分钟。")
    return value


def calc_ratio_theoretical_launch_count(analysis_time_seconds, target_takt, buffer_count=50):
    analysis_seconds = parse_required_float(analysis_time_seconds, "分析时间")
    takt = validate_target_takt(target_takt, required=True)
    buffer_value = parse_int(buffer_count, 0)
    count = int(math.ceil(analysis_seconds / takt)) + buffer_value
    if count <= 0:
        raise ValueError("分析时间过短，按当前目标节拍计算的投车生成台数为 0，请增加分析时间。")
    return count


def build_ratio_pattern(a_value, b_value, c_value):
    return validate_vehicle_values(
        {"A": a_value, "B": b_value, "C": c_value},
        mode_name="比例",
    )


def calc_total_vehicle_count(a_value, b_value, c_value):
    values = validate_vehicle_values(
        {"A": a_value, "B": b_value, "C": c_value},
        mode_name="数量",
    )
    return values["A"] + values["B"] + values["C"]


def is_empty_station_row(row_data):
    row_data = row_data or {}
    return (
        is_blank(row_data.get("display"))
        and is_blank(row_data.get("group"))
        and is_blank(row_data.get("duration_a"))
        and is_blank(row_data.get("duration_b"))
        and is_blank(row_data.get("duration_c"))
    )


def parse_duration_value(value, car_type, row_index, active_by_type):
    if is_blank(value):
        if bool((active_by_type or {}).get(car_type, False)):
            raise ValueError(
                f"第 {row_index + 1} 行『{car_type}工时』不能为空；参与投车车型若跳过该岗位请填写 0。"
            )
        return 0.0

    try:
        return float(value)
    except Exception:
        raise ValueError(f"第 {row_index + 1} 行『{car_type}工时』不是有效数字：{value}")


def parse_device_count(value, row_index):
    try:
        device_count = int(float(value))
    except Exception:
        raise ValueError(f"第 {row_index + 1} 行『设备数量』不是有效值：{value}")

    if device_count not in (1, 2):
        raise ValueError(f"第 {row_index + 1} 行『设备数量』当前仅支持 1 或 2。")

    return device_count


def validate_line_scope(device_count, line_scope, row_index):
    if device_count == 2 and line_scope != "双线":
        raise ValueError(f"第 {row_index + 1} 行设备数量为 2 时，所属线别必须为『双线』。")
    if device_count == 1 and line_scope == "双线":
        raise ValueError(
            f"第 {row_index + 1} 行设备数量为 1 时，所属线别不能为『双线』，请选择 1号线 / 2号线 / 双线共用。"
        )


def resolve_run_mode(device_count, line_scope):
    if device_count == 2:
        return "双线双设备"
    if line_scope == "双线共用":
        return "双线单设备"
    return "单线单设备"


def parse_station_row(row_data, row_index, active_by_type):
    row_data = row_data or {}
    if is_empty_station_row(row_data):
        return None

    seq = normalize_text(row_data.get("seq"))
    display = normalize_text(row_data.get("display"))
    group = normalize_text(row_data.get("group"))
    device_count_value = normalize_text(row_data.get("device_count"))
    line_scope = normalize_text(row_data.get("line_scope"))
    duration_a_value = normalize_text(row_data.get("duration_a"))
    duration_b_value = normalize_text(row_data.get("duration_b"))
    duration_c_value = normalize_text(row_data.get("duration_c"))
    color = normalize_text(row_data.get("color"))

    if not display or not group:
        raise ValueError(f"第 {row_index + 1} 行请填写工程名称和岗位设备。")

    duration_a = parse_duration_value(duration_a_value, "A", row_index, active_by_type)
    duration_b = parse_duration_value(duration_b_value, "B", row_index, active_by_type)
    duration_c = parse_duration_value(duration_c_value, "C", row_index, active_by_type)

    device_count = parse_device_count(device_count_value, row_index)
    validate_line_scope(device_count, line_scope, row_index)

    active_by_type = active_by_type or {}
    if (
        (not active_by_type.get("A", False) or duration_a == 0)
        and (not active_by_type.get("B", False) or duration_b == 0)
        and (not active_by_type.get("C", False) or duration_c == 0)
    ):
        raise ValueError(f"第 {row_index + 1} 行参与投车的车型不能全部跳过该岗位，请至少填写一个大于 0 的工时。")

    capacity = 2 if device_count == 2 else 1
    run_mode = resolve_run_mode(device_count, line_scope)
    seq_int = int(float(seq)) if seq else row_index + 1

    return {
        "seq": seq_int,
        "display": display,
        "group": group,
        "capacity": capacity,
        "durations": [duration_a],
        "color": color,
        "run_mode": run_mode,
        "device_count": device_count,
        "line_scope": line_scope,
        "duration_a": duration_a,
        "duration_b": duration_b,
        "duration_c": duration_c,
    }


def parse_station_rows(row_data_list, active_by_type):
    defs = []
    for row_index, row_data in enumerate(row_data_list or []):
        rec = parse_station_row(row_data, row_index, active_by_type)
        if rec is None:
            continue
        defs.append(rec)

    defs.sort(key=lambda x: x["seq"])
    if not defs:
        raise ValueError("请至少填写一行有效的步骤（工程名称/岗位设备/A工时）")

    return defs


def parse_multi_project_inputs(raw_inputs):
    raw_inputs = raw_inputs or {}

    project = normalize_text(raw_inputs.get("project")) or "工程"
    cars_a = parse_int(raw_inputs.get("cars_a"), 0)
    cars_b = parse_int(raw_inputs.get("cars_b"), 0)
    cars_c = parse_int(raw_inputs.get("cars_c"), 0)
    is_ratio_mode = bool(raw_inputs.get("is_ratio_mode", False))

    target_takt = validate_target_takt(
        raw_inputs.get("target_takt"),
        required=is_ratio_mode,
    )

    if is_ratio_mode:
        sequence_mode = "ratio"
    else:
        sequence_mode_index = parse_int(raw_inputs.get("sequence_mode_index"), 0)
        sequence_mode = "alternate" if sequence_mode_index == 1 else "grouped"

    max_consecutive = parse_int(raw_inputs.get("max_consecutive"), 10)
    if max_consecutive <= 0:
        max_consecutive = 10

    if is_ratio_mode:
        vehicle_values = validate_vehicle_values(
            {"A": cars_a, "B": cars_b, "C": cars_c},
            mode_name="比例",
        )
        analysis_minutes = validate_analysis_minutes(raw_inputs.get("analysis_minutes"))
        analysis_time_seconds = analysis_minutes * 60.0
        theoretical_launch_count = calc_ratio_theoretical_launch_count(
            analysis_time_seconds,
            target_takt,
            buffer_count=50,
        )
        ratio_pattern = build_ratio_pattern(cars_a, cars_b, cars_c)
        cars = theoretical_launch_count
    else:
        vehicle_values = validate_vehicle_values(
            {"A": cars_a, "B": cars_b, "C": cars_c},
            mode_name="数量",
        )
        cars = calc_total_vehicle_count(cars_a, cars_b, cars_c)
        ratio_pattern = None
        analysis_time_seconds = None
        theoretical_launch_count = None

    vehicle_counts = {
        "A": vehicle_values["A"],
        "B": vehicle_values["B"],
        "C": vehicle_values["C"],
    }
    active_by_type = {
        "A": cars_a > 0,
        "B": cars_b > 0,
        "C": cars_c > 0,
    }
    defs = parse_station_rows(raw_inputs.get("station_rows") or [], active_by_type)

    return {
        "project": project,
        "cars": cars,
        "grid_step": 1.0,
        "wait_policy": "before",
        "defs": defs,
        "vehicle_counts": vehicle_counts,
        "sequence_mode": sequence_mode,
        "max_consecutive": max_consecutive,
        "ratio_pattern": ratio_pattern,
        "target_takt": target_takt,
        "analysis_time_seconds": analysis_time_seconds,
        "theoretical_launch_count": theoretical_launch_count,
    }
