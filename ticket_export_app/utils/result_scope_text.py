"""Single source for quantity/ratio model-result scope wording."""

from __future__ import annotations


def _fmt_number(value, default="-"):
    if value is None:
        return default
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    if abs(number - round(number)) < 1e-9:
        return str(int(round(number)))
    return f"{number:.1f}"


def build_result_scope_text(
    *,
    is_ratio_mode,
    output_count,
    analysis_time_seconds=0.0,
    last_output_car_no=None,
    last_output_car_out=None,
):
    """Return all user-facing scope wording for the current result mode."""
    output_count = max(0, int(output_count or 0))
    try:
        analysis_seconds = max(0.0, float(analysis_time_seconds or 0.0))
    except (TypeError, ValueError):
        analysis_seconds = 0.0
    analysis_minutes = analysis_seconds / 60.0 if analysis_seconds > 0 else 0.0

    if is_ratio_mode:
        minutes_text = _fmt_number(analysis_minutes, "0")
        seconds_text = _fmt_number(analysis_seconds, "0")
        title = f"模型结果（分析窗口终值：{minutes_text}分钟）"
        note = f"统计范围：按分析时间{minutes_text}分钟（{seconds_text}s）统计"
        vehicle_definition = "表示：在设定分析时间内，已经完成最后一道工程并下线的车辆数量。"
        vehicle_rule = "计算口径：下线完成时间 ≤ 分析时间。"
        overall_definition = "表示：分析时间内已下线车辆的整体下线节奏。"
        if output_count <= 0:
            vehicle_current = (
                "本次计算：当前分析时间内暂无车辆完成下线"
                f"（窗口{minutes_text}分钟，{seconds_text}s），所以当前下线车辆为0台。"
            )
        elif last_output_car_no is not None and last_output_car_out is not None:
            vehicle_current = (
                f"本次计算：Car#{last_output_car_no}下线完成时间"
                f"{_fmt_number(last_output_car_out)}s ≤ {seconds_text}s，"
                f"所以分析窗口内下线车辆为{output_count}台。"
            )
        else:
            vehicle_current = f"本次计算：分析窗口内共有{output_count}台车辆完成下线。"
    else:
        title = f"模型结果（目标批次终值：{output_count}台）"
        note = f"统计范围：按目标批次共{output_count}台统计"
        vehicle_definition = "表示：目标批次中，已经完成最后一道工程并下线的车辆数量。"
        vehicle_rule = "计算口径：按本次输入的A/B/C目标批次车辆统计。"
        overall_definition = "表示：目标批次全部车辆的整体下线节奏。"
        if output_count <= 0:
            vehicle_current = "本次计算：目标批次暂无车辆完成下线。"
        elif last_output_car_no is not None and last_output_car_out is not None:
            vehicle_current = (
                f"本次计算：目标批次共{output_count}台，最后完成车辆为"
                f"Car#{last_output_car_no}，下线完成时间为{_fmt_number(last_output_car_out)}s。"
            )
        else:
            vehicle_current = f"本次计算：目标批次共{output_count}台车辆完成下线。"

    return {
        "mode": "ratio" if is_ratio_mode else "quantity",
        "title": title,
        "note": note,
        "vehicle_definition": vehicle_definition,
        "vehicle_rule": vehicle_rule,
        "vehicle_current": vehicle_current,
        "overall_definition": overall_definition,
        "analysis_time_minutes": analysis_minutes if is_ratio_mode else 0.0,
        "analysis_time_seconds": analysis_seconds if is_ratio_mode else 0.0,
    }
