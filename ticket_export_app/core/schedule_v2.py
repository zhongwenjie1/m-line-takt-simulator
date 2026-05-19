# -*- coding: utf-8 -*-
"""
排程模型 v2 主算法模块。

职责边界：
- 后续用于承接排程模型 v2 主算法。
- 不负责 Excel 导出。
- 不负责 UI。
- 不负责结果分析；分析逻辑放在 core.analysis。

当前阶段：
- 为避免一次性搬迁导致排程行为变化，先作为安全代理入口。
- 真实且已验证的 schedule 实现暂时仍保留在 core.tickets 中。
- 后续再把 core.tickets 中的真实 schedule 主体完整迁移到本文件。
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple


def schedule(
    step_defs: List[Dict[str, Any]],
    cars: int,
    vehicle_counts: Dict[str, int] | None = None,
    sequence_mode: str = "grouped",
    max_consecutive: int = 10,
    ratio_pattern: Optional[Dict[str, int]] = None,
    launch_takt: Optional[float] = None,
) -> Tuple[List[Dict[str, Any]], float]:
    """
    schedule_v2 安全代理入口。

    当前阶段动态调用 core.tickets 中保留的已验证 schedule 实现，避免循环导入和行为变化。
    完整迁移完成后，本函数将改为真正的 v2 排程主体。
    """
    from core import tickets

    legacy_schedule = getattr(tickets, "_schedule_legacy", None)
    if legacy_schedule is None:
        legacy_schedule = getattr(tickets, "schedule")

    return legacy_schedule(
        step_defs=step_defs,
        cars=cars,
        vehicle_counts=vehicle_counts,
        sequence_mode=sequence_mode,
        max_consecutive=max_consecutive,
        ratio_pattern=ratio_pattern,
        launch_takt=launch_takt,
    )