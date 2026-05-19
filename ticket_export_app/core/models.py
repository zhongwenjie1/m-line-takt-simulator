# -*- coding: utf-8 -*-
from dataclasses import dataclass
from typing import Optional


@dataclass
class StationDef:
    seq: int
    display: str
    group: str
    device_count: int = 1
    line_scope: str = ""
    duration_a: float = 0.0
    duration_b: float = 0.0
    duration_c: float = 0.0


@dataclass
class VehicleFinish:
    car: int
    finish_time: float


@dataclass
class TimeWindowResult:
    planned_output_count: int
    actual_output_count: int
    achievement_rate: float
    planned_n_finish_time: Optional[float]
    actual_n_finish_time: Optional[float]
    finish_delta: Optional[float]
    actual_line_takt: Optional[float]
    result: str
