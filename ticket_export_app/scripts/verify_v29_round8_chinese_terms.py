#!/usr/bin/env python3
"""Verify round-8 user-facing terminology without changing internal row keys."""

from __future__ import annotations

from pathlib import Path


APP_DIR = Path(__file__).resolve().parents[1]
UI_FILE = APP_DIR / "ui" / "export_ticket_window.py"


def main() -> None:
    source = UI_FILE.read_text(encoding="utf-8")
    required = [
        "资源标识",
        "加工完成",
        "离开工程",
        "加工工时",
        "投入等待",
        "完工后等待",
        "相邻下线间隔",
        "IN=实际投车时间",
        "OUT=下线完成时间",
        "WAIT=单车总等待",
        "FLOW=单车贯通时间",
        "SEGMENTS",
    ]
    missing = [term for term in required if term not in source]
    if missing:
        raise AssertionError(f"缺少中文术语：{missing}")

    retained = ['"svc_finish"', '"depart"', '"dur"', '"launch_wait"', '"block_wait"', '"resource_key"']
    absent_internal_keys = [key for key in retained if key not in source]
    if absent_internal_keys:
        raise AssertionError(f"内部字段读取被意外移除：{absent_internal_keys}")

    forbidden_headers = [
        '"资源 key"',
        '"svc_finish"',
        '"depart"',
        '"dur"',
        '"block_wait"',
        '"launch_wait"',
    ]
    columns_block = source[source.index("columns = ["):source.index("output = []", source.index("columns = ["))]
    leaked = [header for header in forbidden_headers if header in columns_block]
    if leaked:
        raise AssertionError(f"结构化日志仍显示内部字段名：{leaked}")

    user_log_block = source[source.index('lines = [', source.index('def _build_schedule_debug_log')):source.index('sorted_car_items =', source.index('def _build_schedule_debug_log'))]
    leaked_log_terms = [term for term in ("svc_finish", "depart", "dur", "launch_wait", "block_wait") if term in user_log_block]
    if leaked_log_terms:
        raise AssertionError(f"车辆过程日志仍显示内部专业字段：{leaked_log_terms}")

    print("round8 terminology checks passed")


if __name__ == "__main__":
    main()
