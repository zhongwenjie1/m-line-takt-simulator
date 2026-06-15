#!/usr/bin/env python3
"""Verify round-10 terminal/process scope labels and unchanged values."""

from __future__ import annotations

from pathlib import Path


APP_DIR = Path(__file__).resolve().parents[1]
UI_FILE = APP_DIR / "ui" / "export_ticket_window.py"
SCOPE_TEXT_FILE = APP_DIR / "utils" / "result_scope_text.py"


def main() -> None:
    ui_source = UI_FILE.read_text(encoding="utf-8")
    wording_source = SCOPE_TEXT_FILE.read_text(encoding="utf-8")
    source = f"{ui_source}\n{wording_source}"
    required = [
        "模型结果（分析窗口终值：",
        "模型结果（目标批次终值：",
        "统计范围：按分析时间",
        "统计范围：按目标批次",
        "当前播放：",
        "当前等待位置",
        "近期节拍（近5个间隔）",
    ]
    missing = [text for text in required if text not in source]
    if missing:
        raise AssertionError(f"缺少范围标注：{missing}")

    forbidden = [
        "当前仿真时间 {current_time}",
        ">模型提示</div>",
        "仿真时间：{current:.1f}s",
    ]
    leaked = [text for text in forbidden if text in ui_source]
    if leaked:
        raise AssertionError(f"仍存在混用文字：{leaked}")

    print("round10 scope label checks passed")


if __name__ == "__main__":
    main()
