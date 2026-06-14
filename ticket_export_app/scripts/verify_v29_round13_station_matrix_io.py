#!/usr/bin/env python3
"""Verify round-13 station-matrix Excel round-trip and validation safety."""

from __future__ import annotations

import importlib.util
import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from PySide6.QtWidgets import QApplication  # noqa: E402
from ui.export_ticket_window import ExportTicketWindow  # noqa: E402


def _load_baseline():
    path = APP_DIR / "scripts" / "verify_v29_real_data_baseline.py"
    spec = importlib.util.spec_from_file_location("v29_baseline", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _matrix_rows(definitions):
    return [
        [
            str(item["seq"]), str(item["display"]), str(item["device_count"]),
            str(item["line_scope"]), str(item["group"]),
            str(item["duration_a"]), str(item["duration_b"]), str(item["duration_c"]),
        ]
        for item in definitions
    ]


def main():
    baseline = _load_baseline()
    app = QApplication.instance() or QApplication([])
    window = ExportTicketWindow()
    assert window.btn_import_matrix.text() == "导入矩阵"
    assert window.btn_export_matrix.text() == "导出矩阵"
    expected = window._validate_station_matrix_rows(_matrix_rows(baseline.station_defs()))
    window._apply_station_matrix_rows(expected)
    window.spn_a_cars.setValue(276)
    window.spn_b_cars.setValue(306)
    window.spn_c_cars.setValue(566)
    before = window._collect_station_matrix_rows()
    parsed = window._parse_multi_inputs_from_raw()
    parsed_defs = parsed["defs"]
    assert len(parsed_defs) == len(baseline.station_defs())
    for actual, source in zip(parsed_defs, baseline.station_defs()):
        assert actual["seq"] == source["seq"]
        assert actual["display"] == source["display"]
        assert actual["device_count"] == source["device_count"]
        assert actual["line_scope"] == source["line_scope"]
        assert actual["duration_a"] == source["duration_a"]
        assert actual["duration_b"] == source["duration_b"]
        assert actual["duration_c"] == source["duration_c"]

    export_path = Path(tempfile.gettempdir()) / "mline_round13_station_matrix.xlsx"
    window._write_station_matrix_xlsx(export_path, expected)
    version, imported = window._read_station_matrix_xlsx(export_path)
    assert version == "v2.9"
    assert imported == expected

    window._apply_station_matrix_rows(imported)
    after = window._collect_station_matrix_rows()
    assert after == before == expected
    assert any(row[5] == "0" or row[6] == "0" or row[7] == "0" for row in imported)

    invalid = [list(row) for row in expected]
    invalid[1][2] = "3"
    try:
        window._validate_station_matrix_rows(invalid)
    except ValueError as exc:
        assert "设备数量" in str(exc)
    else:
        raise AssertionError("无效设备数量未被拒绝")
    assert window._collect_station_matrix_rows() == before

    duplicate = [list(row) for row in expected]
    duplicate[1][1] = duplicate[0][1]
    try:
        window._validate_station_matrix_rows(duplicate)
    except ValueError as exc:
        assert "重复" in str(exc)
    else:
        raise AssertionError("重复工程名未被拒绝")
    assert window._collect_station_matrix_rows() == before

    export_path.unlink(missing_ok=True)
    print({
        "template_version": version,
        "station_rows": len(imported),
        "round_trip_equal": imported == expected,
        "parsed_station_defs_equal": True,
        "zero_duration_preserved": True,
        "invalid_input_kept_current_matrix": True,
    })
    window.close()


if __name__ == "__main__":
    main()
