#!/usr/bin/env python3
"""Verify round-16L frozen sequence export/import hard checks."""

from __future__ import annotations

import copy
import json
import os
import sys
import tempfile
from pathlib import Path


os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from PySide6.QtWidgets import QApplication  # noqa: E402
from core import tickets  # noqa: E402
from ui.export_ticket_window import ExportTicketWindow  # noqa: E402


EXPECTED_SEQUENCE_SHA = "dac6086f3b896255eeb77cb82679b0334e49ab972a5e9ed6aa0cd4e6d347f08f"


def _set_quantity_alternate(window: ExportTicketWindow, *, max_run: int = 5, a: int = 276) -> None:
    window.cmb_launch_mode.setCurrentIndex(0)
    window.cmb_seq.setCurrentIndex(1)
    window.spn_a_cars.setValue(a)
    window.spn_b_cars.setValue(306)
    window.spn_c_cars.setValue(566)
    window.spn_max_run.setValue(max_run)


def _seed_frozen_sequence(window: ExportTicketWindow) -> list[str]:
    vehicle_counts = {"A": 276, "B": 306, "C": 566}
    sequence = tickets.build_vehicle_sequence(
        sum(vehicle_counts.values()),
        vehicle_counts,
        "alternate",
        5,
    )
    sequence_hash = tickets.vehicle_sequence_hash(sequence)
    assert sequence_hash == EXPECTED_SEQUENCE_SHA
    window._frozen_vehicle_sequence = list(sequence)
    window._frozen_vehicle_sequence_signature = window._sequence_freeze_signature()
    window._frozen_vehicle_sequence_hash = sequence_hash
    window._frozen_vehicle_sequence_generated_at = "2026-06-16T00:00:00"
    return list(sequence)


def _rehash_payload(window: ExportTicketWindow, payload: dict) -> dict:
    payload = copy.deepcopy(payload)
    payload.pop("payload_sha256", None)
    payload["payload_sha256"] = window._canonical_json_sha256(payload)
    return payload


def _assert_rejected(window: ExportTicketWindow, payload: dict, expected_text: str) -> None:
    try:
        window._apply_imported_frozen_vehicle_sequence(payload)
    except ValueError as exc:
        assert expected_text in str(exc), str(exc)
        return
    raise AssertionError(f"应拒绝导入：{expected_text}")


def main() -> None:
    app = QApplication.instance() or QApplication([])

    source = ExportTicketWindow()
    _set_quantity_alternate(source)
    sequence = _seed_frozen_sequence(source)
    payload = source._build_sequence_freeze_payload()
    assert payload["app"]["version"] == "2.9.0"
    assert payload["sequence_rule_version"] == tickets.QUANTITY_SEQUENCE_RULE_VERSION
    assert payload["vehicle_counts"] == {"A": 276, "B": 306, "C": 566}
    assert payload["max_consecutive"] == 5
    assert payload["sequence_sha256"] == EXPECTED_SEQUENCE_SHA
    assert payload["actual_max_run"] == 5
    assert payload["preview_first_30"] == "CCCBAAABCCCBBBCCAACCCCBBAACCCC"

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "freeze.mline-sequence.json"
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        loaded = json.loads(path.read_text(encoding="utf-8"))

    imported = ExportTicketWindow()
    _set_quantity_alternate(imported)
    imported_sequence, imported_hash = imported._apply_imported_frozen_vehicle_sequence(loaded)
    assert imported_sequence == sequence
    assert imported_hash == EXPECTED_SEQUENCE_SHA
    assert tickets.vehicle_sequence_hash(imported._frozen_sequence_for_current_inputs("alternate")) == EXPECTED_SEQUENCE_SHA

    version_bad = copy.deepcopy(payload)
    version_bad["app"]["version"] = "2.8.0"
    version_bad = _rehash_payload(imported, version_bad)
    _assert_rejected(imported, version_bad, "程序版本不一致")

    max_run_bad_window = ExportTicketWindow()
    _set_quantity_alternate(max_run_bad_window, max_run=4)
    _assert_rejected(max_run_bad_window, payload, "最大连续台数不一致")

    count_bad_window = ExportTicketWindow()
    _set_quantity_alternate(count_bad_window, a=275)
    _assert_rejected(count_bad_window, payload, "A/B/C数量不一致")

    ratio_window = ExportTicketWindow()
    _set_quantity_alternate(ratio_window)
    ratio_window.cmb_launch_mode.setCurrentIndex(1)
    _assert_rejected(ratio_window, payload, "当前界面必须切换到按数量投车 + 交替混流")

    tampered = copy.deepcopy(payload)
    first_c = tampered["sequence"].index("C")
    first_a = tampered["sequence"].index("A")
    tampered["sequence"][first_c], tampered["sequence"][first_a] = (
        tampered["sequence"][first_a],
        tampered["sequence"][first_c],
    )
    tampered = _rehash_payload(imported, tampered)
    _assert_rejected(imported, tampered, "序列SHA-256校验失败")

    source.close()
    imported.close()
    max_run_bad_window.close()
    count_bad_window.close()
    ratio_window.close()
    app.processEvents()

    print({
        "sequence_sha256": EXPECTED_SEQUENCE_SHA,
        "export_schema": payload["schema"],
        "version_hard_check": True,
        "rule_version_check": True,
        "parameter_check": True,
        "tamper_check": True,
        "assertions": "passed",
    })


if __name__ == "__main__":
    main()
