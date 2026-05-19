# -*- coding: utf-8 -*-
import json
from pathlib import Path


def read_local_version(base_dir: str | Path):
    version_path = Path(base_dir) / "version.json"
    if not version_path.exists():
        return {}
    try:
        return json.loads(version_path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def normalize_version_info(data):
    if not isinstance(data, dict):
        data = {}

    return {
        "version": str(data.get("version", "") or ""),
        "tag": str(data.get("tag", "") or ""),
        "commit": str(data.get("commit", "") or ""),
        "stage": str(data.get("stage", "") or ""),
    }


def compare_version_info(local_info, latest_info):
    local = normalize_version_info(local_info)
    latest = normalize_version_info(latest_info)

    if not latest["version"]:
        return {
            "has_update": False,
            "local": local,
            "latest": latest,
            "reason": "latest_version_missing",
        }

    if not local["version"]:
        return {
            "has_update": True,
            "local": local,
            "latest": latest,
            "reason": "local_version_missing",
        }

    if latest["version"] != local["version"]:
        return {
            "has_update": True,
            "local": local,
            "latest": latest,
            "reason": "version_changed",
        }

    if latest["tag"] and local["tag"] and latest["tag"] != local["tag"]:
        return {
            "has_update": True,
            "local": local,
            "latest": latest,
            "reason": "tag_changed",
        }

    if latest["commit"] and local["commit"] and latest["commit"] != local["commit"]:
        return {
            "has_update": True,
            "local": local,
            "latest": latest,
            "reason": "commit_changed",
        }

    return {
        "has_update": False,
        "local": local,
        "latest": latest,
        "reason": "up_to_date",
    }


def check_update_from_info(base_dir: str | Path, latest_info):
    local_info = read_local_version(base_dir)
    return compare_version_info(local_info, latest_info)


def format_update_message(result):
    if not isinstance(result, dict):
        result = {}

    local = normalize_version_info(result.get("local"))
    latest = normalize_version_info(result.get("latest"))
    reason = str(result.get("reason", "") or "")

    local_version = local["version"] or "未知"
    local_tag = local["tag"] or "无标签"
    latest_version = latest["version"] or "未知"
    latest_tag = latest["tag"] or "无标签"

    if bool(result.get("has_update", False)):
        return (
            "发现新版本：\n"
            f"当前版本：{local_version} ({local_tag})\n"
            f"最新版本：{latest_version} ({latest_tag})\n"
            f"原因：{reason}"
        )

    return f"当前已是最新版本：{local_version} ({local_tag})"
