import json
import math
from numbers import Real
from pathlib import Path

from config import BACKEND

EXPECTATIONS = BACKEND / "data" / "source_expectations.json"
FIELDS = {
    "01": {
        "block_count": "原文块数", "rows": "行数", "city_order_matches": "城市顺序一致行数",
        "cells_written": "写入格数", "cells_empty": "空着的格数",
        "cell_accuracy": "单元格准确率", "row_completeness": "行完整率",
    },
    "03": {
        "block_count": "原文块数", "rows": "行数", "required": "应填格",
        "expected_empty": "应空格", "cells_written": "写入格数", "cells_empty": "空着的格数",
        "correct": "对格数", "false_fills": "误填数", "missed": "漏填数",
        "correct_rate": "对格率", "false_fill_rate": "误填率", "missed_rate": "漏填率",
    },
}
MODEL_FIELDS = {"cells_written", "cells_empty", "cell_accuracy", "row_completeness"}


def load_expectations(path: Path = EXPECTATIONS) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def is_number(value: object) -> bool:
    return isinstance(value, Real) and not isinstance(value, bool) and math.isfinite(value)


def checked_rules(config: dict, mode: str) -> dict:
    exact = config["same-db"]
    if set(exact) != set(FIELDS):
        raise ValueError("验收配置必须同时包含场景 01 和 03")
    for case, fields in FIELDS.items():
        if set(exact[case]) != set(fields):
            raise ValueError(f"{case} 配置字段不完整或含未识别字段")
        for key, value in exact[case].items():
            if not is_number(value):
                raise ValueError(f"{case}.{key} 必须使用精确数值，配置为 {value!r}")
    if mode == "same-db":
        return exact
    intervals = config["rebuilt"]
    if set(intervals) != {"01"} or set(intervals["01"]) != MODEL_FIELDS:
        raise ValueError("重建区间只适用于 01 的写入格数、空格和两个率")
    for key, target in intervals["01"].items():
        valid = {
            point["extraction_id"]: point["01"][key] for point in config.get("observations", [])
            if point.get("extraction_id") and is_number(point.get("01", {}).get(key))
        }
        if (len(valid) < 2 or not isinstance(target, list) or len(target) != 2
                or not all(is_number(value) for value in target)
                or target != [min(valid.values()), max(valid.values())]):
            raise ValueError(f"01.{key} 区间需由至少两个独立重抽的实测值确定，配置为 {target!r}")
    return {**exact, "01": {**exact["01"], **intervals["01"]}}
