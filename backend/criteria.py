import json
import math
from hashlib import sha256
from numbers import Real
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from config import BACKEND

EXPECTATIONS = BACKEND / "data" / "source_expectations.json"
OBSERVATION_FILES = BACKEND / "data" / "deliverables" / "source"
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


def observed_scores(config: dict) -> list[dict]:
    """Bind the historical numeric points to the archived facts and score, rather than labels alone."""
    identities = set()
    observations = []
    for point in config.get("observations", []):
        try:
            extraction_id, archive = point["extraction_id"], point["score_archive"]
            if Path(extraction_id).name != extraction_id or Path(archive).name != archive:
                raise ValueError("归档名无效")
            raw_facts = (OBSERVATION_FILES / (extraction_id + ".json")).read_bytes()
            with ZipFile(OBSERVATION_FILES / archive) as bundle:
                raw_score = bundle.read("source_report.json")
            facts, score = json.loads(raw_facts), json.loads(raw_score)
            identity = (sha256(raw_facts).hexdigest(), sha256(raw_score).hexdigest())
            if (identity != (point["facts_sha256"], point["score_sha256"]) or identity in identities
                    or facts["extraction_id"] != extraction_id or not facts["facts"]
                    or facts["source_sha256"] != score["source_sha256"] or score["run_id"] != point["run_id"]
                    or any(not is_number(score.get(key)) or score[key] != point["01"].get(key) for key in MODEL_FIELDS)):
                raise ValueError("归档摘要、运行 ID、原文或数值不一致")
        except (OSError, KeyError, ValueError, TypeError, BadZipFile) as error:
            raise ValueError(f"重建区间需由至少两个独立重抽的实测值确定。观察点无效：{error}") from error
        identities.add(identity)
        observations.append(score)
    if (len(observations) < 2 or len({row["run_id"] for row in observations}) != len(observations)
            or len({facts_hash for facts_hash, _ in identities}) != len(observations)
            or len({row["source_sha256"] for row in observations}) != 1):
        raise ValueError("重建区间需由至少两个独立重抽的实测值确定")
    return observations


def checked_rules(config: dict, mode: str, *, reference_points: list[dict] | None = None) -> dict:
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
    if config.get("basis") == "reference-workbook":
        if not reference_points or len(reference_points) < 2:
            raise ValueError("参考重建判据需要至少两个已核验的独立观察点")
        observations = reference_points
    else:
        observations = observed_scores(config)
    for key, target in intervals["01"].items():
        valid = [score[key] for score in observations]
        if (not isinstance(target, list) or len(target) != 2
                or not all(is_number(value) for value in target)
                or target != [min(valid), max(valid)]):
            raise ValueError(f"01.{key} 区间需由至少两个独立重抽的实测值确定，配置为 {target!r}")
    return {**exact, "01": {**exact["01"], **intervals["01"]}}
