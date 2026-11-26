import json
import re
from hashlib import file_digest
from pathlib import Path

from store import Store


RUN_EVIDENCE = "acceptance_run.json"
HISTORY_INPUT = "history-input.json"


def digest(path: Path) -> str:
    with path.open("rb") as stream:
        return file_digest(stream, "sha256").hexdigest()


def check_execution(directory: Path, mode: str, reports: dict, *, rebuilt_kind: str = "online") -> None:
    path = directory / RUN_EVIDENCE
    if not path.is_file():
        raise ValueError(f"缺少 {RUN_EVIDENCE}，请生成本轮同库或独立重建记录")
    evidence = json.loads(path.read_bytes())
    if evidence.get("mode") != mode:
        raise ValueError(f"{directory} 记录模式为 {evidence.get('mode')!r}，不能用于 {mode}")
    database = directory / "a23.db"
    db_hash = digest(database)
    if evidence.get("database_sha256") != db_hash or evidence.get("outputs") != {
        case: {key: report[key] for key in ("run_id", "source_sha256", "template_sha256", "workbook_sha256")}
        for case, report in reports.items()
    }:
        raise ValueError(f"{RUN_EVIDENCE} 与当前数据库、原文、模板或填写批次不一致")
    initial_hash = evidence.get("initial_database_sha256")
    if mode == "same-db":
        if not re.fullmatch(r"[0-9a-f]{64}", initial_hash or ""):
            raise ValueError("同库验收缺少运行前已有数据库的摘要")
        return
    if initial_hash is not None or evidence.get("extraction_kind") != rebuilt_kind:
        raise ValueError(f"重建验收需要空库和 {rebuilt_kind} 来源证据")
    run_id = evidence.get("extraction_id", "")
    if rebuilt_kind == "online":
        if not re.fullmatch(r"[0-9a-f]{32}", run_id):
            raise ValueError("重建验收缺少有效的模型运行 ID")
        source = directory / "extraction-runs" / f"{run_id}.json"
        recorded = json.loads(source.read_bytes())
        batches = recorded.get("batches")
        if (recorded.get("run_id") != run_id or recorded.get("provider") != "deepseek" or recorded.get("error")
                or not batches or not all(
                    batch.get("attempts") and batch["attempts"][-1].get("response") for batch in batches
                )
                or recorded.get("input_sha256") != reports["01"]["source_sha256"]):
            raise ValueError("重建验收缺少同一原文的完整 DeepSeek 响应")
        facts = recorded["stored_facts"]
    else:
        source = directory / HISTORY_INPUT
        recorded = json.loads(source.read_bytes())
        if (recorded.get("extraction_id") != run_id
                or recorded.get("source_sha256") != reports["01"]["source_sha256"]):
            raise ValueError("历史重建归档与原文或抽取 ID 不一致")
        facts = recorded["facts"]
    if evidence.get("extraction_sha256") != digest(source) or not facts:
        raise ValueError("抽取记录摘要不一致或缺少事实")
    if Store(database).facts(reports["01"]["doc"]) != facts:
        raise ValueError("数据库中的城市事实与本轮抽取留档不一致")


def check_distinct_runs(same: dict, rebuilt: dict) -> None:
    if same["01"]["database_sha256"] == rebuilt["01"]["database_sha256"]:
        raise ValueError("同库与重建库的 SHA256 完全相同，无法作为独立重建证据")
    for case in same:
        if any(same[case][key] != rebuilt[case][key] for key in ("source_sha256", "template_sha256")):
            raise ValueError(f"{case} 同库与重建库的原文或模板不同")
        if same[case]["run_id"] == rebuilt[case]["run_id"]:
            raise ValueError(f"{case} 同库与重建库指向同一次填写")
