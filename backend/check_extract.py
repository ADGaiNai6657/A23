import argparse
import json
from collections import Counter
from hashlib import sha256
from pathlib import Path

from config import DATA
from extractor import ModelFormatError, extract_document
from filler import write_json
from store import Store


def export_facts(store: Store, paths: list[Path], output: Path) -> list[dict]:
    suffixes = Counter(path.suffix.lower()[1:] for path in paths)
    manifest_path = output / "facts_export_manifest.json"
    prior = json.loads(manifest_path.read_text(encoding="utf-8")) if manifest_path.is_file() else []
    registry = {item["doc"]: item for item in prior}
    owners = {item["file"]: item["doc"] for item in prior}
    manifest = []
    introductory = []
    for path in paths:
        suffix = path.suffix.lower()[1:]
        rows = store.facts(path.name)
        target = output / f"facts_export_{suffix}.json"
        previous = json.loads(target.read_text(encoding="utf-8")) if target.is_file() else []
        if path.name in registry:
            target = output / Path(registry[path.name]["file"]).name
        elif (
            suffixes[suffix] > 1
            or owners.get(target.name, path.name) != path.name
            or any(row["doc"] != path.name for row in previous)
        ):
            target = output / f"facts_export_{suffix}_{path.stem}.json"
        exported = rows[:200] if suffix == "xlsx" else rows
        write_json(target, exported)
        item = {
            "doc": path.name, "file": target.name, "total": len(rows), "exported": len(exported),
        }
        registry[path.name] = item
        owners[target.name] = path.name
        manifest.append(item)
        if suffix in {"txt", "md"}:
            introductory.extend(rows)
    if {"txt", "md"} <= suffixes.keys():
        write_json(output / "facts_export.json", introductory)
    write_json(manifest_path, list(registry.values()))
    return manifest


def recorded_extract(path: Path, content: bytes, metrics: list[str], store: Store, workers: int) -> dict:
    audit = {}
    try:
        extract_document(content, path.name, metrics, store, workers=workers, audit=audit)
    finally:
        if audit:
            write_json(store.path.parent / "extraction-runs" / f"{audit['run_id']}.json", audit)
    return audit


def compare_runs(first: dict, second: dict) -> dict:
    configuration = ("doc", "input_sha256", "mode", "metrics", "batch_size", "workers",
                     "provider", "model", "temperature", "timeout_seconds")
    if any(first[key] != second[key] for key in configuration) or [
        (batch["blocks"], batch["attempts"][0]["prompt_sha256"]) for batch in first["batches"]
    ] != [
        (batch["blocks"], batch["attempts"][0]["prompt_sha256"]) for batch in second["batches"]
    ]:
        raise ValueError("两轮的输入、配置或初始批次 Prompt 不一致")

    def indexed(run: dict) -> Counter:
        return Counter(json.dumps(fact, ensure_ascii=False, sort_keys=True) for fact in run["stored_facts"])

    left, right = indexed(first), indexed(second)
    return {
        "configuration": {key: first[key] for key in configuration},
        "runs": [
            {"run_id": run["run_id"], **run["result"], "traceable": len(run["stored_facts"]),
             "attempts": sum(len(batch["attempts"]) for batch in run["batches"]),
             "model_calls": sum(len(batch["attempts"]) for batch in run["batches"]) if run["provider"] == "deepseek" else 0,
             "seconds": run["seconds"]}
            for run in (first, second)
        ],
        "shared": sum((left & right).values()),
        "first_only": [json.loads(item) for item in (left - right).elements()],
        "second_only": [json.loads(item) for item in (right - left).elements()],
    }


def replay_run(path: Path, content: bytes, original: dict, store: Store) -> dict:
    if (path.name, sha256(content).hexdigest()) != (original["doc"], original["input_sha256"]):
        raise ValueError("重放用的文件名或原文字节与记录不符")
    if not original.get("batches") or any(not batch.get("attempts") for batch in original["batches"]):
        raise ValueError("重放记录缺少批次或请求")
    attempts = iter(attempt for batch in original["batches"] for attempt in batch["attempts"])

    def respond(prompt: str) -> dict:
        attempt = next(attempts, None)
        if attempt is None:
            raise ValueError("重放记录缺少批次响应")
        if sha256(prompt.encode()).hexdigest() != attempt["prompt_sha256"]:
            raise ValueError("重放的 Prompt 与原请求不符")
        payload = attempt["response"]
        try:
            return json.loads(payload) if isinstance(payload, str) else payload
        except json.JSONDecodeError as error:
            raise ModelFormatError("记录的响应无法解析", payload) from error

    audit = {}
    extract_document(content, path.name, original["metrics"], store, respond, audit=audit)
    if next(attempts, None) is not None:
        raise ValueError("记录中还有未重放的响应")
    for key in ("batch_size", "block_count", "sent_blocks", "mode"):
        if original.get(key) != audit[key]:
            raise ValueError(f"重放记录的 {key} 与本次解析不一致")

    def batch_signatures(run: dict) -> list[tuple[list[int], list[str]]]:
        return [(batch["blocks"], [a["prompt_sha256"] for a in batch["attempts"]]) for batch in run["batches"]]

    if batch_signatures(original) != batch_signatures(audit):
        raise ValueError("重放记录的批次覆盖、顺序或请求边界不一致")
    audit.update({"provider": "recorded-response", "original_run_id": original["run_id"]})
    return audit


def main() -> None:
    parser = argparse.ArgumentParser(description="调用 DeepSeek 核对通用模式和指标清单模式")
    parser.add_argument("files", type=Path, nargs="+")
    parser.add_argument("--metrics", help="清单模式指标，以逗号分隔")
    parser.add_argument("--data-dir", type=Path, default=DATA)
    parser.add_argument("--export-dir", type=Path, help="导出目录，默认与数据库相同")
    parser.add_argument("--workers", type=int, choices=range(1, 5), default=1, help="并发批次数")
    parser.add_argument("--export-only", action="store_true", help="从已有库导出，不调用模型")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--repeat-list", action="store_true",
                      help="仅对一份原文做两轮清单抽取，需指定不存在的输出目录，会实际调用模型")
    mode.add_argument("--replay", type=Path, nargs=2, metavar="RECORD",
                      help="用两份原始响应记录重跑后处理，无网络调用")
    args = parser.parse_args()
    metrics = [name.strip() for name in (args.metrics or "").split(",") if name.strip()]
    if not args.export_only and not args.replay and not metrics:
        parser.error("至少指定一个指标")
    if len({path.name for path in args.files}) != len(args.files):
        parser.error("输入文件名必须唯一，避免文档覆盖")
    if args.repeat_list or args.replay:
        if args.export_only or len(args.files) != 1:
            parser.error("重复抽取或重放需要一份原文")
        args.data_dir.mkdir(parents=True, exist_ok=False)
        path, = args.files
        content = path.read_bytes()
        runs = []
        for i in (1, 2):
            trial = args.data_dir / f"run-{i}"
            store = Store(trial / "a23.db")
            if args.replay:
                original = json.loads(args.replay[i - 1].read_text(encoding="utf-8"))
                run = replay_run(path, content, original, store)
                write_json(trial / "replay.json", run)
            else:
                run = recorded_extract(path, content, metrics, store, args.workers)
            runs.append(run)
        first, second = runs
        comparison = compare_runs(first, second)
        target = args.data_dir / "comparison.json"
        write_json(target, comparison)
        print(json.dumps(comparison["runs"], ensure_ascii=False), flush=True)
        print(f"两轮共同事实 {comparison['shared']} 条，差异见 {target}", flush=True)
        return
    store = Store(args.data_dir / "a23.db")
    for path in args.files:
        if args.export_only:
            if store.content(path.name) is None:
                parser.error(f"库中没有文档：{path.name}")
            continue
        content = path.read_bytes()
        for selection in ([], metrics):
            audit = recorded_extract(path, content, selection, store, args.workers)
            print(json.dumps({
                **audit["result"], **{key: audit[key] for key in (
                    "run_id", "metrics", "seconds", "block_count", "model", "temperature", "batch_size",
                )},
            }, ensure_ascii=False), flush=True)
    for item in export_facts(store, args.files, args.export_dir or args.data_dir):
        print(f"{item['doc']}：入库 {item['total']} 条，导出 {item['exported']} 条至 {item['file']}")


if __name__ == "__main__":
    main()
