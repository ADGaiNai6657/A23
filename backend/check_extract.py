import argparse
import json
from collections import Counter
from pathlib import Path
from time import perf_counter

from config import DATA
from extractor import BATCH_SIZE, MODEL, extract_document
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


def main() -> None:
    parser = argparse.ArgumentParser(description="调用 DeepSeek 核对通用模式和指标清单模式")
    parser.add_argument("files", type=Path, nargs="+")
    parser.add_argument("--metrics", help="清单模式指标，以逗号分隔")
    parser.add_argument("--data-dir", type=Path, default=DATA)
    parser.add_argument("--export-dir", type=Path, help="导出目录，默认与数据库相同")
    parser.add_argument("--workers", type=int, choices=range(1, 5), default=1, help="并发批次数")
    parser.add_argument("--export-only", action="store_true", help="从已有库导出，不调用模型")
    args = parser.parse_args()
    metrics = [name.strip() for name in (args.metrics or "").split(",") if name.strip()]
    if not args.export_only and not metrics:
        parser.error("至少指定一个指标")
    if len({path.name for path in args.files}) != len(args.files):
        parser.error("输入文件名必须唯一，避免文档覆盖")
    store = Store(args.data_dir / "a23.db")
    for path in args.files:
        if args.export_only:
            if store.content(path.name) is None:
                parser.error(f"库中没有文档：{path.name}")
            continue
        content = path.read_bytes()
        for selection in ([], metrics):
            started = perf_counter()
            result = extract_document(content, path.name, selection, store, workers=args.workers)
            print(json.dumps({
                **result, "metrics": selection, "seconds": round(perf_counter() - started, 2),
                "blocks": len(store.blocks(path.name)), "model": MODEL, "temperature": 0, "batch_size": BATCH_SIZE,
            }, ensure_ascii=False), flush=True)
    for item in export_facts(store, args.files, args.export_dir or args.data_dir):
        print(f"{item['doc']}：入库 {item['total']} 条，导出 {item['exported']} 条至 {item['file']}")


if __name__ == "__main__":
    main()
