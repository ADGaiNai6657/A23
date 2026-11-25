import argparse
import json
from pathlib import Path

from config import DATA
from extractor import extract_document
from store import Store


def main() -> None:
    parser = argparse.ArgumentParser(description="调用 DeepSeek 核对通用模式和指标清单模式")
    parser.add_argument("files", type=Path, nargs="+")
    parser.add_argument("--metrics", required=True, help="清单模式指标，以逗号分隔")
    args = parser.parse_args()
    metrics = [name.strip() for name in args.metrics.split(",") if name.strip()]
    if not metrics:
        parser.error("至少指定一个指标")
    store = Store()
    for path in args.files:
        content = path.read_bytes()
        for selection in ([], metrics):
            print(json.dumps(extract_document(content, path.name, selection, store), ensure_ascii=False))
        rows = store.facts(path.name)
        count = len(rows)
        if path.suffix.lower() == ".xlsx":
            rows = rows[:200]
        output = DATA / f"facts_export_{path.suffix.lower().lstrip('.')}.json"
        output.write_text(json.dumps(rows, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"{path.name}：入库 {count} 条，导出 {len(rows)} 条至 {output.name}")


if __name__ == "__main__":
    main()
