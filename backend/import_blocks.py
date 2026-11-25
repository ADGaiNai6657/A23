import argparse
from pathlib import Path

from parsers import parse_document
from store import Store


def main() -> None:
    parser = argparse.ArgumentParser(description="导入样例文档块，无模型调用")
    parser.add_argument("files", type=Path, nargs="+")
    args = parser.parse_args()
    store = Store()
    for path in args.files:
        content = path.read_bytes()
        parsed = parse_document(content, path.name)
        store.save(parsed, content)
        print(f"{path.name}：{len(parsed.blocks)} 块")


if __name__ == "__main__":
    main()
