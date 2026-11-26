from pathlib import Path
from zipfile import ZipFile

from config import ROOT


def unpack_dataset(archive: Path, output: Path) -> None:
    with ZipFile(archive, metadata_encoding="gb18030") as compressed:
        compressed.extractall(output)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="按发布压缩包的文件名编码解压测试集")
    parser.add_argument("--output", type=Path, default=ROOT / "dataset" / "测试集")
    args = parser.parse_args()
    if args.output.exists():
        parser.error("请指定不存在的目录，以保留已解压的资料")
    unpack_dataset(ROOT / "dataset" / "测试集.zip", args.output)
    print(f"已解压到 {args.output}")
