import argparse
import hashlib
import json
import re
from collections import Counter
from datetime import date, datetime
from itertools import zip_longest
from pathlib import Path
from time import perf_counter

from docx import Document
from openpyxl import load_workbook

from config import BACKEND, ROOT
from evaluate import matches
from extractor import MODEL, extract_document
from fill_spec import CITY_OUTPUT, COVID_OUTPUT
from filler import fill_workbook, write_json
from parsers import parse_document
from store import Store

DATASET = ROOT / "dataset" / "测试集"
OUTPUT = BACKEND / "data" / "actual"
CITY_FOLDER = "2025年中国城市经济百强全景报告"
CITY_SOURCE = f"{CITY_FOLDER}.docx"
CITY_TEMPLATE = f"{CITY_FOLDER}-模板.xlsx"
CITY_METRICS = [
    "GDP总量", "常住人口", "人均GDP", "一般公共预算收入",
    "GDP", "地区生产总值", "人口基数", "人口严控至", "万人口", "人口",
]
DAILY_FOLDER = "COVID-19数据集"
DAILY_SOURCE = "COVID-19全球数据集（节选）.xlsx"
DAILY_TEMPLATE = "COVID-19 模板.xlsx"
DAILY_CONTEXT = "中国COVID-19新冠疫情情况.docx"
DAILY_FIELDS = ("国家/地区", "大洲", "人均GDP", "人口", "每日检测数", "病例数")
VALUE = r"(\d[\d,]*(?:\.\d+)?)"
CORPUS_ANCHORS = {
    "2024年药品流通行业运行统计分析报告.txt": {"text_line": 70},
    "2023年文化和旅游发展统计公报.md": {"text_line": 68},
    "南京市2021~2024年国民经济和社会发展统计公报汇总.docx": {"paragraph": 271, "table_cell": 946},
}


def city_expected(path: Path) -> list[tuple]:
    paragraphs = [
        paragraph.text for paragraph in Document(path).paragraphs
        if "一般公共预算收入" in paragraph.text
    ]
    expected = []
    for index, text in enumerate(paragraphs):
        name = re.split(r" GDP 总量|以|紧随|凭借|在转型|作为", text, maxsplit=1)[0]
        if not re.fullmatch(r"[\u4e00-\u9fff]{2,12}", name):
            raise ValueError(f"01 原文核对无法确定城市名，正文行 {index + 1}")
        # This audit reads the source quantities directly instead of using model facts.
        revenue = text.split("一般公共预算收入", 1)[1]
        numbers = []
        for section, unit in ((text, "亿元"), (text, "万(?!元|亿)"), (text, "元"), (revenue, "亿元")):
            found = re.search(VALUE + r"\s*" + unit, section)
            if not found:
                raise ValueError(f"01 原文核对缺少指标数量，城市 {name}")
            numbers.append(float(found[1].replace(",", "")))
        expected.append((name, *numbers))
    return expected


def daily_expected(path: Path) -> list[tuple]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        rows = workbook.active.values
        headers = next(rows)
        indices = [headers.index(field) for field in ("日期", *DAILY_FIELDS)]
        selected = {}
        for row in rows:
            if not isinstance(row[indices[0]], datetime):
                continue
            day = row[indices[0]].date()
            if date(2020, 7, 1) <= day <= date(2020, 8, 31):
                key = (day, row[indices[1]])
                if key in selected:
                    raise ValueError(f"03 原文核对发现重复日期和国家键 {key}")
                selected[key] = (day, *(row[index] for index in indices[1:]))
        return [selected[key] for key in sorted(selected)]
    finally:
        workbook.close()


def verify_workbook(path: Path, expected: list[tuple], width: int) -> dict:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.active
        rows = list(sheet.iter_rows(min_row=2, max_col=width, values_only=True))
        diffs = []
        nonempty = correct = false_fills = missed = 0
        for row, (actual, target) in enumerate(
            zip_longest(rows, expected, fillvalue=(None,) * width), 2,
        ):
            for column, (value, truth) in enumerate(zip(actual, target), 1):
                nonempty += truth is not None
                correct += truth is not None and matches(value, truth)
                false_fills += truth is None and value is not None
                missed += truth is not None and value is None
                if not matches(value, truth):
                    diffs.append({"row": row, "column": column, "actual": value, "source_value": truth})
        return {
            "rows": len(rows), "columns": sheet.max_column, "source_nonempty": nonempty,
            "source_empty": len(expected) * width - nonempty, "correct_nonempty": correct,
            "false_fills": false_fills, "missed": missed, "differences": diffs,
        }
    finally:
        workbook.close()


def verify_trace(store: Store, path: Path, rows: int, columns: int) -> None:
    report = json.loads(path.read_text(encoding="utf-8"))
    by_number = {block["number"]: block for block in store.blocks(report["doc"])}
    coordinates = {(item["row"], item["column"]) for item in report["cells"]}
    expected = {(row, column) for row in range(2, rows + 2) for column in range(1, columns + 1)}
    if coordinates != expected or len(report["cells"]) != len(expected):
        raise ValueError(f"逐格记录数量或坐标不符：{path.name}")
    for item in report["cells"]:
        source = item["source"]
        if source is None or not item["reason"]:
            raise ValueError(f"逐格记录缺少原文位置或原因：{path.name} {item['row']},{item['column']}")
        block = by_number.get(source["block_no"])
        if not block or source["doc"] != report["doc"] or source["locator"] != block["locator"]:
            raise ValueError(f"逐格记录源定位无法回查：{path.name} {item['row']},{item['column']}")


def parse_corpus(root: Path, store: Store) -> list[dict]:
    inventory = []
    for category, suffix, count in (("Excel", ".xlsx", 5), ("md", ".md", 3), ("txt", ".txt", 3), ("word", ".docx", 5)):
        paths = sorted((root / category).glob(f"*{suffix}"))
        if len(paths) != count:
            raise ValueError(f"{category} 预期 {count} 份样本，实际 {len(paths)}")
        for path in paths:
            start = perf_counter()
            data = path.read_bytes()
            parsed = parse_document(data, path.name)
            blocks_by_type = dict(Counter(block.type for block in parsed.blocks))
            expected = CORPUS_ANCHORS.get(path.name)
            if not blocks_by_type or (expected is not None and blocks_by_type != expected):
                raise ValueError(f"测试语料的块数与预期不符：{path.name} {blocks_by_type}")
            store.save(parsed, data)
            inventory.append({
                "file": path.name, "blocks": len(parsed.blocks), "seconds": round(perf_counter() - start, 3),
                "types": blocks_by_type,
            })
    return inventory


def check_dataset(root: Path, output: Path, extract: bool) -> dict:
    templates = root / "包含模板文件"
    city_source = templates / CITY_FOLDER / CITY_SOURCE
    city_template = city_source.with_name(CITY_TEMPLATE)
    daily_source = templates / DAILY_FOLDER / DAILY_SOURCE
    daily_template = daily_source.with_name(DAILY_TEMPLATE)
    store = Store(output / "a23.db")
    if extract:
        started = perf_counter()
        print(f"{MODEL} 正在抽取 01 城市原文", flush=True)
        extraction = extract_document(city_source.read_bytes(), CITY_SOURCE, CITY_METRICS, store)
        extraction["seconds"] = round(perf_counter() - started, 2)
        extraction["model"] = MODEL
        write_json(output / "city_extraction.json", extraction)
    if not store.facts(CITY_SOURCE):
        raise ValueError("还没有 01 事实，请加 --extract 从城市报告抽取")
    inventory = parse_corpus(root, store)
    daily_content = daily_source.read_bytes()
    parsed = parse_document(daily_content, DAILY_SOURCE)
    store.save(parsed, daily_content)
    context_content = daily_source.with_name(DAILY_CONTEXT).read_bytes()
    context = parse_document(context_content, DAILY_CONTEXT)
    store.save(context, context_content)
    originals = {
        path: hashlib.sha256(path.read_bytes()).hexdigest() for path in (city_template, daily_template)
    }
    expected_city = city_expected(city_source)
    expected_daily = daily_expected(daily_source)
    if len(expected_city) != 100 or len(expected_daily) != 2542:
        raise ValueError("独立原文核对得出的预期行数不符合任务书")
    if sum(value is None for row in expected_daily for value in row) != 330:
        raise ValueError("逐日原文的空格数与任务书的 330 不一致")

    results = {}
    for name, source, template, workbook, trace, expected, width in (
        ("01", CITY_SOURCE, city_template, CITY_OUTPUT, "fill_report.json", expected_city, 5),
        ("03", DAILY_SOURCE, daily_template, COVID_OUTPUT, "fill_report_covid19.json", expected_daily, 7),
    ):
        previous = None
        timings = []
        for _ in range(2):
            started = perf_counter()
            outcome = fill_workbook(store, source, template, output)
            timings.append(round(perf_counter() - started, 3))
            audit = verify_workbook(output / workbook, expected, width)
            verify_trace(store, output / trace, len(expected), width)
            current = {**outcome, **audit}
            if previous is not None and previous != current:
                raise ValueError(f"{name} 同库两次填写结果不一致")
            previous = current
        results[name] = {**current, "fill_seconds": timings, "block_count": len(store.blocks(source))}
        if hashlib.sha256(template.read_bytes()).hexdigest() != originals[template]:
            raise ValueError(f"模板原件出现变更：{template.name}")

    report = {
        "basis": "independent-source-check",
        "corpus": inventory,
        "scenarios": results,
        "daily_context": {
            "file": DAILY_CONTEXT, "blocks": len(context.blocks),
            "used_for_fill": False,
            "reason": "单日省级记述与模板需要的国家或地区逐日统计口径不同",
        },
    }
    write_json(output / "source_audit.json", report)
    if any(
        item["differences"] or item["rows"] != rows or item["columns"] != columns
        for item, rows, columns in ((results["01"], 100, 5), (results["03"], 2542, 7))
    ):
        raise ValueError(f"存在原文核对差异，详见 {output / 'source_audit.json'}")
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description="用测试集原文独立核对 01 和 03，输出本地审计记录")
    parser.add_argument("--root", type=Path, default=DATASET)
    parser.add_argument("--data-dir", type=Path, default=OUTPUT)
    parser.add_argument("--extract", action="store_true", help="调用 DeepSeek 重新抽取 01 城市报告")
    args = parser.parse_args()
    try:
        report = check_dataset(args.root, args.data_dir, args.extract)
    except (OSError, ValueError) as error:
        parser.exit(1, f"核对失败：{error}\n")
    print(json.dumps(report["scenarios"], ensure_ascii=False, indent=2))
    print(f"{len(report['corpus'])} 份语料解析完成，报告：{args.data_dir / 'source_audit.json'}")


if __name__ == "__main__":
    main()
