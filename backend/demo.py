from datetime import date, datetime
from pathlib import Path

from docx import Document
from openpyxl import Workbook

from config import DATA
from evaluate import evaluate
from extractor import extract_document
from fill_spec import CITY_HEADERS, CITY_OUTPUT, COVID_OUTPUT, REFERENCE_FILES
from filler import fill_workbook
from parsers import parse_document
from store import Store

CITY_DOC = "合成城市报告.docx"
CITY_TEMPLATE = "模板_合成城市.xlsx"
DAILY_DOC = "合成逐日统计.xlsx"
DAILY_TEMPLATE = "模板_合成逐日.xlsx"
SOURCE_HEADERS = ("date", "location", "new_cases", "new_deaths", "total_cases", "total_deaths", "population")
DAILY_ROWS = (
    (datetime(2020, 7, 2), "Beta", None, 0, 5, 0, 200),
    (datetime(2020, 7, 1), "Beta", 0, None, 5, 0, 200),
    (datetime(2020, 7, 1), "Alpha", 2, 0, 3, 0, 100),
)
CITY_SENTENCE = "甲城：2025年GDP总量100亿元，常住人口20万人，人均GDP50000元，一般公共预算收入10亿元。"


def make_workbook(path: Path, headers: tuple[str, ...], rows=()) -> None:
    workbook = Workbook()
    workbook.active.append(headers)
    for row in rows:
        workbook.active.append(row)
    workbook.save(path)
    workbook.close()


def demo_response(_prompt: str) -> dict:
    return {"facts": [
        {"block_no": 1, "entity": "甲城", "metric": metric, "value": value, "unit": unit,
         "snippet": f"{metric}{value}{unit}"}
        for metric, value, unit in (
            ("GDP总量", "100", "亿元"), ("常住人口", "20", "万人"),
            ("人均GDP", "50000", "元"), ("一般公共预算收入", "10", "亿元"),
        )
    ]}


def prepare_demo_inputs(inputs: Path) -> None:
    inputs.mkdir(parents=True, exist_ok=True)
    document = Document()
    document.add_paragraph(CITY_SENTENCE)
    document.save(inputs / CITY_DOC)
    make_workbook(inputs / CITY_TEMPLATE, CITY_HEADERS)
    make_workbook(inputs / REFERENCE_FILES[CITY_OUTPUT], CITY_HEADERS, [("甲城", 100, 20, 50000, 10)])
    make_workbook(inputs / DAILY_DOC, SOURCE_HEADERS, (*DAILY_ROWS, (date(2020, 6, 30), "Alpha", 99, 9, 99, 9, 100)))
    make_workbook(inputs / DAILY_TEMPLATE, SOURCE_HEADERS)
    make_workbook(inputs / REFERENCE_FILES[COVID_OUTPUT], SOURCE_HEADERS, sorted(DAILY_ROWS))


def fill_demo(inputs: Path, outputs: Path) -> tuple[dict, dict]:
    store = Store(outputs / "a23.db")
    reports = []
    for doc, template, file, report in (
        (CITY_DOC, CITY_TEMPLATE, CITY_OUTPUT, "eval_report.json"),
        (DAILY_DOC, DAILY_TEMPLATE, COVID_OUTPUT, "eval_report_covid19.json"),
    ):
        fill_workbook(store, doc, inputs / template, outputs)
        reports.append(evaluate(
            outputs / file, inputs / REFERENCE_FILES[file], outputs / report, doc, file == COVID_OUTPUT,
            block_count=len(store.blocks(doc)), database=store.path,
        ))
    return reports[0], reports[1]


def build_demo(root: Path) -> tuple[Path, Path, dict, dict]:
    inputs, outputs = root / "doc", root / "data"
    prepare_demo_inputs(inputs)
    outputs.mkdir(parents=True, exist_ok=True)
    store = Store(outputs / "a23.db")
    extract_document((inputs / CITY_DOC).read_bytes(), CITY_DOC, [], store, demo_response)
    source = (inputs / DAILY_DOC).read_bytes()
    store.save(parse_document(source, DAILY_DOC), source)
    city, daily = fill_demo(inputs, outputs)
    return inputs, outputs, city, daily


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="准备无模型调用的合成演示")
    parser.add_argument("--inputs-only", action="store_true", help="只写原文、模板和参考表，供 03 空库演示")
    args = parser.parse_args()
    root = DATA / "demo"
    if args.inputs_only:
        prepare_demo_inputs(root / "doc")
        print(f"合成样例：{root / 'doc'}")
    else:
        inputs, outputs, city, daily = build_demo(root)
        print(f"合成样例：{inputs}")
        print(f"生成结果：{outputs}")
        print(f"01 准确率 {city['cell_accuracy']}，03 对格率 {daily['correct_rate']}")
