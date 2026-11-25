from datetime import date, datetime
from pathlib import Path

from docx import Document
from openpyxl import Workbook

from config import DATA
from evaluate import evaluate
from extractor import extract_document
from fill_spec import CITY_HEADERS, CITY_OUTPUT, COVID_OUTPUT
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


def build_demo(root: Path) -> tuple[Path, Path, dict, dict]:
    inputs, outputs = root / "doc", root / "data"
    inputs.mkdir(parents=True, exist_ok=True)
    outputs.mkdir(parents=True, exist_ok=True)
    store = Store(outputs / "a23.db")

    document = Document()
    document.add_paragraph(CITY_SENTENCE)
    city_path = inputs / CITY_DOC
    document.save(city_path)
    extract_document(city_path.read_bytes(), CITY_DOC, [], store, demo_response)
    make_workbook(inputs / CITY_TEMPLATE, CITY_HEADERS)
    city_reference = inputs / "参考_合成城市.xlsx"
    make_workbook(city_reference, CITY_HEADERS, [("甲城", 100, 20, 50000, 10)])
    fill_workbook(store, CITY_DOC, inputs / CITY_TEMPLATE, outputs)
    city = evaluate(
        outputs / CITY_OUTPUT, city_reference, outputs / "eval_report.json", CITY_DOC, False,
        block_count=len(store.blocks(CITY_DOC)),
    )

    daily_path = inputs / DAILY_DOC
    make_workbook(daily_path, SOURCE_HEADERS, (*DAILY_ROWS, (date(2020, 6, 30), "Alpha", 99, 9, 99, 9, 100)))
    source = daily_path.read_bytes()
    store.save(parse_document(source, DAILY_DOC), source)
    make_workbook(inputs / DAILY_TEMPLATE, SOURCE_HEADERS)
    daily_reference = inputs / "参考_合成逐日.xlsx"
    make_workbook(daily_reference, SOURCE_HEADERS, sorted(DAILY_ROWS))
    fill_workbook(store, DAILY_DOC, inputs / DAILY_TEMPLATE, outputs)
    daily = evaluate(
        outputs / COVID_OUTPUT, daily_reference, outputs / "eval_report_covid19.json", DAILY_DOC, True,
        block_count=len(store.blocks(DAILY_DOC)),
    )
    return inputs, outputs, city, daily


if __name__ == "__main__":
    inputs, outputs, city, daily = build_demo(DATA / "demo")
    print(f"合成样例：{inputs}")
    print(f"生成结果：{outputs}")
    print(f"01 准确率 {city['cell_accuracy']}，03 对格率 {daily['correct_rate']}")
