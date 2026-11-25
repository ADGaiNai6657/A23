import json
import re
from collections import defaultdict
from collections.abc import Callable
from datetime import date, datetime
from pathlib import Path
from tempfile import NamedTemporaryFile

from openpyxl import load_workbook

from fill_spec import (
    CITY_COLUMNS, CITY_HEADERS, CITY_OUTPUT, COUNTRY_KEYS, COVID_END, COVID_OUTPUT,
    COVID_START, DATE_KEYS, choose_value, metric_key, normalize, number,
)
from parser_xlsx import workbook_rows
from store import Store


def atomic_file(path: Path, write: Callable[[Path], None]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(dir=path.parent, suffix=path.suffix, delete=False) as temp:
        temporary = Path(temp.name)
    try:
        write(temporary)
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def write_json(path: Path, result: dict) -> None:
    atomic_file(path, lambda target: target.write_text(
        json.dumps(result, ensure_ascii=False, indent=2, default=str) + "\n", encoding="utf-8",
    ))


def as_date(value: object) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    try:
        return date.fromisoformat(str(value).strip())
    except ValueError:
        return None


def cell(value: object, reason: str, source: dict | None = None) -> dict:
    return {
        "value": value, "decision": "filled" if value is not None else "refused",
        "reason": reason, "source": source,
    }


def resolve_facts(
    facts: list[dict], entity: str, indicator: str, period: str, scope: str,
    preferred_doc: str | None = None,
) -> dict:
    candidates = [
        fact for fact in facts
        if fact["entity"] == entity and fact["metric_key"] == metric_key(indicator)
        and fact["period"] == period and fact["scope"] == scope
    ]
    comparable = [
        {"value": (number(fact["value_num"]), normalize(fact["unit"])), "fact": fact}
        for fact in candidates
        if fact["value_num"] is not None
    ]
    if preferred_doc:
        comparable = [item for item in comparable if item["fact"]["doc"] == preferred_doc]
    chosen, reason = choose_value(comparable)
    if chosen and preferred_doc:
        reason = f"采用指定来源文档 {preferred_doc}"
    return {"candidates": candidates, "selected": chosen["fact"] if chosen else None, "reason": reason}


def city_rows(store: Store, doc: str) -> list[list[dict]]:
    by_block = defaultdict(list)
    for fact in store.facts(doc):
        by_block[fact["block_no"]].append(fact)
    target_keys = {column.key for column in CITY_COLUMNS}
    rows = []
    for block in store.blocks(doc):
        if block["type"] != "paragraph":
            continue
        index = block["number"]
        aliases_present = {
            metric_key(match.group()) for match in re.finditer(
                r"人均\s*(?:GDP|地区生产总值|生产总值)|GDP\s*总量|地区生产总值|生产总值|GDP|"
                r"常住人口|人口基数|人口严控至|万人口|人口|一般公共预算收入|公共预算收入",
                block["text"], re.IGNORECASE,
            )
        }
        if len(aliases_present & target_keys) < 3:
            continue
        # City paragraphs start with a single name directly followed by prose or punctuation.
        match = re.match(
            r"^(?:\s*\d+[.、．]\s*)?([\u4e00-\u9fff]{2,12}?)(?=以|凭借|作为|在|紧随|的|，|,|：|:|（|\(|\s|GDP|地区生产总值|生产总值)",
            block["text"],
        )
        if not match or any(char in match[1] for char in "与和及"):
            continue
        city = match[1]
        source = {"doc": doc, "block_no": index, "locator": block["locator"]}
        values = [cell(city, "城市正文段首", source)]
        for column in CITY_COLUMNS:
            candidates = [
                {"value": number(fact["value_num"]), "fact": fact}
                for fact in by_block[index]
                if fact["metric_key"] == column.key and fact["unit"] in column.units
                and fact["entity"].removesuffix("市") == city.removesuffix("市")
                and fact["value_num"] is not None
            ]
            chosen, reason = choose_value(candidates)
            values.append(cell(chosen["value"] if chosen else None, reason, source if chosen else None))
        rows.append(values)
    if not rows:
        raise ValueError("库里没有可用的城市正文段落")
    return rows


def covid_rows(store: Store, doc: str, headers: tuple[str, ...]) -> list[list[dict]]:
    content = store.content(doc)
    if content is None or not doc.lower().endswith(".xlsx"):
        raise ValueError("03 模板需要已导入的逐日 XLSX 数据")
    candidates = defaultdict(list)
    block_numbers = {
        (block["locator"]["sheet"], block["locator"]["row"]): block["number"]
        for block in store.blocks(doc)
    }
    for sheet, row, values in workbook_rows(content):
        keys = {normalize(name): name for name in values}
        date_key = next((keys[key] for key in DATE_KEYS if key in keys), None)
        country_key = next((keys[key] for key in COUNTRY_KEYS if key in keys), None)
        if not date_key or not country_key:
            continue
        day = as_date(values[date_key])
        country = str(values[country_key] or "").strip()
        if day is None or not COVID_START <= day <= COVID_END or not country:
            continue
        source = {"doc": doc, "block_no": block_numbers.get((sheet, row)), "locator": {"sheet": sheet, "row": row}}
        record = {metric_key(name): value for name, value in values.items()}
        candidates[day, country].append((record, source))
    if not candidates:
        raise ValueError("逐日数据没有与日期区间、日期列和国家列匹配的记录")

    result = []
    for (day, country), records in sorted(candidates.items()):
        values = []
        for header in headers:
            normalized = normalize(header)
            if normalized in DATE_KEYS:
                values.append(cell(day, "日期列", records[0][1]))
                continue
            if normalized in COUNTRY_KEYS:
                values.append(cell(country, "国家或地区列", records[0][1]))
                continue
            key = metric_key(header)
            options = [
                {"value": value, "source": source}
                for record, source in records
                if (value := number(record.get(key))) is not None
            ]
            chosen, reason = choose_value(options)
            values.append(cell(chosen["value"] if chosen else None, reason, chosen["source"] if chosen else None))
        result.append(values)
    return result


def fill_workbook(store: Store, doc: str, template: Path, output_dir: Path) -> dict:
    if store.content(doc) is None:
        raise ValueError("库里没有这份文档")
    workbook = load_workbook(template)
    try:
        sheet = workbook.active
        headers = tuple(str(value or "").strip() for value in next(sheet.values))
        if doc.lower().endswith(".docx"):
            if headers != CITY_HEADERS:
                raise ValueError("01 模板表头与字段映射不一致")
            rows = city_rows(store, doc)
            filename, report_name = CITY_OUTPUT, "fill_report.json"
        elif doc.lower().endswith(".xlsx"):
            if len(headers) == 6 and not any(normalize(value) in DATE_KEYS for value in headers):
                sheet.insert_cols(1)
                sheet.cell(1, 1, "日期")
                headers = ("日期", *headers)
            if len(headers) != 7 or not any(normalize(value) in DATE_KEYS for value in headers):
                raise ValueError("03 模板须包含日期列及六个目标字段")
            if not any(normalize(value) in COUNTRY_KEYS for value in headers):
                raise ValueError("03 模板缺少国家或地区列")
            rows = covid_rows(store, doc, headers)
            filename, report_name = COVID_OUTPUT, "fill_report_covid19.json"
        else:
            raise ValueError("填写仅支持 01 DOCX 或 03 XLSX 数据源")

        cells = []
        written = 0
        for row_no, row in enumerate(rows, 2):
            for col_no, item in enumerate(row, 1):
                target = sheet.cell(row_no, col_no)
                target.value = item["value"]
                if isinstance(target.value, str) and target.data_type == "f":
                    target.data_type = "s"
                written += item["value"] is not None
                cells.append({"row": row_no, "column": col_no, **item})

        if sheet.max_row > len(rows) + 1:
            sheet.delete_rows(len(rows) + 2, sheet.max_row - len(rows) - 1)
        result = {"rows": len(rows), "cells_written": written, "cells_empty": len(cells) - written}
        atomic_file(output_dir / filename, workbook.save)
        write_json(output_dir / report_name, {"doc": doc, "file": filename, **result, "cells": cells})
    finally:
        workbook.close()
    return result
