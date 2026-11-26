import json
import re
from collections import defaultdict
from collections.abc import Callable
from datetime import date, datetime
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from tempfile import NamedTemporaryFile
from uuid import uuid4

from openpyxl import load_workbook

from fill_spec import (
    CITY_COLUMNS, CITY_HEADERS, CITY_OUTPUT, COUNTRY_KEYS, COVID_END, COVID_OUTPUT,
    COVID_START, DAILY_METRIC_KEYS, DATE_KEYS, CityColumn, choose_value, metric_key,
    normalize, number, spreadsheet_key,
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


def write_json(path: Path, result: object) -> None:
    atomic_file(path, lambda target: target.write_text(
        json.dumps(result, ensure_ascii=False, indent=2, default=str) + "\n", encoding="utf-8",
    ))


def json_digest(value: object) -> str:
    payload = json.dumps(value, ensure_ascii=False, sort_keys=True, default=str)
    return sha256(payload.encode("utf-8")).hexdigest()


def block_digest(block: dict) -> str:
    return json_digest((block["number"], block["type"], block["text"], block["locator"]))


def source_for(block: dict) -> dict:
    return {
        "doc": block["doc"], "block_no": block["number"], "locator": block["locator"],
        "block_sha256": block_digest(block),
    }


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


def city_value(fact: dict, column: CityColumn):
    units = {normalize(unit) for unit in column.units}
    raw = re.match(r"\s*([+-]?\d[\d,，]*(?:\.\d+)?)\s*(.*)", fact["value"])
    if not raw:
        return None
    unit = normalize(fact["unit"])
    suffix = normalize(raw[2])
    if (unit and unit not in units) or (suffix and suffix not in units):
        return None
    value = number(raw[1])
    if not unit and not suffix and value is not None:
        digits = re.sub(r"[\s,，]", "", raw[1])
        quote = re.sub(r"[\s,，]", "", fact["snippet"])
        observed = re.search(
            rf"(?<![\d.]){re.escape(digits)}(万亿元|亿元|万元|万人口|万人|万|元[/／]人|元|%)",
            quote,
        )
        if observed and normalize(observed[1]) not in units:
            return None
    parsed = number(fact["value_num"])
    return parsed if parsed is not None and parsed == value else value


def city_rows(blocks: list[dict], facts: list[dict]) -> list[list[dict]]:
    by_block = defaultdict(list)
    for fact in facts:
        by_block[fact["block_no"]].append(fact)
    target_keys = {column.key for column in CITY_COLUMNS}
    rows = []
    for block in blocks:
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
        if not match or re.search(r"[\u4e00-\u9fff]{2,}[与及、和][\u4e00-\u9fff]{2,}", match[1]):
            continue
        city = match[1]
        source = source_for(block)
        values = [cell(city, "城市正文段首", source)]
        for column in CITY_COLUMNS:
            candidates = [
                {"value": value, "fact": fact}
                for fact in by_block[index]
                if fact["metric_key"] == column.key
                and fact["entity"].removesuffix("市") == city.removesuffix("市")
                and (value := city_value(fact, column)) is not None
            ]
            chosen, reason = choose_value(candidates)
            values.append(cell(chosen["value"] if chosen else None, reason, source if chosen else None))
        rows.append(values)
    if not rows:
        raise ValueError("库里没有可用的城市正文段落")
    return rows


def covid_rows(content: bytes, blocks: list[dict], headers: tuple[str, ...]) -> list[list[dict]]:
    candidates = defaultdict(list)
    sources = {
        (block["locator"]["sheet"], block["locator"]["row"]): source_for(block)
        for block in blocks
    }
    source_keys = set()
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
        source = sources.get((sheet, row))
        if source is None:
            raise ValueError("逐日数据的入库位置缺失，请重新导入文档")
        record = defaultdict(list)
        for name, value in values.items():
            record[spreadsheet_key(name)].append((name, value))
        source_keys.update(record)
        candidates[day, country].append((record, source))
    if not candidates:
        raise ValueError("逐日数据没有与日期区间、日期列和国家列匹配的记录")
    for header in headers:
        normalized = normalize(header)
        if normalized in DATE_KEYS or normalized in COUNTRY_KEYS:
            continue
        key = spreadsheet_key(header)
        if key not in DAILY_METRIC_KEYS or key not in source_keys:
            raise ValueError(f"03 模板列无法映射到逐日数据：{header}")

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
            key = spreadsheet_key(header)
            options = [
                {"value": value, "source": {**source, "field": field}}
                for record, source in records
                for field, raw in record[key]
                if (value := covid_value(key, raw)) is not None
            ]
            if not options:
                values.append(cell(None, "匹配的来源行在该字段没有可用值", records[0][1]))
                continue
            chosen, reason = choose_value(options)
            item = cell(chosen["value"] if chosen else None, reason, chosen["source"] if chosen else None)
            if chosen is None:
                item["conflicts"] = options
            values.append(item)
        result.append(values)
    return result


def covid_value(key: str, value: object) -> object:
    if key == "continent":
        return (value.strip() or None) if isinstance(value, str) else None
    return number(value)


def fill_workbook(
    store: Store, doc: str, template: Path, output_dir: Path, source_sha256: str | None = None,
) -> dict:
    city = doc.lower().endswith(".docx")
    snapshot = store.snapshot(doc, with_facts=city)
    if snapshot is None:
        raise ValueError("库里没有这份文档")
    content, blocks, facts = snapshot
    document_digest = sha256(content).hexdigest()
    if source_sha256 is not None and source_sha256 != document_digest:
        raise ValueError("上传文件与库中同名文档内容不同，请先重新导入或抽取")
    template_content = template.read_bytes()
    workbook = load_workbook(BytesIO(template_content))
    try:
        sheet = workbook.active
        headers = tuple(str(value or "").strip() for value in next(sheet.values))
        if city:
            if headers != CITY_HEADERS:
                raise ValueError("01 模板表头与字段映射不一致")
            if not facts:
                raise ValueError("城市报告缺少可用事实，请先重新抽取事实")
            rows = city_rows(blocks, facts)
            filename, report_name = CITY_OUTPUT, "fill_report.json"
        elif doc.lower().endswith(".xlsx"):
            if len(headers) == 6 and not any(normalize(value) in DATE_KEYS for value in headers):
                sheet.insert_cols(1)
                sheet.cell(1, 1, "日期")
                headers = ("日期", *headers)
            if len(headers) != 7 or sum(normalize(value) in DATE_KEYS for value in headers) != 1:
                raise ValueError("03 模板须包含日期列及六个目标字段")
            if sum(normalize(value) in COUNTRY_KEYS for value in headers) != 1:
                raise ValueError("03 模板缺少国家或地区列")
            if len({spreadsheet_key(header) for header in headers}) != 7:
                raise ValueError("03 模板列重复，请核对表头")
            rows = covid_rows(content, blocks, headers)
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
        trace = {
            "doc": doc, "file": filename, **result, "cells": cells,
            "source_sha256": document_digest, "template_sha256": sha256(template_content).hexdigest(),
            "block_count": len(blocks),
        }
        result_id = json_digest(trace)
        output = output_dir / filename
        # Save each workbook and its byte hash from the same local buffer.
        written_workbook = BytesIO()
        workbook.save(written_workbook)
        content = written_workbook.getvalue()
        atomic_file(output, lambda target: target.write_bytes(content))
        write_json(output_dir / report_name, {
            **trace, "result_id": result_id, "run_id": uuid4().hex,
            "workbook_sha256": sha256(content).hexdigest(),
        })
    finally:
        workbook.close()
    return {**result, "result_id": result_id}
