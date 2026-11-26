import json
import re
from collections import Counter, defaultdict
from collections.abc import Callable
from datetime import date, datetime
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from tempfile import NamedTemporaryFile
from time import perf_counter
from uuid import uuid4

from openpyxl import load_workbook

from criteria import checked_rules, load_expectations
from extractor import CITY_TITLE, evidence_error, period_for, periods
from fill_spec import (
    CITY_COLUMNS, CITY_DOCUMENTS, CITY_HEADERS, CITY_OUTPUT, COUNTRY_KEYS, COVID_END, COVID_OUTPUT,
    COVID_START, DAILY_METRIC_KEYS, DATE_KEYS, CityColumn, choose_value, metric_key,
    normalize, number, spreadsheet_key,
)
from parser_xlsx import workbook_rows
from schema import Block, RawFact
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


def cell(value: object, reason: str, source: dict | None = None, reason_code: str | None = None) -> dict:
    return {
        "value": value, "decision": "filled" if value is not None else "refused",
        "reason": reason, "source": source, **({"reason_code": reason_code} if reason_code else {}),
    }


def resolve_facts(
    facts: list[dict], entity: str, indicator: str, period: str, scope: str,
    preferred_doc: str | None = None,
) -> dict:
    candidates = [
        fact for fact in facts
        if fact["entity"] == entity and metric_key(fact["metric"]) == metric_key(indicator)
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
    if value is not None:
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


def city_fact_evidence(fact: dict) -> dict:
    return {key: fact[key] for key in ("doc", "block_no", "metric", "value", "unit", "snippet",
                                      "period", "period_block_no", "mode", "scope")}


def annual_scope(scope: str, period: str) -> str:
    label = normalize(scope)
    if label in {"", "全年", "年度", period, f"{period}年", f"{period}年度", f"{period}年全年"}:
        return ""
    return label


def city_rows(blocks: list[dict], facts: list[dict], target_period: str | None = None) -> list[list[dict]]:
    by_block = defaultdict(list)
    parsed = [(block["number"], Block.model_validate(block)) for block in blocks]
    context = periods(parsed)
    parsed_by_number = dict(parsed)
    title_year = next((match[1] for _, block in parsed[:1] if (match := CITY_TITLE.match(block.text))), None)
    if target_period and title_year and target_period != title_year:
        raise ValueError(f"01 目标期次 {target_period} 与原文标题 {title_year} 不一致")
    target_period = target_period or title_year
    for fact in facts:
        index = fact["block_no"]
        raw = RawFact.model_validate(fact)
        if evidence_error(raw, parsed_by_number):
            continue
        period, anchor = period_for(raw, parsed_by_number[index], context[index])
        by_block[index].append({**fact, "period": period or fact["period"],
                                "period_block_no": anchor if period else fact["period_block_no"],
                                "metric_key": metric_key(fact["metric"])})
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
        city_facts = [fact for fact in by_block[index] if fact["metric_key"] in target_keys
                      and fact["entity"].removesuffix("市") == city.removesuffix("市")]
        known_periods = {fact["period"] for fact in city_facts if fact["period"]}
        period_ambiguous = target_period is None and len(known_periods) > 1
        selected_period = target_period or (next(iter(known_periods)) if len(known_periods) == 1 else "")
        for column in CITY_COLUMNS:
            matching = [fact for fact in city_facts if fact["metric_key"] == column.key]
            comparable = [
                fact for fact in matching
                if not period_ambiguous and (fact["period"] == selected_period or (not target_period and not fact["period"]))
                and not annual_scope(fact["scope"], selected_period)
            ]
            candidates = [
                {"value": value, "fact": fact}
                for fact in comparable if (value := city_value(fact, column)) is not None
            ]
            chosen, reason = choose_value(candidates)
            if chosen:
                values.append({**cell(chosen["value"], reason, source), "fact": city_fact_evidence(chosen["fact"])})
            elif candidates:
                item = cell(None, reason, source, "conflict")
                item["conflicts"] = [
                    {"value": candidate["value"], "fact": city_fact_evidence(candidate["fact"]), "source": source}
                    for candidate in candidates
                ]
                values.append(item)
            elif comparable:
                values.append(cell(None, "抽到同城同指标事实，但数值或单位与模板列不符", source,
                                   "incompatible_value"))
            elif matching:
                values.append({**cell(None, "缺少可确认的目标期次或全年口径事实", source, "period_or_scope"),
                               "conflicts": [{"value": fact["value"], "fact": city_fact_evidence(fact),
                                              "source": source} for fact in matching]})
            else:
                values.append(cell(None, "没有抽到来源段中的同城同指标事实", source, "missing_fact"))
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
                candidates_have_value = any(
                    raw is not None and str(raw).strip() for record, _ in records for _, raw in record[key]
                )
                values.append(cell(
                    None, "匹配的来源行在该字段没有可用值", records[0][1],
                    "incompatible_value" if candidates_have_value else "source_blank",
                ))
                continue
            chosen, reason = choose_value(options)
            item = cell(chosen["value"] if chosen else None, reason, chosen["source"] if chosen else None)
            if chosen is None:
                item["reason_code"] = "conflict"
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
    *, target_period: str | None = None, formal_city: bool = False,
) -> dict:
    started = perf_counter()
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
            rows = city_rows(blocks, facts, target_period)
            if formal_city or doc in CITY_DOCUMENTS:
                expected_rows = checked_rules(load_expectations(), "same-db")["01"]["rows"]
                names = {row[0]["value"].removesuffix("市") for row in rows}
                if len(rows) != expected_rows or len(names) != expected_rows:
                    raise ValueError(f"01 正式输出应覆盖 {expected_rows} 个不同城市，实际 {len(rows)} 行、{len(names)} 个城市")
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
            "refusal_counts": dict(Counter(item["reason_code"] for item in cells if "reason_code" in item)),
        }
        result_id = json_digest(trace)
        output = output_dir / filename
        # Save each workbook and its byte hash from the same local buffer.
        written_workbook = BytesIO()
        workbook.save(written_workbook)
        content = written_workbook.getvalue()
        previous = output.read_bytes() if output.is_file() else None
        atomic_file(output, lambda target: target.write_bytes(content))
        seconds = round(perf_counter() - started, 4)
        try:
            write_json(output_dir / report_name, {
                **trace, "result_id": result_id, "run_id": uuid4().hex,
                "workbook_sha256": sha256(content).hexdigest(), "fill_seconds": seconds,
            })
        except (OSError, ValueError, TypeError):
            if previous is None:
                output.unlink(missing_ok=True)
            else:
                atomic_file(output, lambda target: target.write_bytes(previous))
            raise
    finally:
        workbook.close()
    return {**result, "result_id": result_id, "fill_seconds": seconds}
