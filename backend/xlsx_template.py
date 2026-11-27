"""Confirmed, auditable XLSX target mapping for the public 03 data shape."""

import json
from collections import OrderedDict, defaultdict
from decimal import Decimal
from hashlib import sha256
from io import BytesIO
from threading import Lock
from time import monotonic
from uuid import uuid4

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from fill_spec import COUNTRY_KEYS, DATE_KEYS, normalize, spreadsheet_key
from filler import covid_rows
from operations import MAX_FILE_BYTES, MAX_PLANS, PLAN_LIFETIME, prune_expired, store_output
from parser_xlsx import compatible_workbook, workbook_rows
from parsers import parse_document

FIELDS = ("date", "country", "continent", "gdp_per_capita", "population", "每日检测数", "病例数")
LABELS = {"date": "日期", "country": "国家/地区", "continent": "大洲",
          "gdp_per_capita": "人均GDP", "population": "人口", "每日检测数": "每日检测数", "病例数": "病例数"}
plans: OrderedDict[str, dict] = OrderedDict()
lock = Lock()


def field_for(header: object) -> str | None:
    name = str(header or "").strip()
    normalized = normalize(name)
    if normalized in DATE_KEYS:
        return "date"
    if normalized in COUNTRY_KEYS:
        return "country"
    key = spreadsheet_key(name)
    return key if key in FIELDS else None


def target_sheets(content: bytes) -> list[dict]:
    workbook = load_workbook(compatible_workbook(content), read_only=True, data_only=True)
    try:
        result = []
        for sheet in workbook:
            sheet.reset_dimensions()
            headers = next(sheet.iter_rows(values_only=True), ())
            result.append({"name": sheet.title, "headers": [
                {"column": get_column_letter(column), "label": str(value or "").strip(),
                 "suggested": field_for(value)}
                for column, value in enumerate(headers, 1)
            ]})
        return result
    finally:
        workbook.close()


def source_headers(content: bytes) -> tuple[list[dict], dict[str, dict[str, tuple[int, str]]]]:
    workbook = load_workbook(compatible_workbook(content), read_only=True, data_only=True)
    result, positions = [], {}
    try:
        for sheet in workbook:
            sheet.reset_dimensions()
            headers = next(sheet.iter_rows(values_only=True), ())
            fields = defaultdict(list)
            for column, value in enumerate(headers, 1):
                if key := field_for(value):
                    fields[key].append((column, str(value).strip()))
            for key in FIELDS:
                if len(fields[key]) > 1:
                    columns = "、".join(get_column_letter(column) for column, _ in fields[key])
                    raise ValueError(f"源工作表「{sheet.title}」字段「{LABELS[key]}」重复，冲突列：{columns}")
            positions[sheet.title] = {key: found[0] for key, found in fields.items() if found}
            result.append({"name": sheet.title, "headers": [str(value or "").strip() for value in headers]})
    finally:
        workbook.close()
    return result, positions


def inspect(source: bytes, source_name: str, template: bytes, template_name: str) -> dict:
    if not source or not template or max(len(source), len(template)) > MAX_FILE_BYTES:
        raise ValueError("源文件或模板为空，或超过 12 MB")
    if not source_name.lower().endswith(".xlsx") or not template_name.lower().endswith(".xlsx"):
        raise ValueError("此链路需要 XLSX 源文件和 XLSX 目标模板")
    targets = target_sheets(template)
    sources, _ = source_headers(source)
    if not targets or not sources:
        raise ValueError("源文件或目标模板缺少工作表")
    plan_id = uuid4().hex
    with lock:
        prune_expired(plans, monotonic())
        while len(plans) >= MAX_PLANS:
            plans.popitem(last=False)
        plans[plan_id] = {"source_sha256": sha256(source).hexdigest(),
                          "template_sha256": sha256(template).hexdigest(), "created": monotonic()}
    return {"plan_id": plan_id, "source_sheets": sources, "target_sheets": targets,
            "fields": [{"key": key, "label": LABELS[key]} for key in FIELDS]}


def validate_mapping(headers: list[dict], mapping: dict) -> list[str]:
    if not isinstance(mapping, dict) or not 6 <= len(headers) <= 7:
        raise ValueError("目标模板须有 6 或 7 个字段，请选择正确工作表")
    expected_columns = [item["column"] for item in headers]
    if set(mapping) != set(expected_columns) or any(value not in FIELDS for value in mapping.values()):
        raise ValueError("字段映射须覆盖目标表头的每一列，且只能选择 03 支持的字段")
    keys = [mapping[column] for column in expected_columns]
    if len(set(keys)) != len(keys):
        raise ValueError("同一来源字段不能映射到多个目标列")
    required = set(FIELDS)
    if len(keys) == 6 and "date" not in keys:
        required.remove("date")
    if set(keys) != required:
        raise ValueError("目标字段缺少日期、国家/地区或 03 指标；请核对映射")
    return keys


def fill(source: bytes, source_name: str, template: bytes, template_name: str,
         plan_id: str, sheet_name: str, mapping_json: str) -> dict:
    with lock:
        prune_expired(plans, monotonic())
        plan = plans.get(plan_id)
        if plan is None:
            raise ValueError("XLSX 模板计划已失效，请重新预览")
        if sha256(source).hexdigest() != plan["source_sha256"] or sha256(template).hexdigest() != plan["template_sha256"]:
            raise ValueError("源文件或目标模板版本与已确认计划不符，请重新预览")
        plans.pop(plan_id, None)
    try:
        return _fill_claimed(source, source_name, template, template_name, sheet_name, mapping_json, plan)
    except Exception:
        with lock:
            if monotonic() - plan["created"] <= PLAN_LIFETIME:
                prune_expired(plans, monotonic())
                while len(plans) >= MAX_PLANS:
                    plans.popitem(last=False)
                plans[plan_id] = plan
        raise


def _fill_claimed(source: bytes, source_name: str, template: bytes, template_name: str,
                  sheet_name: str, mapping_json: str, plan: dict) -> dict:
    try:
        mapping = json.loads(mapping_json)
    except (TypeError, ValueError) as error:
        raise ValueError("字段映射 JSON 无效，请重新选择目标列") from error
    target_previews = target_sheets(template)
    chosen = next((item for item in target_previews if item["name"] == sheet_name), None)
    if chosen is None:
        raise ValueError(f"目标工作表「{sheet_name}」不存在，请重新选择")
    keys = validate_mapping(chosen["headers"], mapping)
    _, source_columns = source_headers(source)
    parsed = parse_document(source, source_name)
    blocks = [
        {"doc": source_name, "number": number, "type": block.type,
         "text": block.text, "locator": block.locator}
        for number, block in parsed.numbered_blocks()
    ]
    output_keys = tuple(keys)
    add_date = "date" not in keys
    if add_date:
        output_keys = ("date", *output_keys)
    mapped_headers = tuple(LABELS[key] for key in output_keys)
    rows = covid_rows(source, blocks, mapped_headers)
    needed_rows = {(item["source"]["locator"]["sheet"], item["source"]["locator"]["row"])
                   for row in rows for item in row if item.get("source")}
    raw_rows = {(sheet, row): values for sheet, row, values in workbook_rows(source)
                if (sheet, row) in needed_rows}
    book = load_workbook(compatible_workbook(template))
    try:
        sheet = book[sheet_name]
        if add_date:
            sheet.insert_cols(1)
            sheet.cell(1, 1, "日期")
        cells = []
        for row_number, row in enumerate(rows, 2):
            for column_number, item in enumerate(row, 1):
                value = item["value"]
                if isinstance(value, Decimal):
                    value = int(value) if value == value.to_integral_value() else float(value)
                target = sheet.cell(row_number, column_number)
                target.value = value
                if isinstance(value, str) and target.data_type == "f":
                    target.data_type = "s"
                key = output_keys[column_number - 1]
                source_info = item.get("source")
                provenance = {}
                if source_info:
                    locator = source_info["locator"]
                    source_sheet, source_row = locator["sheet"], locator["row"]
                    coordinate = source_columns[source_sheet].get(key)
                    if coordinate:
                        source_column, source_header = coordinate
                        provenance = {"source_sheet": source_sheet, "source_row": source_row,
                                      "source_column": get_column_letter(source_column),
                                      "source_value": raw_rows[(source_sheet, source_row)].get(source_header)}
                cells.append({"row": row_number, "column": column_number, "field": key,
                              "value": value, "reason": item["reason"],
                              "reason_code": item.get("reason_code"), **provenance})
        if sheet.max_row > len(rows) + 1:
            sheet.delete_rows(len(rows) + 2, sheet.max_row - len(rows) - 1)
        output = BytesIO()
        book.save(output)
    finally:
        book.close()
    result_id = store_output(f"填写结果_{template_name}", output.getvalue())
    return {"result_id": result_id, "sheet": sheet_name, "rows": len(rows),
            "cells_written": sum(cell["value"] is not None for cell in cells),
            "cells_empty": sum(cell["value"] is None for cell in cells),
            "cells": cells, "source_sha256": plan["source_sha256"],
            "template_sha256": plan["template_sha256"]}
