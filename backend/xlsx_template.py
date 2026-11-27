"""Confirmed, auditable XLSX target mapping for the public 03 data shape."""

import json
from copy import copy
from collections import OrderedDict, defaultdict
from datetime import date, datetime
from decimal import Decimal
from hashlib import sha256
from io import BytesIO
from math import isfinite
from threading import Lock
from time import monotonic
from uuid import uuid4

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.cell_range import CellRange

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
            headers = list(next(sheet.iter_rows(values_only=True), ()))
            while headers and (headers[-1] is None or str(headers[-1]).strip() == ""):
                headers.pop()
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


def source_data_rows(content: bytes, name: str) -> list[list[dict]]:
    parsed = parse_document(content, name)
    blocks = [
        {"doc": name, "number": number, "type": block.type,
         "text": block.text, "locator": block.locator}
        for number, block in parsed.numbered_blocks()
    ]
    return covid_rows(content, blocks, tuple(LABELS[key] for key in FIELDS))


def validate_write_area(sheet, rows_count: int, header_count: int) -> str:
    """Only blank, unmerged data cells may be written; preserve all other cells."""
    width = len(FIELDS)
    last_row = rows_count + 1
    if sheet.protection.sheet:
        raise ValueError(f"目标工作表「{sheet.title}」已保护，无法填写 A2:G{last_row}")
    for merged in sheet.merged_cells.ranges:
        touches_write = (merged.min_row <= last_row and merged.max_row >= 2
                         and merged.min_col <= width)
        if touches_write or header_count == width - 1:
            raise ValueError(f"目标工作表「{sheet.title}」{merged} 含合并单元格，无法安全填写")
    for cell in sheet._cells.values():
        if cell.value is None and cell.comment is None:
            continue
        if header_count == width - 1:
            writable_header = cell.row == 1 and 1 <= cell.column <= header_count
            if not writable_header:
                raise ValueError(
                    f"目标工作表「{sheet.title}」{cell.coordinate} 已有内容；插入日期列会移动该位置，请改用空白模板或七列表头"
                )
        elif 2 <= cell.row <= last_row and 1 <= cell.column <= width:
            kind = "公式" if cell.data_type == "f" else "批注" if cell.comment else "内容"
            raise ValueError(f"目标工作表「{sheet.title}」{cell.coordinate} 已有{kind}，位于预定写入范围 A2:G{last_row}")
    if header_count == width - 1:
        validate_column_rules(sheet)
    return f"A2:G{last_row}"


def validate_column_rules(sheet) -> None:
    """Check metadata that openpyxl does not shift with insert_cols."""
    if sheet.tables or sheet.conditional_formatting or sheet.print_area:
        raise ValueError(f"目标工作表「{sheet.title}」含表格、条件格式或打印区域，无法安全迁移列规则；请改用七列表头")
    for validation in sheet.data_validations.dataValidation:
        if validation.formula1 and not (validation.type == "list" and
                                        validation.formula1.startswith('"') and
                                        validation.formula1.endswith('"')):
            raise ValueError(f"目标工作表「{sheet.title}」数据验证 {validation.sqref} 使用公式，无法安全迁移；请改用七列表头")
        if validation.formula2:
            raise ValueError(f"目标工作表「{sheet.title}」数据验证 {validation.sqref} 使用公式，无法安全迁移；请改用七列表头")
        if any(area.max_col > 6 for area in validation.sqref.ranges):
            raise ValueError(f"目标工作表「{sheet.title}」数据验证 {validation.sqref} 超出六列，无法安全迁移；请改用七列表头")
    if sheet.auto_filter.ref:
        area = CellRange(sheet.auto_filter.ref)
        if (area.min_col, area.min_row, area.max_col) != (1, 1, 6) or sheet.auto_filter.filterColumn or sheet.auto_filter.sortState:
            raise ValueError(f"目标工作表「{sheet.title}」自动筛选 {sheet.auto_filter.ref} 无法安全迁移；请改用七列表头")
    for letter, dimension in sheet.column_dimensions.items():
        index = dimension.index or letter
        if not index.isalpha() or (dimension.min and dimension.max and dimension.min != dimension.max):
            raise ValueError(f"目标工作表「{sheet.title}」列宽规则 {letter} 无法安全迁移；请改用七列表头")


def shift_column_rules(sheet) -> None:
    """Move six-column field metadata after inserting the date column."""
    for validation in sheet.data_validations.dataValidation:
        moved = []
        for area in validation.sqref.ranges:
            shifted = copy(area)
            shifted.shift(col_shift=1)
            moved.append(str(shifted))
        validation.sqref = " ".join(moved)
    if sheet.auto_filter.ref:
        area = CellRange(sheet.auto_filter.ref)
        area.max_col += 1
        sheet.auto_filter.ref = str(area)
    dimensions = list(sheet.column_dimensions.items())
    sheet.column_dimensions.clear()
    for letter, dimension in dimensions:
        shifted = copy(dimension)
        shifted.index = get_column_letter(column_index_from_string(letter) + 1)
        if shifted.min:
            shifted.min += 1
        if shifted.max:
            shifted.max += 1
        sheet.column_dimensions[shifted.index] = shifted


def safe_excel_number(value: Decimal) -> int | float | None:
    """Reject numbers Excel cannot preserve under the 15 significant digit policy."""
    if len(value.normalize().as_tuple().digits) > 15 or abs(value) >= 10**15:
        return None
    converted = int(value) if value == value.to_integral_value() else float(value)
    if isinstance(converted, float) and not isfinite(converted):
        return None
    return converted if Decimal(str(converted)) == value else None


def inspect(source: bytes, source_name: str, template: bytes, template_name: str) -> dict:
    if not source or not template or max(len(source), len(template)) > MAX_FILE_BYTES:
        raise ValueError("源文件或模板为空，或超过 12 MB")
    if not source_name.lower().endswith(".xlsx") or not template_name.lower().endswith(".xlsx"):
        raise ValueError("此链路需要 XLSX 源文件和 XLSX 目标模板")
    sources, _ = source_headers(source)
    targets = target_sheets(template)
    if not targets or not sources:
        raise ValueError("源文件或目标模板缺少工作表")
    rows_count = len(source_data_rows(source, source_name))
    workbook = load_workbook(compatible_workbook(template))
    try:
        first_error = None
        for target in targets:
            if len(target["headers"]) not in (6, 7):
                continue
            try:
                target["write_range"] = validate_write_area(
                    workbook[target["name"]], rows_count, len(target["headers"])
                )
                target["write_rows"] = rows_count
            except ValueError as error:
                target["blocked_reason"] = str(error)
                first_error = first_error or str(error)
        if not any(target.get("write_range") for target in targets):
            raise ValueError(first_error or "目标模板没有可填写的 6 列或 7 列工作表")
    finally:
        workbook.close()
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
    output_keys = tuple(keys)
    add_date = "date" not in keys
    if add_date:
        output_keys = ("date", *output_keys)
    rows = source_data_rows(source, source_name)
    by_field = {key: index for index, key in enumerate(FIELDS)}
    rows = [[row[by_field[key]] for key in output_keys] for row in rows]
    source_infos = [source_info for row in rows for item in row
                    for source_info in ([item["source"]] if item.get("source") else [])
                    + [option["source"] for option in item.get("conflicts", [])]]
    needed_rows = {(info["locator"]["sheet"], info["locator"]["row"])
                   for info in source_infos}
    raw_rows = {(sheet, row): values for sheet, row, values in workbook_rows(source)
                if (sheet, row) in needed_rows}
    book = load_workbook(compatible_workbook(template))
    try:
        sheet = book[sheet_name]
        validate_write_area(sheet, len(rows), len(chosen["headers"]))
        if add_date:
            sheet.insert_cols(1)
            shift_column_rules(sheet)
            sheet.cell(1, 1, "日期")
        def provenance_for(source_info, key):
            if not source_info:
                return {}
            locator = source_info["locator"]
            source_sheet, source_row = locator["sheet"], locator["row"]
            coordinate = source_columns[source_sheet].get(key)
            if coordinate is None:
                return {}
            source_column, source_header = coordinate
            raw = raw_rows[(source_sheet, source_row)].get(source_info.get("field", source_header))
            return {"source_sheet": source_sheet, "source_row": source_row,
                    "source_column": get_column_letter(source_column), "source_value": raw}

        def report_number(value):
            if not isinstance(value, Decimal):
                return value
            safe = safe_excel_number(value)
            return safe if safe is not None else str(value)

        cells = []
        for row_number, row in enumerate(rows, 2):
            for column_number, item in enumerate(row, 1):
                value = item["value"]
                reason, reason_code = item["reason"], item.get("reason_code")
                if isinstance(value, Decimal) and safe_excel_number(value) is None:
                    value = None
                    reason = "源数值超出 Excel 可可靠保存的范围或 15 位有效数字，已拒填并保留原值"
                    reason_code = "precision_loss"
                else:
                    value = report_number(value)
                target = sheet.cell(row_number, column_number)
                target.value = value
                if isinstance(value, str) and target.data_type == "f":
                    target.data_type = "s"
                key = output_keys[column_number - 1]
                provenance = provenance_for(item.get("source"), key)
                conflicts = [
                    {"value": report_number(option["value"]),
                     **provenance_for(option["source"], key)}
                    for option in item.get("conflicts", [])
                ]
                cells.append({"row": row_number, "column": column_number, "field": key,
                              "value": value, "decision": "filled" if value is not None else "refused",
                              "reason": reason, "reason_code": reason_code,
                              **provenance, **({"conflicts": conflicts} if conflicts else {})})
        output = BytesIO()
        book.save(output)
    finally:
        book.close()
    saved = load_workbook(BytesIO(output.getvalue()), data_only=True)
    try:
        for item in cells:
            actual = saved[sheet_name].cell(item["row"], item["column"]).value
            expected = item["value"]
            if isinstance(actual, (date, datetime)) and isinstance(expected, (date, datetime)):
                actual = actual.date() if isinstance(actual, datetime) else actual
                expected = expected.date() if isinstance(expected, datetime) else expected
            if actual != expected:
                raise ValueError(f"目标工作表「{sheet_name}」{get_column_letter(item['column'])}{item['row']} 落盘值与报告不一致，未发布副本")
    finally:
        saved.close()
    result_id = store_output(f"填写结果_{template_name}", output.getvalue())
    return {"result_id": result_id, "sheet": sheet_name, "rows": len(rows),
            "cells_written": sum(cell["value"] is not None for cell in cells),
            "cells_empty": sum(cell["value"] is None for cell in cells),
            "cells": cells, "source_sha256": plan["source_sha256"],
            "template_sha256": plan["template_sha256"]}
