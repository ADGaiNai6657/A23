"""Auditable first DOCX template chain for the public air-quality workbook."""

from collections import OrderedDict, defaultdict
from copy import deepcopy
from hashlib import sha256
from io import BytesIO
from threading import Lock
from time import monotonic
from uuid import uuid4

from docx import Document
from docx.oxml.ns import qn
from docx.table import _Cell
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from parser_xlsx import compatible_workbook
from operations import MAX_PLANS, PLAN_LIFETIME, prune_expired, store_output

CITIES = ("德州市", "潍坊市", "临沂市")
PERIOD = "2025-11-25 09:00:00.0"
HEADERS = ("城市", "区", "站点名称", "空气质量指数", "PM10监测值", "PM2.5监测值", "首要污染物", "污染类型")
plans: OrderedDict[str, dict] = OrderedDict()
lock = Lock()


def source_rows(content: bytes) -> tuple[list[tuple[str, int, dict]], dict[str, dict[str, int]]]:
    book = load_workbook(compatible_workbook(content), read_only=True, data_only=True)
    selected = []
    columns = {}
    try:
        for sheet in book:
            sheet.reset_dimensions()
            rows = sheet.iter_rows(values_only=True)
            names = next(rows, ())
            positions = defaultdict(list)
            for column, value in enumerate(names):
                if value is not None:
                    positions[str(value).strip()].append(column)
            needed = {"监测时间", *HEADERS}
            if not needed.issubset(positions):
                continue
            for name in needed:
                if len(positions[name]) > 1:
                    conflict = "、".join(get_column_letter(column + 1) for column in positions[name])
                    raise ValueError(f"工作表「{sheet.title}」必要字段「{name}」重复，冲突列：{conflict}；请明确选择列")
            lookup = {name: columns[0] for name, columns in positions.items()}
            columns[sheet.title] = {name: lookup[name] + 1 for name in HEADERS}
            for row_number, values in enumerate(rows, 2):
                if (len(values) > lookup["城市"] and len(values) > lookup["监测时间"]
                        and values[lookup["监测时间"]] == PERIOD
                        and values[lookup["城市"]] in CITIES):
                    selected.append((sheet.title, row_number, {
                        name: values[lookup[name]] if lookup[name] < len(values) else None
                        for name in HEADERS
                    }))
    finally:
        book.close()
    if not selected:
        raise ValueError("源 XLSX 未找到指定时间和城市的空气质量记录，请核对监测时间及表头")
    counts = {city: sum(values["城市"] == city for _, _, values in selected) for city in CITIES}
    if any(count == 0 for count in counts.values()):
        raise ValueError(f"源 XLSX 缺少指定时间的城市记录：{counts}")
    return selected, columns


def template_tables(content: bytes):
    document = Document(BytesIO(content))
    if len(document.tables) != 3:
        raise ValueError(f"模板须包含三张表，当前为 {len(document.tables)} 张")
    for number, (city, table) in enumerate(zip(CITIES, document.tables), 1):
        preceding = table._tbl.getprevious()
        while preceding is not None:
            text = "".join(node.text or "" for node in preceding.iter(qn("w:t")))
            if text.strip():
                break
            preceding = preceding.getprevious()
        if preceding is None or city not in text or "2025年11月25日09:00" not in text:
            raise ValueError(f"第 {number} 张表前缺少 {city} 的指定时间段落")
        headers = tuple(cell.text.strip() for cell in table.rows[0].cells)
        if headers != HEADERS:
            raise ValueError(f"第 {number} 张表表头与空气质量字段不符：{headers}")
        if any(cell.text.strip() for row in table.rows[1:] for cell in row.cells):
            raise ValueError(f"第 {number} 张表已有数据，请使用空白模板")
    return document


def physical_cells(table, table_number: int, row_index: int):
    """Return a rectangular physical row or report its first unmappable grid column."""
    row = table.rows[row_index]
    width = len(HEADERS)
    if row.grid_cols_before:
        raise ValueError(f"第 {table_number} 张表第 {row_index + 1} 行第 1 列省略，无法一格一值填写")
    cells = row._tr.tc_lst
    if row.grid_cols_after:
        column = width - row.grid_cols_after + 1
        raise ValueError(f"第 {table_number} 张表第 {row_index + 1} 行第 {column} 列省略，无法一格一值填写")
    for column, tc in enumerate(cells, 1):
        merged = tc.grid_span != 1 or tc.vMerge is not None
        objects = any(node.tag in {qn("w:drawing"), qn("w:object"), qn("w:pict")}
                      for node in tc.iter())
        if merged or objects:
            reason = "合并单元格" if merged else "内嵌对象"
            raise ValueError(f"第 {table_number} 张表第 {row_index + 1} 行第 {column} 列含{reason}，无法一格一值填写")
    if len(cells) != width:
        column = min(len(cells) + 1, width + 1)
        raise ValueError(f"第 {table_number} 张表第 {row_index + 1} 行第 {column} 列行宽异常，实际 {len(cells)} 格、预期 {width} 格")
    return [_Cell(tc, table) for tc in cells]


def validate_target_cells(document, grouped) -> None:
    """Validate rows by physical w:tc position, never the table-wide flat index."""
    for table_number, (city, table) in enumerate(zip(CITIES, document.tables), 1):
        count = len(grouped[city])
        capacity = len(table.rows) - 1
        rows_to_check = range(1, min(count, capacity) + 1)
        for row_index in rows_to_check:
            physical_cells(table, table_number, row_index)


def set_cell_text(cell, value: str) -> None:
    """Keep the existing paragraph style, alignment and cell formatting."""
    paragraph = cell.paragraphs[0]
    if paragraph.runs:
        paragraph.runs[0].text = value
    else:
        paragraph.add_run(value)


def inspect(source: bytes, template: bytes) -> dict:
    rows, columns = source_rows(source)
    document = template_tables(template)
    grouped = defaultdict(list)
    for sheet, row, values in rows:
        grouped[values["城市"]].append((sheet, row, values))
    validate_target_cells(document, grouped)
    summaries = []
    for city, table in zip(CITIES, document.tables):
        count = len(grouped[city])
        capacity = len(table.rows) - 1
        summaries.append({"city": city, "source_rows": count, "template_rows": capacity,
                          "added_rows": max(0, count - capacity)})
    plan_id = uuid4().hex
    with lock:
        prune_expired(plans, monotonic())
        while len(plans) >= MAX_PLANS:
            plans.popitem(last=False)
        plans[plan_id] = {"source_sha256": sha256(source).hexdigest(),
                          "template_sha256": sha256(template).hexdigest(), "created": monotonic()}
    return {"plan_id": plan_id, "period": PERIOD, "tables": summaries,
            "source_columns": {sheet: {name: get_column_letter(column) for name, column in mapping.items()}
                               for sheet, mapping in columns.items()}}


def fill(source: bytes, template: bytes, plan_id: str) -> dict:
    with lock:
        prune_expired(plans, monotonic())
        plan = plans.get(plan_id)
        if plan is None:
            raise ValueError("模板计划已失效，请重新预览")
        if sha256(source).hexdigest() != plan["source_sha256"] or sha256(template).hexdigest() != plan["template_sha256"]:
            raise ValueError("源文件或模板版本与已确认计划不符，请重新预览")
        plans.pop(plan_id, None)
    try:
        return _fill_claimed(source, template, plan)
    except Exception:
        with lock:
            if monotonic() - plan["created"] <= PLAN_LIFETIME:
                prune_expired(plans, monotonic())
                while len(plans) >= MAX_PLANS:
                    plans.popitem(last=False)
                plans[plan_id] = plan
        raise


def _fill_claimed(source: bytes, template: bytes, plan: dict) -> dict:
    rows, columns = source_rows(source)
    document = template_tables(template)
    grouped = defaultdict(list)
    for sheet, row, values in rows:
        grouped[values["城市"]].append((sheet, row, values))
    validate_target_cells(document, grouped)
    cells = []
    for table_number, (city, table) in enumerate(zip(CITIES, document.tables), 1):
        records = grouped[city]
        for _ in range(max(0, len(records) - len(table.rows) + 1)):
            table._tbl.append(deepcopy(table.rows[-1]._tr))
        for target_row, (sheet, source_row, values) in enumerate(records, 2):
            targets = physical_cells(table, table_number, target_row - 1)
            for target_column, name in enumerate(HEADERS, 1):
                original = values[name]
                missing = original is None or original == "" or original == "—"
                written = "—" if original is None or original == "" else str(original)
                set_cell_text(targets[target_column - 1], written)
                cells.append({"table": table_number, "city": city, "row": target_row,
                              "column": target_column, "field": name, "value": written,
                              "source_sheet": sheet, "source_row": source_row,
                              "source_column": get_column_letter(columns[sheet][name]),
                              "source_value": original, "source_missing": missing})
    output = BytesIO()
    document.save(output)
    result_id = store_output("空气质量填写结果.docx", output.getvalue())
    return {"period": PERIOD, "tables": [len(table.rows) for table in document.tables],
            "cells": cells, "source_sha256": plan["source_sha256"], "result_id": result_id}
