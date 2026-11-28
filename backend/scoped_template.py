"""Cross-document filling for regular DOCX tables and XLSX sheets."""

from collections import OrderedDict
from decimal import Decimal, InvalidOperation
from hashlib import sha256
from io import BytesIO
from threading import Lock
from time import monotonic, perf_counter
from uuid import uuid4

from docx import Document
from docx.oxml.ns import qn
from docx.table import _Cell
from openpyxl import load_workbook

from air_quality import set_cell_text
from fill_spec import metric_key, normalize
from filler import resolve_facts
from operations import MAX_FILE_BYTES, MAX_PLANS, PLAN_LIFETIME, prune_expired, store_output
from parser_xlsx import compatible_workbook
from parsers import leaf_name
from store import Store
from xlsx_template import safe_excel_number

plans: OrderedDict[str, dict] = OrderedDict()
lock = Lock()
MAX_ROWS = 500
MAX_COLUMNS = 20


def physical_cells(table, table_number: int, row_number: int, width: int):
    """Map a Word row to physical cells, rejecting merged or omitted positions."""
    row = table.rows[row_number - 1]
    if row.grid_cols_before or row.grid_cols_after:
        raise ValueError(f"第 {table_number} 张表第 {row_number} 行有省略列，无法一格一值填写")
    cells = row._tr.tc_lst
    if len(cells) != width:
        raise ValueError(f"第 {table_number} 张表第 {row_number} 行宽异常：需 {width} 格，实际 {len(cells)} 格")
    for column, tc in enumerate(cells, 1):
        if tc.grid_span != 1 or tc.vMerge is not None:
            raise ValueError(f"第 {table_number} 张表第 {row_number} 行第 {column} 列为合并格")
        if any(node.tag in {qn("w:drawing"), qn("w:object"), qn("w:pict")} for node in tc.iter()):
            raise ValueError(f"第 {table_number} 张表第 {row_number} 行第 {column} 列含内嵌对象")
    return [_Cell(tc, table) for tc in cells]


def structure(content: bytes, name: str) -> list[dict]:
    """Read only regular, bounded template targets; leave mapping for the user."""
    if not content or len(content) > MAX_FILE_BYTES:
        raise ValueError("模板为空或超过 12 MB")
    extension = leaf_name(name).lower().rsplit(".", 1)[-1]
    targets = []
    if extension == "docx":
        document = Document(BytesIO(content))
        for number, table in enumerate(document.tables, 1):
            if not table.rows:
                continue
            width = len(table.rows[0]._tr.tc_lst)
            if not 2 <= width <= MAX_COLUMNS or len(table.rows) - 1 > MAX_ROWS:
                raise ValueError(f"第 {number} 张表超出规整模板支持的列数或行数")
            headers = [cell.text.strip() for cell in physical_cells(table, number, 1, width)]
            for row_number in range(2, len(table.rows) + 1):
                physical_cells(table, number, row_number, width)
            targets.append({"id": f"table:{number}", "label": f"第 {number} 张表",
                            "headers": headers, "rows": len(table.rows) - 1})
    elif extension == "xlsx":
        book = load_workbook(compatible_workbook(content))
        try:
            for sheet in book:
                width = 0
                while width < sheet.max_column and sheet.cell(1, width + 1).value is not None:
                    width += 1
                if width < 2:
                    continue
                rows = 0
                while (rows + 2 <= sheet.max_row and
                       any(sheet.cell(rows + 2, column).value is not None
                           for column in range(1, width + 1))):
                    rows += 1
                if not 2 <= width <= MAX_COLUMNS or rows > MAX_ROWS:
                    raise ValueError(f"工作表「{sheet.title}」超出规整模板支持的列数或行数")
                if any(merged.min_row <= rows + 1 and merged.max_row >= 1
                       and merged.min_col <= width for merged in sheet.merged_cells.ranges):
                    raise ValueError(f"工作表「{sheet.title}」含合并格，请使用规整模板")
                targets.append({"id": sheet.title, "label": sheet.title,
                                "headers": [str(sheet.cell(1, col).value or "").strip()
                                            for col in range(1, width + 1)],
                                "rows": rows})
        finally:
            book.close()
    else:
        raise ValueError("仅支持 DOCX 表格或 XLSX 工作表模板")
    if not targets:
        raise ValueError("模板没有可填写的表格或工作表")
    return targets


def validate_mapping(target: dict, entity_column: int, mapping: dict) -> dict[int, dict]:
    width = len(target["headers"])
    if type(entity_column) is not int or not 1 <= entity_column <= width:
        raise ValueError("请选择明确的实体列")
    if not isinstance(mapping, dict) or set(mapping) != {str(col) for col in range(1, width + 1)
                                                          if col != entity_column}:
        raise ValueError("请为实体列之外的每个目标列确认指标映射")
    result = {}
    for column, raw in mapping.items():
        if not isinstance(raw, dict) or not isinstance(raw.get("metric"), str) or not raw["metric"].strip():
            raise ValueError(f"第 {column} 列缺少指标名称")
        if any(not isinstance(raw.get(field, ""), str) for field in ("period", "scope", "unit")):
            raise ValueError(f"第 {column} 列期次、口径或单位须为文本")
        result[int(column)] = {field: raw.get(field, "").strip()
                               for field in ("metric", "period", "scope", "unit")}
    return result


def rows_for(content: bytes, name: str, target: str, entity_column: int,
             mapping: dict[int, dict], rows_count: int) -> list[dict]:
    """Inspect every target position before publishing a plan."""
    positions = []
    if name.lower().endswith(".docx"):
        document = Document(BytesIO(content))
        table_number = int(target.split(":", 1)[1])
        table = document.tables[table_number - 1]
        width = len(table.rows[0]._tr.tc_lst)
        for row_number in range(2, len(table.rows) + 1):
            cells = physical_cells(table, table_number, row_number, width)
            entity = cells[entity_column - 1].text.strip()
            if not entity:
                raise ValueError(f"第 {table_number} 张表第 {row_number} 行实体列为空")
            for column in mapping:
                if cells[column - 1].text.strip():
                    raise ValueError(f"第 {table_number} 张表第 {row_number} 行第 {column} 列已有内容")
                positions.append({"row": row_number, "column": column, "entity": entity})
    else:
        book = load_workbook(compatible_workbook(content))
        try:
            sheet = book[target]
            for row_number in range(2, rows_count + 2):
                entity = sheet.cell(row_number, entity_column).value
                if not isinstance(entity, str) or not entity.strip():
                    raise ValueError(f"工作表「{target}」第 {row_number} 行实体列须为非空文本")
                for column in mapping:
                    cell = sheet.cell(row_number, column)
                    if cell.value is not None:
                        raise ValueError(f"工作表「{target}」{cell.coordinate} 已有值或公式")
                    positions.append({"row": row_number, "column": column,
                                      "entity": entity.strip()})
        finally:
            book.close()
    if not positions:
        raise ValueError("目标没有数据行，请先在模板中列出实体")
    return positions


def source_of(fact: dict) -> dict:
    return {key: fact[key] for key in ("doc", "source_sha256", "publication_id", "block_no",
                                      "locator", "snippet", "value", "unit")}


def preview(store: Store, content: bytes, name: str, docs: list[str], target_id: str,
            entity_column: int, mapping: dict) -> dict:
    targets = structure(content, name)
    target = next((item for item in targets if item["id"] == target_id), None)
    if target is None:
        raise ValueError("请选择预览中列出的目标表或工作表")
    if not isinstance(docs, list) or not docs or len(docs) > 16 or len(set(docs)) != len(docs):
        raise ValueError("请选择 1 至 16 份互不重复的资料")
    columns = validate_mapping(target, entity_column, mapping)
    positions = rows_for(content, name, target_id, entity_column, columns, target["rows"])
    versions, facts = store.scoped_facts(docs)
    cells = []
    for position in positions:
        field = columns[position["column"]]
        same_metric = [fact for fact in facts if fact["entity"] == position["entity"]
                       and metric_key(fact["metric"]) == metric_key(field["metric"])
                       and fact["period"] == field["period"] and fact["scope"] == field["scope"]]
        matching = [fact for fact in facts if not field["unit"]
                    or normalize(fact["unit"]) == normalize(field["unit"])]
        decision = resolve_facts(matching, position["entity"], field["metric"],
                                 field["period"], field["scope"])
        selected = decision["selected"]
        value = None
        reason = decision["reason"]
        code = None
        if selected is not None:
            try:
                numeric = Decimal(selected["value_num"])
                value = safe_excel_number(numeric)
            except (InvalidOperation, TypeError):
                pass
            if value is None:
                reason, code = "数值超出 Excel 安全精度，保留来源原值", "precision_loss"
        elif decision["candidates"]:
            code = "conflict" if "不同值" in reason else "incompatible_value"
        elif same_metric:
            reason, code = "候选事实的单位与目标列映射不一致", "incompatible_value"
            decision["candidates"] = same_metric
        else:
            reason, code = "所选资料范围内没有同实体、指标、期次和口径的事实", "missing_fact"
        cells.append({**position, "metric": field["metric"], "period": field["period"],
                      "scope": field["scope"], "unit": field["unit"], "value": value,
                      "decision": "filled" if value is not None else "refused",
                      "reason": reason, "reason_code": code,
                      "source": source_of(selected) if value is not None else None,
                      "candidates": [source_of(fact) for fact in decision["candidates"]]})
    plan_id = uuid4().hex
    plan = {"created": monotonic(), "name": name, "template_sha256": sha256(content).hexdigest(),
            "target": target_id, "entity_column": entity_column, "mapping": columns,
            "rows": target["rows"], "docs": docs, "versions": versions, "cells": cells}
    with lock:
        prune_expired(plans, monotonic())
        while len(plans) >= MAX_PLANS:
            plans.popitem(last=False)
        plans[plan_id] = plan
    return {"plan_id": plan_id, "template_sha256": plan["template_sha256"],
            "target": target, "docs": docs, "versions": versions, "cells": cells,
            "cells_written": sum(item["decision"] == "filled" for item in cells),
            "cells_empty": sum(item["decision"] == "refused" for item in cells)}


def fill(store: Store, content: bytes, name: str, plan_id: str) -> dict:
    start = perf_counter()
    with lock:
        prune_expired(plans, monotonic())
        plan = plans.pop(plan_id, None)
    if plan is None:
        raise ValueError("填写计划已失效，请重新预览")
    try:
        if name != plan["name"] or sha256(content).hexdigest() != plan["template_sha256"]:
            raise ValueError("上传模板与已确认计划版本不符")
        current, _ = store.scoped_facts(plan["docs"])
        if current != plan["versions"]:
            raise ValueError("所选资料的事实版本已变化，请重新预览")
        rows_for(content, name, plan["target"], plan["entity_column"], plan["mapping"], plan["rows"])
        if name.lower().endswith(".docx"):
            document = Document(BytesIO(content))
            number = int(plan["target"].split(":", 1)[1])
            table = document.tables[number - 1]
            for item in plan["cells"]:
                if item["decision"] == "filled":
                    cell = physical_cells(table, number, item["row"], len(table.rows[0]._tr.tc_lst))[item["column"] - 1]
                    set_cell_text(cell, str(item["value"]))
            stream = BytesIO()
            document.save(stream)
        else:
            book = load_workbook(compatible_workbook(content))
            try:
                sheet = book[plan["target"]]
                for item in plan["cells"]:
                    if item["decision"] == "filled":
                        sheet.cell(item["row"], item["column"], item["value"])
                stream = BytesIO()
                book.save(stream)
            finally:
                book.close()
        result_id = store_output(f"跨文档结果_{name}", stream.getvalue())
    except Exception:
        with lock:
            if monotonic() - plan["created"] <= PLAN_LIFETIME:
                prune_expired(plans, monotonic())
                while len(plans) >= MAX_PLANS:
                    plans.popitem(last=False)
                plans[plan_id] = plan
        raise
    return {"result_id": result_id, "target": plan["target"], "docs": plan["docs"],
            "versions": plan["versions"], "cells": plan["cells"],
            "cells_written": sum(item["decision"] == "filled" for item in plan["cells"]),
            "cells_empty": sum(item["decision"] == "refused" for item in plan["cells"]),
            "fill_seconds": round(perf_counter() - start, 4)}
