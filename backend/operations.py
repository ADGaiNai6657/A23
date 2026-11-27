"""Validated natural-language document operations for DOCX and XLSX copies."""

from collections import OrderedDict
from copy import copy
from hashlib import sha256
from io import BytesIO
from threading import Lock
from time import monotonic
from uuid import uuid4

from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries
from docx.oxml.ns import qn

from extractor import ModelReply, model_response
from parser_docx import read_document
from parser_xlsx import compatible_workbook
from parsers import parse_document

MAX_FILE_BYTES = 12 * 1024 * 1024
MAX_CONTEXT = 40000
MAX_OPERATIONS = 12
PLAN_LIFETIME = 30 * 60
MAX_PLANS = 100
MAX_OUTPUTS = 20
plans: OrderedDict[str, dict] = OrderedDict()
outputs: OrderedDict[str, dict] = OrderedDict()
registry_lock = Lock()


def prune_expired(registry, now: float) -> None:
    for key, item in list(registry.items()):
        if now - item["created"] > PLAN_LIFETIME:
            registry.pop(key, None)


def store_output(name: str, content: bytes) -> str:
    """Publish a completed copy through one bounded, expiring registry."""
    result_id = uuid4().hex
    now = monotonic()
    with registry_lock:
        prune_expired(outputs, now)
        while len(outputs) >= MAX_OUTPUTS:
            outputs.popitem(last=False)
        outputs[result_id] = {"name": name, "content": content, "created": now}
    return result_id
TYPES = {
    ".docx": {"docx_replace", "docx_format", "docx_extract"},
    ".xlsx": {"xlsx_set", "xlsx_format", "xlsx_extract"},
}


def source_context(content: bytes, name: str) -> str:
    parsed = parse_document(content, name)
    lines = [f"{block.locator}: {block.text}" for block in parsed.blocks]
    if name.lower().endswith(".xlsx"):
        book = load_workbook(compatible_workbook(content), read_only=True, data_only=True)
        try:
            lines.insert(0, "工作表：" + "、".join(book.sheetnames))
            for sheet in book:
                for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 5),
                                           max_col=min(sheet.max_column, 12)):
                    lines.append(f"{sheet.title} " + "；".join(
                        f"{cell.coordinate}={cell.value}" for cell in row if cell.value is not None
                    ))
        finally:
            book.close()
    context = "\n".join(lines)
    if len(context) > MAX_CONTEXT:
        context = context[:MAX_CONTEXT] + "\n（后续内容已截断，请指定更明确的位置）"
    return context


def instruction_prompt(content: bytes, name: str, instruction: str) -> str:
    return (
        "你把中文文档操作指令转换成 JSON，不执行任何操作。"
        "只返回 {\"operations\": [...]}。按用户提到的顺序列出全部操作。"
        "支持 docx_replace(paragraph,old,new)、docx_format(paragraph,bold,align)、"
        "docx_extract(paragraph,quote)、xlsx_set(sheet,cell,value)、"
        "xlsx_format(sheet,range,bold,align)、xlsx_extract(sheet,cell,quote)。"
        "paragraph 为从 1 开始的正文段号，align 只可为 center、left、right，bold 为布尔值。"
        "格式操作只提供用户指定的 bold 或 align，也可同时提供。xlsx_set 必须包含 value，null 表示清空。"
        "提取操作的 quote 必须是原文连续片段。工作表不明确时不要猜测，可省略 sheet，由系统请用户补充。"
        "无法转换的操作请仍放进列表并给出原始 type，由系统报告不支持。"
        f"\n文件名：{name}\n文档内容与位置：\n{source_context(content, name)}"
        f"\n用户指令：{instruction}"
    )


def parse_plan(payload: object, extension: str) -> list[dict]:
    if isinstance(payload, ModelReply):
        payload = payload.payload
    if not isinstance(payload, dict) or not isinstance(payload.get("operations"), list):
        raise ValueError("模型返回的操作列表格式无效，请修改指令后重试")
    operations = payload["operations"]
    if not 1 <= len(operations) <= MAX_OPERATIONS:
        raise ValueError(f"指令须包含 1 至 {MAX_OPERATIONS} 项操作")
    for index, operation in enumerate(operations, 1):
        if not isinstance(operation, dict) or not isinstance(operation.get("type"), str):
            raise ValueError(f"第 {index} 项缺少操作类型")
        if operation["type"] not in TYPES[extension]:
            raise ValueError(f"第 {index} 项操作 {operation['type']} 不受支持，请改用当前文件的编辑、排版或提取指令")
    return operations


def paragraph_at(document, operation: dict, index: int):
    number = operation.get("paragraph")
    if type(number) is not int:
        raise ValueError(f"第 {index} 项请补充明确的段落号")
    if not 1 <= number <= len(document.paragraphs):
        raise ValueError(f"第 {index} 项第 {number} 段越界，正文共 {len(document.paragraphs)} 段")
    return document.paragraphs[number - 1]


def sheet_at(workbook, operation: dict, index: int):
    name = operation.get("sheet")
    if not isinstance(name, str) or not name.strip():
        if len(workbook.sheetnames) != 1:
            raise ValueError(f"第 {index} 项请补充工作表名称，可选：{'、'.join(workbook.sheetnames)}")
        name = workbook.sheetnames[0]
    if name not in workbook:
        raise ValueError(f"第 {index} 项找不到工作表「{name}」，可选：{'、'.join(workbook.sheetnames)}")
    return workbook[name]


def cell_range(sheet, address: object, index: int, *, single: bool):
    if not isinstance(address, str):
        raise ValueError(f"第 {index} 项请补充明确的单元格或范围")
    try:
        bounds = range_boundaries(address.upper())
    except ValueError as error:
        raise ValueError(f"第 {index} 项单元格或范围「{address}」无效") from error
    left, top, right, bottom = bounds
    if None in bounds or left < 1 or top < 1 or right < left or bottom < top:
        raise ValueError(f"第 {index} 项单元格或范围「{address}」无效")
    if single and (left != right or top != bottom):
        raise ValueError(f"第 {index} 项只能指定一个单元格")
    if right > sheet.max_column or bottom > sheet.max_row:
        raise ValueError(f"第 {index} 项「{address}」超出工作表 {sheet.title} 已有范围")
    if (right - left + 1) * (bottom - top + 1) > 1000:
        raise ValueError(f"第 {index} 项格式范围超过 1000 格，请缩小范围")
    return left, top, right, bottom


def replace_runs(paragraph, old: str, new: str, index: int, paragraph_number: int) -> None:
    # Mutate w:t nodes directly: assigning run.text rebuilds the run and drops drawings.
    nodes = [node for run in paragraph.runs for node in run._r.iter(qn("w:t"))]
    original = "".join(node.text or "" for node in nodes)
    if original != paragraph.text:
        raise ValueError(f"第 {index} 项第 {paragraph_number} 段含域、换行或特殊文字节点，无法安全替换")
    if old not in original:
        raise ValueError(f"第 {index} 项第 {paragraph_number} 段找不到「{old}」")
    positions = []
    offset = 0
    for node in nodes:
        positions.append((offset, offset + len(node.text or "")))
        offset += len(node.text or "")
    matches = []
    cursor = 0
    while (found := original.find(old, cursor)) >= 0:
        matches.append((found, found + len(old)))
        cursor = found + len(old)
    for start, end in reversed(matches):
        first = next(i for i, (_, right) in enumerate(positions) if right > start)
        last = next(i for i, (_, right) in enumerate(positions) if right >= end)
        if first != last:
            sequence = list(paragraph._p.iter())
            between = sequence[sequence.index(nodes[first]) + 1:sequence.index(nodes[last])]
            if any(node.tag in {qn("w:drawing"), qn("w:object"), qn("w:pict")} for node in between):
                raise ValueError(f"第 {index} 项第 {paragraph_number} 段替换文字跨越内嵌图片或对象，无法安全替换")
        first_left = positions[first][0]
        last_left = positions[last][0]
        if first == last:
            value = nodes[first].text or ""
            nodes[first].text = value[:start - first_left] + new + value[end - first_left:]
            continue
        nodes[first].text = (nodes[first].text or "")[:start - first_left] + new
        for middle in nodes[first + 1:last]:
            middle.text = ""
        nodes[last].text = (nodes[last].text or "")[end - last_left:]


def alignment(value: object, index: int):
    options = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
               "right": WD_ALIGN_PARAGRAPH.RIGHT}
    if value not in options:
        raise ValueError(f"第 {index} 项对齐方式只支持 left、center、right")
    return options[value]


def apply_docx(document, source, operation: dict, index: int) -> dict:
    kind = operation["type"]
    paragraph = paragraph_at(source if kind == "docx_extract" else document, operation, index)
    number = operation["paragraph"]
    before = paragraph.text
    if kind == "docx_replace":
        old, new = operation.get("old"), operation.get("new")
        if not isinstance(old, str) or not old or not isinstance(new, str):
            raise ValueError(f"第 {index} 项请提供要替换的原文和新文字")
        replace_runs(paragraph, old, new, index, number)
        result = {"before": before, "after": paragraph.text}
    elif kind == "docx_format":
        if not ({"bold", "align"} & operation.keys()):
            raise ValueError(f"第 {index} 项请明确加粗或对齐方式")
        if "bold" in operation and type(operation["bold"]) is not bool:
            raise ValueError(f"第 {index} 项加粗状态须为布尔值")
        if not paragraph.runs:
            raise ValueError(f"第 {index} 项第 {number} 段为空，无法排版")
        if "bold" in operation:
            for run in paragraph.runs:
                run.bold = operation["bold"]
        if "align" in operation:
            paragraph.alignment = alignment(operation["align"], index)
        result = {"before": before, "after": "，".join(f"{key}={operation[key]}" for key in ("bold", "align") if key in operation)}
    else:
        quote = operation.get("quote")
        if not isinstance(quote, str) or not quote or quote not in before:
            raise ValueError(f"第 {index} 项第 {number} 段找不到提取原文，请核对引用")
        result = {"before": before, "after": quote, "quote": quote}
    return {"index": index, "type": kind, "location": f"第 {number} 段", **result}


def apply_xlsx(workbook, source, operation: dict, index: int) -> dict:
    kind = operation["type"]
    sheet = sheet_at(source if kind == "xlsx_extract" else workbook, operation, index)
    address = operation.get("range") if kind == "xlsx_format" else operation.get("cell")
    left, top, right, bottom = cell_range(sheet, address, index, single=kind != "xlsx_format")
    location = f"{sheet.title}!{address.upper()}"
    if kind == "xlsx_set":
        if "value" not in operation:
            raise ValueError(f"第 {index} 项缺少写入值 value；如需清空请明确指定 null")
        value = operation.get("value")
        if value is not None and type(value) not in (str, int, float, bool):
            raise ValueError(f"第 {index} 项写入值类型无效")
        cell = sheet.cell(top, left)
        before = cell.value
        cell.value = value
        return {"index": index, "type": kind, "location": location, "before": before, "after": value}
    if kind == "xlsx_extract":
        original = sheet.cell(top, left).value
        quote = operation.get("quote")
        if not isinstance(quote, str) or not quote or original is None or quote not in str(original):
            raise ValueError(f"第 {index} 项 {location} 找不到提取原文，请核对引用")
        return {"index": index, "type": kind, "location": location, "before": original,
                "after": quote, "quote": quote}
    if not ({"bold", "align"} & operation.keys()):
        raise ValueError(f"第 {index} 项请明确加粗或对齐方式")
    if "bold" in operation and type(operation["bold"]) is not bool:
        raise ValueError(f"第 {index} 项加粗状态须为布尔值")
    direction = operation.get("align")
    if "align" in operation and direction not in {"left", "center", "right"}:
        raise ValueError(f"第 {index} 项对齐方式只支持 left、center、right")
    for row in sheet.iter_rows(min_row=top, max_row=bottom, min_col=left, max_col=right):
        for cell in row:
            if "bold" in operation:
                font = copy(cell.font)
                font.bold = operation["bold"]
                cell.font = font
            if "align" in operation:
                aligned = copy(cell.alignment)
                aligned.horizontal = direction
                cell.alignment = aligned
    return {"index": index, "type": kind, "location": location,
            "before": f"{(right-left+1)*(bottom-top+1)} 格原格式",
            "after": "，".join(f"{key}={operation[key]}" for key in ("bold", "align") if key in operation)}


def run_operations(content: bytes, name: str, operations: list[dict], *, save: bool) -> tuple[list[dict], bytes | None]:
    extension = "." + name.rsplit(".", 1)[-1].lower()
    if extension == ".docx":
        source = read_document(content)
        document = read_document(content)
        apply = apply_docx
    else:
        source = load_workbook(compatible_workbook(content))
        document = load_workbook(compatible_workbook(content))
        apply = apply_xlsx
    try:
        results = [apply(document, source, operation, index)
                   for index, operation in enumerate(operations, 1)]
        changed = any(not operation["type"].endswith("_extract") for operation in operations)
        if not save or not changed:
            return results, None
        output = BytesIO()
        document.save(output)
        return results, output.getvalue()
    finally:
        if extension == ".xlsx":
            source.close()
            document.close()


def preview(content: bytes, name: str, instruction: str) -> dict:
    extension = "." + name.rsplit(".", 1)[-1].lower()
    if extension not in TYPES:
        raise ValueError("自然语言操作仅支持 DOCX 和 XLSX")
    if not content or len(content) > MAX_FILE_BYTES:
        raise ValueError("文件为空或超过 12 MB，请选择较小的 DOCX 或 XLSX")
    if not instruction.strip():
        raise ValueError("请输入中文操作指令")
    if extension == ".xlsx":
        workbook = load_workbook(compatible_workbook(content), read_only=True)
        try:
            if len(workbook.sheetnames) > 1 and not any(
                sheet in instruction for sheet in workbook.sheetnames
            ):
                raise ValueError(f"请补充工作表名称，可选：{'、'.join(workbook.sheetnames)}")
        finally:
            workbook.close()
    prompt = instruction_prompt(content, name, instruction.strip())
    try:
        payload = model_response(prompt)
    except ValueError as error:
        raise ValueError(f"指令解析失败：{error}") from error
    operations = parse_plan(payload, extension)
    results, _ = run_operations(content, name, operations, save=False)
    plan_id = uuid4().hex
    plan = {"name": name, "sha256": sha256(content).hexdigest(), "operations": operations,
            "preview": results, "created": monotonic()}
    with registry_lock:
        prune_expired(plans, monotonic())
        while len(plans) >= MAX_PLANS:
            plans.popitem(last=False)
        plans[plan_id] = plan
    return {"plan_id": plan_id, "file": name, "source_sha256": plan["sha256"],
            "operations": results}


def execute(content: bytes, name: str, plan_id: str) -> dict:
    with registry_lock:
        prune_expired(plans, monotonic())
        plan = plans.get(plan_id)
        if plan is None:
            raise ValueError("操作计划已失效，请重新预览并确认")
        if plan["name"] != name or plan["sha256"] != sha256(content).hexdigest():
            raise ValueError("上传文件与已确认计划的原文件版本不符，请重新预览")
        plans.pop(plan_id, None)
    try:
        results, output = run_operations(content, name, plan["operations"], save=True)
        result_id = store_output(f"操作结果_{name}", output) if output is not None else None
    except Exception:
        with registry_lock:
            if monotonic() - plan["created"] <= PLAN_LIFETIME:
                prune_expired(plans, monotonic())
                while len(plans) >= MAX_PLANS:
                    plans.popitem(last=False)
                plans[plan_id] = plan
        raise
    return {"plan_id": plan_id, "source_sha256": plan["sha256"], "operations": results,
            "result_id": result_id}


def download(result_id: str) -> dict | None:
    with registry_lock:
        prune_expired(outputs, monotonic())
        return outputs.get(result_id)
