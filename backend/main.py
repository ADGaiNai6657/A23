import json
from pathlib import Path
from typing import Annotated
from zipfile import BadZipFile

from fastapi import FastAPI, File, Form, HTTPException, Query, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from openpyxl.utils.exceptions import InvalidFileException

from config import DATA, DIST, INPUTS
from extractor import extract_document
from fill_spec import CITY_OUTPUT, COVID_OUTPUT
from filler import fill_workbook, resolve_facts
from parsers import leaf_name, parse_document
from schema import ExtractResult, FillRequest, ParsedDoc
from store import Store

app = FastAPI(title="A23 文档理解与数据融合")
store = Store()
REPORTS = {
    CITY_OUTPUT: ("fill_report.json", "eval_report.json"),
    COVID_OUTPUT: ("fill_report_covid19.json", "eval_report_covid19.json"),
}


def find_template(name: str) -> Path:
    leaf_name(name)
    if not name.lower().endswith(".xlsx") or not name.startswith("模板"):
        raise ValueError("请选择任务材料中的 XLSX 模板")
    matches = [path for path in INPUTS.rglob("*") if path.name == name and path.is_file()
               and path.resolve().is_relative_to(INPUTS)]
    if len(matches) != 1:
        raise ValueError("任务材料中找不到唯一匹配的模板")
    return matches[0]


def read_report(name: str) -> dict:
    try:
        return json.loads((DATA / name).read_text(encoding="utf-8"))
    except FileNotFoundError as error:
        raise HTTPException(404, "没有对应报告，请先运行自检或填写") from error
    except (OSError, ValueError) as error:
        raise HTTPException(500, "报告无法读取，请重新运行自检") from error


@app.get("/health")
def health() -> dict:
    return {"ok": True}


@app.post("/parse")
def parse(file: Annotated[UploadFile, File()]) -> ParsedDoc:
    try:
        return parse_document(file.file.read(), file.filename or "")
    except ValueError as error:
        raise HTTPException(400, str(error)) from error


@app.post("/ingest")
def ingest(file: Annotated[UploadFile, File()]) -> dict:
    try:
        content = file.file.read()
        parsed = parse_document(content, file.filename or "")
        store.save(parsed, content)
    except ValueError as error:
        raise HTTPException(400, str(error)) from error
    return {"doc": parsed.file, "blocks": len(parsed.blocks)}


@app.post("/extract", response_model=ExtractResult)
def extract(file: Annotated[UploadFile, File()], metrics: Annotated[str, Form()] = "") -> dict:
    try:
        chosen = [part.strip() for part in metrics.split(",") if part.strip()]
        return extract_document(file.file.read(), file.filename or "", chosen, store)
    except ValueError as error:
        raise HTTPException(400, str(error)) from error


@app.get("/facts")
def facts(
    entity: str, metric: str, period: str = "", scope: str = "", preferred_doc: str | None = None,
) -> dict:
    return resolve_facts(store.facts(), entity, metric, period, scope, preferred_doc)


@app.post("/fill")
def fill(request: FillRequest) -> dict:
    try:
        doc = leaf_name(request.doc)
        return fill_workbook(store, doc, find_template(request.template), DATA)
    except (BadZipFile, InvalidFileException, KeyError) as error:
        raise HTTPException(400, "无法读取 XLSX 模板，请检查文件内容") from error
    except (OSError, ValueError) as error:
        raise HTTPException(400, str(error)) from error


@app.get("/reports")
def reports(doc: str, file: str = CITY_OUTPUT) -> dict:
    if file not in REPORTS:
        raise HTTPException(400, "未识别的结果文件")
    report = read_report(REPORTS[file][1])
    if report.get("doc") != doc:
        raise HTTPException(404, "没有这份文档的自检报告")
    return report


@app.get("/cells")
def cell_detail(
    doc: str, file: str, row: Annotated[int, Query(ge=2)], column: Annotated[int, Query(ge=1)],
) -> dict:
    if file not in REPORTS:
        raise HTTPException(400, "未识别的结果文件")
    fill_name, eval_name = REPORTS[file]
    fill_report = read_report(fill_name)
    if fill_report.get("doc") != doc:
        raise HTTPException(404, "没有这份文档的填写记录")
    if row > fill_report["rows"] + 1 or column > (5 if file == CITY_OUTPUT else 7):
        raise HTTPException(404, "结果中没有这个格子")
    entry = next(
        (item for item in fill_report["cells"] if item["row"] == row and item["column"] == column), None,
    )
    if entry is None:
        return {"decision": "missing", "reason": "这格没有填写记录"}
    source = entry.get("source")
    block = store.block(source["doc"], source["block_no"]) if source and source["block_no"] else None
    graded = None
    if (DATA / eval_name).is_file():
        evaluation = read_report(eval_name)
        if evaluation.get("doc") == doc:
            diff = next((item for item in evaluation["diff"]
                         if item["row"] == row and item["column"] == column), None)
            graded = diff["decision"] if diff else "与最近自检一致"
    return {**entry, "text": block["text"] if block else None, "evaluation": graded}


@app.get("/files")
def files(name: str) -> FileResponse:
    try:
        leaf_name(name)
        if not name.startswith("filled_") or not name.endswith(".xlsx"):
            raise ValueError("只允许下载填好的 XLSX")
    except ValueError as error:
        raise HTTPException(400, str(error)) from error
    path = (DATA / name).resolve()
    if path.parent != DATA or not path.is_file():
        raise HTTPException(404, "结果文件不存在")
    return FileResponse(path, filename=name)


if (DIST / "index.html").is_file():
    app.mount("/", StaticFiles(directory=DIST, html=True), name="frontend")
else:
    @app.get("/")
    def index() -> dict:
        return {"detail": "前端尚未构建，接口说明位于 /docs"}
