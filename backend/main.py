import json
import logging
from hashlib import sha256
from pathlib import Path
from threading import Lock
from typing import Annotated
from urllib.parse import quote
from zipfile import BadZipFile

from fastapi import FastAPI, File, Form, HTTPException, Query, Response, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from openpyxl.utils.exceptions import InvalidFileException

from check_dataset import score_published_fill
from config import DATA, DIST, INPUTS, REFERENCE_DIR, SOURCE_CHECK_ROOT
from evaluate import evaluate
from extractor import extract_document
from fill_spec import CITY_OUTPUT, COVID_OUTPUT, REFERENCE_FILES, REPORT_FILES
from filler import block_digest, fill_workbook, resolve_facts
from parsers import leaf_name, parse_document
from schema import ExtractResult, FillRequest, ParsedDoc
from store import Store

app = FastAPI(title="A23 文档理解与数据融合")
store = Store()
fill_lock = Lock()
logger = logging.getLogger(__name__)


def find_template(name: str) -> Path:
    leaf_name(name)
    filename = Path(name)
    if filename.suffix.lower() != ".xlsx" or not (
        filename.stem.startswith("模板") or filename.stem.endswith("模板")
    ):
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
def parse(file: Annotated[UploadFile, File()], response: Response) -> ParsedDoc:
    try:
        content = file.file.read()
        parsed = parse_document(content, file.filename or "")
    except ValueError as error:
        raise HTTPException(400, str(error)) from error
    response.headers["X-A23-Source-SHA256"] = sha256(content).hexdigest()
    return parsed


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
def fill(request: FillRequest, response: Response) -> dict:
    try:
        doc = leaf_name(request.doc)
        with fill_lock:
            result = fill_workbook(store, doc, find_template(request.template), DATA, request.source_sha256)
            try:
                score_filled(doc)
            except (OSError, ValueError, BadZipFile, InvalidFileException, KeyError, StopIteration):
                logger.exception("填写已完成，自检报告生成失败")
                response.headers["X-A23-Score-Status"] = "failed"
        response.headers["X-A23-Result-ID"] = result.pop("result_id")
        response.headers["X-A23-Fill-Seconds"] = str(result.pop("fill_seconds"))
        return result
    except (BadZipFile, InvalidFileException, KeyError) as error:
        raise HTTPException(400, "无法读取 XLSX 模板，请检查文件内容") from error
    except (OSError, ValueError) as error:
        raise HTTPException(400, str(error)) from error


def score_filled(doc: str) -> None:
    file = CITY_OUTPUT if doc.lower().endswith(".docx") else COVID_OUTPUT
    reference = REFERENCE_DIR / REFERENCE_FILES[file] if REFERENCE_DIR else None
    if reference is not None and reference.is_file():
        evaluate(DATA / file, reference, DATA / REPORT_FILES[file][1], doc, file == COVID_OUTPUT,
                 block_count=len(store.blocks(doc)), database=store.path)
    elif SOURCE_CHECK_ROOT:
        score_published_fill(SOURCE_CHECK_ROOT, DATA, store, file, doc)


def current_trace(doc: str | None, file: str, result_id: str | None, content: bytes | None = None) -> dict:
    if file not in REPORT_FILES:
        raise HTTPException(400, "未识别的结果文件")
    trace = read_report(REPORT_FILES[file][0])
    if doc is not None and trace.get("doc") != doc:
        raise HTTPException(404, "没有这份文档的填写记录")
    if result_id and trace.get("result_id") != result_id:
        raise HTTPException(409, "填写结果已更新，请重新填写或查询")
    if not trace.get("result_id"):
        raise HTTPException(409, "旧版填写记录无法核对版本，请重新填写")
    try:
        actual_hash = sha256((DATA / file).read_bytes() if content is None else content).hexdigest()
    except FileNotFoundError as error:
        raise HTTPException(404, "填写结果文件已缺失，请重新填写") from error
    if trace.get("workbook_sha256") != actual_hash:
        raise HTTPException(409, "填写工作簿与逐格记录版本不同，请重新填写")
    return trace


def matching_score(file: str, trace: dict) -> dict | None:
    for name in REPORT_FILES[file][1:]:
        if (DATA / name).is_file():
            report = read_report(name)
            if report.get("file") == file and trace.get("run_id") and all(
                report.get(key) == trace.get(key)
                for key in ("doc", "result_id", "run_id", "workbook_sha256", "source_sha256", "template_sha256")
            ):
                return report
    return None


@app.get("/reports")
def reports(doc: str, file: str = CITY_OUTPUT, result_id: str | None = None) -> dict:
    score = matching_score(file, current_trace(doc, file, result_id))
    if score is None:
        raise HTTPException(404, "本次填写尚无匹配的自检报告")
    return score


@app.get("/cells")
def cell_detail(
    doc: str, file: str, row: Annotated[int, Query(ge=2)], column: Annotated[int, Query(ge=1)],
    result_id: str | None = None,
) -> dict:
    fill_report = current_trace(doc, file, result_id)
    if row > fill_report["rows"] + 1 or column > (5 if file == CITY_OUTPUT else 7):
        raise HTTPException(404, "结果中没有这个格子")
    entry = next(
        (item for item in fill_report["cells"] if item["row"] == row and item["column"] == column), None,
    )
    if entry is None:
        return {"decision": "missing", "reason": "这格没有填写记录"}
    source = entry.get("source")
    block = store.block(source["doc"], source["block_no"]) if source and source["block_no"] else None
    verified = (
        block is not None
        and block["document_digest"].rsplit(":", 1)[-1] == fill_report["source_sha256"]
        and block_digest(block) == source.get("block_sha256")
    )
    graded = None
    evaluation = matching_score(file, fill_report)
    if evaluation:
        diff = next((item for item in evaluation["diff"]
                     if item["row"] == row and item["column"] == column), None)
        graded = diff["decision"] if diff else "与本次核对依据一致"
    return {
        **entry, "text": block["text"] if verified else None, "evaluation": graded,
        "evidence_status": "stale" if source and not verified else "available" if verified else "none",
    }


@app.get("/files")
def files(name: str, result_id: str | None = None) -> Response:
    try:
        leaf_name(name)
        if not name.startswith("filled_") or not name.endswith(".xlsx"):
            raise ValueError("只允许下载填好的 XLSX")
    except ValueError as error:
        raise HTTPException(400, str(error)) from error
    path = (DATA / name).resolve()
    if path.parent != DATA or not path.is_file():
        raise HTTPException(404, "结果文件不存在")
    if result_id is not None:
        content = path.read_bytes()
        current_trace(None, name, result_id, content)
        return Response(
            content, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(name)}"},
        )
    return FileResponse(path, filename=name)


if (DIST / "index.html").is_file():
    app.mount("/", StaticFiles(directory=DIST, html=True), name="frontend")
else:
    @app.get("/")
    def index() -> dict:
        return {"detail": "前端尚未构建，接口说明位于 /docs"}
