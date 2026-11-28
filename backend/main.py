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
from config import DATA, DIST, INPUTS, REFERENCE_DIR, SOURCE_CHECK_ROOT, SYNTHETIC_REFERENCE
from criteria import checked_rules, load_expectations
from evaluate import evaluate
from extractor import extract_document
from fill_spec import CITY_OUTPUT, COVID_OUTPUT, REFERENCE_FILES, REPORT_FILES
from filler import block_digest, fill_workbook, resolve_facts
import operations
import air_quality
import xlsx_template
import scoped_template
import task_queue
from parsers import leaf_name, parse_document
from schema import ExtractResult, FactDecisionRequest, FillRequest, ParsedDoc
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
        content = file.file.read()
        name = leaf_name(file.filename or "")
        chosen = [part.strip() for part in metrics.split(",") if part.strip()]
        request_id = store.register_extraction(name, content)
        result = extract_document(content, name, chosen, store,
                                  publish=lambda parsed, data, mode, facts, summary:
                                  store.publish_extraction(parsed, data, mode, facts, request_id, summary))
        if not result.pop("published"):
            raise HTTPException(409, "同名文档的较新版本已发布，本次旧版抽取未覆盖事实")
        return result
    except ValueError as error:
        raise HTTPException(400, str(error)) from error


@app.post("/tasks", status_code=202)
def submit_extraction_task(file: Annotated[UploadFile, File()],
                           metrics: Annotated[str, Form()] = "") -> dict:
    try:
        return task_queue.submit(store, file.file.read(), file.filename or "", metrics)
    except ValueError as error:
        raise HTTPException(400, str(error)) from error


@app.get("/tasks/{task_id}")
def extraction_task(task_id: str) -> dict:
    if len(task_id) != 32 or any(char not in "0123456789abcdef" for char in task_id):
        raise HTTPException(404, "没有这条识别任务")
    task = task_queue.get(store, task_id)
    if task is None:
        raise HTTPException(404, "没有这条识别任务")
    return task


@app.post("/workspace/import", status_code=202)
def batch_import(files: Annotated[list[UploadFile], File()],
                 metrics: Annotated[str, Form()] = "") -> dict:
    if not 1 <= len(files) <= 8:
        raise HTTPException(400, "每批请选择 1 至 8 份资料")
    prepared = []
    for file in files:
        try:
            name = leaf_name(file.filename or "")
            content = file.file.read()
            if not content or len(content) > operations.MAX_FILE_BYTES:
                raise ValueError("文件为空或超过 12 MB")
            parse_document(content, name)
        except (ValueError, BadZipFile, InvalidFileException) as error:
            raise HTTPException(400, f"资料「{file.filename}」无法解析：{error}") from error
        prepared.append((name, content))
    if len({name for name, _ in prepared}) != len(prepared):
        raise HTTPException(400, "同批资料文件名不能重复，请先重命名")
    tasks = []
    for name, content in prepared:
        try:
            tasks.append({"doc": name, **task_queue.submit(store, content, name, metrics)})
        except ValueError as error:
            raise HTTPException(400, f"资料「{name}」提交失败：{error}；已提交任务可继续查询") from error
    return {"tasks": tasks}


@app.get("/workspace/sources")
def workspace_sources() -> dict:
    return {"sources": store.sources()}


@app.post("/workspace/templates/inspect")
def inspect_scoped_template(template: Annotated[UploadFile, File()]) -> dict:
    try:
        content = template.file.read()
        name = leaf_name(template.filename or "")
        return {"name": name, "template_sha256": sha256(content).hexdigest(),
                "targets": scoped_template.structure(content, name)}
    except (ValueError, BadZipFile, InvalidFileException) as error:
        raise HTTPException(400, str(error)) from error
    except Exception as error:
        logger.exception("跨文档模板结构预览失败")
        raise HTTPException(400, "无法读取 Word 或 Excel 模板，请检查文件格式") from error


@app.post("/workspace/templates/preview")
def preview_scoped_template(template: Annotated[UploadFile, File()],
                            docs: Annotated[str, Form()], target: Annotated[str, Form()],
                            entity_column: Annotated[int, Form()], mapping: Annotated[str, Form()]) -> dict:
    try:
        return scoped_template.preview(store, template.file.read(), leaf_name(template.filename or ""),
                                       json.loads(docs), target, entity_column, json.loads(mapping))
    except (ValueError, TypeError, BadZipFile, InvalidFileException) as error:
        raise HTTPException(400, f"跨文档填写预览失败：{error}") from error
    except Exception as error:
        logger.exception("跨文档填写计划预览失败")
        raise HTTPException(400, "无法生成跨文档填写计划，请核对目标表和映射") from error


@app.post("/workspace/templates/fill")
def fill_scoped_template(template: Annotated[UploadFile, File()], plan_id: Annotated[str, Form()]) -> dict:
    try:
        return scoped_template.fill(store, template.file.read(), leaf_name(template.filename or ""), plan_id)
    except (ValueError, BadZipFile, InvalidFileException) as error:
        raise HTTPException(409 if "版本" in str(error) or "失效" in str(error) else 400, str(error)) from error
    except Exception as error:
        logger.exception("跨文档模板填写失败")
        raise HTTPException(400, "模板处理失败，未生成结果副本") from error


@app.post("/operations/preview")
def preview_operations(file: Annotated[UploadFile, File()],
                       instruction: Annotated[str, Form()]) -> dict:
    try:
        return operations.preview(file.file.read(), leaf_name(file.filename or ""), instruction)
    except ValueError as error:
        raise HTTPException(400, str(error)) from error
    except Exception as error:
        logger.exception("操作预览失败")
        raise HTTPException(400, "无法解析文件或生成操作预览，请核对文件格式和指令") from error


@app.post("/operations")
def apply_operations(file: Annotated[UploadFile, File()],
                     plan_id: Annotated[str, Form()]) -> dict:
    try:
        return operations.execute(file.file.read(), leaf_name(file.filename or ""), plan_id)
    except ValueError as error:
        raise HTTPException(409 if "版本不符" in str(error) or "已失效" in str(error) else 400, str(error)) from error
    except Exception as error:
        logger.exception("文档操作失败")
        raise HTTPException(400, "文件处理失败，未生成结果副本，请核对目标位置和文件格式") from error


@app.get("/operations/files/{result_id}")
def operation_file(result_id: str) -> Response:
    result = operations.download(result_id)
    if result is None:
        raise HTTPException(404, "操作结果不存在或已过期，请重新执行")
    media_type = ("application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                  if result["name"].lower().endswith(".docx")
                  else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return Response(result["content"], media_type=media_type,
                    headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(result['name'])}"})


@app.post("/templates/inspect")
def inspect_template(source: Annotated[UploadFile, File()], template: Annotated[UploadFile, File()]) -> dict:
    try:
        return air_quality.inspect(source.file.read(), template.file.read())
    except ValueError as error:
        raise HTTPException(400, str(error)) from error
    except Exception as error:
        logger.exception("空气质量模板预览失败")
        raise HTTPException(400, "无法读取空气质量源表或 Word 模板，请检查文件格式") from error


@app.post("/templates/fill")
def fill_template(source: Annotated[UploadFile, File()], template: Annotated[UploadFile, File()],
                  plan_id: Annotated[str, Form()]) -> dict:
    try:
        return air_quality.fill(source.file.read(), template.file.read(), plan_id)
    except ValueError as error:
        raise HTTPException(409 if "版本" in str(error) or "失效" in str(error) else 400, str(error)) from error
    except Exception as error:
        logger.exception("空气质量模板填写失败")
        raise HTTPException(400, "模板处理失败，未生成下载副本") from error


@app.post("/templates/xlsx/inspect")
def inspect_xlsx_template(source: Annotated[UploadFile, File()], template: Annotated[UploadFile, File()]) -> dict:
    try:
        return xlsx_template.inspect(source.file.read(), leaf_name(source.filename or ""),
                                     template.file.read(), leaf_name(template.filename or ""))
    except ValueError as error:
        raise HTTPException(400, str(error)) from error
    except Exception as error:
        logger.exception("XLSX 模板预览失败")
        raise HTTPException(400, "无法读取源 XLSX 或目标模板") from error


@app.post("/templates/xlsx/fill")
def fill_xlsx_template(source: Annotated[UploadFile, File()], template: Annotated[UploadFile, File()],
                       plan_id: Annotated[str, Form()], sheet: Annotated[str, Form()],
                       mapping: Annotated[str, Form()]) -> dict:
    try:
        return xlsx_template.fill(source.file.read(), leaf_name(source.filename or ""),
                                  template.file.read(), leaf_name(template.filename or ""),
                                  plan_id, sheet, mapping)
    except ValueError as error:
        raise HTTPException(409 if "版本" in str(error) or "失效" in str(error) else 400, str(error)) from error
    except Exception as error:
        logger.exception("XLSX 模板填写失败")
        raise HTTPException(400, "XLSX 模板处理失败，未生成下载副本") from error


@app.get("/facts")
def facts(
    entity: str, metric: str, period: str = "", scope: str = "", preferred_doc: str | None = None,
) -> dict:
    return resolve_facts(store.facts(), entity, metric, period, scope, preferred_doc)


@app.post("/facts/decisions", status_code=201)
def record_fact_decision(request: FactDecisionRequest) -> dict:
    criteria = request.model_dump()
    return store.record_fact_decision(criteria, lambda candidates: resolve_facts(
        candidates, request.entity, request.metric, request.period, request.scope, request.preferred_doc,
    ))


@app.get("/facts/decisions/{decision_id}")
def fact_decision(decision_id: int) -> dict:
    record = store.fact_decision(decision_id)
    if record is None:
        raise HTTPException(404, "没有这条多源裁决记录")
    return record


@app.post("/fill")
def fill(request: FillRequest, response: Response) -> dict:
    try:
        doc = leaf_name(request.doc)
        with fill_lock:
            result = fill_workbook(store, doc, find_template(request.template), DATA, request.source_sha256,
                                   target_period=request.target_period,
                                   expected_publication_id=request.expected_publication_id)
            try:
                score_filled(doc)
            except (OSError, ValueError, BadZipFile, InvalidFileException, KeyError, StopIteration):
                logger.exception("填写已完成，自检报告生成失败")
                response.headers["X-A23-Score-Status"] = "failed"
        response.headers["X-A23-Result-ID"] = result.pop("result_id")
        response.headers["X-A23-Fill-Seconds"] = str(result.pop("fill_seconds"))
        publication_id = result.pop("publication_id")
        if publication_id is not None:
            response.headers["X-A23-Publication-ID"] = str(publication_id)
        return result
    except (BadZipFile, InvalidFileException, KeyError) as error:
        raise HTTPException(400, "无法读取 XLSX 模板，请检查文件内容") from error
    except (OSError, ValueError) as error:
        raise HTTPException(409 if "事实版本冲突" in str(error) else 400, str(error)) from error


def score_filled(doc: str) -> None:
    file = CITY_OUTPUT if doc.lower().endswith(".docx") else COVID_OUTPUT
    reference = REFERENCE_DIR / REFERENCE_FILES[file] if REFERENCE_DIR else None
    if reference is not None and reference.is_file():
        rows = None if SYNTHETIC_REFERENCE else checked_rules(load_expectations(), "same-db")[
            "03" if file == COVID_OUTPUT else "01"
        ]["rows"]
        evaluate(DATA / file, reference, DATA / REPORT_FILES[file][1], doc, file == COVID_OUTPUT,
                 expected_rows=rows)
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
                for key in ("doc", "result_id", "run_id", "workbook_sha256", "source_sha256",
                            "template_sha256", "publication_id")
            ):
                return report
    return None


@app.get("/reports")
def reports(doc: str, file: str = CITY_OUTPUT, result_id: str | None = None) -> dict:
    with fill_lock:
        score = matching_score(file, current_trace(doc, file, result_id))
    if score is None:
        raise HTTPException(404, "本次填写尚无匹配的自检报告")
    return score


@app.get("/cells")
def cell_detail(
    doc: str, file: str, row: Annotated[int, Query(ge=2)], column: Annotated[int, Query(ge=1)],
    result_id: str | None = None,
) -> dict:
    with fill_lock:
        fill_report = current_trace(doc, file, result_id)
        evaluation = matching_score(file, fill_report)
    if row > fill_report["rows"] + 1 or column > (5 if file == CITY_OUTPUT else 7):
        raise HTTPException(404, "结果中没有这个格子")
    entry = next(
        (item for item in fill_report["cells"] if item["row"] == row and item["column"] == column), None,
    )
    if entry is None:
        return {"decision": "missing", "reason": "这格没有填写记录"}
    graded = None
    if evaluation:
        diff = next((item for item in evaluation["diff"]
                     if item["row"] == row and item["column"] == column), None)
        graded = diff["decision"] if diff else "与本次核对依据一致"
    return {
        **entry, **read_cell_evidence(entry, fill_report["source_sha256"]), "evaluation": graded,
        **({"conflicts": [
            {**option, **read_cell_evidence(option, fill_report["source_sha256"])}
            for option in entry["conflicts"]
        ]} if "conflicts" in entry else {}),
    }


def read_cell_evidence(entry: dict, source_sha256: str) -> dict:
    source = entry.get("source")
    block = store.block(source["doc"], source["block_no"]) if source and source["block_no"] else None
    verified = (block is not None and block["document_digest"].rsplit(":", 1)[-1] == source_sha256
                and block_digest(block) == source.get("block_sha256"))
    return {"text": block["text"] if verified else None,
            "evidence_status": "stale" if source and not verified else "available" if verified else "none"}


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
    if result_id is not None or name in REPORT_FILES:
        with fill_lock:
            try:
                content = path.read_bytes()
            except FileNotFoundError as error:
                raise HTTPException(404, "结果文件不存在") from error
            if result_id is not None:
                current_trace(None, name, result_id, content)
        return Response(
            content, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(name)}"},
        )
    return FileResponse(path, filename=name)


FRONTEND_BUILT = (DIST / "index.html").is_file()
if FRONTEND_BUILT:
    app.mount("/", StaticFiles(directory=DIST, html=True), name="frontend")
else:
    @app.get("/")
    def index() -> dict:
        return {"detail": "前端尚未构建，接口说明位于 /docs"}
