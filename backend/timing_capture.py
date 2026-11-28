"""Client timing records for recognition and supported template workflows."""

import argparse
import json
from io import BytesIO
from queue import Empty, Queue
from datetime import datetime, timezone
from hashlib import sha256
from pathlib import Path
from threading import Event, Thread
from time import monotonic, sleep

import httpx
from openpyxl import load_workbook


class DeadlineExceeded(TimeoutError):
    """The workflow's monotonic deadline elapsed while waiting for I/O."""


def response_detail(response) -> object:
    """Keep error bodies readable even when a service sends an array or HTML."""
    try:
        body = response.json()
    except ValueError:
        return response.text
    return body.get("detail", body) if isinstance(body, dict) else body


def object_response(response, *fields: str) -> tuple[dict | None, str | None]:
    """Validate the shape before a timing workflow reads protocol fields."""
    try:
        body = response.json()
    except ValueError:
        return None, "响应不是 JSON 对象"
    if not isinstance(body, dict):
        return None, "响应不是 JSON 对象"
    missing = [field for field in fields if field not in body]
    if missing:
        return None, f"响应缺少字段：{', '.join(missing)}"
    return body, None


def valid_xlsx(content: bytes) -> bool:
    """Require an openable workbook before recording a download as successful."""
    try:
        book = load_workbook(BytesIO(content), read_only=True, data_only=False)
        book.close()
        return True
    except Exception:
        # openpyxl may raise several ZIP/XML errors for a corrupt workbook.
        return False


def before_deadline(action, deadline: float, clock=monotonic):
    """Bound a blocking HTTP operation by the workflow deadline, not an idle-read timeout."""
    remaining = deadline - clock()
    if remaining <= 0:
        raise DeadlineExceeded
    completed = Queue(maxsize=1)

    def run():
        try:
            completed.put((True, action(remaining)))
        except Exception as error:
            completed.put((False, error))

    Thread(target=run, daemon=True).start()
    try:
        success, result = completed.get(timeout=remaining)
    except Empty as error:
        raise DeadlineExceeded from error
    if clock() >= deadline:
        raise DeadlineExceeded
    if not success:
        raise result
    return result


def request_before_deadline(client, method: str, address: str, deadline: float,
                            clock=monotonic, **kwargs):
    """Read a whole HTTP response with cancellation between streamed chunks."""
    if not callable(getattr(client, "stream", None)) or (
        method != "GET" and not isinstance(client, httpx.Client)
    ):
        call = client.get if method == "GET" else client.post
        return before_deadline(lambda remaining: call(address, timeout=remaining, **kwargs), deadline, clock)

    events, acknowledgements = Queue(), Queue(maxsize=1)
    cancelled = Event()
    active = {}

    def run():
        try:
            with client.stream(method, address, timeout=max(deadline - clock(), 0.001),
                               **kwargs) as response:
                active["response"] = response
                events.put(("headers", (response.status_code, dict(getattr(response, "headers", {})))))
                stream = iter(response.iter_bytes())
                while not cancelled.is_set():
                    try:
                        chunk = next(stream)
                    except StopIteration:
                        break
                    events.put(("chunk", chunk))
                    while not cancelled.is_set():
                        try:
                            acknowledgements.get(timeout=0.01)
                            break
                        except Empty:
                            continue
                    if cancelled.is_set():
                        return
                if not cancelled.is_set():
                    events.put(("done", None))
        except Exception as error:
            if not cancelled.is_set():
                events.put(("error", error))

    def abort():
        cancelled.set()
        response = active.get("response")
        close = getattr(response, "close", None)
        if callable(close):
            # Closing a socket can block; do not extend the caller's deadline.
            def close_safely():
                try:
                    close()
                except Exception:
                    pass

            Thread(target=close_safely, daemon=True).start()

    Thread(target=run, daemon=True).start()
    chunks = []
    status = None
    headers = {}
    while True:
        remaining = deadline - clock()
        if remaining <= 0:
            abort()
            raise DeadlineExceeded
        try:
            kind, value = events.get(timeout=remaining)
        except Empty as error:
            abort()
            raise DeadlineExceeded from error
        if clock() >= deadline:
            abort()
            raise DeadlineExceeded
        if kind == "error":
            raise value
        if kind == "headers":
            status, headers = value
        if kind == "chunk":
            chunks.append(value)
            acknowledgements.put_nowait(True)
        if kind == "done":
            return httpx.Response(status, headers=headers, content=b"".join(chunks))


def download_before_deadline(client, address: str, deadline: float, clock=monotonic,
                             params: dict | None = None) -> tuple[int, object]:
    """Read a download under the same deadline and retain failed HTTP details."""
    response = request_before_deadline(client, "GET", address, deadline, clock, params=params)
    return response.status_code, response.content if response.status_code == 200 else response


def seconds_between(start: str | None, end: str | None) -> float | None:
    if not start or not end:
        return None
    return round((datetime.fromisoformat(end) - datetime.fromisoformat(start)).total_seconds(), 4)


def capture(client, source: bytes, source_name: str, *, metrics: str = "",
            template_name: str | None = None, output_name: str | None = None,
            timeout: float = 300, poll_interval: float = 0.2,
            clock=monotonic, pause=sleep) -> dict:
    """Measure first response, result observation, and optional workbook download."""
    if bool(template_name) != bool(output_name):
        raise ValueError("计时需同时提供模板名和输出文件名")
    if timeout <= 0:
        raise ValueError("识别计时的 --timeout 必须大于 0 秒")
    started_at = datetime.now(timezone.utc).isoformat(timespec="milliseconds")
    start = clock()
    deadline = start + timeout
    record = {
        "schema_version": 1, "scope": "fixed_template_client" if template_name else "extraction_client",
        "source_name": source_name, "source_sha256": sha256(source).hexdigest(),
        "task_id": None, "status": "running", "task_status": None,
        "publication_id": None, "fill_publication_id": None,
        "client_started_at": started_at,
        "task_source_sha256": None, "current_source_sha256": None, "is_current": None,
        "submit_request_seconds": None, "first_response_seconds": None,
        "poll_elapsed_seconds": None, "terminal_observed_seconds": None,
        "queue_seconds": None, "run_seconds": None,
        "fill_response_seconds": None, "fill_request_seconds": None,
        "download_complete_seconds": None, "download_request_seconds": None,
        "download": None, "error": None,
    }

    def failure(stage: str, response) -> dict:
        record["status"] = "failed"
        record["error"] = {"code": "http_error", "stage": stage,
                           "http_status": response.status_code, "detail": response_detail(response),
                           "elapsed_seconds": round(clock() - start, 4)}
        return record

    def invalid(stage: str, detail: str) -> dict:
        record["status"] = "failed"
        record["error"] = {"code": "invalid_response", "stage": stage, "detail": detail,
                           "elapsed_seconds": round(clock() - start, 4)}
        return record

    def timed_out(stage: str) -> dict:
        record["status"] = "timed_out"
        record["error"] = {"code": "timeout", "stage": stage,
                           "elapsed_seconds": round(clock() - start, 4),
                           "detail": f"计时超过 {timeout} 秒，停止于{stage}阶段"}
        return record

    def disconnected(stage: str, error: httpx.RequestError) -> dict:
        record["status"] = "failed"
        record["error"] = {"code": "transport_error", "stage": stage,
                           "elapsed_seconds": round(clock() - start, 4),
                           "detail": f"HTTP 连接中断：{error}"}
        return record

    try:
        response = request_before_deadline(client, "POST", "/tasks", deadline, clock,
                                           files={"file": (source_name, source)},
                                           data={"metrics": metrics})
    except (DeadlineExceeded, httpx.TimeoutException):
        record["submit_request_seconds"] = round(clock() - start, 4)
        return timed_out("submit")
    except httpx.RequestError as error:
        record["submit_request_seconds"] = round(clock() - start, 4)
        return disconnected("submit", error)
    record["submit_request_seconds"] = record["first_response_seconds"] = round(clock() - start, 4)
    if response.status_code != 202:
        return failure("submit", response)
    submitted, problem = object_response(response, "task_id")
    if problem or not isinstance(submitted["task_id"], str) or not submitted["task_id"].strip():
        return invalid("submit", problem or "任务 ID 必须是非空字符串")
    task_id = submitted["task_id"]
    record["task_id"] = task_id
    while True:
        remaining = deadline - clock()
        if remaining <= 0:
            record["poll_elapsed_seconds"] = round(clock() - start, 4)
            return timed_out("poll")
        try:
            status_response = request_before_deadline(client, "GET", f"/tasks/{task_id}",
                                                      deadline, clock)
        except (DeadlineExceeded, httpx.TimeoutException):
            record["poll_elapsed_seconds"] = round(clock() - start, 4)
            return timed_out("poll")
        except httpx.RequestError as error:
            record["poll_elapsed_seconds"] = round(clock() - start, 4)
            return disconnected("poll", error)
        record["poll_elapsed_seconds"] = round(clock() - start, 4)
        if status_response.status_code != 200:
            return failure("poll", status_response)
        task, problem = object_response(status_response, "status")
        if (problem or not isinstance(task["status"], str)
            or task["status"] not in {"queued", "running", "succeeded", "superseded", "failed"}):
            return invalid("poll", problem or "任务状态字段无效")
        if task["status"] in {"succeeded", "superseded", "failed"}:
            required = ("publication_id", "source_sha256", "current_source_sha256", "is_current",
                        "submitted_at", "started_at", "finished_at", "error")
            task, problem = object_response(status_response, *required)
            if problem:
                return invalid("poll", problem)
            if (not isinstance(task["source_sha256"], str)
                or not isinstance(task["is_current"], bool)
                or (task["current_source_sha256"] is not None
                    and not isinstance(task["current_source_sha256"], str))
                or (task["publication_id"] is not None and type(task["publication_id"]) is not int)
                or (task["error"] is not None and not isinstance(task["error"], dict))
                or any(value is not None and not isinstance(value, str)
                       for value in (task["submitted_at"], task["started_at"], task["finished_at"]))):
                return invalid("poll", "任务版本或时间字段类型无效")
            try:
                seconds_between(task["submitted_at"], task["started_at"])
                seconds_between(task["started_at"], task["finished_at"])
            except ValueError:
                return invalid("poll", "任务时间字段格式无效")
            record["terminal_observed_seconds"] = record["poll_elapsed_seconds"]
            break
        pause(min(poll_interval, max(0, deadline - clock())))
    record.update({
        "status": task["status"], "task_status": task["status"],
        "publication_id": task["publication_id"],
        "task_source_sha256": task["source_sha256"],
        "current_source_sha256": task["current_source_sha256"],
        "is_current": task["is_current"],
        "queue_seconds": seconds_between(task["submitted_at"], task["started_at"]),
        "run_seconds": seconds_between(task["started_at"], task["finished_at"]),
        "error": task["error"],
    })
    if record["task_source_sha256"] != record["source_sha256"]:
        record["status"] = "failed"
        record["error"] = {"code": "source_mismatch", "stage": "poll",
                           "detail": f"任务 {task_id} 的原文摘要与提交文件不一致"}
        return record
    if task["status"] == "failed" and isinstance(task["error"], dict):
        record["error"] = {"stage": "extraction", **task["error"]}
    if template_name and task["status"] == "succeeded" and not task["is_current"]:
        record["status"] = "version_conflict"
        record["error"] = {"code": "version_conflict", "stage": "fill",
                           "message": "识别任务成功后，事实批次已被替换，请重新识别并填写"}
        return record
    if template_name and task["status"] == "superseded":
        record["status"] = "version_conflict"
        record["error"] = {"code": "version_conflict", "stage": "fill",
                           "message": "识别任务已被较新版本取代，请重新识别并填写"}
        return record
    if template_name and task["status"] == "succeeded":
        fill_start = clock()
        try:
            filled = request_before_deadline(client, "POST", "/fill", deadline, clock,
                json={
                    "doc": source_name, "template": template_name,
                    "source_sha256": record["source_sha256"],
                    "expected_publication_id": record["publication_id"],
                })
        except (DeadlineExceeded, httpx.TimeoutException):
            record["fill_request_seconds"] = round(clock() - fill_start, 4)
            return timed_out("fill")
        except httpx.RequestError as error:
            record["fill_request_seconds"] = round(clock() - fill_start, 4)
            return disconnected("fill", error)
        record["fill_request_seconds"] = round(clock() - fill_start, 4)
        record["fill_response_seconds"] = round(clock() - start, 4)
        if filled.status_code == 409:
            record["status"] = "version_conflict"
            record["error"] = {"code": "version_conflict", "stage": "fill",
                               "message": response_detail(filled)}
            return record
        if filled.status_code != 200:
            return failure("fill", filled)
        fill_body, problem = object_response(filled, "rows", "cells_written", "cells_empty")
        if (problem or any(type(fill_body[field]) is not int or fill_body[field] < 0
                           for field in ("rows", "cells_written", "cells_empty"))):
            return invalid("fill", problem or "填写响应的数量字段无效")
        publication = filled.headers.get("X-A23-Publication-ID")
        result_id = filled.headers.get("X-A23-Result-ID")
        if not publication or not publication.isdecimal() or not result_id or not result_id.strip():
            return invalid("fill", "填写响应缺少有效发布号或结果版本头")
        record["fill_publication_id"] = int(publication)
        if record["fill_publication_id"] != record["publication_id"]:
            record["status"] = "version_conflict"
            record["error"] = {"code": "version_conflict", "stage": "fill",
                               "message": "填写结果与识别任务的事实批次不一致"}
            return record
        download_start = clock()
        try:
            status, payload = download_before_deadline(
                client, "/files", deadline, clock,
                params={"name": output_name, "result_id": result_id})
        except (DeadlineExceeded, httpx.TimeoutException):
            record["download_request_seconds"] = round(clock() - download_start, 4)
            return timed_out("download")
        except httpx.RequestError as error:
            record["download_request_seconds"] = round(clock() - download_start, 4)
            return disconnected("download", error)
        record["download_request_seconds"] = round(clock() - download_start, 4)
        if status == 409:
            record["status"] = "version_conflict"
            record["error"] = {"code": "version_conflict", "stage": "download",
                               "message": response_detail(payload)}
            return record
        if status != 200:
            return failure("download", payload)
        if not valid_xlsx(payload):
            return invalid("download", "下载内容不是有效 XLSX 工作簿")
        record["download_complete_seconds"] = round(clock() - start, 4)
        record["download"] = {"name": output_name, "result_id": result_id,
                              "publication_id": record["fill_publication_id"],
                              "bytes": len(payload),
                              "sha256": sha256(payload).hexdigest()}
    return record


def capture_uploaded_xlsx(client, source: bytes, source_name: str, template: bytes,
                          template_name: str, sheet: str, mapping: dict[str, str], *,
                          download_to: Path | None = None, timeout: float = 300,
                          clock=monotonic) -> dict:
    """Time a confirmed 03-style template from upload preview through download."""
    if timeout <= 0:
        raise ValueError("上传模板计时的 --timeout 必须大于 0 秒")
    started_at = datetime.now(timezone.utc).isoformat(timespec="milliseconds")
    start = clock()
    deadline = start + timeout
    files = {"source": (source_name, source), "template": (template_name, template)}
    record = {
        "schema_version": 1, "scope": "uploaded_xlsx03_client", "client_started_at": started_at,
        "source_name": source_name, "source_sha256": sha256(source).hexdigest(),
        "template_name": template_name, "template_sha256": sha256(template).hexdigest(),
        "sheet": sheet, "mapping": mapping, "status": "running", "error": None,
        "timeout_seconds": timeout,
        "preview_request_seconds": None, "preview_response_seconds": None,
        "fill_complete_seconds": None,
        "fill_request_seconds": None, "download_complete_seconds": None,
        "download_request_seconds": None, "fill": None, "download": None,
    }

    def failure(stage: str, response) -> dict:
        record["status"] = "failed"
        record["error"] = {"stage": stage, "sheet": sheet,
                           "http_status": response.status_code, "detail": response_detail(response),
                           "elapsed_seconds": round(clock() - start, 4)}
        return record

    def invalid(stage: str, detail: str) -> dict:
        record["status"] = "failed"
        record["error"] = {"code": "invalid_response", "stage": stage, "sheet": sheet,
                           "detail": detail, "elapsed_seconds": round(clock() - start, 4)}
        return record

    def timed_out(stage: str) -> dict:
        elapsed = round(clock() - start, 4)
        record["status"] = "timed_out"
        record["error"] = {"code": "timeout", "stage": stage, "sheet": sheet,
                           "elapsed_seconds": elapsed,
                           "detail": f"上传模板计时超过 {timeout} 秒，停止于{stage}阶段"}
        return record

    def disconnected(stage: str, error: httpx.RequestError) -> dict:
        record["status"] = "failed"
        record["error"] = {"code": "transport_error", "stage": stage, "sheet": sheet,
                           "elapsed_seconds": round(clock() - start, 4),
                           "detail": f"HTTP 连接中断：{error}"}
        return record

    def remaining() -> float:
        return deadline - clock()

    if remaining() <= 0:
        return timed_out("inspect")
    try:
        preview = request_before_deadline(client, "POST", "/templates/xlsx/inspect",
                                          deadline, clock, files=files)
    except (DeadlineExceeded, httpx.TimeoutException):
        record["preview_request_seconds"] = round(clock() - start, 4)
        return timed_out("inspect")
    except httpx.RequestError as error:
        record["preview_request_seconds"] = round(clock() - start, 4)
        return disconnected("inspect", error)
    record["preview_request_seconds"] = record["preview_response_seconds"] = round(clock() - start, 4)
    if remaining() <= 0:
        return timed_out("inspect")
    if preview.status_code != 200:
        return failure("inspect", preview)
    plan, problem = object_response(preview, "plan_id", "target_sheets")
    if (problem or not isinstance(plan["plan_id"], str) or not plan["plan_id"].strip()
        or not isinstance(plan["target_sheets"], list)
        or any(not isinstance(item, dict) or not isinstance(item.get("name"), str)
               for item in plan["target_sheets"])):
        return invalid("inspect", problem or "预览计划或工作表列表无效")
    target = next((item for item in plan["target_sheets"] if item["name"] == sheet), None)
    if target is None:
        record["status"] = "failed"
        record["error"] = {"code": "sheet_not_found", "stage": "selection", "sheet": sheet,
                           "detail": f"目标模板没有工作表「{sheet}」"}
        return record
    if target.get("blocked_reason"):
        record["status"] = "failed"
        record["error"] = {"code": "blocked_sheet", "stage": "selection", "sheet": sheet,
                           "detail": target["blocked_reason"]}
        return record
    headers = target.get("headers")
    if not isinstance(headers, list) or any(
        not isinstance(item, dict) or not isinstance(item.get("column"), str) for item in headers
    ):
        return invalid("inspect", f"工作表「{sheet}」的表头结构无效")
    columns = {item["column"] for item in headers}
    if set(mapping) != columns:
        record["status"] = "failed"
        record["error"] = {"code": "mapping_mismatch", "stage": "selection", "sheet": sheet,
                           "detail": f"工作表「{sheet}」的列映射与预览表头不一致"}
        return record
    if remaining() <= 0:
        return timed_out("fill")
    fill_start = clock()
    try:
        filled = request_before_deadline(client, "POST", "/templates/xlsx/fill",
            deadline, clock, files=files, data={
                "plan_id": plan["plan_id"], "sheet": sheet,
                "mapping": json.dumps(mapping, ensure_ascii=False),
            })
    except (DeadlineExceeded, httpx.TimeoutException):
        record["fill_request_seconds"] = round(clock() - fill_start, 4)
        return timed_out("fill")
    except httpx.RequestError as error:
        record["fill_request_seconds"] = round(clock() - fill_start, 4)
        return disconnected("fill", error)
    record["fill_request_seconds"] = round(clock() - fill_start, 4)
    if remaining() <= 0:
        return timed_out("fill")
    if filled.status_code != 200:
        return failure("fill", filled)
    record["fill_complete_seconds"] = round(clock() - start, 4)
    result, problem = object_response(filled, "result_id", "rows", "cells_written", "cells_empty")
    if (problem or not isinstance(result["result_id"], str) or not result["result_id"].strip()
        or any(type(result[field]) is not int or result[field] < 0
               for field in ("rows", "cells_written", "cells_empty"))):
        return invalid("fill", problem or "填写响应的结果版本或数量字段无效")
    record["fill"] = {key: result[key] for key in ("result_id", "rows", "cells_written", "cells_empty")}
    if remaining() <= 0:
        return timed_out("download")
    download_start = clock()

    def download_timing() -> None:
        record["download_request_seconds"] = round(clock() - download_start, 4)

    try:
        status, payload = download_before_deadline(
            client, f"/operations/files/{result['result_id']}", deadline, clock)
    except (DeadlineExceeded, httpx.TimeoutException):
        download_timing()
        return timed_out("download")
    except httpx.RequestError as error:
        download_timing()
        return disconnected("download", error)
    download_timing()
    if remaining() <= 0:
        return timed_out("download")
    if status != 200:
        return failure("download", payload)
    if not valid_xlsx(payload):
        return invalid("download", "下载内容不是有效 XLSX 工作簿")
    record["download_complete_seconds"] = round(clock() - start, 4)
    record["download"] = {"bytes": len(payload), "sha256": sha256(payload).hexdigest()}
    if download_to is not None:
        download_to.parent.mkdir(parents=True, exist_ok=True)
        download_to.write_bytes(payload)
        record["download"]["path"] = str(download_to)
    record["status"] = "succeeded"
    return record


def main() -> int:
    parser = argparse.ArgumentParser(description="记录异步识别或受支持模板上传至下载的客户端耗时")
    parser.add_argument("source", type=Path)
    parser.add_argument("--base-url", default="http://127.0.0.1:8000")
    parser.add_argument("--metrics", default="")
    parser.add_argument("--template-name")
    parser.add_argument("--output-name")
    parser.add_argument("--uploaded-xlsx-template", type=Path)
    parser.add_argument("--template-sheet")
    parser.add_argument("--mapping-json", type=Path)
    parser.add_argument("--download-to", type=Path)
    parser.add_argument("--record", type=Path, required=True)
    parser.add_argument("--timeout", type=float, default=300)
    args = parser.parse_args()
    uploaded = args.uploaded_xlsx_template is not None
    if uploaded and (not args.template_sheet or not args.mapping_json or args.template_name or args.output_name):
        parser.error("上传 XLSX 计时需提供 --template-sheet 和 --mapping-json，且不能混用固定模板参数")
    if not uploaded and (args.template_sheet or args.mapping_json or args.download_to):
        parser.error("--template-sheet、--mapping-json 和 --download-to 只用于上传 XLSX 计时")
    with httpx.Client(base_url=args.base_url, timeout=120) as client:
        if uploaded:
            mapping = json.loads(args.mapping_json.read_text(encoding="utf-8"))
            if not isinstance(mapping, dict) or not all(
                isinstance(key, str) and isinstance(value, str) for key, value in mapping.items()
            ):
                parser.error("映射 JSON 须为目标列字母到字段键的字符串对象")
            result = capture_uploaded_xlsx(
                client, args.source.read_bytes(), args.source.name,
                args.uploaded_xlsx_template.read_bytes(), args.uploaded_xlsx_template.name,
                args.template_sheet, mapping, download_to=args.download_to,
                timeout=args.timeout,
            )
        else:
            result = capture(client, args.source.read_bytes(), args.source.name, metrics=args.metrics,
                             template_name=args.template_name, output_name=args.output_name,
                             timeout=args.timeout)
    args.record.parent.mkdir(parents=True, exist_ok=True)
    args.record.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["status"] == "succeeded" else 2 if result["status"] == "timed_out" else 1


if __name__ == "__main__":
    raise SystemExit(main())
