"""Client timing records for recognition and supported template workflows."""

import argparse
import json
from datetime import datetime, timezone
from hashlib import sha256
from pathlib import Path
from time import monotonic, sleep

import httpx


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
    record = {
        "schema_version": 1, "scope": "fixed_template_client" if template_name else "extraction_client",
        "source_name": source_name, "source_sha256": sha256(source).hexdigest(),
        "task_id": None, "status": "running", "task_status": None,
        "publication_id": None, "fill_publication_id": None,
        "client_started_at": started_at,
        "task_source_sha256": None, "current_source_sha256": None, "is_current": None,
        "first_response_seconds": None, "terminal_observed_seconds": None,
        "queue_seconds": None, "run_seconds": None,
        "fill_response_seconds": None, "fill_request_seconds": None,
        "download_complete_seconds": None, "download_request_seconds": None,
        "download": None, "error": None,
    }

    def failure(stage: str, response) -> dict:
        try:
            detail = response.json().get("detail", response.text)
        except ValueError:
            detail = response.text
        record["status"] = "failed"
        record["error"] = {"code": "http_error", "stage": stage,
                           "http_status": response.status_code, "detail": detail}
        return record

    def timed_out(stage: str) -> dict:
        record["status"] = "timed_out"
        record["error"] = {"code": "timeout", "stage": stage,
                           "elapsed_seconds": round(clock() - start, 4),
                           "detail": f"计时超过 {timeout} 秒，停止于{stage}阶段"}
        return record

    try:
        response = client.post("/tasks", files={"file": (source_name, source)},
                               data={"metrics": metrics})
    except httpx.TimeoutException:
        record["first_response_seconds"] = round(clock() - start, 4)
        return timed_out("submit")
    record["first_response_seconds"] = round(clock() - start, 4)
    if response.status_code != 202:
        return failure("submit", response)
    task_id = response.json()["task_id"]
    record["task_id"] = task_id
    while True:
        remaining = timeout - (clock() - start)
        if remaining <= 0:
            record["terminal_observed_seconds"] = round(clock() - start, 4)
            return timed_out("poll")
        try:
            status_response = client.get(f"/tasks/{task_id}", timeout=remaining)
        except httpx.TimeoutException:
            record["terminal_observed_seconds"] = round(clock() - start, 4)
            return timed_out("poll")
        record["terminal_observed_seconds"] = round(clock() - start, 4)
        if status_response.status_code != 200:
            return failure("poll", status_response)
        task = status_response.json()
        if task["status"] in {"succeeded", "superseded", "failed"}:
            break
        pause(min(poll_interval, max(0, timeout - (clock() - start))))
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
            filled = client.post("/fill", json={"doc": source_name, "template": template_name,
                                                 "source_sha256": record["source_sha256"],
                                                 "expected_publication_id": record["publication_id"]})
        except httpx.TimeoutException:
            record["fill_request_seconds"] = round(clock() - fill_start, 4)
            record["fill_response_seconds"] = round(clock() - start, 4)
            return timed_out("fill")
        record["fill_request_seconds"] = round(clock() - fill_start, 4)
        record["fill_response_seconds"] = round(clock() - start, 4)
        if filled.status_code == 409:
            record["status"] = "version_conflict"
            record["error"] = {"code": "version_conflict", "message": filled.json().get("detail", "事实版本冲突")}
            return record
        if filled.status_code != 200:
            return failure("fill", filled)
        record["fill_publication_id"] = int(filled.headers["X-A23-Publication-ID"])
        if record["fill_publication_id"] != record["publication_id"]:
            record["status"] = "version_conflict"
            record["error"] = {"code": "version_conflict", "stage": "fill",
                               "message": "填写结果与识别任务的事实批次不一致"}
            return record
        result_id = filled.headers["X-A23-Result-ID"]
        download_start = clock()
        try:
            downloaded = client.get("/files", params={"name": output_name, "result_id": result_id})
        except httpx.TimeoutException:
            record["download_request_seconds"] = round(clock() - download_start, 4)
            record["download_complete_seconds"] = round(clock() - start, 4)
            return timed_out("download")
        record["download_request_seconds"] = round(clock() - download_start, 4)
        record["download_complete_seconds"] = round(clock() - start, 4)
        if downloaded.status_code == 409:
            record["status"] = "version_conflict"
            record["error"] = {"code": "version_conflict",
                               "message": downloaded.json().get("detail", "填写结果版本冲突")}
            return record
        if downloaded.status_code != 200:
            return failure("download", downloaded)
        record["download"] = {"name": output_name, "result_id": result_id,
                              "publication_id": record["fill_publication_id"],
                              "bytes": len(downloaded.content),
                              "sha256": sha256(downloaded.content).hexdigest()}
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
    files = {"source": (source_name, source), "template": (template_name, template)}
    record = {
        "schema_version": 1, "scope": "uploaded_xlsx03_client", "client_started_at": started_at,
        "source_name": source_name, "source_sha256": sha256(source).hexdigest(),
        "template_name": template_name, "template_sha256": sha256(template).hexdigest(),
        "sheet": sheet, "mapping": mapping, "status": "running", "error": None,
        "timeout_seconds": timeout,
        "preview_response_seconds": None, "fill_complete_seconds": None,
        "fill_request_seconds": None, "download_complete_seconds": None,
        "download_request_seconds": None, "fill": None, "download": None,
    }

    def failure(stage: str, response) -> dict:
        try:
            detail = response.json().get("detail", response.text)
        except ValueError:
            detail = response.text
        record["status"] = "failed"
        record["error"] = {"stage": stage, "sheet": sheet,
                           "http_status": response.status_code, "detail": detail}
        return record

    def timed_out(stage: str) -> dict:
        elapsed = round(clock() - start, 4)
        record["status"] = "timed_out"
        record["error"] = {"code": "timeout", "stage": stage, "sheet": sheet,
                           "elapsed_seconds": elapsed,
                           "detail": f"上传模板计时超过 {timeout} 秒，停止于{stage}阶段"}
        return record

    def remaining() -> float:
        return timeout - (clock() - start)

    if remaining() <= 0:
        return timed_out("inspect")
    try:
        preview = client.post("/templates/xlsx/inspect", files=files, timeout=remaining())
    except httpx.TimeoutException:
        record["preview_response_seconds"] = round(clock() - start, 4)
        return timed_out("inspect")
    record["preview_response_seconds"] = round(clock() - start, 4)
    if remaining() <= 0:
        return timed_out("inspect")
    if preview.status_code != 200:
        return failure("inspect", preview)
    plan = preview.json()
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
    columns = {item["column"] for item in target["headers"]}
    if set(mapping) != columns:
        record["status"] = "failed"
        record["error"] = {"code": "mapping_mismatch", "stage": "selection", "sheet": sheet,
                           "detail": f"工作表「{sheet}」的列映射与预览表头不一致"}
        return record
    if remaining() <= 0:
        return timed_out("fill")
    fill_start = clock()
    try:
        filled = client.post("/templates/xlsx/fill", files=files, data={
            "plan_id": plan["plan_id"], "sheet": sheet,
            "mapping": json.dumps(mapping, ensure_ascii=False),
        }, timeout=remaining())
    except httpx.TimeoutException:
        record["fill_request_seconds"] = round(clock() - fill_start, 4)
        record["fill_complete_seconds"] = round(clock() - start, 4)
        return timed_out("fill")
    record["fill_request_seconds"] = round(clock() - fill_start, 4)
    record["fill_complete_seconds"] = round(clock() - start, 4)
    if remaining() <= 0:
        return timed_out("fill")
    if filled.status_code != 200:
        return failure("fill", filled)
    result = filled.json()
    record["fill"] = {key: result[key] for key in ("result_id", "rows", "cells_written", "cells_empty")}
    if remaining() <= 0:
        return timed_out("download")
    download_start = clock()

    def download_timing() -> None:
        record["download_request_seconds"] = round(clock() - download_start, 4)
        record["download_complete_seconds"] = round(clock() - start, 4)

    try:
        address = f"/operations/files/{result['result_id']}"
        if callable(getattr(client, "stream", None)):
            # A read timeout limits silence between chunks, not the total transfer.
            with client.stream("GET", address, timeout=remaining()) as downloaded:
                if downloaded.status_code != 200:
                    downloaded.read()
                    download_timing()
                    return timed_out("download") if remaining() <= 0 else failure("download", downloaded)
                chunks = []
                for chunk in downloaded.iter_bytes():
                    if remaining() <= 0:
                        download_timing()
                        return timed_out("download")
                    chunks.append(chunk)
                content = b"".join(chunks)
        else:
            # In-process model substitutes may only implement get().
            downloaded = client.get(address, timeout=remaining())
            if downloaded.status_code != 200:
                download_timing()
                return timed_out("download") if remaining() <= 0 else failure("download", downloaded)
            content = downloaded.content
    except httpx.TimeoutException:
        download_timing()
        return timed_out("download")
    download_timing()
    if remaining() <= 0:
        return timed_out("download")
    record["download"] = {"bytes": len(content), "sha256": sha256(content).hexdigest()}
    if download_to is not None:
        download_to.parent.mkdir(parents=True, exist_ok=True)
        download_to.write_bytes(content)
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
