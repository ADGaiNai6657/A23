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
    started_at = datetime.now(timezone.utc).isoformat(timespec="milliseconds")
    start = clock()
    response = client.post("/tasks", files={"file": (source_name, source)}, data={"metrics": metrics})
    first_response = round(clock() - start, 4)
    response.raise_for_status()
    task_id = response.json()["task_id"]
    while True:
        if clock() - start > timeout:
            raise TimeoutError(f"任务 {task_id} 在 {timeout} 秒内未返回终态")
        status_response = client.get(f"/tasks/{task_id}")
        status_response.raise_for_status()
        task = status_response.json()
        if task["status"] in {"succeeded", "superseded", "failed"}:
            break
        pause(poll_interval)
    observed = round(clock() - start, 4)
    record = {
        "schema_version": 1, "scope": "fixed_template_client" if template_name else "extraction_client",
        "source_name": source_name, "source_sha256": sha256(source).hexdigest(),
        "task_id": task_id, "status": task["status"], "client_started_at": started_at,
        "task_source_sha256": task["source_sha256"],
        "current_source_sha256": task["current_source_sha256"],
        "is_current": task["is_current"],
        "first_response_seconds": first_response, "terminal_observed_seconds": observed,
        "queue_seconds": seconds_between(task["submitted_at"], task["started_at"]),
        "run_seconds": seconds_between(task["started_at"], task["finished_at"]),
        "fill_response_seconds": None, "fill_request_seconds": None,
        "download_complete_seconds": None, "download_request_seconds": None,
        "download": None, "error": task["error"],
    }
    if record["task_source_sha256"] != record["source_sha256"]:
        raise ValueError(f"任务 {task_id} 的原文摘要与提交文件不一致")
    if template_name and task["status"] == "succeeded" and task["is_current"]:
        fill_start = clock()
        filled = client.post("/fill", json={"doc": source_name, "template": template_name,
                                             "source_sha256": record["source_sha256"]})
        record["fill_request_seconds"] = round(clock() - fill_start, 4)
        record["fill_response_seconds"] = round(clock() - start, 4)
        filled.raise_for_status()
        result_id = filled.headers["X-A23-Result-ID"]
        download_start = clock()
        downloaded = client.get("/files", params={"name": output_name, "result_id": result_id})
        record["download_request_seconds"] = round(clock() - download_start, 4)
        record["download_complete_seconds"] = round(clock() - start, 4)
        downloaded.raise_for_status()
        record["download"] = {"name": output_name, "result_id": result_id,
                              "bytes": len(downloaded.content),
                              "sha256": sha256(downloaded.content).hexdigest()}
    return record


def capture_uploaded_xlsx(client, source: bytes, source_name: str, template: bytes,
                          template_name: str, sheet: str, mapping: dict[str, str], *,
                          download_to: Path | None = None, clock=monotonic) -> dict:
    """Time a confirmed 03-style template from upload preview through download."""
    started_at = datetime.now(timezone.utc).isoformat(timespec="milliseconds")
    start = clock()
    files = {"source": (source_name, source), "template": (template_name, template)}
    record = {
        "schema_version": 1, "scope": "uploaded_xlsx03_client", "client_started_at": started_at,
        "source_name": source_name, "source_sha256": sha256(source).hexdigest(),
        "template_name": template_name, "template_sha256": sha256(template).hexdigest(),
        "sheet": sheet, "mapping": mapping, "status": "running", "error": None,
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
        record["error"] = {"stage": stage, "http_status": response.status_code, "detail": detail}
        return record

    preview = client.post("/templates/xlsx/inspect", files=files)
    record["preview_response_seconds"] = round(clock() - start, 4)
    if preview.status_code != 200:
        return failure("inspect", preview)
    plan = preview.json()
    target = next((item for item in plan["target_sheets"] if item["name"] == sheet), None)
    if target is None:
        raise ValueError(f"目标模板没有工作表「{sheet}」")
    if target.get("blocked_reason"):
        raise ValueError(target["blocked_reason"])
    columns = {item["column"] for item in target["headers"]}
    if set(mapping) != columns:
        raise ValueError(f"工作表「{sheet}」的列映射与预览表头不一致")
    fill_start = clock()
    filled = client.post("/templates/xlsx/fill", files=files, data={
        "plan_id": plan["plan_id"], "sheet": sheet,
        "mapping": json.dumps(mapping, ensure_ascii=False),
    })
    record["fill_request_seconds"] = round(clock() - fill_start, 4)
    record["fill_complete_seconds"] = round(clock() - start, 4)
    if filled.status_code != 200:
        return failure("fill", filled)
    result = filled.json()
    record["fill"] = {key: result[key] for key in ("result_id", "rows", "cells_written", "cells_empty")}
    download_start = clock()
    downloaded = client.get(f"/operations/files/{result['result_id']}")
    record["download_request_seconds"] = round(clock() - download_start, 4)
    record["download_complete_seconds"] = round(clock() - start, 4)
    if downloaded.status_code != 200:
        return failure("download", downloaded)
    record["download"] = {"bytes": len(downloaded.content),
                          "sha256": sha256(downloaded.content).hexdigest()}
    if download_to is not None:
        download_to.parent.mkdir(parents=True, exist_ok=True)
        download_to.write_bytes(downloaded.content)
        record["download"]["path"] = str(download_to)
    record["status"] = "succeeded"
    return record


def main() -> None:
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
            )
        else:
            result = capture(client, args.source.read_bytes(), args.source.name, metrics=args.metrics,
                             template_name=args.template_name, output_name=args.output_name,
                             timeout=args.timeout)
    args.record.parent.mkdir(parents=True, exist_ok=True)
    args.record.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
