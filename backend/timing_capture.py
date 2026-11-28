"""Client-side timing record for asynchronous recognition and supported fixed fills."""

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


def main() -> None:
    parser = argparse.ArgumentParser(description="记录异步识别首响、排队、执行及可选固定模板下载耗时")
    parser.add_argument("source", type=Path)
    parser.add_argument("--base-url", default="http://127.0.0.1:8000")
    parser.add_argument("--metrics", default="")
    parser.add_argument("--template-name")
    parser.add_argument("--output-name")
    parser.add_argument("--record", type=Path, required=True)
    parser.add_argument("--timeout", type=float, default=300)
    args = parser.parse_args()
    with httpx.Client(base_url=args.base_url, timeout=120) as client:
        result = capture(client, args.source.read_bytes(), args.source.name, metrics=args.metrics,
                         template_name=args.template_name, output_name=args.output_name,
                         timeout=args.timeout)
    args.record.parent.mkdir(parents=True, exist_ok=True)
    args.record.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
