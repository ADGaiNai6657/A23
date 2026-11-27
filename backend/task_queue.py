"""Single-process asynchronous extraction with queryable SQLite task status."""

import json
import logging
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from hashlib import sha256
from pathlib import Path
from threading import Lock
from uuid import uuid4

from extractor import extract_document
from operations import MAX_FILE_BYTES
from parsers import leaf_name
from store import Store

logger = logging.getLogger(__name__)
executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix="a23-extract")
lock = Lock()
initialized: set[Path] = set()
MAX_ACTIVE = 16


def now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds")


def prepare(store: Store) -> None:
    """Mark tasks interrupted by a previous process as failed on first access."""
    path = store.path.resolve()
    if path in initialized:
        return
    with store.connect() as db:
        db.execute("""
            CREATE TABLE IF NOT EXISTS extraction_tasks (
                id TEXT PRIMARY KEY, doc TEXT NOT NULL, source_sha256 TEXT NOT NULL,
                status TEXT NOT NULL, submitted_at TEXT NOT NULL, started_at TEXT,
                finished_at TEXT, result TEXT, error TEXT
            )
        """)
        db.execute("""
            UPDATE extraction_tasks SET status = 'failed', finished_at = ?,
                error = ? WHERE status IN ('queued', 'running')
        """, (now(), json.dumps({"code": "interrupted", "message": "服务重启中断了任务，请重新提交"},
                             ensure_ascii=False)))
    initialized.add(path)


def submit(store: Store, content: bytes, filename: str, metrics: str) -> dict:
    name = leaf_name(filename)
    if not name.lower().endswith((".txt", ".md", ".docx", ".xlsx")):
        raise ValueError("异步识别仅支持 TXT、Markdown、DOCX 和 XLSX")
    if not content or len(content) > MAX_FILE_BYTES:
        raise ValueError("识别文件为空或超过 12 MB")
    chosen = [part.strip() for part in metrics.split(",") if part.strip()]
    task_id, submitted = uuid4().hex, now()
    with lock:
        prepare(store)
        with store.connect() as db:
            active = db.execute("SELECT COUNT(*) FROM extraction_tasks WHERE status IN ('queued', 'running')").fetchone()[0]
            if active >= MAX_ACTIVE:
                raise ValueError("待识别任务已达 16 份上限，请稍后重试")
            db.execute("""
                INSERT INTO extraction_tasks(id, doc, source_sha256, status, submitted_at)
                VALUES (?, ?, ?, 'queued', ?)
            """, (task_id, name, sha256(content).hexdigest(), submitted))
        try:
            executor.submit(run, store, task_id, content, name, chosen)
        except RuntimeError as error:
            finish(store, task_id, "failed", error={"code": "queue_unavailable", "message": "任务队列不可用，请重试"})
            raise ValueError("任务队列不可用，请重试") from error
    return {"task_id": task_id, "status": "queued", "submitted_at": submitted}


def finish(store: Store, task_id: str, status: str, *, result: dict | None = None,
           error: dict | None = None) -> None:
    with store.connect() as db:
        db.execute("""
            UPDATE extraction_tasks SET status = ?, finished_at = ?, result = ?, error = ? WHERE id = ?
        """, (status, now(), json.dumps(result, ensure_ascii=False) if result is not None else None,
              json.dumps(error, ensure_ascii=False) if error is not None else None, task_id))


def run(store: Store, task_id: str, content: bytes, name: str, metrics: list[str]) -> None:
    audit: dict = {}
    try:
        with store.connect() as db:
            db.execute("UPDATE extraction_tasks SET status = 'running', started_at = ? WHERE id = ?",
                       (now(), task_id))
        result = extract_document(content, name, metrics, store, audit=audit)
        finish(store, task_id, "succeeded", result=result)
    except Exception as exc:
        logger.exception("异步识别失败：task_id=%s doc=%s", task_id, name)
        failed_batch = next((index for index, batch in enumerate(audit.get("batches", []), 1)
                             if any("error" in attempt for attempt in batch.get("attempts", []))), None)
        detail = {"code": "extraction_failed", "message": str(exc) if isinstance(exc, ValueError)
                  else "识别任务处理失败，请查看服务日志"}
        if failed_batch is not None:
            detail["batch_index"] = failed_batch
            detail["blocks"] = audit["batches"][failed_batch - 1]["blocks"]
        finish(store, task_id, "failed", error=detail)


def get(store: Store, task_id: str) -> dict | None:
    with lock:
        prepare(store)
    with store.connect() as db:
        row = db.execute("SELECT * FROM extraction_tasks WHERE id = ?", (task_id,)).fetchone()
    if row is None:
        return None
    return {"task_id": row["id"], "doc": row["doc"], "source_sha256": row["source_sha256"],
            "status": row["status"], "submitted_at": row["submitted_at"],
            "started_at": row["started_at"], "finished_at": row["finished_at"],
            "result": json.loads(row["result"]) if row["result"] else None,
            "error": json.loads(row["error"]) if row["error"] else None}
