"""Asynchronous extraction API with a model substitute and persistent status."""

from threading import Event
from time import monotonic, sleep

from fastapi.testclient import TestClient

import extractor
import main
import task_queue
from store import Store

SAMPLE = "2024年甲市 GDP 90 亿元。\n".encode()


def completed(client, task_id):
    deadline = monotonic() + 5
    while monotonic() < deadline:
        response = client.get(f"/tasks/{task_id}")
        assert response.status_code == 200
        if response.json()["status"] in {"succeeded", "superseded", "failed"}:
            return response.json()
        sleep(0.01)
    raise AssertionError("异步识别任务没有在 5 秒内结束")


def test_task_returns_id_before_model_finishes_then_saves_result(tmp_path, monkeypatch):
    database = Store(tmp_path / "tasks.db")
    monkeypatch.setattr(main, "store", database)
    entered, release = Event(), Event()

    def model_substitute(content, name, metrics, store, *, audit, publish):
        entered.set()
        assert release.wait(5)
        return extractor.extract_document(content, name, metrics, store, audit=audit, publish=publish,
                                          respond=lambda prompt: {"facts": [{
                                              "block_no": 1, "entity": "甲市", "metric": "GDP",
                                              "value": "90", "unit": "亿元", "snippet": "GDP 90 亿元",
                                          }]})

    monkeypatch.setattr(task_queue, "extract_document", model_substitute)
    client = TestClient(main.app)
    try:
        start = monotonic()
        submitted = client.post("/tasks", files={"file": ("合成.txt", SAMPLE)}, data={"metrics": "GDP"})
        assert submitted.status_code == 202 and monotonic() - start < 1
        task_id = submitted.json()["task_id"]
        assert entered.wait(2)
        running = client.get(f"/tasks/{task_id}").json()
        assert running["status"] == "running" and running["result"] is None
        assert running["started_at"] and running["finished_at"] is None
    finally:
        release.set()
    finished = completed(client, task_id)
    assert finished["status"] == "succeeded"
    assert finished["result"]["kept"] == 1 and finished["result"]["doc"] == "合成.txt"
    assert finished["result"]["classification"] == {
        "schema_version": 1, "method": "parser_structure_v1", "document_format": "txt",
        "content_family": "text", "block_types": {"text_line": 1},
        "extracted_value_types": {"number": 1, "other": 0}, "semantic_category": None,
    }
    assert finished["error"] is None and finished["finished_at"]
    assert database.facts("合成.txt")[0]["value"] == "90"


def test_task_failure_locates_batch_and_survives_query(tmp_path, monkeypatch):
    database = Store(tmp_path / "tasks.db")
    monkeypatch.setattr(main, "store", database)

    def failing_model(content, name, metrics, store, *, audit, publish):
        raise extractor.BatchFailure(1, [1], "模型返回的事实字段不符合约定")

    monkeypatch.setattr(task_queue, "extract_document", failing_model)
    client = TestClient(main.app)
    response = client.post("/tasks", files={"file": ("合成.txt", SAMPLE)})
    assert response.status_code == 202
    task_id = response.json()["task_id"]
    failed = completed(client, task_id)
    assert failed["status"] == "failed" and failed["result"] is None
    assert failed["error"] == {"code": "extraction_failed", "message": "模型返回的事实字段不符合约定",
                               "batch_index": 1, "blocks": [1]}
    assert task_queue.get(Store(database.path), task_id)["error"] == failed["error"]
    assert client.get("/tasks/invalid").status_code == 404
    assert client.get("/tasks/" + "0" * 32).status_code == 404


def test_invalid_upload_and_interrupted_task(tmp_path, monkeypatch):
    database = Store(tmp_path / "tasks.db")
    monkeypatch.setattr(main, "store", database)
    client = TestClient(main.app)
    assert client.post("/tasks", files={"file": ("bad.pdf", b"data")}).status_code == 400
    assert client.post("/tasks", files={"file": ("empty.txt", b"")}).status_code == 400
    with task_queue.lock:
        task_queue.prepare(database)
        with database.connect() as db:
            db.execute("""INSERT INTO extraction_tasks(id, doc, source_sha256, status, submitted_at)
                          VALUES (?, '合成.txt', 'hash', 'queued', ?)""", ("f" * 32, task_queue.now()))
        task_queue.initialized.remove(database.path.resolve())
    recovered = client.get("/tasks/" + "f" * 32).json()
    assert recovered["status"] == "failed"
    assert recovered["error"]["code"] == "interrupted"
