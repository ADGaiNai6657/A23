"""Version ordering, failure location, and crash consistency for extraction tasks."""

import os
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor
from hashlib import sha256
from pathlib import Path
from threading import Event

from fastapi.testclient import TestClient

import extractor
import main
import task_queue
from store import Store
from test_tasks import completed

NAME = "同名.txt"


def sample(value):
    return f"2024年甲市 GDP {value} 亿元。\n".encode()


def substitute(content, name, metrics, store, *, audit=None, publish=None):
    value = "120" if b"120" in content else "90" if b"90" in content else "55"
    return extractor.extract_document(content, name, metrics, store, audit=audit, publish=publish,
                                      respond=lambda _: {"facts": [{
                                          "block_no": 1, "entity": "甲市", "metric": "GDP",
                                          "value": value, "unit": "亿元", "snippet": f"GDP {value} 亿元",
                                      }]})


def test_old_async_result_cannot_replace_newer_success(tmp_path, monkeypatch):
    database = Store(tmp_path / "versions.db")
    monkeypatch.setattr(main, "store", database)
    waiting, release = Event(), Event()

    def delayed(content, name, metrics, store, *, audit, publish):
        if b"90" in content:
            waiting.set()
            assert release.wait(5)
        return substitute(content, name, metrics, store, audit=audit, publish=publish)

    monkeypatch.setattr(task_queue, "extract_document", delayed)
    client = TestClient(main.app)
    old = client.post("/tasks", files={"file": (NAME, sample(90))}).json()["task_id"]
    assert waiting.wait(2)
    new = client.post("/tasks", files={"file": (NAME, sample(120))}).json()["task_id"]
    current = completed(client, new)
    assert current["status"] == "succeeded" and current["is_current"]
    assert current["source_sha256"] == current["current_source_sha256"] == sha256(sample(120)).hexdigest()
    release.set()
    older = completed(client, old)
    assert older["status"] == "superseded" and not older["is_current"]
    assert older["source_sha256"] == sha256(sample(90)).hexdigest()
    assert older["current_source_sha256"] == sha256(sample(120)).hexdigest()
    assert database.content(NAME) == sample(120)
    assert [fact["value"] for fact in database.facts(NAME)] == ["120"]


def test_sync_and_async_share_version_order(tmp_path, monkeypatch):
    database = Store(tmp_path / "mixed.db")
    monkeypatch.setattr(main, "store", database)
    monkeypatch.setattr(main, "extract_document", substitute)
    waiting, release = Event(), Event()

    def delayed(content, name, metrics, store, *, audit, publish):
        if b"90" in content:
            waiting.set()
            assert release.wait(5)
        return substitute(content, name, metrics, store, audit=audit, publish=publish)

    monkeypatch.setattr(task_queue, "extract_document", delayed)
    client = TestClient(main.app)
    old = client.post("/tasks", files={"file": (NAME, sample(90))}).json()["task_id"]
    assert waiting.wait(2)
    synced = client.post("/extract", files={"file": (NAME, sample(120))})
    assert synced.status_code == 200 and synced.json()["kept"] == 1
    release.set()
    assert completed(client, old)["status"] == "superseded"
    assert database.content(NAME) == sample(120)
    assert [fact["value"] for fact in database.facts(NAME)] == ["120"]

    # A synchronous old request can also finish after a newer asynchronous task.
    waiting.clear()
    release.clear()
    def delayed_sync(content, name, metrics, store, *, publish):
        if b"90" in content:
            waiting.set()
            assert release.wait(5)
        return substitute(content, name, metrics, store, publish=publish)
    monkeypatch.setattr(main, "extract_document", delayed_sync)
    with ThreadPoolExecutor(max_workers=1) as pool:
        future = pool.submit(client.post, "/extract", files={"file": (NAME, sample(90))})
        assert waiting.wait(2)
        new = client.post("/tasks", files={"file": (NAME, sample(120))}).json()["task_id"]
        assert completed(client, new)["status"] == "succeeded"
        release.set()
        assert future.result(timeout=5).status_code == 409
    assert database.content(NAME) == sample(120)
    assert [fact["value"] for fact in database.facts(NAME)] == ["120"]


def test_final_failure_batch_after_successful_retry(tmp_path, monkeypatch):
    database = Store(tmp_path / "batches.db")
    monkeypatch.setattr(main, "store", database)
    source = "".join(f"第{i}行 GDP {i} 亿元\n" for i in range(1, 32)).encode()
    audits = []
    fail_second = True

    def two_batch_model(content, name, metrics, store, *, audit, publish):
        audits.append(audit)
        first_calls = 0
        def respond(prompt):
            nonlocal first_calls
            if "【块号 31】" in prompt:
                if fail_second:
                    raise ValueError("第 31 块处理失败")
                return {"facts": []}
            first_calls += 1
            return {"facts": "错误格式"} if first_calls == 1 else {"facts": []}
        return extractor.extract_document(content, name, metrics, store, respond=respond,
                                          audit=audit, publish=publish)

    monkeypatch.setattr(task_queue, "extract_document", two_batch_model)
    client = TestClient(main.app)
    failed_id = client.post("/tasks", files={"file": ("31行.txt", source)}).json()["task_id"]
    failed = completed(client, failed_id)
    assert failed["status"] == "failed"
    assert failed["error"]["batch_index"] == 2 and failed["error"]["blocks"] == [31]
    assert len(audits[0]["batches"][0]["attempts"]) == 2
    assert "error" in audits[0]["batches"][0]["attempts"][0]
    assert "error" not in audits[0]["batches"][0]["attempts"][1]

    fail_second = False
    successful_id = client.post("/tasks", files={"file": ("31行.txt", source)}).json()["task_id"]
    successful = completed(client, successful_id)
    assert successful["status"] == "succeeded" and successful["error"] is None
    assert len(audits[1]["batches"][0]["attempts"]) == 2


CHILD = r'''
import os
import sys
from pathlib import Path
import extractor
import task_queue
from store import Store

database = Store(Path(sys.argv[1]))
phase = sys.argv[2]
data = "2024年甲市 GDP 55 亿元。\n".encode()
name = "复原.txt"
with task_queue.lock:
    task_queue.prepare(database)
    with database.connect() as db:
        db.execute("BEGIN IMMEDIATE")
        request = db.execute("INSERT INTO extraction_requests(doc, source_sha256) VALUES (?, ?)",
                             (name, __import__('hashlib').sha256(data).hexdigest()))
        request_id = request.lastrowid
        db.execute("""INSERT INTO extraction_tasks(id, doc, source_sha256, request_id, status, submitted_at)
                      VALUES (?, ?, ?, ?, 'queued', ?)""",
                   ('e' * 32, name, __import__('hashlib').sha256(data).hexdigest(), request_id, task_queue.now()))
original = Store.publish_extraction
def crash(self, *args, **kwargs):
    if phase == 'before':
        os._exit(71)
    result = original(self, *args, **kwargs)
    os._exit(72)
Store.publish_extraction = crash
def model(content, name, metrics, store, *, audit, publish):
    return extractor.extract_document(content, name, metrics, store, audit=audit, publish=publish,
                                      respond=lambda _: {"facts": [{"block_no": 1, "entity": "甲市",
                                          "metric": "GDP", "value": "55", "unit": "亿元",
                                          "snippet": "GDP 55 亿元"}]})
task_queue.extract_document = model
task_queue.run(database, 'e' * 32, request_id, data, name, ['GDP'])
'''


def test_crash_before_and_after_atomic_publication(tmp_path, monkeypatch):
    backend = Path(__file__).resolve().parent
    for phase, expected_code in (("before", 71), ("after", 72)):
        database = Store(tmp_path / phase / "crash.db")
        process = subprocess.run([sys.executable, "-c", CHILD, str(database.path), phase],
                                 cwd=backend.parent, capture_output=True, text=True,
                                 env={**os.environ, "PYTHONPATH": str(backend)}, timeout=10)
        assert process.returncode == expected_code, process.stderr
        recovered = task_queue.get(database, "e" * 32)
        if phase == "before":
            assert recovered["status"] == "failed" and recovered["error"]["code"] == "interrupted"
            assert database.content("复原.txt") is None and database.facts("复原.txt") == []
        else:
            assert recovered["status"] == "succeeded" and recovered["is_current"]
            assert database.content("复原.txt") == sample(55)
            assert [fact["value"] for fact in database.facts("复原.txt")] == ["55"]
        monkeypatch.setattr(task_queue, "extract_document", substitute)
        client = TestClient(main.app)
        monkeypatch.setattr(main, "store", database)
        repeated = client.post("/tasks", files={"file": ("复原.txt", sample(55))},
                               data={"metrics": "GDP"})
        assert repeated.status_code == 202
        assert completed(client, repeated.json()["task_id"])["status"] == "succeeded"
        assert database.content("复原.txt") == sample(55)
        assert [fact["value"] for fact in database.facts("复原.txt")] == ["55"]
