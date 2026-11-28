"""Client timing boundaries using the actual API and a model substitute."""

import json
import sys
import time
from contextlib import contextmanager
from hashlib import sha256
from io import BytesIO
from queue import Empty, Queue
from threading import Event

import httpx
import pytest
from fastapi.testclient import TestClient
from docx import Document
from openpyxl import load_workbook

import extractor
import main
import task_queue
from demo import CITY_DOC, CITY_TEMPLATE, DAILY_DOC, DAILY_TEMPLATE, build_demo, prepare_demo_inputs
from fill_spec import CITY_OUTPUT, COVID_OUTPUT
from store import Store
from test_xlsx_template import seven_column_template, synthetic_source
import timing_capture
from timing_capture import capture, capture_uploaded_xlsx, seconds_between


XLSX_MAPPING = {"A": "date", "B": "country", "C": "continent", "D": "gdp_per_capita",
                "E": "population", "F": "每日检测数", "G": "病例数"}


def test_stream_cancel_when_ack_queue_is_full(monkeypatch):
    """The deadline can arrive after enqueueing True but before the reader consumes it."""
    class Clock:
        value = 0.0

        def __call__(self):
            return self.value

    clock = Clock()
    release = Event()
    queued = Event()
    closed = Event()

    class PausedAck(Queue):
        def put_nowait(self, item):
            super().put_nowait(item)
            clock.value = 1.0
            queued.set()

        def get(self, *args, **kwargs):
            if not release.wait(1):
                raise Empty
            return super().get(*args, **kwargs)

    monkeypatch.setattr(timing_capture, "Queue", lambda maxsize=0: PausedAck(maxsize) if maxsize else Queue())

    class Response:
        status_code = 200
        closed = False

        def iter_bytes(self):
            yield b"first"
            yield b"should-not-read"

        def close(self):
            self.closed = True
            closed.set()

    response = Response()

    class Client:
        @contextmanager
        def stream(self, *_args, **_kwargs):
            yield response

    try:
        with pytest.raises(timing_capture.DeadlineExceeded):
            timing_capture.request_before_deadline(Client(), "GET", "/files", 1, clock)
        assert queued.is_set()
        assert closed.wait(1) and response.closed
    finally:
        release.set()


def test_malformed_timing_responses_keep_stage_and_no_download(tmp_path, monkeypatch):
    inputs, outputs, _ = city_timing_inputs(tmp_path, monkeypatch)
    source = (inputs / CITY_DOC).read_bytes()
    inner = TestClient(main.app)

    class ModifiedClient:
        def __init__(self, path, response):
            self.path, self.response = path, response

        def post(self, path, **kwargs):
            return self.response if path == self.path else inner.post(path, **kwargs)

        def get(self, path, **kwargs):
            return self.response if path.startswith(self.path) else inner.get(path, **kwargs)

    for response, code in ((httpx.Response(503, json=["繁忙"]), "http_error"),
                           (httpx.Response(202, json={}), "invalid_response")):
        record = capture(ModifiedClient("/tasks", response), source, CITY_DOC)
        assert record["status"] == "failed" and record["error"]["code"] == code
        assert record["error"]["stage"] == "submit" and record["task_id"] is None
        assert record["error"]["elapsed_seconds"] >= 0 and record["download"] is None

    missing_version = httpx.Response(200, json={}, headers={"X-A23-Publication-ID": "1"})
    record = capture(ModifiedClient("/fill", missing_version), source, CITY_DOC, metrics="GDP",
                     template_name=CITY_TEMPLATE, output_name=CITY_OUTPUT, poll_interval=0.01)
    assert record["status"] == "failed" and record["error"]["stage"] == "fill"
    assert record["error"]["code"] == "invalid_response" and record["task_id"]
    assert record["download"] is None and record["fill_request_seconds"] is not None

    html = httpx.Response(200, text="<html>wrong file</html>", headers={"content-type": "text/html"})
    malformed_fill = httpx.Response(200, text="<html>not JSON</html>", headers={
        "X-A23-Publication-ID": "1", "X-A23-Result-ID": "result",
    })
    record = capture(ModifiedClient("/fill", malformed_fill), source, CITY_DOC, metrics="GDP",
                     template_name=CITY_TEMPLATE, output_name=CITY_OUTPUT, poll_interval=0.01)
    assert record["status"] == "failed" and record["error"]["stage"] == "fill"
    assert record["error"]["code"] == "invalid_response" and record["task_id"]

    record = capture(ModifiedClient("/tasks/", html), source, CITY_DOC, metrics="GDP",
                     template_name=CITY_TEMPLATE, output_name=CITY_OUTPUT, poll_interval=0.01)
    assert record["status"] == "failed" and record["error"]["stage"] == "poll"
    assert record["error"]["code"] == "invalid_response" and record["task_id"]

    record = capture(ModifiedClient("/files", html), source, CITY_DOC, metrics="GDP",
                     template_name=CITY_TEMPLATE, output_name=CITY_OUTPUT, poll_interval=0.01)
    assert record["status"] == "failed" and record["error"]["stage"] == "download"
    assert record["error"]["code"] == "invalid_response" and record["task_id"]
    assert record["download"] is None and not (outputs / "invalid.xlsx").exists()

    destination = tmp_path / "html.xlsx"
    record = capture_uploaded_xlsx(ModifiedClient("/templates/xlsx/inspect", html),
                                   synthetic_source(), "源.xlsx", seven_column_template(),
                                   "目标.xlsx", "目标", XLSX_MAPPING, download_to=destination)
    assert record["status"] == "failed" and record["error"]["stage"] == "inspect"
    assert record["error"]["code"] == "invalid_response" and not destination.exists()

    record = capture_uploaded_xlsx(ModifiedClient("/templates/xlsx/fill", httpx.Response(200, json={"rows": 1})),
                                   synthetic_source(), "源.xlsx", seven_column_template(),
                                   "目标.xlsx", "目标", XLSX_MAPPING, download_to=destination)
    assert record["status"] == "failed" and record["error"]["stage"] == "fill"
    assert record["error"]["code"] == "invalid_response" and not destination.exists()

    record = capture_uploaded_xlsx(ModifiedClient("/operations/files/", html), synthetic_source(),
                                   "源.xlsx", seven_column_template(), "目标.xlsx", "目标", XLSX_MAPPING,
                                   download_to=destination)
    assert record["status"] == "failed" and record["error"]["stage"] == "download"
    assert record["download"] is None and not destination.exists()


def test_invalid_download_cli_writes_failed_json(tmp_path, monkeypatch):
    inputs, _, _ = city_timing_inputs(tmp_path, monkeypatch)
    inner = TestClient(main.app)
    source = inputs / CITY_DOC
    record_path = tmp_path / "invalid-download.json"

    class LocalClient:
        def __init__(self, **_kwargs):
            pass

        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return None

        def post(self, path, **kwargs):
            return inner.post(path, **kwargs)

        def get(self, path, **kwargs):
            if path == "/files":
                return httpx.Response(200, text="<html>not an Excel file</html>")
            return inner.get(path, **kwargs)

    monkeypatch.setattr(timing_capture.httpx, "Client", LocalClient)
    monkeypatch.setattr(sys, "argv", ["timing_capture.py", str(source), "--metrics", "GDP",
                                     "--template-name", CITY_TEMPLATE, "--output-name", CITY_OUTPUT,
                                     "--record", str(record_path)])
    assert timing_capture.main() == 1
    receipt = json.loads(record_path.read_text())
    assert receipt["status"] == "failed" and receipt["error"]["stage"] == "download"
    assert receipt["error"]["code"] == "invalid_response" and receipt["task_id"]
    assert receipt["download"] is None and receipt["download_request_seconds"] is not None


def test_timing_capture_covers_first_response_result_and_download(tmp_path, monkeypatch):
    inputs, outputs, _, _ = build_demo(tmp_path)
    monkeypatch.setattr(main, "store", Store(outputs / "a23.db"))
    monkeypatch.setattr(main, "INPUTS", inputs)
    monkeypatch.setattr(main, "DATA", outputs)
    monkeypatch.setattr(main, "REFERENCE_DIR", inputs)
    monkeypatch.setattr(main, "SYNTHETIC_REFERENCE", True)

    def model_substitute(content, name, metrics, store, *, audit, publish):
        return extractor.extract_document(content, name, metrics, store, audit=audit, publish=publish,
                                          respond=lambda _: {"facts": []})

    monkeypatch.setattr(task_queue, "extract_document", model_substitute)
    source = (inputs / DAILY_DOC).read_bytes()
    record = capture(TestClient(main.app), source, DAILY_DOC,
                     template_name=DAILY_TEMPLATE, output_name=COVID_OUTPUT, poll_interval=0.01)
    assert record["status"] == "succeeded" and record["scope"] == "fixed_template_client"
    assert record["source_sha256"] == sha256(source).hexdigest()
    assert record["task_source_sha256"] == record["current_source_sha256"] == record["source_sha256"]
    assert record["is_current"] is True
    assert 0 <= record["first_response_seconds"] <= record["terminal_observed_seconds"]
    assert record["terminal_observed_seconds"] <= record["fill_response_seconds"] <= record["download_complete_seconds"]
    assert 0 <= record["fill_request_seconds"] <= record["fill_response_seconds"]
    assert 0 <= record["download_request_seconds"] <= record["download_complete_seconds"]
    assert record["queue_seconds"] >= 0 and record["run_seconds"] >= 0
    assert record["download"]["bytes"] == (outputs / COVID_OUTPUT).stat().st_size
    assert record["download"]["sha256"] == sha256((outputs / COVID_OUTPUT).read_bytes()).hexdigest()


def test_timing_capture_failure_has_no_download(tmp_path, monkeypatch):
    monkeypatch.setattr(main, "store", Store(tmp_path / "failed.db"))
    def failing(content, name, metrics, store, *, audit, publish):
        raise extractor.BatchFailure(1, [1], "模型替身失败")
    monkeypatch.setattr(task_queue, "extract_document", failing)
    record = capture(TestClient(main.app), b"GDP 1", "失败.txt", poll_interval=0.01)
    assert record["status"] == "failed" and record["download"] is None
    assert record["error"]["batch_index"] == 1
    assert record["fill_response_seconds"] is None
    assert record["fill_request_seconds"] is None and record["download_request_seconds"] is None


def city_model_substitute(content, name, metrics, store, *, audit=None, publish=None):
    is_gdp = "GDP" in metrics
    fact = {"block_no": 1, "entity": "甲城", "metric": "GDP总量" if is_gdp else "常住人口",
            "value": "100" if is_gdp else "20", "unit": "亿元" if is_gdp else "万人",
            "snippet": "GDP总量100亿元" if is_gdp else "常住人口20万人"}
    return extractor.extract_document(content, name, metrics, store, audit=audit, publish=publish,
                                      respond=lambda _: {"facts": [fact]})


def city_timing_inputs(tmp_path, monkeypatch):
    inputs, outputs = tmp_path / "doc", tmp_path / "data"
    prepare_demo_inputs(inputs)
    outputs.mkdir()
    database = Store(outputs / "a23.db")
    monkeypatch.setattr(main, "store", database)
    monkeypatch.setattr(main, "INPUTS", inputs)
    monkeypatch.setattr(main, "DATA", outputs)
    monkeypatch.setattr(main, "REFERENCE_DIR", inputs)
    monkeypatch.setattr(main, "SYNTHETIC_REFERENCE", True)
    monkeypatch.setattr(task_queue, "extract_document", city_model_substitute)
    monkeypatch.setattr(main, "extract_document", city_model_substitute)
    return inputs, outputs, database


def test_fixed_timing_receipt_matches_fact_batch_and_download(tmp_path, monkeypatch):
    inputs, outputs, database = city_timing_inputs(tmp_path, monkeypatch)
    source = (inputs / CITY_DOC).read_bytes()
    record = capture(TestClient(main.app), source, CITY_DOC, metrics="GDP",
                     template_name=CITY_TEMPLATE, output_name=CITY_OUTPUT, poll_interval=0.01)
    assert record["status"] == "succeeded"
    assert record["publication_id"] == record["fill_publication_id"] == record["download"]["publication_id"]
    assert record["publication_id"] == database.snapshot(CITY_DOC, with_facts=True)[3]
    book = load_workbook(outputs / CITY_OUTPUT)
    try:
        assert book.active["B2"].value == 100 and book.active["C2"].value is None
    finally:
        book.close()
    assert record["download"]["sha256"] == sha256((outputs / CITY_OUTPUT).read_bytes()).hexdigest()
    assert json.loads((outputs / "fill_report.json").read_text())["publication_id"] == record["publication_id"]


def test_fixed_timing_rejects_same_sha_newer_fact_batch(tmp_path, monkeypatch):
    inputs, outputs, database = city_timing_inputs(tmp_path, monkeypatch)
    source = (inputs / CITY_DOC).read_bytes()
    inner = TestClient(main.app)

    class InterleavingClient:
        def post(self, path, **kwargs):
            if path == "/fill":
                newer = inner.post("/extract", files={"file": (CITY_DOC, source)}, data={"metrics": "人口"})
                assert newer.status_code == 200, newer.text
            return inner.post(path, **kwargs)

        def get(self, path, **kwargs):
            return inner.get(path, **kwargs)

    record = capture(InterleavingClient(), source, CITY_DOC, metrics="GDP",
                     template_name=CITY_TEMPLATE, output_name=CITY_OUTPUT, poll_interval=0.01)
    assert record["task_status"] == "succeeded" and record["status"] == "version_conflict"
    assert record["error"]["code"] == "version_conflict"
    assert record["fill_publication_id"] is None and record["download"] is None
    assert not (outputs / CITY_OUTPUT).exists()
    assert {fact["metric"] for fact in database.facts(CITY_DOC)} == {"常住人口"}


def test_fixed_timing_superseded_before_first_poll_is_not_success(tmp_path, monkeypatch, capsys):
    inputs, outputs, database = city_timing_inputs(tmp_path, monkeypatch)
    source = (inputs / CITY_DOC).read_bytes()
    inner = TestClient(main.app)

    class InterleavingClient:
        def __init__(self):
            self.replaced = False

        def post(self, path, **kwargs):
            return inner.post(path, **kwargs)

        def get(self, path, **kwargs):
            if path.startswith("/tasks/") and not self.replaced:
                from test_tasks import completed
                assert completed(inner, path.rsplit("/", 1)[-1])["status"] == "succeeded"
                newer = inner.post("/extract", files={"file": (CITY_DOC, source)},
                                   data={"metrics": "人口"})
                assert newer.status_code == 200
                self.replaced = True
            return inner.get(path, **kwargs)

    record = capture(InterleavingClient(), source, CITY_DOC, metrics="GDP",
                     template_name=CITY_TEMPLATE, output_name=CITY_OUTPUT, poll_interval=0.01)
    assert record["task_status"] == "succeeded" and not record["is_current"]
    assert record["status"] == "version_conflict" and record["error"]["stage"] == "fill"
    assert record["download"] is None and not (outputs / CITY_OUTPUT).exists()

    # The CLI must write the same terminal receipt and return a nonzero status.
    class LocalClient:
        def __init__(self, **kwargs):
            self.client = InterleavingClient()

        def __enter__(self):
            return self.client

        def __exit__(self, *args):
            return None

    source_path = inputs / CITY_DOC
    record_path = tmp_path / "early-conflict.json"
    monkeypatch.setattr(timing_capture.httpx, "Client", LocalClient)
    monkeypatch.setattr(sys, "argv", ["timing_capture.py", str(source_path),
                                     "--metrics", "GDP", "--template-name", CITY_TEMPLATE,
                                     "--output-name", CITY_OUTPUT, "--record", str(record_path)])
    assert timing_capture.main() == 1
    assert capsys.readouterr().out
    assert json.loads(record_path.read_text())["status"] == "version_conflict"
    standalone = capture(InterleavingClient(), source, CITY_DOC, metrics="GDP", poll_interval=0.01)
    assert standalone["status"] == standalone["task_status"] == "succeeded"
    assert standalone["is_current"] is False and standalone["download"] is None


def test_score_uses_fill_snapshot_after_new_document_import(tmp_path, monkeypatch):
    inputs, outputs, database = city_timing_inputs(tmp_path, monkeypatch)
    source = (inputs / CITY_DOC).read_bytes()
    assert TestClient(main.app).post("/extract", files={"file": (CITY_DOC, source)},
                                     data={"metrics": "GDP"}).status_code == 200
    publication_id = database.snapshot(CITY_DOC, with_facts=True)[3]
    newer = Document(BytesIO(source))
    newer.add_paragraph("新版补充段落。")
    stream = BytesIO()
    newer.save(stream)
    client = TestClient(main.app)
    original_score = main.score_filled

    def interleaved_score(doc):
        imported = client.post("/ingest", files={"file": (CITY_DOC, stream.getvalue())})
        assert imported.status_code == 200
        original_score(doc)

    monkeypatch.setattr(main, "score_filled", interleaved_score)
    filled = client.post("/fill", json={"doc": CITY_DOC, "template": CITY_TEMPLATE,
                                        "source_sha256": sha256(source).hexdigest(),
                                        "expected_publication_id": publication_id})
    assert filled.status_code == 200, filled.text
    trace = json.loads((outputs / "fill_report.json").read_text())
    score = json.loads((outputs / "eval_report.json").read_text())
    assert len(database.blocks(CITY_DOC)) == 2
    assert trace["block_count"] == score["block_count"] == 1
    assert trace["publication_id"] == score["publication_id"] == publication_id
    assert trace["source_sha256"] == score["source_sha256"] == sha256(source).hexdigest()
    assert "database_sha256" not in score
    downloaded = client.get("/files", params={"name": CITY_OUTPUT,
                                               "result_id": filled.headers["X-A23-Result-ID"]})
    assert downloaded.status_code == 200
    book = load_workbook(BytesIO(downloaded.content))
    try:
        assert book.active["B2"].value == 100 and book.active["C2"].value is None
    finally:
        book.close()


def test_fixed_cli_records_fill_download_and_poll_failures(tmp_path, monkeypatch, capsys):
    inputs, outputs, _ = city_timing_inputs(tmp_path, monkeypatch)
    inner = TestClient(main.app)
    source = inputs / CITY_DOC

    class LocalClient:
        fail_download = False
        slow_poll = False

        def __init__(self, **kwargs):
            pass

        def __enter__(self):
            return self

        def __exit__(self, *args):
            return None

        def post(self, path, **kwargs):
            return inner.post(path, **kwargs)

        def get(self, path, **kwargs):
            if self.fail_download and path == "/files":
                return httpx.Response(404, json={"detail": "受控下载失败"})
            if self.slow_poll and path.startswith("/tasks/"):
                time.sleep(0.05)
            return inner.get(path, **kwargs)

    monkeypatch.setattr(timing_capture.httpx, "Client", LocalClient)
    for label, template, timeout, expected_status, expected_stage, expected_exit in (
        ("fill", "不存在.xlsx", "10", "failed", "fill", 1),
        ("download", CITY_TEMPLATE, "10", "failed", "download", 1),
        ("poll", CITY_TEMPLATE, "0.02", "timed_out", "poll", 2),
    ):
        LocalClient.fail_download = label == "download"
        LocalClient.slow_poll = label == "poll"
        record_path = tmp_path / f"{label}.json"
        monkeypatch.setattr(sys, "argv", ["timing_capture.py", str(source), "--metrics", "GDP",
                                         "--template-name", template, "--output-name", CITY_OUTPUT,
                                         "--timeout", timeout, "--record", str(record_path)])
        assert timing_capture.main() == expected_exit
        record = json.loads(record_path.read_text())
        assert record["status"] == expected_status and record["error"]["stage"] == expected_stage
        assert record["task_id"] and record["first_response_seconds"] is not None
        if label == "fill":
            assert record["error"]["http_status"] == 400
            assert record["fill_request_seconds"] is not None and record["download"] is None
        elif label == "download":
            assert record["error"]["http_status"] == 404
            assert record["download_request_seconds"] is not None and record["download"] is None
        else:
            assert record["poll_elapsed_seconds"] is not None
            assert record["terminal_observed_seconds"] is None and record["download"] is None
    assert capsys.readouterr().out


def test_transport_disconnect_is_located_in_all_timing_stages(tmp_path, monkeypatch):
    inputs, _, _ = city_timing_inputs(tmp_path, monkeypatch)
    source = (inputs / CITY_DOC).read_bytes()
    inner = TestClient(main.app)

    class BrokenClient:
        def __init__(self, broken_path):
            self.broken_path = broken_path

        def post(self, path, **kwargs):
            if path == self.broken_path:
                raise httpx.RemoteProtocolError("受控断连")
            return inner.post(path, **kwargs)

        def get(self, path, **kwargs):
            if path.startswith(self.broken_path):
                raise httpx.RemoteProtocolError("受控断连")
            return inner.get(path, **kwargs)

    for stage, path in (("submit", "/tasks"), ("poll", "/tasks/"),
                        ("fill", "/fill"), ("download", "/files")):
        record = capture(BrokenClient(path), source, CITY_DOC, metrics="GDP",
                         template_name=CITY_TEMPLATE, output_name=CITY_OUTPUT,
                         poll_interval=0.01)
        assert record["status"] == "failed" and record["error"]["code"] == "transport_error"
        assert record["error"]["stage"] == stage and record["error"]["elapsed_seconds"] >= 0
        assert bool(record["task_id"]) == (stage != "submit")
        assert record["download"] is None

    for stage, path in (("inspect", "/templates/xlsx/inspect"),
                        ("fill", "/templates/xlsx/fill"),
                        ("download", "/operations/files/")):
        record = capture_uploaded_xlsx(BrokenClient(path), synthetic_source(), "源.xlsx",
                                       seven_column_template(), "目标.xlsx", "目标", XLSX_MAPPING)
        assert record["status"] == "failed" and record["error"]["code"] == "transport_error"
        assert record["error"]["stage"] == stage and record["download"] is None


def test_fixed_timing_rejects_replaced_download_even_with_identical_cells(tmp_path, monkeypatch):
    inputs, outputs, database = city_timing_inputs(tmp_path, monkeypatch)
    source = (inputs / CITY_DOC).read_bytes()
    inner = TestClient(main.app)
    newer_id = old_id = None

    class InterleavingClient:
        def post(self, path, **kwargs):
            return inner.post(path, **kwargs)

        def get(self, path, **kwargs):
            nonlocal newer_id, old_id
            if path == "/files":
                old_id = kwargs["params"]["result_id"]
                newer = inner.post("/extract", files={"file": (CITY_DOC, source)},
                                   data={"metrics": "GDP"})
                assert newer.status_code == 200
                replacement = inner.post("/fill", json={"doc": CITY_DOC, "template": CITY_TEMPLATE,
                                                        "source_sha256": sha256(source).hexdigest()})
                assert replacement.status_code == 200
                newer_id = replacement.headers["X-A23-Result-ID"]
            return inner.get(path, **kwargs)

    record = capture(InterleavingClient(), source, CITY_DOC, metrics="GDP",
                     template_name=CITY_TEMPLATE, output_name=CITY_OUTPUT, poll_interval=0.01)
    assert record["task_status"] == "succeeded" and record["status"] == "version_conflict"
    assert record["error"]["code"] == "version_conflict" and record["download"] is None
    assert newer_id != old_id
    assert json.loads((outputs / "fill_report.json").read_text())["publication_id"] != record["publication_id"]
    book = load_workbook(outputs / CITY_OUTPUT)
    try:
        assert book.active["B2"].value == 100
    finally:
        book.close()


def test_server_timing_durations_have_explicit_units():
    assert seconds_between("2026-01-01T00:00:00+00:00", "2026-01-01T00:00:01.250+00:00") == 1.25
    assert seconds_between(None, "2026-01-01T00:00:00+00:00") is None


def test_uploaded_template_timing_starts_at_preview_and_reopens_download(tmp_path):
    source, template = synthetic_source(conflict=True), seven_column_template()
    destination = tmp_path / "计时下载.xlsx"
    record = capture_uploaded_xlsx(TestClient(main.app), source, "源.xlsx", template, "目标.xlsx",
                                   "目标", XLSX_MAPPING, download_to=destination)
    assert record["status"] == "succeeded" and record["scope"] == "uploaded_xlsx03_client"
    assert record["source_sha256"] == sha256(source).hexdigest()
    assert record["template_sha256"] == sha256(template).hexdigest()
    assert 0 <= record["preview_response_seconds"] <= record["fill_complete_seconds"]
    assert record["fill_complete_seconds"] <= record["download_complete_seconds"]
    assert 0 <= record["fill_request_seconds"] <= record["fill_complete_seconds"]
    assert 0 <= record["download_request_seconds"] <= record["download_complete_seconds"]
    assert (record["fill"]["rows"], record["fill"]["cells_written"],
            record["fill"]["cells_empty"]) == (1, 6, 1)
    assert record["download"]["sha256"] == sha256(destination.read_bytes()).hexdigest()
    workbook = load_workbook(destination)
    try:
        assert workbook["目标"]["G2"].value is None
        assert workbook["目标"]["A5"].value == "人工备注"
        assert workbook["目标"]["H6"].value == "=SUM(A6:G6)"
        assert workbook["其他"]["B3"].value == "保持原样"
    finally:
        workbook.close()


def test_uploaded_template_timing_records_preview_rejection(tmp_path):
    destination = tmp_path / "不应生成.xlsx"
    record = capture_uploaded_xlsx(TestClient(main.app), synthetic_source(), "源.xlsx",
                                   seven_column_template(blocked=True), "目标.xlsx", "目标",
                                   XLSX_MAPPING, download_to=destination)
    assert record["status"] == "failed" and record["error"]["stage"] == "inspect"
    assert record["error"]["http_status"] == 400 and "E2" in record["error"]["detail"]
    assert record["fill"] is None and record["download"] is None
    assert not destination.exists()


def test_uploaded_template_timing_records_blocked_selected_sheet(tmp_path):
    book = load_workbook(BytesIO(seven_column_template(blocked=True)))
    valid = book.copy_worksheet(book["目标"])
    valid.title = "可填"
    valid["E2"] = None
    output = BytesIO()
    book.save(output)
    book.close()
    destination = tmp_path / "不应生成.xlsx"
    record = capture_uploaded_xlsx(TestClient(main.app), synthetic_source(), "源.xlsx",
                                   output.getvalue(), "双表.xlsx", "目标", XLSX_MAPPING,
                                   download_to=destination)
    assert record["status"] == "failed" and record["error"] == {
        "code": "blocked_sheet", "stage": "selection", "sheet": "目标",
        "detail": record["error"]["detail"],
    }
    assert "E2" in record["error"]["detail"] and record["preview_response_seconds"] is not None
    assert record["fill"] is None and not destination.exists()


def test_uploaded_template_timeout_covers_each_stage(tmp_path):
    source, template = synthetic_source(), seven_column_template()

    class FakeClock:
        value = 0.0

        def __call__(self):
            return self.value

    for stage, slow_path in (("inspect", "/templates/xlsx/inspect"),
                             ("fill", "/templates/xlsx/fill"),
                             ("download", "/operations/files/")):
        clock = FakeClock()
        inner = TestClient(main.app)

        class SlowClient:
            def post(self, path, **kwargs):
                if path == slow_path:
                    clock.value += 2.0
                return inner.post(path, **kwargs)

            def get(self, path, **kwargs):
                if path.startswith(slow_path):
                    clock.value += 2.0
                return inner.get(path, **kwargs)

        destination = tmp_path / f"{stage}.xlsx"
        record = capture_uploaded_xlsx(SlowClient(), source, "源.xlsx", template, "目标.xlsx",
                                       "目标", XLSX_MAPPING, download_to=destination,
                                       timeout=1, clock=clock)
        assert record["status"] == "timed_out" and record["error"]["stage"] == stage
        assert record["error"]["code"] == "timeout"
        assert record["error"]["elapsed_seconds"] == 2.0
        assert record["download"] is None and not destination.exists()


def test_uploaded_template_transport_timeout_records_stage(tmp_path):
    source, template = synthetic_source(), seven_column_template()
    inner = TestClient(main.app)
    for stage, path in (("inspect", "/templates/xlsx/inspect"),
                        ("fill", "/templates/xlsx/fill"),
                        ("download", "/operations/files/")):
        class TimeoutClient:
            def post(self, url, **kwargs):
                if url == path:
                    raise httpx.ReadTimeout("受控超时")
                return inner.post(url, **kwargs)

            def get(self, url, **kwargs):
                if url.startswith(path):
                    raise httpx.ReadTimeout("受控超时")
                return inner.get(url, **kwargs)

        destination = tmp_path / f"transport-{stage}.xlsx"
        record = capture_uploaded_xlsx(TimeoutClient(), source, "源.xlsx", template,
                                       "目标.xlsx", "目标", XLSX_MAPPING,
                                       download_to=destination, timeout=1)
        assert record["status"] == "timed_out"
        assert record["error"]["stage"] == stage
        assert record["download"] is None and not destination.exists()


def test_uploaded_template_stream_stops_at_total_deadline(tmp_path):
    class FakeClock:
        value = 0.0

        def __call__(self):
            return self.value

    clock = FakeClock()
    received = []
    inner = TestClient(main.app)

    class StreamingResponse:
        status_code = 200

        def iter_bytes(self):
            for index in range(6):
                clock.value += 0.05
                received.append(index)
                yield b"part"

    class StreamingClient:
        def post(self, path, **kwargs):
            return inner.post(path, **kwargs)

        @contextmanager
        def stream(self, method, path, **kwargs):
            assert method == "GET" and path.startswith("/operations/files/")
            yield StreamingResponse()

    destination = tmp_path / "不应下载.xlsx"
    record = capture_uploaded_xlsx(StreamingClient(), synthetic_source(), "源.xlsx",
                                   seven_column_template(), "目标.xlsx", "目标", XLSX_MAPPING,
                                   download_to=destination, timeout=0.15, clock=clock)
    assert record["status"] == "timed_out" and record["error"]["stage"] == "download"
    assert 2 <= len(received) < 6 and record["download"] is None
    assert not destination.exists()


def test_uploaded_template_cli_writes_blocked_sheet_record(tmp_path, monkeypatch, capsys):
    book = load_workbook(BytesIO(seven_column_template(blocked=True)))
    valid = book.copy_worksheet(book["目标"])
    valid.title = "可填"
    valid["E2"] = None
    template = tmp_path / "双表.xlsx"
    book.save(template)
    book.close()
    source = tmp_path / "源.xlsx"
    source.write_bytes(synthetic_source())
    mapping = tmp_path / "mapping.json"
    mapping.write_text(json.dumps(XLSX_MAPPING), encoding="utf-8")
    record_path = tmp_path / "failure.json"

    class LocalClient:
        def __init__(self, **kwargs):
            self.client = TestClient(main.app)

        def __enter__(self):
            return self.client

        def __exit__(self, *args):
            self.client.close()

    monkeypatch.setattr(timing_capture.httpx, "Client", LocalClient)
    monkeypatch.setattr(sys, "argv", ["timing_capture.py", str(source),
                                     "--uploaded-xlsx-template", str(template),
                                     "--template-sheet", "目标", "--mapping-json", str(mapping),
                                     "--record", str(record_path)])
    assert timing_capture.main() == 1
    assert capsys.readouterr().out
    record = json.loads(record_path.read_text(encoding="utf-8"))
    assert record["status"] == "failed"
    assert record["error"]["stage"] == "selection"
    assert record["error"]["sheet"] == "目标"
    assert "E2" in record["error"]["detail"]
