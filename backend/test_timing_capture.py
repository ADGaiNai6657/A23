"""Client timing boundaries using the actual API and a model substitute."""

from hashlib import sha256

from fastapi.testclient import TestClient

import extractor
import main
import task_queue
from demo import DAILY_DOC, DAILY_TEMPLATE, build_demo
from fill_spec import COVID_OUTPUT
from store import Store
from timing_capture import capture, seconds_between


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


def test_server_timing_durations_have_explicit_units():
    assert seconds_between("2026-01-01T00:00:00+00:00", "2026-01-01T00:00:01.250+00:00") == 1.25
    assert seconds_between(None, "2026-01-01T00:00:00+00:00") is None
