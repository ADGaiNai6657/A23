"""Client timing boundaries using the actual API and a model substitute."""

from hashlib import sha256

from fastapi.testclient import TestClient
from openpyxl import load_workbook

import extractor
import main
import task_queue
from demo import DAILY_DOC, DAILY_TEMPLATE, build_demo
from fill_spec import COVID_OUTPUT
from store import Store
from test_xlsx_template import seven_column_template, synthetic_source
from timing_capture import capture, capture_uploaded_xlsx, seconds_between


XLSX_MAPPING = {"A": "date", "B": "country", "C": "continent", "D": "gdp_per_capita",
                "E": "population", "F": "每日检测数", "G": "病例数"}


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
