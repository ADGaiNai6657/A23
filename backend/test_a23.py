import json
from datetime import datetime
from io import BytesIO

import pytest
from docx import Document
from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

import main
from acceptance import check
from demo import CITY_DOC, CITY_TEMPLATE, DAILY_DOC, build_demo
from extractor import extract_document
from fill_spec import CITY_OUTPUT, COVID_OUTPUT
from filler import resolve_facts
from parsers import parse_document
from store import Store


@pytest.fixture(scope="module")
def demo(tmp_path_factory):
    return build_demo(tmp_path_factory.mktemp("a23-demo"))


@pytest.fixture
def client(monkeypatch, demo):
    inputs, outputs, _, _ = demo
    monkeypatch.setattr(main, "INPUTS", inputs)
    monkeypatch.setattr(main, "DATA", outputs)
    monkeypatch.setattr(main, "store", Store(outputs / "a23.db"))
    return TestClient(main.app)


def test_health(client):
    assert client.get("/health").json() == {"ok": True}


@pytest.mark.parametrize("extension", ["txt", "md"])
def test_text_keeps_line_numbers(client, extension):
    response = client.post("/parse", files={"file": (f"原文.{extension}", b"alpha\n\n  beta  \n")})
    assert response.status_code == 200
    assert [block["locator"]["line"] for block in response.json()["blocks"]] == [1, 3]


def test_docx_order_and_merged_cells():
    document = Document()
    document.add_paragraph("开头")
    table = document.add_table(rows=1, cols=2)
    table.cell(0, 0).merge(table.cell(0, 1)).text = "合并格"
    document.add_paragraph("结尾")
    stream = BytesIO()
    document.save(stream)
    parsed = parse_document(stream.getvalue(), "合并.docx")
    assert [(block.type, block.text) for block in parsed.blocks] == [
        ("paragraph", "开头"), ("table_cell", "合并格"), ("paragraph", "结尾"),
    ]
    assert parsed.blocks[1].locator == {"table": 1, "row": 1, "column": 1}


def test_xlsx_stream_rows_keep_headers():
    workbook = Workbook()
    workbook.active.append(["日期", "国家", "数量"])
    workbook.active.append([None, None, None])
    workbook.active.append([datetime(2020, 7, 1), "Beta", 0])
    stream = BytesIO()
    workbook.save(stream)
    parsed = parse_document(stream.getvalue(), "数据.xlsx")
    assert len(parsed.blocks) == 1
    assert parsed.blocks[0].locator["row"] == 3
    assert "数量: 0" in parsed.blocks[0].text


@pytest.mark.parametrize("path,name", [("/parse", "坏.docx"), ("/extract", "无.csv")])
def test_bad_uploads(client, path, name):
    response = client.post(path, files={"file": (name, b"not an office file")})
    assert response.status_code == 400
    assert isinstance(response.json()["detail"], str)


@pytest.mark.parametrize("extension", ["txt", "md"])
def test_fact_evidence_period_and_idempotence(tmp_path, extension):
    store = Store(tmp_path / "a23.db")
    content = "2024年\n\n甲城GDP10亿元。".encode()
    doc = f"甲.{extension}"
    valid = {"block_no": 3, "entity": "甲城", "metric": "GDP", "value": "10",
             "unit": "亿元", "snippet": "GDP10亿元"}

    def reply(prompt):
        assert "【块号 3】" in prompt and "【块号 2】" not in prompt
        return {"facts": [valid, {**valid, "block_no": 2}, {**valid, "value": "99"}]}

    for _ in range(2):
        response = extract_document(content, doc, [], store, reply)
        assert (response["count"], response["kept"]) == (3, 1)
    fact, = store.facts(doc)
    assert (fact["period"], fact["period_block_no"]) == ("2024", 1)
    assert store.block(doc, fact["block_no"])["locator"] == {"line": 3}
    assert store.block(doc, 2) is None


def test_fusion_does_not_merge_periods_or_hide_conflict(tmp_path):
    store = Store(tmp_path / "a23.db")
    for doc, year, value in (("甲.txt", 2024, "10"), ("乙.txt", 2024, "11"), ("丙.txt", 2025, "21")):
        snippet = f"GDP{value}亿元"
        raw = f"甲城{year}年{snippet}。".encode()
        fact = {"block_no": 1, "entity": "甲城", "metric": "GDP", "value": value,
                "unit": "亿元", "snippet": snippet}
        extract_document(raw, doc, [], store, lambda _prompt: {"facts": [fact]})
    conflict = resolve_facts(store.facts(), "甲城", "GDP总量", "2024", "")
    assert len(conflict["candidates"]) == 2
    assert conflict["selected"] is None
    preferred = resolve_facts(store.facts(), "甲城", "GDP", "2024", "", preferred_doc="乙.txt")
    assert preferred["selected"]["doc"] == "乙.txt"
    assert preferred["reason"] == "采用指定来源文档 乙.txt"
    assert resolve_facts(store.facts(), "甲城", "GDP", "2025", "")["selected"]["doc"] == "丙.txt"


def test_failed_batch_preserves_previous_facts(tmp_path):
    store = Store(tmp_path / "a23.db")
    raw = b"GDP10"
    fact = {"block_no": 1, "metric": "GDP", "value": "10", "snippet": "GDP10"}
    extract_document(raw, "a.txt", [], store, lambda _prompt: {"facts": [fact]})
    with pytest.raises(ValueError, match="字段不符合"):
        extract_document(raw, "a.txt", [], store, lambda _prompt: {"facts": 42})
    assert len(store.facts("a.txt")) == 1


def test_city_report_and_cell(client):
    report = client.get("/reports", params={"doc": CITY_DOC}).json()
    assert report["cell_accuracy"] == report["row_completeness"] == 1
    response = client.get("/cells", params={"doc": CITY_DOC, "file": CITY_OUTPUT, "row": 2, "column": 2})
    assert response.status_code == 200
    assert response.json()["decision"] == "filled"
    assert "GDP总量100亿元" in response.json()["text"]


def test_daily_zero_and_true_blank(client, demo):
    path = demo[1] / COVID_OUTPUT
    workbook = load_workbook(path)
    try:
        assert isinstance(workbook.active["A2"].value, datetime)
        assert workbook.active["B2"].value == "Alpha"
        assert workbook.active["C3"].value == 0
        assert workbook.active["D3"].value is None
    finally:
        workbook.close()
    response = client.get("/cells", params={"doc": DAILY_DOC, "file": COVID_OUTPUT, "row": 3, "column": 4})
    assert response.json()["decision"] == "refused"
    report = client.get("/reports", params={"doc": DAILY_DOC, "file": COVID_OUTPUT}).json()
    assert (report["required"], report["expected_empty"], report["correct"]) == (19, 2, 19)


def test_download_validation(client):
    assert client.get("/files", params={"name": "../.env"}).status_code == 400
    assert client.get("/files", params={"name": CITY_OUTPUT}).status_code == 200
    assert client.get("/files", params={"name": "filled_missing.xlsx"}).status_code == 404


def test_built_frontend_is_served(client):
    if not (main.DIST / "index.html").is_file():
        pytest.skip("先在 frontend/ 构建页面")
    response = client.get("/")
    assert response.status_code == 200
    assert "文档理解与填表" in response.text


def test_unknown_document_cannot_fill(client):
    response = client.post("/fill", json={"doc": "不存在.docx", "template": CITY_TEMPLATE})
    assert response.status_code == 400
    assert response.json()["detail"] == "库里没有这份文档"


def test_ingest_and_fill_endpoint(client, demo):
    before = (demo[0] / CITY_TEMPLATE).read_bytes()
    response = client.post("/fill", json={"doc": CITY_DOC, "template": CITY_TEMPLATE})
    assert response.status_code == 200
    assert response.json() == {"rows": 1, "cells_written": 5, "cells_empty": 0}
    assert (demo[0] / CITY_TEMPLATE).read_bytes() == before
    report = json.loads((demo[1] / "fill_report.json").read_text())
    assert all(cell["source"] for cell in report["cells"])
    result = client.post("/ingest", files={"file": ("演示.txt", b"sample")})
    assert result.json() == {"doc": "演示.txt", "blocks": 1}


@pytest.mark.parametrize("actual", [{"rows": 2}, {"incorrect_key": 1}, None])
def test_acceptance_detects_changed_value_key_or_missing_report(actual):
    failures = check({"01": actual}, {"01": {"rows": 1}})
    assert len(failures) == 1
    assert failures[0].startswith("01")
