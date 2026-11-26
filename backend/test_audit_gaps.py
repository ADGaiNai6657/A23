import json
from datetime import datetime
from pathlib import Path

import pytest
from docx import Document
from fastapi.testclient import TestClient
from openpyxl import load_workbook
from openpyxl.styles import Font

import check_dataset
import filler
import main
from acceptance import DELIVERABLES, checked_outputs
from config import ROOT
from criteria import checked_rules, load_expectations
from demo import CITY_DOC, CITY_TEMPLATE, build_demo, make_workbook
from evaluate import evaluate
from extractor import extract_document
from fill_spec import CITY_HEADERS, CITY_OUTPUT, COVID_OUTPUT, REPORT_FILES, metric_key
from filler import fill_workbook
from parsers import parse_document
from rebuild_evidence import check_distinct_runs, check_execution
from store import Store
from unpack_dataset import unpack_dataset


def test_complete_quantity_and_declared_metric_aliases(tmp_path):
    content = "甲城GDP总量100亿元，另有支出1,234.50亿元，温度−20度。".encode()
    base = {"block_no": 1, "entity": "甲城", "metric": "GDP总量", "value": "100", "snippet": "GDP总量100亿元"}
    store = Store(tmp_path / "a23.db")
    facts = [
        base, base, {**base, "value": "10"}, {**base, "value": "10", "snippet": "GDP总量10"},
        {**base, "metric": "GDP增速"},
        {**base, "metric": "支出", "value": "1234.50", "snippet": "支出1,234.50亿元"},
        {**base, "metric": "温度", "value": "-20", "snippet": "温度−20度"},
    ]
    result = extract_document(content, "统计.txt", ["地区生产总值", "支出", "温度"], store, lambda _: {"facts": facts})
    assert result["kept"] == 3
    assert result["rejected"] == {"重复事实": 1, "原话缺少数值": 2, "指标不在清单内": 1}
    assert metric_key("新增病例") == metric_key("new_cases")
    assert metric_key("new_cases_per_million") != metric_key("new_cases")
    assert metric_key("GDP 总量（亿元）") == metric_key("地区生产总值")


@pytest.mark.parametrize("extension", ["txt", "md", "docx", "xlsx"])
def test_extract_http_accepts_all_formats_with_model_substitute(extension, tmp_path, monkeypatch):
    path = tmp_path / f"指标.{extension}"
    if extension == "xlsx":
        make_workbook(path, ("GDP",), [(100,)])
        content = path.read_bytes()
    elif extension == "docx":
        doc = Document()
        doc.add_paragraph("GDP100亿元")
        doc.save(path)
        content = path.read_bytes()
    else:
        content = "GDP100亿元".encode()
    block = parse_document(content, path.name).blocks[0]
    reply = {"facts": [{"block_no": 1, "metric": "GDP", "value": "100", "snippet": block.text}]}
    store = Store(tmp_path / "a23.db")
    monkeypatch.setattr(main, "store", store)
    monkeypatch.setattr(main, "extract_document", lambda data, name, metrics, target: extract_document(
        data, name, metrics, target, lambda _: reply,
    ))
    response = TestClient(main.app).post("/extract", files={"file": (path.name, content)})
    assert response.status_code == 200 and response.json()["kept"] == 1
    assert store.block(path.name, 1)["locator"] == block.locator


def city_with_periods(tmp_path: Path) -> tuple[Path, dict]:
    doc, template = tmp_path / "城市.docx", tmp_path / CITY_TEMPLATE
    document = Document()
    document.add_paragraph("2024 年中国城市经济百强全景报告")
    document.add_paragraph("甲城：2023年GDP总量90亿元，2024年GDP总量100亿元，2024年GDP总量25亿元。"
                           "2024年统计如下。常住人口20万人，人均GDP50000元，一般公共预算收入10亿元。")
    document.save(doc)
    make_workbook(template, CITY_HEADERS)
    rows = [
        ("GDP总量", "90", "亿元", "GDP总量90亿元", ""),
        ("GDP总量", "100", "亿元", "GDP总量100亿元", ""),
        ("GDP总量", "25", "亿元", "GDP总量25亿元", "一季度"),
        ("常住人口", "20", "万人", "常住人口20万人", ""),
        ("人均GDP", "50000", "元", "人均GDP50000元", ""),
        ("一般公共预算收入", "10", "亿元", "一般公共预算收入10亿元", ""),
    ]
    response = {"facts": [
        {"block_no": 2, "entity": "甲城", "metric": metric, "value": value, "unit": unit,
         "snippet": snippet, "scope": scope}
        for metric, value, unit, snippet, scope in rows
    ]}
    return doc, response


def test_city_filters_period_and_scope_and_exposes_selected_fact(tmp_path, monkeypatch):
    doc, response = city_with_periods(tmp_path)
    store = Store(tmp_path / "a23.db")
    extract_document(doc.read_bytes(), doc.name, [], store, lambda _: response)
    facts = {fact["value"]: fact for fact in store.facts(doc.name)}
    assert facts["90"]["period"] == "2023"
    assert (facts["20"]["period"], facts["20"]["period_block_no"]) == ("2024", 2)
    fill_workbook(store, doc.name, tmp_path / CITY_TEMPLATE, tmp_path)
    monkeypatch.setattr(main, "DATA", tmp_path)
    monkeypatch.setattr(main, "store", store)
    cell = TestClient(main.app).get("/cells", params={"doc": doc.name, "file": CITY_OUTPUT, "row": 2, "column": 2}).json()
    assert cell["value"] == "100" and cell["fact"]["snippet"] == "GDP总量100亿元"
    assert cell["fact"]["period"] == "2024" and cell["fact"]["mode"] == "general"
    with pytest.raises(ValueError, match="正式输出应覆盖"):
        fill_workbook(store, doc.name, tmp_path / CITY_TEMPLATE, tmp_path, formal_city=True)


@pytest.mark.parametrize("issue", ["empty", "columns", "short", "duplicate", "valid"])
def test_formal_city_reference_rejects_wrong_shape_and_missing_truth(tmp_path, issue):
    filled, reference = tmp_path / "filled.xlsx", tmp_path / "reference.xlsx"
    rows = [("城市" + str(i), 100, 10, 20, 30) for i in range(100)]
    actual = rows.copy()
    expected = rows.copy()
    headers = CITY_HEADERS
    if issue == "empty":
        actual[0] = expected[0] = (rows[0][0], None, 10, 20, 30)
    elif issue == "columns":
        headers = (headers[0], headers[2], headers[1], *headers[3:])
    elif issue == "short":
        expected.pop()
    elif issue == "duplicate":
        expected[1] = expected[0]
    make_workbook(filled, CITY_HEADERS, actual)
    make_workbook(reference, headers, expected)
    if issue == "valid":
        assert evaluate(filled, reference, tmp_path / "report.json", CITY_DOC, False, expected_rows=100)["cell_accuracy"] == 1
    else:
        cause = {"empty": "缺少有效应填值", "columns": "列名或列顺序", "short": "行数",
                 "duplicate": "重复出现城市"}[issue]
        with pytest.raises(ValueError, match=cause):
            evaluate(filled, reference, tmp_path / "report.json", CITY_DOC, False, expected_rows=100)


def test_api_reference_profile_must_explicitly_allow_short_synthetic_reference(tmp_path, monkeypatch):
    inputs, outputs, _, _ = build_demo(tmp_path)
    monkeypatch.setattr(main, "store", Store(outputs / "a23.db"))
    monkeypatch.setattr(main, "DATA", outputs)
    monkeypatch.setattr(main, "INPUTS", inputs)
    monkeypatch.setattr(main, "REFERENCE_DIR", inputs)
    monkeypatch.setattr(main, "SYNTHETIC_REFERENCE", False)
    client = TestClient(main.app)
    payload = {"doc": CITY_DOC, "template": CITY_TEMPLATE}
    response = client.post("/fill", json=payload)
    assert response.status_code == 200 and response.headers["X-A23-Score-Status"] == "failed"
    assert client.get("/reports", params={"doc": CITY_DOC}).status_code == 404
    monkeypatch.setattr(main, "SYNTHETIC_REFERENCE", True)
    assert "X-A23-Score-Status" not in client.post("/fill", json=payload).headers
    assert client.get("/reports", params={"doc": CITY_DOC}).json()["cell_accuracy"] == 1


def test_rebuild_honors_numeric_ranges_and_rejects_relabelled_same_database(tmp_path, monkeypatch):
    corpus = tmp_path / "corpus"
    unpack_dataset(ROOT / "dataset" / "测试集.zip", corpus)
    original = json.loads((DELIVERABLES / "source" / "city-original-historical.json").read_bytes())
    original["extraction_id"] = "controlled-missing-quantity"
    original["facts"] = [fact for fact in original["facts"]
                         if not (fact["entity"] == "上海" and fact["metric_key"] == "gdp")]
    archive = tmp_path / "test-facts.json"
    archive.write_text(json.dumps(original, ensure_ascii=False), encoding="utf-8")
    rules = checked_rules(load_expectations(), "rebuilt")
    rules["01"].update({"cells_written": [499, 500], "cells_empty": [0, 1],
                        "cell_accuracy": [.998, 1], "row_completeness": [.99, 1]})
    parse_rules = check_dataset.checked_rules
    monkeypatch.setattr(check_dataset, "checked_rules", lambda config, mode: (
        rules if mode == "rebuilt" else parse_rules(config, mode)
    ))
    output = tmp_path / "rebuilt"
    generated = check_dataset.check_dataset(corpus, output, False, archive)
    assert generated["scenarios"]["01"]["differences"] == [
        {"row": 2, "column": 2, "actual": None, "source_value": 56708.71, "decision": "漏填"}
    ]
    reports = checked_outputs(output, "source")
    check_execution(output, "rebuilt", reports, rebuilt_kind="history")
    with pytest.raises(ValueError, match="online"):
        check_execution(output, "rebuilt", reports)
    with pytest.raises(ValueError, match="不能用于 same-db"):
        check_execution(output, "same-db", reports)
    with pytest.raises(ValueError, match="SHA256 完全相同"):
        check_distinct_runs(reports, reports)
    with pytest.raises(ValueError, match="01.cells_written"):
        check_dataset.check_dataset(corpus, output, False)


def test_daily_conflict_detail_exposes_both_verified_sources(tmp_path, monkeypatch):
    source, template = tmp_path / "来源.xlsx", tmp_path / "模板.xlsx"
    keys = ("date", "location", "new_cases", "new_deaths", "total_cases", "total_deaths", "population")
    make_workbook(source, (*keys, "新增病例"), [(datetime(2020, 7, 1), "Alpha", 10, 0, 100, 0, 1000, 20)])
    make_workbook(template, keys)
    store = Store(tmp_path / "a23.db")
    data = source.read_bytes()
    store.save(parse_document(data, source.name), data)
    fill_workbook(store, source.name, template, tmp_path)
    monkeypatch.setattr(main, "store", store)
    monkeypatch.setattr(main, "DATA", tmp_path)
    cell = TestClient(main.app).get("/cells", params={"doc": source.name, "file": COVID_OUTPUT, "row": 2, "column": 3})
    assert cell.status_code == 200 and cell.json()["decision"] == "refused"
    options = cell.json()["conflicts"]
    assert {(item["value"], item["source"]["field"]) for item in options} == {("10", "new_cases"), ("20", "新增病例")}
    assert all(item["evidence_status"] == "available" and item["source"]["locator"]["row"] == 2 for item in options)


def test_trace_write_failure_keeps_last_readable_result(tmp_path, monkeypatch):
    inputs, outputs, city, _ = build_demo(tmp_path)
    store = Store(outputs / "a23.db")
    template = inputs / CITY_TEMPLATE
    workbook = load_workbook(template)
    workbook.active["A1"].font = Font(bold=True)
    workbook.save(template)
    workbook.close()
    published = (outputs / CITY_OUTPUT, outputs / REPORT_FILES[CITY_OUTPUT][0])
    before = [path.read_bytes() for path in published]

    def fail(path, result):
        raise PermissionError("模拟逐格文件保存失败")

    monkeypatch.setattr(filler, "write_json", fail)
    with pytest.raises(PermissionError, match="模拟逐格"):
        fill_workbook(store, CITY_DOC, template, outputs)
    assert [path.read_bytes() for path in published] == before
    monkeypatch.setattr(main, "store", store)
    monkeypatch.setattr(main, "DATA", outputs)
    query = {"doc": CITY_DOC, "result_id": city["result_id"]}
    assert TestClient(main.app).get("/reports", params=query).status_code == 200
