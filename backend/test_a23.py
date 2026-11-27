import json
import os
import re
from datetime import datetime
from hashlib import sha256
from io import BytesIO
from pathlib import Path

import pytest
from docx import Document
from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook
from openpyxl.utils.datetime import to_excel

import main
from acceptance import DEMO_DELIVERY, check, checked_outputs
from check_extract import compare_runs, export_facts, replay_run
from criteria import checked_rules, load_expectations
from demo import CITY_DOC, CITY_TEMPLATE, DAILY_DOC, DAILY_TEMPLATE, build_demo, demo_response, make_workbook, prepare_demo_inputs
from evaluate import evaluate
from config import ROOT
from extractor import evidence_error, extract_document
from fill_spec import CITY_COLUMNS, CITY_OUTPUT, COVID_OUTPUT, REFERENCE_FILES, REPORT_FILES
from filler import city_value, fill_workbook, resolve_facts, write_json
from parsers import parse_document
from schema import Block, ParsedDoc, RawFact
from store import Store
from unpack_dataset import unpack_dataset


@pytest.fixture
def demo(tmp_path_factory):
    return build_demo(tmp_path_factory.mktemp("a23-demo"))


@pytest.fixture
def client(monkeypatch, demo):
    inputs, outputs, _, _ = demo
    monkeypatch.setattr(main, "INPUTS", inputs)
    monkeypatch.setattr(main, "DATA", outputs)
    monkeypatch.setattr(main, "store", Store(outputs / "a23.db"))
    monkeypatch.setattr(main, "REFERENCE_DIR", None)
    monkeypatch.setattr(main, "SOURCE_CHECK_ROOT", None)
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


@pytest.mark.parametrize("path,name", [
    ("/parse", "坏.docx"), ("/extract", "无.csv"), ("/extract", "坏.xlsx"), ("/extract", "坏.docx"),
])
def test_bad_uploads(client, path, name):
    response = client.post(path, files={"file": (name, b"not an office file")})
    assert response.status_code == 400
    assert isinstance(response.json()["detail"], str)


def test_text_decoding_exceptions_and_empty_extract(client):
    for path in ("/parse", "/ingest", "/extract"):
        response = client.post(path, files={"file": ("bad.txt", "无法解码".encode("gbk"))})
        assert response.status_code == 400 and "无法读取 .txt" in response.json()["detail"]
    assert client.post("/extract", files={"file": ("空.txt", b"\n ")}).status_code == 400
    assert client.post("/extract").status_code == 422
    result = client.post("/parse", files={"file": ("bom.txt", "测试".encode("utf-8-sig"))})
    assert result.json()["blocks"][0]["text"] == "测试"


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
    if not main.FRONTEND_BUILT and os.getenv("A23_REQUIRE_FRONTEND") != "1":
        pytest.skip("先在 frontend/ 构建页面")
    response = client.get("/")
    assert response.status_code == 200
    assert "文档理解与填表" in response.text
    assets = re.findall(r'(?:src|href)="(/assets/[^"]+)"', response.text)
    assert assets and any(asset.endswith(".js") for asset in assets)
    assert all(client.get(asset).status_code == 200 for asset in assets)


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


def test_same_xlsx_bytes_refresh_old_blocks_and_remove_dependent_facts(tmp_path):
    store = Store(tmp_path / "a23.db")
    path = tmp_path / "数据.xlsx"
    make_workbook(path, ("B", None), [(11, 22)])
    content = path.read_bytes()
    block = Block(type="spreadsheet_row", text="B: 22", locator={"sheet": "Sheet", "row": 2})
    old = ParsedDoc(file=path.name, blocks=[block])
    fact = (1, "甲", "B", "b", "22", "22", "", "B: 22", "", None, "")
    store.save(old, content, "general", [fact])
    parsed = parse_document(content, path.name)
    store.save(parsed, content)
    store.save(parsed, content)
    assert [item["text"] for item in store.blocks(path.name)] == ["B: 11｜列 B: 22"]
    assert store.facts(path.name) == []


@pytest.mark.parametrize("index,value,number,unit,quote,expected", [
    (0, "100", "100", "", "GDP 100 亿元", 100),
    (1, "1,295.80 万人口", None, "", "1,295.80 万人口", 1295.8),
    (0, "100", None, "万元", "GDP 100 万元", None),
    (0, "100", None, "", "GDP 100 万元", None),
])
def test_city_unit_evidence_and_number_fallback(index, value, number, unit, quote, expected):
    fact = {"value": value, "value_num": number, "unit": unit, "snippet": quote}
    result = city_value(fact, CITY_COLUMNS[index])
    assert (float(result) if result is not None else None) == expected


@pytest.mark.parametrize("alias,expected", [(None, 10), (2, None), (10, 10)])
def test_daily_synonym_conflict_is_traced_without_per_million_collision(tmp_path, alias, expected):
    source = tmp_path / "原文.xlsx"
    template = tmp_path / "模板.xlsx"
    headers = ("location", "new_cases", "new_deaths", "total_cases", "total_deaths", "population")
    make_workbook(source, ("date", *headers, "新增病例", "new_cases_per_million"),
                  [(datetime(2020, 7, 1), "Alpha", 10, 0, 100, 0, 300, alias, 2)])
    make_workbook(template, headers)
    store = Store(tmp_path / "a23.db")
    data = source.read_bytes()
    store.save(parse_document(data, source.name), data)
    fill_workbook(store, source.name, template, tmp_path)
    workbook = load_workbook(tmp_path / COVID_OUTPUT)
    assert workbook.active["C2"].value == expected
    workbook.close()
    report = json.loads((tmp_path / "fill_report_covid19.json").read_text())
    item = report["cells"][2]
    if expected is None:
        assert {option["source"]["field"] for option in item["conflicts"]} == {"new_cases", "新增病例"}


def test_bad_daily_mapping_returns_readable_error(client, demo):
    template = demo[0] / "模板_无关.xlsx"
    make_workbook(template, ("location", "额外", "new_deaths", "total_cases", "total_deaths", "population"))
    response = client.post("/fill", json={"doc": DAILY_DOC, "template": template.name})
    assert response.status_code == 400
    assert "无法映射" in response.json()["detail"]


def test_unknown_document_has_no_score(client):
    assert client.get("/reports", params={"doc": "不存在.docx"}).status_code == 404


def test_parse_and_fill_reject_different_version_and_hide_stale_evidence(tmp_path, monkeypatch):
    inputs, outputs, original, _ = build_demo(tmp_path)
    store = Store(outputs / "a23.db")
    monkeypatch.setattr(main, "INPUTS", inputs)
    monkeypatch.setattr(main, "DATA", outputs)
    monkeypatch.setattr(main, "store", store)
    client = TestClient(main.app)
    document = Document(inputs / CITY_DOC)
    document.paragraphs[0].text = document.paragraphs[0].text.replace("GDP总量100", "GDP总量200")
    buffer = BytesIO()
    document.save(buffer)
    content = buffer.getvalue()
    response = client.post("/parse", files={"file": (CITY_DOC, content)})
    digest = response.headers["X-A23-Source-SHA256"]
    assert digest == sha256(content).hexdigest()
    response = client.post("/fill", json={"doc": CITY_DOC, "template": CITY_TEMPLATE, "source_sha256": digest})
    assert response.status_code == 400 and "内容不同" in response.json()["detail"]
    fact = {"block_no": 1, "entity": "甲城", "metric": "GDP总量", "value": "200",
            "unit": "亿元", "snippet": "GDP总量200亿元"}
    extract_document(content, CITY_DOC, [], store, lambda _: {"facts": [fact]})
    query = {"doc": CITY_DOC, "file": CITY_OUTPUT, "row": 2, "column": 2}
    old = client.get("/cells", params=query).json()
    assert old["text"] is None and old["evidence_status"] == "stale"
    latest = client.post("/fill", json={"doc": CITY_DOC, "template": CITY_TEMPLATE, "source_sha256": digest})
    assert latest.status_code == 200
    assert latest.headers["X-A23-Result-ID"] != original["result_id"]
    assert client.get("/reports", params={"doc": CITY_DOC}).status_code == 404
    assert client.get("/cells", params={**query, "result_id": original["result_id"]}).status_code == 409
    new = client.get("/cells", params=query).json()
    assert "GDP总量200亿元" in new["text"] and new["evaluation"] is None


def test_wrong_refusal_and_missing_cell_are_distinct(tmp_path, monkeypatch):
    inputs, outputs, _, _ = build_demo(tmp_path)
    monkeypatch.setattr(main, "DATA", outputs)
    monkeypatch.setattr(main, "store", Store(outputs / "a23.db"))
    client = TestClient(main.app)
    reference = inputs / REFERENCE_FILES[CITY_OUTPUT]
    workbook = load_workbook(reference)
    workbook.active["B2"] = 999
    workbook.save(reference)
    workbook.close()
    evaluate(outputs / CITY_OUTPUT, reference, outputs / "eval_report.json", CITY_DOC, False)
    city = {"doc": CITY_DOC, "file": CITY_OUTPUT, "row": 2, "column": 2}
    wrong = client.get("/cells", params=city).json()
    assert wrong["evaluation"] == "值不符"
    daily = {"doc": DAILY_DOC, "file": COVID_OUTPUT, "row": 3, "column": 4}
    refused = client.get("/cells", params=daily).json()
    assert refused["decision"] == "refused" and refused["evidence_status"] == "available"
    trace_path = outputs / "fill_report.json"
    trace = json.loads(trace_path.read_text())
    trace["cells"] = [cell for cell in trace["cells"] if (cell["row"], cell["column"]) != (2, 2)]
    write_json(trace_path, trace)
    assert client.get("/cells", params=city).json()["decision"] == "missing"


def test_reference_excel_serial_is_compared_in_the_date_column(tmp_path):
    inputs, outputs, _, _ = build_demo(tmp_path)
    reference = inputs / REFERENCE_FILES[COVID_OUTPUT]
    workbook = load_workbook(reference)
    value = workbook.active["A2"].value
    workbook.active["A2"] = to_excel(value, workbook.epoch)
    workbook.active["A2"].number_format = "General"
    workbook.save(reference)
    workbook.close()
    report = evaluate(outputs / COVID_OUTPUT, reference, outputs / "eval_report_covid19.json", DAILY_DOC, True)
    assert report["correct_rate"] == 1 and report["diff"] == []


def test_formal_rules_refuse_partial_and_unbounded_fixed_counts():
    config = load_expectations()
    with pytest.raises(ValueError, match="同时包含"):
        checked_rules({**config, "same-db": {"01": config["same-db"]["01"]}}, "rebuilt")
    exact = config["same-db"]["01"]
    count = exact["rows"]
    exact["rows"] = [1, count]
    with pytest.raises(ValueError, match=r"01.rows 必须使用精确"):
        checked_rules(config, "rebuilt")
    exact["rows"] = count
    config["rebuilt"]["01"]["cell_accuracy"] = [0.0, 1.0]
    with pytest.raises(ValueError, match="两个独立重抽"):
        checked_rules(config, "rebuilt")
    config = load_expectations()
    config["observations"][1]["extraction_id"] = config["observations"][0]["extraction_id"]
    with pytest.raises(ValueError, match="两个独立重抽"):
        checked_rules(config, "rebuilt")


def test_valid_extraction_survives_a_malformed_fact(tmp_path):
    store = Store(tmp_path / "a23.db")
    valid = {"block_no": 1, "metric": "GDP", "value": 10, "snippet": "GDP10", "unit": None}
    result = extract_document(b"GDP10", "输入.txt", [], store, lambda _: {"facts": [valid, {"block_no": 1}]})
    assert result["kept"] == 1 and result["rejected"] == {"事实字段不符合约定": 1}


def test_fact_export_preserves_two_documents_with_the_same_suffix(tmp_path):
    store = Store(tmp_path / "a23.db")
    for name in ("空结果.txt", "另一份.txt"):
        content = name.encode()
        store.save(parse_document(content, name), content)
    first, = export_facts(store, [tmp_path / "空结果.txt"], tmp_path)
    second, = export_facts(store, [tmp_path / "另一份.txt"], tmp_path)
    assert first["file"] != second["file"]
    saved = json.loads((tmp_path / "facts_export_manifest.json").read_text())
    assert saved == [first, second]


@pytest.mark.parametrize("case,key", [
    (case, key) for case, rules in load_expectations()["demo"].items() for key in rules
])
def test_delivered_report_predicate(case, key):
    directory = Path(os.getenv("A23_TEST_REPORT_DIR", str(DEMO_DELIVERY)))
    rules = load_expectations()["demo"]
    reports = checked_outputs(directory, "reference", require_database=False)
    assert not (failures := check(reports, {case: {key: rules[case][key]}})), "\n".join(failures)


def test_extraction_audit_replays_retries_and_preserves_fact_distinctions(tmp_path):
    content = "主营业务收入总额10亿元，营业收入11亿元。".encode()
    first = {"block_no": 1, "metric": "主营业务收入总额", "value": "10", "snippet": "主营业务收入总额10亿元"}
    separate = {"block_no": 1, "metric": "营业收入", "value": "11", "snippet": "营业收入11亿元"}
    replies = iter([{"facts": False}, {"facts": [first, separate, {**first, "block_no": 2}]}])
    original = {}
    extract_document(content, "输入.txt", ["主营业务收入"], Store(tmp_path / "original.db"),
                     lambda _: next(replies), audit=original)
    batch, = original["batches"]
    assert len(batch["attempts"]) == 2 and batch["attempts"][0]["error"]
    assert [item["rejected"] for item in batch["decisions"]] == [None, "指标不在清单内", "块号不在本批"]
    replays = [replay_run(tmp_path / "输入.txt", content, original, Store(tmp_path / f"replay-{i}.db")) for i in range(2)]
    comparison = compare_runs(*replays)
    assert comparison["shared"] == 1 and not comparison["first_only"] and not comparison["second_only"]
    assert all(run["traceable"] == 1 for run in comparison["runs"])


def test_same_content_refill_cannot_return_prior_run_score(client, demo):
    old = demo[2]
    filled = client.post("/fill", json={"doc": CITY_DOC, "template": CITY_TEMPLATE})
    assert filled.headers["X-A23-Result-ID"] == old["result_id"]
    assert float(filled.headers["X-A23-Fill-Seconds"]) >= 0
    assert client.get("/reports", params={"doc": CITY_DOC, "result_id": old["result_id"]}).status_code == 404
    result = client.get("/cells", params={"doc": CITY_DOC, "file": CITY_OUTPUT, "row": 2, "column": 2}).json()
    assert result["evaluation"] is None
    assert client.get("/files", params={"name": CITY_OUTPUT, "result_id": old["result_id"]}).status_code == 200
    assert client.get("/files", params={"name": CITY_OUTPUT, "result_id": "old"}).status_code == 409


def test_daily_empty_database_upload_fill_score_download_and_evidence(tmp_path, monkeypatch):
    inputs, outputs = tmp_path / "doc", tmp_path / "out"
    prepare_demo_inputs(inputs)
    store = Store(outputs / "a23.db")
    monkeypatch.setattr(main, "DATA", outputs)
    monkeypatch.setattr(main, "INPUTS", inputs)
    monkeypatch.setattr(main, "REFERENCE_DIR", inputs)
    monkeypatch.setattr(main, "SYNTHETIC_REFERENCE", True)
    monkeypatch.setattr(main, "store", store)
    client = TestClient(main.app)
    upload = {"file": (DAILY_DOC, (inputs / DAILY_DOC).read_bytes())}
    assert not store.path.exists()
    parsed = client.post("/parse", files=upload)
    assert parsed.status_code == 200 and not store.path.exists()
    assert client.post("/ingest", files=upload).status_code == 200
    filled = client.post("/fill", json={
        "doc": DAILY_DOC, "template": DAILY_TEMPLATE, "source_sha256": parsed.headers["X-A23-Source-SHA256"],
    })
    assert filled.status_code == 200
    query = {"doc": DAILY_DOC, "file": COVID_OUTPUT, "result_id": filled.headers["X-A23-Result-ID"]}
    report = client.get("/reports", params=query).json()
    assert report["refusal_counts"] == {"source_blank": report["expected_empty"]}
    assert client.get("/files", params={"name": COVID_OUTPUT, "result_id": query["result_id"]}).status_code == 200
    detail = client.get("/cells", params={**query, "row": 3, "column": 4}).json()
    assert detail["decision"] == "refused" and detail["reason_code"] == "source_blank"
    assert detail["source"]["locator"] == {"sheet": "Sheet", "row": 3}


def test_real_corpus_fusion_conflict(tmp_path, monkeypatch):
    data = json.loads((DEMO_DELIVERY.parent / "source" / "fusion-real.json").read_bytes())
    corpus = tmp_path / "corpus"
    unpack_dataset(ROOT / "dataset" / "测试集.zip", corpus)
    store = Store(tmp_path / "a23.db")
    entity, metric, period, scope, unit = (data[key] for key in ("entity", "metric", "period", "scope", "unit"))
    for candidate in data["candidates"]:
        path = corpus / candidate["path"]
        content = path.read_bytes()
        doc = parse_document(content, path.name)
        blocks = dict(doc.numbered_blocks())
        block_no, period_block_no, value, snippet = (
            candidate[key] for key in ("block_no", "period_block_no", "value", "snippet")
        )
        fact = RawFact(block_no=block_no, metric=metric, value=value, snippet=snippet)
        assert evidence_error(fact, blocks) is None and period in blocks[period_block_no].text
        store.save(doc, content, "list", [
            (block_no, entity, metric, metric, value, value, unit, snippet, period, period_block_no, scope),
        ])
    facts = store.facts()
    evidence = resolve_facts(facts, entity, metric, period, scope)
    assert {fact["value"] for fact in evidence["candidates"]} == {"1037", "1029.8"}
    assert evidence["selected"] is None
    preferred = resolve_facts(facts, entity, metric, period, scope, preferred_doc=Path(data["candidates"][1]["path"]).name)
    assert preferred["selected"]["value"] == "1029.8" and len(preferred["candidates"]) == 2
    monkeypatch.setattr(main, "store", store)
    client = TestClient(main.app)
    criteria = {"entity": entity, "metric": metric, "period": period, "scope": scope,
                "preferred_doc": preferred["selected"]["doc"]}
    response = client.post("/facts/decisions", json=criteria)
    assert response.status_code == 201
    receipt = response.json()
    assert receipt["criteria"] == criteria and receipt["selected"]["value"] == "1029.8"
    assert all(len(candidate["source_sha256"]) == 64 for candidate in receipt["candidates"])
    raw = "新版统计公报".encode()
    store.save(parse_document(raw, criteria["preferred_doc"]), raw)
    history = client.get(f"/facts/decisions/{receipt['decision_id']}").json()
    assert history["selected"] == receipt["selected"]
    assert client.get(f"/facts/decisions/{receipt['decision_id'] + 1}").status_code == 404


def test_city_refusals_explain_missing_fact_and_incompatible_unit(tmp_path):
    inputs, output, _, _ = build_demo(tmp_path)
    store = Store(output / "a23.db")
    extract_document((inputs / CITY_DOC).read_bytes(), CITY_DOC, [], store, lambda _: {"facts": [
        {"block_no": 1, "metric": "GDP总量", "entity": "甲城", "value": "100",
         "snippet": "GDP总量100亿元", "unit": "万元"},
    ]})
    fill_workbook(store, CITY_DOC, inputs / CITY_TEMPLATE, output)
    report = evaluate(output / CITY_OUTPUT, inputs / REFERENCE_FILES[CITY_OUTPUT],
                      output / "eval_report.json", CITY_DOC, False)
    assert report["refusal_counts"] == {"incompatible_value": 1, "missing_fact": 3}
    assert report["missed"] == 4 and report["wrong_values"] == 0
    trace = json.loads((output / REPORT_FILES[CITY_OUTPUT][0]).read_bytes())
    assert all(cell["source"]["locator"] == {"paragraph": 1} for cell in trace["cells"])


def test_city_empty_database_extract_with_model_substitute_and_fill(tmp_path, monkeypatch):
    inputs, outputs = tmp_path / "doc", tmp_path / "out"
    prepare_demo_inputs(inputs)
    store = Store(outputs / "a23.db")
    monkeypatch.setattr(main, "DATA", outputs)
    monkeypatch.setattr(main, "INPUTS", inputs)
    monkeypatch.setattr(main, "REFERENCE_DIR", inputs)
    monkeypatch.setattr(main, "SYNTHETIC_REFERENCE", True)
    monkeypatch.setattr(main, "store", store)
    monkeypatch.setattr(main, "extract_document", lambda data, name, metrics, target, *, publish: extract_document(
        data, name, metrics, target, demo_response, publish=publish,
    ))
    client = TestClient(main.app)
    request = {"doc": CITY_DOC, "template": CITY_TEMPLATE}
    assert not store.path.exists()
    assert client.post("/fill", json=request).status_code == 400
    upload = {"file": (CITY_DOC, (inputs / CITY_DOC).read_bytes())}
    response = client.post("/extract", files=upload)
    assert response.status_code == 200 and response.json()["kept"] == len(store.facts(CITY_DOC))
    response = client.post("/fill", json=request)
    assert response.status_code == 200
    assert client.get("/reports", params={"doc": CITY_DOC}).json()["cell_accuracy"] == 1


def test_score_error_does_not_hide_a_successful_fill(client, demo, monkeypatch):
    inputs, _, _, _ = demo
    monkeypatch.setattr(main, "REFERENCE_DIR", inputs)
    (inputs / REFERENCE_FILES[CITY_OUTPUT]).write_bytes(b"broken reference")
    response = client.post("/fill", json={"doc": CITY_DOC, "template": CITY_TEMPLATE})
    assert response.status_code == 200 and response.headers["X-A23-Score-Status"] == "failed"
    assert client.get("/files", params={"name": CITY_OUTPUT, "result_id": response.headers["X-A23-Result-ID"]}).status_code == 200
    assert client.get("/reports", params={"doc": CITY_DOC}).status_code == 404
