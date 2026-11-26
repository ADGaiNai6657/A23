import json
from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document
from openpyxl import load_workbook

from acceptance import DELIVERABLES, checked_outputs
from check_dataset import CITY_FOLDER, CITY_SOURCE, check_dataset
from config import ROOT
from criteria import MODEL_FIELDS, load_expectations
from demo import CITY_TEMPLATE, make_workbook
from evaluate import evaluate
from extractor import extract_document
from fill_spec import CITY_HEADERS, CITY_OUTPUT, COVID_OUTPUT, REPORT_FILES
from filler import fill_workbook
from rebuild_evidence import RUN_EVIDENCE, check_execution, digest
from store import Store
from unpack_dataset import unpack_dataset
from verify import verify_samples


@pytest.mark.parametrize("clause, quote, value, scope, code", [
    ("GDP总量100亿元，一般公共预算收入10亿元", "一般公共预算收入10亿元", "10", "", "metric_mismatch"),
    ("GDP总量100亿元，一般公共预算收入10亿元", "10亿元", "10", "", "metric_mismatch"),
    ("GDP总量100亿元，一般公共预算收入10亿元", "GDP总量100亿元，一般公共预算收入10亿元", "10", "", "metric_mismatch"),
    ("GDP总量100亿元", "GDP总量100亿元", "100", "", None),
    ("实现100亿元GDP总量", "100亿元GDP总量", "100", "", None),
    ("2024年一季度GDP总量25亿元", "GDP总量25亿元", "25", "", "period_or_scope"),
    ("2024年上半年GDP总量25亿元", "GDP总量25亿元", "25", "全年", "period_or_scope"),
    ("截至2024年6月，GDP总量25亿元", "GDP总量25亿元", "25", "", "period_or_scope"),
    ("一季度GDP总量25亿元，全年GDP总量100亿元", "GDP总量100亿元", "100", "", None),
])
def test_city_quantity_is_bound_to_source_metric_and_scope(tmp_path, clause, quote, value, scope, code):
    doc, template = tmp_path / "城市.docx", tmp_path / CITY_TEMPLATE
    document = Document()
    document.add_paragraph("2024年中国城市经济百强全景报告")
    document.add_paragraph(f"甲城：{clause}。常住人口20万人，人均GDP50000元，一般公共预算收入10亿元。")
    document.save(doc)
    make_workbook(template, CITY_HEADERS)
    store = Store(tmp_path / "a23.db")
    content = {"facts": [{"block_no": 2, "entity": "甲城", "metric": "GDP总量",
                          "value": value, "unit": "亿元", "snippet": quote, "scope": scope}]}
    extract_document(doc.read_bytes(), doc.name, [], store, lambda _: content)
    fill_workbook(store, doc.name, template, tmp_path)
    trace = json.loads((tmp_path / REPORT_FILES[CITY_OUTPUT][0]).read_bytes())
    city_gdp = trace["cells"][1]
    assert city_gdp.get("reason_code") == code
    if code:
        assert city_gdp["value"] is None and city_gdp["conflicts"][0]["source"]["block_no"] == 2
        assert bool(city_gdp["conflicts"][0]["fact"]["source_scope"]) == (code == "period_or_scope")
    else:
        assert city_gdp["value"] == value


@pytest.fixture(scope="module")
def corpus(tmp_path_factory):
    folder = tmp_path_factory.mktemp("audit-corpus")
    unpack_dataset(ROOT / "dataset" / "测试集.zip", folder)
    return folder


@pytest.mark.parametrize("damage", ["none", "missing-batch", "duplicate-batch", "no-receipt", "wrong-document"])
def test_online_package_checks_real_batch_chain_even_if_digest_is_updated(tmp_path, corpus, damage):
    with ZipFile(DELIVERABLES / "follow-up" / "city-online.zip") as bundle:
        bundle.extractall(tmp_path)
    reports = checked_outputs(tmp_path, "source", require_database=False)
    source = corpus / "包含模板文件" / CITY_FOLDER / CITY_SOURCE
    manifest = tmp_path / RUN_EVIDENCE
    evidence = json.loads(manifest.read_bytes())
    log = tmp_path / "extraction-runs" / f"{evidence['extraction_id']}.json"
    record = json.loads(log.read_bytes())
    if damage == "none":
        check_execution(tmp_path, "rebuilt", reports, archived_input=source)
        return
    if damage == "no-receipt":
        manifest.unlink()
        error = RUN_EVIDENCE
    else:
        if damage == "missing-batch":
            record["batches"].pop()
            error = "缺少批次响应"
        elif damage == "duplicate-batch":
            record["batches"][-1]["blocks"] = record["batches"][0]["blocks"]
            error = "批次覆盖"
        else:
            record["doc"] = "另一份.docx"
            error = "原文字节"
        log.write_text(json.dumps(record, ensure_ascii=False), encoding="utf-8")
        evidence["extraction_sha256"] = digest(log)
        manifest.write_text(json.dumps(evidence), encoding="utf-8")
    with pytest.raises(ValueError, match=error):
        check_execution(tmp_path, "rebuilt", reports, archived_input=source)


def test_independent_reference_profile_with_point_998_and_version_mismatch(tmp_path, corpus):
    same, rebuilt = tmp_path / "same", tmp_path / "rebuilt"
    historical = DELIVERABLES / "source"
    check_dataset(corpus, same, False, historical / "city-original-historical.json")
    check_dataset(corpus, same, False)
    check_dataset(corpus, rebuilt, False, historical / "city-rebuilt-historical.json")
    fields = load_expectations()["same-db"]
    references = {}
    for case, filename in (("01", CITY_OUTPUT), ("03", COVID_OUTPUT)):
        trace_name, score_name, _ = REPORT_FILES[filename]
        reference = tmp_path / f"reference-{case}.xlsx"
        book = load_workbook(same / filename)
        if case == "01":
            book.active["B2"] = book.active["B2"].value + 1
        book.save(reference)
        book.close()
        references[case] = digest(reference)
        for directory in (same, rebuilt):
            trace = json.loads((directory / trace_name).read_bytes())
            score = evaluate(directory / filename, reference, directory / score_name, trace["doc"], case == "03",
                             expected_rows=fields[case]["rows"], block_count=trace["block_count"],
                             database=directory / "a23.db")
            if case == "01":
                assert score["cell_accuracy"] == .998
    fields["01"]["cell_accuracy"] = .998
    config = {
        "basis": "reference-workbook", "reference_sha256": references, "same-db": fields,
        "rebuilt": {"01": {key: [fields["01"][key]] * 2 for key in MODEL_FIELDS}},
        "observations": [{"data_dir": "same", "mode": "same-db"},
                         {"data_dir": "rebuilt", "mode": "rebuilt", "rebuilt_kind": "history"}],
    }
    profile = tmp_path / "reference-rules.json"
    profile.write_text(json.dumps(config), encoding="utf-8")
    verify_samples(same, rebuilt, "reference", "history", profile)
    score_path = rebuilt / REPORT_FILES[CITY_OUTPUT][1]
    original = score_path.read_bytes()
    changed = json.loads(original)
    changed["reference_sha256"] = "a" * 64
    score_path.write_text(json.dumps(changed), encoding="utf-8")
    with pytest.raises(ValueError, match="参考表 SHA256"):
        verify_samples(same, rebuilt, "reference", "history", profile)
    score_path.write_bytes(original)
    with pytest.raises(ValueError, match="basis"):
        verify_samples(same, rebuilt, "source", "history", profile)
