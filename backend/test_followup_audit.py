import json
from pathlib import Path
from shutil import copytree
from zipfile import ZipFile

import pytest
from docx import Document
from fastapi.testclient import TestClient
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from acceptance import DELIVERABLES, checked_outputs, sample_rules
from check_dataset import CITY_FOLDER, CITY_SOURCE, CITY_TEMPLATE as PUBLIC_CITY_TEMPLATE, check_dataset
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
import main


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
    ("2024年GDP总量100亿元，GDP同比增加10亿元", "GDP同比增加10亿元", "10", "", "metric_mismatch"),
    ("GDP增速6%", "GDP增速6%", "6", "", "metric_mismatch"),
    ("GDP比上年减少10亿元", "GDP比上年减少10亿元", "10", "", "metric_mismatch"),
    ("GDP增长至100亿元", "GDP增长至100亿元", "100", "", None),
    ("2024年第三季GDP总量25亿元", "第三季GDP总量25亿元", "25", "", "period_or_scope"),
    ("2024年一至六月GDP总量25亿元", "一至六月GDP总量25亿元", "25", "", "period_or_scope"),
    ("2024年1—6月GDP总量25亿元", "GDP总量25亿元", "25", "全年", "period_or_scope"),
    ("全年GDP总量100亿元（其中一季度为25亿元）", "全年GDP总量100亿元", "100", "全年", None),
    ("全年GDP总量100亿元（第三季GDP总量25亿元）", "全年GDP总量100亿元", "100", "", None),
    ("全年GDP总量100亿元（第三季GDP总量25亿元）", "第三季GDP总量25亿元", "25", "", "period_or_scope"),
    ("GDP总量25亿元（第三季）", "GDP总量25亿元", "25", "", "period_or_scope"),
    ("全年GDP总量100亿元及第三季GDP总量25亿元", "全年GDP总量100亿元", "100", "", None),
    ("第三季GDP总量25亿元及全年GDP总量100亿元", "全年GDP总量100亿元", "100", "", None),
    ("第三季GDP总量25亿元（全年GDP总量100亿元）", "全年GDP总量100亿元", "100", "", None),
    ("GDP于2024年全年达到100亿元", "GDP于2024年全年达到100亿元", "100", "", None),
    ("2024年GDP总量100亿元，GDP净增10亿元", "GDP净增10亿元", "10", "", "metric_mismatch"),
    ("2024年GDP总量100亿元，GDP比2023年多10亿元", "GDP比2023年多10亿元", "10", "", "metric_mismatch"),
    ("人口净减10万人，净增GDP10亿元", "净增GDP10亿元", "10", "", "metric_mismatch"),
    ("GDP净增由5亿元提高到10亿元", "GDP净增由5亿元提高到10亿元", "10", "", "metric_mismatch"),
    ("2024年GDP从90亿元增长至100亿元", "GDP从90亿元增长至100亿元", "100", "", None),
    ("2024年GDP从90亿元增长至100亿元", "GDP从90亿元增长至100亿元", "90", "", "metric_mismatch"),
    ("GDP由90亿元同比升至100亿元", "100亿元", "100", "", None),
    ("GDP从90亿元增长至100%", "100", "100", "", "metric_mismatch"),
    ("实际由90亿元升至100亿元的GDP总量", "由90亿元升至100亿元的GDP总量", "100", "", None),
    ("GDP介于90亿元至100亿元之间", "GDP介于90亿元至100亿元之间", "90", "", "metric_mismatch"),
    ("GDP介于90亿元至100亿元之间", "GDP介于90亿元至100亿元之间", "100", "", "metric_mismatch"),
    ("2024年前三个季度GDP总量25亿元", "前三个季度GDP总量25亿元", "25", "", "period_or_scope"),
    ("2024年前3个季度GDP总量25亿元", "GDP总量25亿元", "25", "全年", "period_or_scope"),
    ("2024年截至9月30日，GDP总量25亿元", "GDP总量25亿元", "25", "", "period_or_scope"),
    ("截止到2024年九月三十号，GDP总量25亿元", "GDP总量25亿元", "25", "", "period_or_scope"),
    ("全年GDP总量100亿元（第三季数据未公布）", "全年GDP总量100亿元", "100", "全年", None),
    ("全年GDP总量100亿元（第三季统计，另行公布）", "全年GDP总量100亿元", "100", "全年", None),
    ("GDP总量25亿元（统计口径为第三季）", "GDP总量25亿元", "25", "", "period_or_scope"),
    ("第三季GDP总量25亿元（全年数据未公布）", "第三季GDP总量25亿元", "25", "", "period_or_scope"),
    ("截至9月30日，全年GDP总量100亿元", "全年GDP总量100亿元", "100", "", None),
])
def test_city_quantity_is_bound_to_source_metric_and_scope(tmp_path, clause, quote, value, scope, code):
    cells, report = fill_city_claims(tmp_path, clause, [(quote, value, scope)])
    city_gdp = report["B2"]
    assert cells["B2"] == (None if code else float(value))
    assert city_gdp.get("reason_code") == code
    if code:
        assert city_gdp["value"] is None and city_gdp["conflicts"][0]["source"]["block_no"] == 2
        assert bool(city_gdp["conflicts"][0]["fact"]["source_scope"]) == (code == "period_or_scope")
    else:
        assert city_gdp["value"] == value


def fill_city_claims(tmp_path, clause, gdp_claims, budget_claims=()):
    doc, template = tmp_path / "城市.docx", tmp_path / CITY_TEMPLATE
    document = Document()
    document.add_paragraph("2024年中国城市经济百强全景报告")
    document.add_paragraph(f"甲城：{clause}。常住人口20万人，人均GDP50000元。")
    document.save(doc)
    make_workbook(template, CITY_HEADERS)
    store = Store(tmp_path / "a23.db")
    content = {"facts": [{"block_no": 2, "entity": "甲城", "metric": metric,
                          "value": value, "unit": "亿元", "snippet": quote, "scope": scope}
                         for metric, claims in (("GDP总量", gdp_claims), ("一般公共预算收入", budget_claims))
                         for quote, value, scope in claims]}
    extract_document(doc.read_bytes(), doc.name, [], store, lambda _: content)
    fill_workbook(store, doc.name, template, tmp_path)
    trace = json.loads((tmp_path / REPORT_FILES[CITY_OUTPUT][0]).read_bytes())
    workbook = load_workbook(tmp_path / CITY_OUTPUT, read_only=True)
    try:
        values = {address: workbook.active[address].value for address in ("B2", "E2")}
    finally:
        workbook.close()
    return values, {f'{get_column_letter(item["column"])}{item["row"]}': item for item in trace["cells"]}


def test_city_total_remains_usable_when_change_and_partial_period_are_also_candidates(tmp_path):
    cells, report = fill_city_claims(
        tmp_path, "全年GDP总量100亿元（第三季GDP总量25亿元），GDP同比增加10亿元",
        [("GDP同比增加10亿元", "10", ""), ("第三季GDP总量25亿元", "25", "全年"),
         ("全年GDP总量100亿元", "100", "")],
    )
    assert cells["B2"] == 100
    assert report["B2"]["value"] == "100"
    assert report["B2"]["fact"]["snippet"] == "全年GDP总量100亿元"


def test_city_uses_the_transition_destination_when_all_roles_are_candidates(tmp_path):
    cells, report = fill_city_claims(
        tmp_path, "全年GDP从90亿元增长至100亿元，GDP净增10亿元",
        [("GDP从90亿元增长至100亿元", "90", ""), ("GDP净增10亿元", "10", ""),
         ("GDP从90亿元增长至100亿元", "100", "")],
    )
    assert cells["B2"] == 100
    assert report["B2"]["value"] == "100"
    assert report["B2"]["fact"]["snippet"] == "GDP从90亿元增长至100亿元"


@pytest.mark.parametrize("clause, quote, code, reason", [
    ("GDP目标100亿元", "GDP目标100亿元", "goal", "目标"),
    ("预计2024年全年GDP达100亿元", "预计2024年全年GDP达100亿元", "forecast", "预测"),
    ("预计2024年全年GDP达100亿元", "GDP达100亿元", "forecast", "预测"),
    ("2024年全年GDP预计达100亿元", "GDP预计达100亿元", "forecast", "预测"),
])
def test_city_rejects_non_actual_gdp(tmp_path, clause, quote, code, reason):
    cells, report = fill_city_claims(tmp_path, clause, [(quote, "100", "全年")])
    detail = report["B2"]
    assert cells["B2"] is None
    assert detail["value"] is None
    assert detail["reason_code"] == code
    assert reason in detail["reason"]
    assert detail["conflicts"][0]["fact"]["snippet"] == quote
    assert detail["conflicts"][0]["source"]["block_no"] == 2


@pytest.mark.parametrize("goal, actual", [
    ("全年GDP目标100亿元", "实际GDP总量90亿元"),
    ("预计全年GDP达100亿元", "经初步核算，全年GDP总量90亿元"),
    ("", "初步核算，全年GDP总量90亿元"),
])
def test_city_selects_reported_gdp(tmp_path, goal, actual):
    claims = [(goal, "100", "")] if goal else []
    claims.append((actual, "90", ""))
    cells, report = fill_city_claims(tmp_path, "，".join(filter(None, (goal, actual))), claims)
    detail = report["B2"]
    assert cells["B2"] == 90
    assert detail["value"] == "90"
    assert "reason_code" not in detail
    assert detail["fact"]["snippet"] == actual
    assert detail["fact"]["block_no"] == 2


def test_city_inherits_date_across_a_reporting_aside(tmp_path):
    quote = "GDP总量25亿元"
    cells, report = fill_city_claims(tmp_path, "2024年截至9月30日，初步核算，" + quote,
                                    [(quote, "25", "全年")])
    detail = report["B2"]
    assert cells["B2"] is None
    assert detail["value"] is None
    assert detail["reason_code"] == "period_or_scope"
    assert "全年口径" in detail["reason"]
    assert detail["conflicts"][0]["fact"]["source_scope"] == "截至9月30日"
    assert detail["conflicts"][0]["fact"]["snippet"] == quote


def test_city_does_not_apply_annual_goal_scope_to_partial_actual(tmp_path):
    goal, actual = "全年GDP目标100亿元", "实际GDP总量25亿元"
    cells, report = fill_city_claims(tmp_path, f"截至9月30日，{goal}，{actual}",
                                    [(goal, "100", ""), (actual, "25", "")])
    detail = report["B2"]
    assert cells["B2"] is None
    assert detail["value"] is None
    assert detail["reason_code"] == "period_or_scope"
    assert {item["fact"]["snippet"]: item["fact"]["source_scope"] for item in detail["conflicts"]} == {
        goal: "全年", actual: "截至9月30日",
    }


@pytest.mark.parametrize("clause, budget_quote, budget_value, budget_scope", [
    ("2024年截至9月30日，GDP总量25亿元、一般公共预算收入10亿元",
     "一般公共预算收入10亿元", None, "截至9月30日"),
    ("2024年截至9月30日GDP总量25亿元、一般公共预算收入10亿元",
     "一般公共预算收入10亿元", None, "截至9月30日"),
    ("2024年截至9月30日，GDP总量25亿元、全年一般公共预算收入10亿元",
     "全年一般公共预算收入10亿元", 10, "全年"),
    ("2024年截至9月30日，GDP总量25亿元，一般公共预算收入10亿元（全年）",
     "一般公共预算收入10亿元（全年）", 10, "全年"),
    ("2024年截至9月30日，GDP总量25亿元、全年常住人口20万人、一般公共预算收入10亿元",
     "一般公共预算收入10亿元", 10, "全年"),
    ("2024年截至9月30日，GDP总量25亿元、上半年一般公共预算收入10亿元",
     "上半年一般公共预算收入10亿元", None, "上半年"),
    ("2023年截至9月30日，GDP总量25亿元、2024年一般公共预算收入10亿元",
     "2024年一般公共预算收入10亿元", 10, None),
    ("2024年截至9月30日，GDP总量25亿元。一般公共预算收入10亿元",
     "一般公共预算收入10亿元", 10, None),
])
def test_city_applies_date_to_coordinated_metrics(tmp_path, clause, budget_quote, budget_value, budget_scope):
    gdp_quote = "GDP总量25亿元"
    cells, report = fill_city_claims(tmp_path, clause, [(gdp_quote, "25", "")], [(budget_quote, "10", "")])
    assert cells == {"B2": None, "E2": budget_value}
    assert report["B2"]["value"] is None
    assert report["B2"]["reason_code"] == "period_or_scope"
    assert report["B2"]["conflicts"][0]["fact"]["source_scope"] == "截至9月30日"
    assert report["B2"]["conflicts"][0]["fact"]["snippet"] == gdp_quote
    budget = report["E2"]
    if budget_value is None:
        assert budget["value"] is None
        assert budget["reason_code"] == "period_or_scope"
        assert "全年口径" in budget["reason"]
        budget_fact = budget["conflicts"][0]["fact"]
    else:
        assert budget["value"] == "10"
        assert "reason_code" not in budget
        budget_fact = budget["fact"]
    assert budget_fact["source_scope"] == budget_scope
    assert budget_fact["snippet"] == budget_quote
    assert budget_fact["block_no"] == 2


@pytest.mark.parametrize("suffix, code", [
    ("（目标值）", "goal"), ("（预测值）", "forecast"),
    ("目标值", "goal"), ("预测值", "forecast"),
])
def test_city_rejects_postfixed_non_actual_role(tmp_path, suffix, code):
    quote = f"2024年全年GDP总量100亿元{suffix}"
    cells, report = fill_city_claims(tmp_path, quote, [(quote, "100", "")])
    assert cells["B2"] is None
    assert report["B2"]["reason_code"] == code
    assert report["B2"]["conflicts"][0]["fact"]["snippet"] == quote


def test_city_shared_forecast_and_explicit_actual(tmp_path):
    clause = "预计2024年全年GDP达100亿元、一般公共预算收入达10亿元"
    cells, report = fill_city_claims(tmp_path, clause,
                                    [("GDP达100亿元", "100", "")],
                                    [("一般公共预算收入达10亿元", "10", "")])
    assert cells == {"B2": None, "E2": None}
    assert report["B2"]["reason_code"] == report["E2"]["reason_code"] == "forecast"

    clause += "、实际一般公共预算收入10亿元"
    cells, report = fill_city_claims(tmp_path, clause,
                                    [("GDP达100亿元", "100", "")],
                                    [("一般公共预算收入达10亿元", "10", ""),
                                     ("实际一般公共预算收入10亿元", "10", "")])
    assert cells == {"B2": None, "E2": 10}
    assert report["E2"]["fact"]["snippet"] == "实际一般公共预算收入10亿元"


def test_city_ignores_unrelated_prediction_note(tmp_path):
    quote = "全年GDP总量（预测值仅作比较）90亿元"
    cells, report = fill_city_claims(tmp_path, quote, [(quote, "90", "")])
    assert cells["B2"] == 90
    assert report["B2"]["fact"]["snippet"] == quote


@pytest.mark.parametrize("budget", [
    "一般公共预算收入在2024年达到10亿元",
    "2024年一般公共预算收入达到10亿元",
])
def test_city_new_year_resets_old_cutoff(tmp_path, budget):
    gdp = "GDP总量25亿元"
    clause = f"截至2023年9月30日，{gdp}、{budget}"
    cells, report = fill_city_claims(tmp_path, clause,
                                    [(gdp, "25", "")], [(budget, "10", "")])
    assert cells == {"B2": None, "E2": 10}
    assert report["B2"]["reason_code"] == "period_or_scope"
    assert report["B2"]["conflicts"][0]["fact"]["source_scope"] == "截至2023年9月30日"
    assert report["E2"]["fact"]["period"] == "2024"
    assert report["E2"]["fact"]["source_scope"] is None


def test_city_combined_reporting_aside_inherits_cutoff(tmp_path):
    budget = "一般公共预算收入10亿元"
    clause = f"截至9月30日，GDP总量25亿元，经初步核算口径调整后，{budget}"
    cells, report = fill_city_claims(tmp_path, clause, [("GDP总量25亿元", "25", "")],
                                    [(budget, "10", "")])
    assert cells == {"B2": None, "E2": None}
    assert report["E2"]["reason_code"] == "period_or_scope"
    assert report["E2"]["conflicts"][0]["fact"]["source_scope"] == "截至9月30日"


def test_city_local_forecast_does_not_cover_next_metric(tmp_path):
    clause = "2024年全年GDP总量100亿元（预测值）、一般公共预算收入达10亿元"
    cells, report = fill_city_claims(tmp_path, clause,
                                    [("GDP总量100亿元（预测值）", "100", "")],
                                    [("一般公共预算收入达10亿元", "10", "")])
    assert cells == {"B2": None, "E2": 10}
    assert report["B2"]["reason_code"] == "forecast"
    assert report["E2"]["fact"]["snippet"] == "一般公共预算收入达10亿元"


def test_city_comparison_phrase_keeps_actual_gdp(tmp_path):
    clause = "2024年全年GDP低于预期但达到90亿元"
    cells, report = fill_city_claims(tmp_path, clause, [("GDP低于预期但达到90亿元", "90", "")])
    assert cells["B2"] == 90
    assert report["B2"]["fact"]["snippet"] == "GDP低于预期但达到90亿元"


def test_city_comparison_year_does_not_reset_cutoff(tmp_path):
    clause = "截至2023年9月30日，GDP总量25亿元、一般公共预算收入10亿元（与2024年目标比较）"
    cells, report = fill_city_claims(tmp_path, clause,
                                    [("GDP总量25亿元", "25", "")],
                                    [("一般公共预算收入10亿元", "10", "")])
    assert cells == {"B2": None, "E2": None}
    assert report["B2"]["reason_code"] == report["E2"]["reason_code"] == "period_or_scope"
    assert report["E2"]["conflicts"][0]["fact"]["source_scope"] == "截至2023年9月30日"


def test_city_shared_forecast_survives_local_postfix(tmp_path):
    clause = "预计2024年全年GDP达100亿元（预测值）、一般公共预算收入达10亿元"
    cells, report = fill_city_claims(tmp_path, clause,
                                    [("GDP达100亿元（预测值）", "100", "")],
                                    [("一般公共预算收入达10亿元", "10", "")])
    assert cells == {"B2": None, "E2": None}
    assert report["B2"]["reason_code"] == report["E2"]["reason_code"] == "forecast"


def test_city_exceeding_expectation_is_actual(tmp_path):
    clause = "2024年全年GDP超出预期达到90亿元"
    cells, report = fill_city_claims(tmp_path, clause, [("GDP超出预期达到90亿元", "90", "")])
    assert cells["B2"] == 90
    assert report["B2"]["fact"]["snippet"] == "GDP超出预期达到90亿元"


def test_city_comparison_growth_note_keeps_old_cutoff(tmp_path):
    clause = "截至2023年9月30日，GDP总量25亿元、一般公共预算收入10亿元（较2024年同口径增加5%）"
    cells, report = fill_city_claims(tmp_path, clause,
                                    [("GDP总量25亿元", "25", "")],
                                    [("一般公共预算收入10亿元", "10", "")])
    assert cells == {"B2": None, "E2": None}
    assert report["E2"]["reason_code"] == "period_or_scope"
    assert report["E2"]["conflicts"][0]["fact"]["source_scope"] == "截至2023年9月30日"


def test_city_above_expectation_is_reported_actual(tmp_path):
    quote = "GDP高于预期并达到90亿元"
    cells, report = fill_city_claims(tmp_path, f"2024年全年{quote}", [(quote, "90", "")])
    assert cells["B2"] == 90
    assert report["B2"]["fact"]["snippet"] == quote


def test_city_pending_actual_note_does_not_cancel_shared_forecast(tmp_path):
    clause = "预计2024年全年GDP达100亿元、一般公共预算收入（实际值待核）达10亿元"
    cells, report = fill_city_claims(tmp_path, clause,
                                    [("GDP达100亿元", "100", "")],
                                    [("一般公共预算收入（实际值待核）达10亿元", "10", "")])
    assert cells == {"B2": None, "E2": None}
    assert report["E2"]["reason_code"] == "forecast"


def test_city_comparison_growth_synonym_keeps_old_cutoff(tmp_path):
    clause = "截至2023年9月30日，GDP总量25亿元、一般公共预算收入10亿元（较2024年同口径增长5%）"
    cells, report = fill_city_claims(tmp_path, clause,
                                    [("GDP总量25亿元", "25", "")],
                                    [("一般公共预算收入10亿元", "10", "")])
    assert cells == {"B2": None, "E2": None}
    assert report["E2"]["reason_code"] == "period_or_scope"
    assert report["E2"]["conflicts"][0]["fact"]["source_scope"] == "截至2023年9月30日"


def test_goal_refusal_is_available_in_reports_and_cells(tmp_path, monkeypatch):
    source = tmp_path / "城市.docx"
    template = tmp_path / CITY_TEMPLATE
    document = Document()
    document.add_paragraph("2024年中国城市经济百强全景报告")
    document.add_paragraph("甲城：GDP目标100亿元，常住人口20万人，人均GDP50000元。")
    document.save(source)
    make_workbook(template, CITY_HEADERS)
    make_workbook(tmp_path / "参考_01.xlsx", CITY_HEADERS, [("甲城", 90, 20, 50000, 10)])
    database = Store(tmp_path / "a23.db")
    extract_document(source.read_bytes(), source.name, [], database, lambda _: {"facts": [{
        "block_no": 2, "entity": "甲城", "metric": "GDP总量", "value": "100",
        "unit": "亿元", "snippet": "GDP目标100亿元",
    }]})
    monkeypatch.setattr(main, "store", database)
    monkeypatch.setattr(main, "DATA", tmp_path)
    monkeypatch.setattr(main, "INPUTS", tmp_path)
    monkeypatch.setattr(main, "REFERENCE_DIR", tmp_path)
    monkeypatch.setattr(main, "SYNTHETIC_REFERENCE", True)
    client = TestClient(main.app)
    filled = client.post("/fill", json={"doc": source.name, "template": template.name})
    assert filled.status_code == 200, filled.text
    query = {"doc": source.name, "file": CITY_OUTPUT, "result_id": filled.headers["X-A23-Result-ID"]}
    report = client.get("/reports", params=query)
    cell = client.get("/cells", params={**query, "row": 2, "column": 2})
    assert report.status_code == cell.status_code == 200
    assert report.json()["refusal_counts"]["goal"] == 1
    assert cell.json()["reason_code"] == "goal"
    assert "目标" in cell.json()["reason"]


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

    def score_with_reference(directory, case, filename, reference):
        trace_name, score_name, _ = REPORT_FILES[filename]
        trace = json.loads((directory / trace_name).read_bytes())
        score = evaluate(directory / filename, reference, directory / score_name, trace["doc"], case == "03",
                         expected_rows=fields[case]["rows"], block_count=trace["block_count"],
                         database=directory / "a23.db")
        if case == "01":
            assert score["cell_accuracy"] == .998

    for case, filename in (("01", CITY_OUTPUT), ("03", COVID_OUTPUT)):
        reference = tmp_path / f"reference-{case}.xlsx"
        book = load_workbook(same / filename)
        if case == "01":
            book.active["B2"] = book.active["B2"].value + 1
        book.save(reference)
        book.close()
        references[case] = digest(reference)
        for directory in (same, rebuilt):
            score_with_reference(directory, case, filename, reference)
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

    copied_corpus = tmp_path / "styled-corpus"
    copytree(corpus, copied_corpus)
    template = copied_corpus / "包含模板文件" / CITY_FOLDER / PUBLIC_CITY_TEMPLATE
    book = load_workbook(template)
    book.active["A1"].font = Font(italic=True)
    book.save(template)
    book.close()
    styled_same, styled_rebuilt = tmp_path / "styled-same", tmp_path / "styled-rebuilt"
    check_dataset(copied_corpus, styled_same, False, historical / "city-original-historical.json")
    check_dataset(copied_corpus, styled_same, False)
    check_dataset(copied_corpus, styled_rebuilt, False, historical / "city-rebuilt-historical.json")
    for case, filename in (("01", CITY_OUTPUT), ("03", COVID_OUTPUT)):
        for directory in (styled_same, styled_rebuilt):
            score_with_reference(directory, case, filename, tmp_path / f"reference-{case}.xlsx")
    reports = checked_outputs(styled_rebuilt, "reference")
    with pytest.raises(ValueError, match=r"01\.template_sha256 与参考区间观察点不一致"):
        sample_rules(config, "rebuilt", "reference", tmp_path, reports=reports)
    with pytest.raises(ValueError, match=r"01\.template_sha256 与参考区间观察点不一致"):
        verify_samples(styled_same, styled_rebuilt, "reference", "history", profile)
