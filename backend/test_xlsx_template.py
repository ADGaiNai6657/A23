"""Uploaded 03-style XLSX target mapping and downloaded-copy checks."""

import json
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.utils import column_index_from_string
from openpyxl.worksheet.datavalidation import DataValidation

import main
from unpack_dataset import unpack_dataset
from xlsx_template import LABELS

ARCHIVE = Path(__file__).resolve().parents[1] / "dataset/测试集.zip"


@pytest.fixture(scope="module")
def public_03(tmp_path_factory):
    root = tmp_path_factory.mktemp("public-03-xlsx")
    unpack_dataset(ARCHIVE, root)
    folder = next(root.rglob("COVID-19 模板.xlsx")).parent
    return next(folder.glob("COVID-19全球数据集*.xlsx")), folder / "COVID-19 模板.xlsx"


def upload(client, source_name, source, template_name, template):
    files = {"source": (source_name, source), "template": (template_name, template)}
    preview = client.post("/templates/xlsx/inspect", files=files)
    assert preview.status_code == 200, preview.text
    sheet, = preview.json()["target_sheets"]
    mapping = {item["column"]: item["suggested"] for item in sheet["headers"]}
    filled = client.post("/templates/xlsx/fill", files=files, data={
        "plan_id": preview.json()["plan_id"], "sheet": sheet["name"],
        "mapping": json.dumps(mapping, ensure_ascii=False),
    })
    assert filled.status_code == 200, filled.text[:400]
    downloaded = client.get(f"/operations/files/{filled.json()['result_id']}")
    assert downloaded.status_code == 200
    return preview.json(), filled.json(), downloaded.content


def adjusted_template(original):
    workbook = load_workbook(BytesIO(original))
    sheet = workbook.active
    order = ("病例数", "国家/地区", "每日检测数", "大洲", "人口", "人均GDP")
    for column, name in enumerate(order, 1):
        sheet.cell(1, column, name)
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def check_report_against_download(report, content, source):
    downloaded = load_workbook(BytesIO(content), read_only=True, data_only=True)
    source_book = load_workbook(BytesIO(source), read_only=True, data_only=True)
    try:
        sheet = downloaded[report["sheet"]]
        assert sheet.max_row == report["rows"] + 1
        headers = [sheet.cell(1, column).value for column in range(1, 8)]
        first_row = sorted((item for item in report["cells"] if item["row"] == 2),
                           key=lambda item: item["column"])
        assert headers == [LABELS[item["field"]] for item in first_row]
        result_rows = list(sheet.iter_rows(min_row=2, values_only=True))
        source_rows = {row_number: row for row_number, row in enumerate(
            source_book["Data"].iter_rows(values_only=True), 1)}
        assert len(report["cells"]) == report["rows"] * 7
        for item in report["cells"]:
            actual = result_rows[item["row"] - 2][item["column"] - 1]
            if item["field"] == "date":
                assert isinstance(actual, (date, datetime))
                actual_date = actual.date() if isinstance(actual, datetime) else actual
                assert actual_date.isoformat() == item["value"]
            else:
                assert actual == item["value"]
            raw = source_rows[item["source_row"]][column_index_from_string(item["source_column"]) - 1]
            if isinstance(raw, (date, datetime)):
                assert raw.isoformat() == item["source_value"]
            else:
                assert raw == item["source_value"]
        assert sum(value is None for row in result_rows for value in row) == report["cells_empty"]
    finally:
        downloaded.close()
        source_book.close()


def test_public_03_uploaded_template_and_reordered_target(public_03):
    source_path, template_path = public_03
    source, template = source_path.read_bytes(), template_path.read_bytes()
    client = TestClient(main.app)
    preview, baseline, baseline_file = upload(client, source_path.name, source, template_path.name, template)
    assert preview["source_sheets"][0]["name"] == "Data"
    assert preview["target_sheets"][0]["name"] == "Sheet1"
    assert (baseline["rows"], baseline["cells_written"], baseline["cells_empty"]) == (2542, 17464, 330)
    assert all("source_column" in item for item in baseline["cells"])
    assert all(item["reason_code"] == "source_blank" for item in baseline["cells"] if item["value"] is None)
    check_report_against_download(baseline, baseline_file, source)

    reordered = adjusted_template(template)
    _, changed, changed_file = upload(client, source_path.name, source, "调整列序模板.xlsx", reordered)
    assert (changed["rows"], changed["cells_written"], changed["cells_empty"]) == (2542, 17464, 330)
    check_report_against_download(changed, changed_file, source)
    original_by_field = {(item["row"], item["field"]): item["value"] for item in baseline["cells"]}
    assert all(original_by_field[item["row"], item["field"]] == item["value"] for item in changed["cells"])
    assert template_path.read_bytes() == template


def test_mapping_must_be_confirmed_and_file_version_matches(public_03):
    source_path, template_path = public_03
    source, template = source_path.read_bytes(), template_path.read_bytes()
    client = TestClient(main.app)
    files = {"source": (source_path.name, source), "template": (template_path.name, template)}
    preview = client.post("/templates/xlsx/inspect", files=files).json()
    sheet = preview["target_sheets"][0]
    mapping = {item["column"]: item["suggested"] for item in sheet["headers"]}
    missing = client.post("/templates/xlsx/fill", files=files, data={
        "plan_id": preview["plan_id"], "sheet": sheet["name"],
        "mapping": json.dumps({"A": "country"}),
    })
    assert missing.status_code == 400
    changed = client.post("/templates/xlsx/fill", files={
        "source": (source_path.name, source), "template": (template_path.name, adjusted_template(template)),
    }, data={"plan_id": preview["plan_id"], "sheet": sheet["name"],
             "mapping": json.dumps(mapping, ensure_ascii=False)})
    assert changed.status_code == 409
    accepted = client.post("/templates/xlsx/fill", files=files, data={
        "plan_id": preview["plan_id"], "sheet": sheet["name"],
        "mapping": json.dumps(mapping, ensure_ascii=False),
    })
    assert accepted.status_code == 200


def synthetic_source(*, conflict=False, precise=True):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append([LABELS[key] for key in ("date", "country", "continent", "gdp_per_capita",
                                          "population", "每日检测数", "病例数")])
    sheet.append([date(2020, 7, 1), "甲国", "亚洲",
                  "123.5" if precise else "12345.678901234567",
                  1000 if precise else "9007199254740993", 7, 1])
    if conflict:
        sheet.append([date(2020, 7, 1), "甲国", "亚洲", "123.5", 1000, 7, 4])
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def seven_column_template(*, blocked=False):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "目标"
    sheet.append([LABELS[key] for key in ("date", "country", "continent", "gdp_per_capita",
                                          "population", "每日检测数", "病例数")])
    if blocked:
        sheet["E2"] = "=2+3"
    sheet["A5"] = "人工备注"
    sheet["A5"].comment = Comment("勿覆盖", "审核人")
    sheet["H6"] = "=SUM(A6:G6)"
    sheet["H2"] = "范围外"
    workbook.create_sheet("其他")["B3"] = "保持原样"
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def mapped_upload(client, source, template):
    files = {"source": ("源.xlsx", source), "template": ("目标.xlsx", template)}
    preview = client.post("/templates/xlsx/inspect", files=files)
    if preview.status_code != 200:
        return preview, None, None
    target = next(item for item in preview.json()["target_sheets"] if item["name"] == "目标")
    mapping = {item["column"]: item["suggested"] for item in target["headers"]}
    filled = client.post("/templates/xlsx/fill", files=files, data={
        "plan_id": preview.json()["plan_id"], "sheet": "目标",
        "mapping": json.dumps(mapping, ensure_ascii=False),
    })
    assert filled.status_code == 200, filled.text
    downloaded = client.get(f"/operations/files/{filled.json()['result_id']}")
    assert downloaded.status_code == 200
    return preview, filled.json(), downloaded.content


def test_formula_in_write_area_is_located_and_outside_content_survives():
    client = TestClient(main.app)
    source = synthetic_source()
    blocked = seven_column_template(blocked=True)
    preview, report, content = mapped_upload(client, source, blocked)
    assert preview.status_code == 400
    assert "目标" in preview.json()["detail"] and "E2" in preview.json()["detail"]
    assert "公式" in preview.json()["detail"]
    assert report is content is None

    original = seven_column_template()
    preview, report, content = mapped_upload(client, source, original)
    assert preview.status_code == 200
    target = next(item for item in preview.json()["target_sheets"] if item["name"] == "目标")
    assert target["write_range"] == "A2:G2" and target["write_rows"] == 1
    revised = load_workbook(BytesIO(content))
    try:
        sheet = revised["目标"]
        assert sheet["E2"].value == 1000
        assert sheet["A5"].value == "人工备注"
        assert sheet["A5"].comment.text == "勿覆盖"
        assert sheet["H6"].value == "=SUM(A6:G6)"
        assert sheet["H2"].value == "范围外"
        assert revised["其他"]["B3"].value == "保持原样"
        assert sheet.max_row >= 6
    finally:
        revised.close()
    assert len(report["cells"]) == 7
    assert load_workbook(BytesIO(original))["目标"]["E2"].value is None


def test_six_column_template_rejects_content_that_date_insertion_would_move():
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "目标"
    sheet.append([LABELS[key] for key in ("country", "continent", "gdp_per_capita",
                                          "population", "每日检测数", "病例数")])
    sheet["A5"] = "保留备注"
    sheet["H6"] = "=SUM(A6:G6)"
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    preview, _, _ = mapped_upload(TestClient(main.app), synthetic_source(), output.getvalue())
    assert preview.status_code == 400
    assert "目标" in preview.json()["detail"] and "A5" in preview.json()["detail"]


def test_conflicting_candidates_keep_both_source_coordinates():
    client = TestClient(main.app)
    preview, report, content = mapped_upload(client, synthetic_source(conflict=True), seven_column_template())
    assert preview.status_code == 200
    target = next(item for item in report["cells"] if item["field"] == "病例数")
    assert (target["row"], target["column"], target["value"]) == (2, 7, None)
    assert (target["decision"], target["reason_code"]) == ("refused", "conflict")
    assert {(item["value"], item["source_sheet"], item["source_column"],
             item["source_row"], item["source_value"]) for item in target["conflicts"]} == {
        (1, "Data", "G", 2, 1), (4, "Data", "G", 3, 4),
    }
    revised = load_workbook(BytesIO(content))
    try:
        assert revised["目标"]["G2"].value is None
    finally:
        revised.close()


def test_excess_precision_refused_with_original_value():
    client = TestClient(main.app)
    preview, report, content = mapped_upload(client, synthetic_source(precise=False), seven_column_template())
    assert preview.status_code == 200
    by_field = {item["field"]: item for item in report["cells"]}
    assert (report["cells_written"], report["cells_empty"]) == (5, 2)
    assert by_field["gdp_per_capita"]["source_value"] == "12345.678901234567"
    assert by_field["population"]["source_value"] == "9007199254740993"
    for field in ("gdp_per_capita", "population"):
        assert by_field[field]["value"] is None
        assert by_field[field]["decision"] == "refused"
        assert by_field[field]["reason_code"] == "precision_loss"
    revised = load_workbook(BytesIO(content), data_only=True)
    try:
        assert revised["目标"]["D2"].value is None
        assert revised["目标"]["E2"].value is None
        assert revised["目标"]["F2"].value == by_field["每日检测数"]["value"] == 7
    finally:
        revised.close()


def six_column_rules(*, formula=False, frozen=False, unsafe_print=False):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "目标"
    sheet.append([LABELS[key] for key in ("country", "continent", "gdp_per_capita",
                                          "population", "每日检测数", "病例数")])
    rule = DataValidation(type="list", formula1="=$H$1:$H$2" if formula else '"亚洲,欧洲"')
    sheet.add_data_validation(rule)
    rule.add("B2:B20")
    sheet.auto_filter.ref = "A1:F20"
    if frozen:
        sheet.freeze_panes = "B2"
        sheet.print_title_cols = "G:G" if unsafe_print else "A:A"
    sheet.column_dimensions["A"].width = 31
    sheet.column_dimensions["B"].width = 13
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def test_six_column_rules_follow_business_fields_after_insertion():
    client = TestClient(main.app)
    preview, report, content = mapped_upload(client, synthetic_source(), six_column_rules())
    assert preview.status_code == 200
    assert report["cells_written"] == 7
    revised = load_workbook(BytesIO(content))
    try:
        sheet = revised["目标"]
        assert sheet["A1"].value == "日期" and sheet["B1"].value == "国家/地区"
        assert sheet["C1"].value == "大洲" and sheet["C2"].value == "亚洲"
        assert str(sheet.data_validations.dataValidation[0].sqref) == "C2:C20"
        assert sheet.auto_filter.ref == "A1:G20"
        assert sheet.column_dimensions["B"].width == 31
        assert sheet.column_dimensions["C"].width == 13
    finally:
        revised.close()


def test_six_column_formula_validation_requires_seven_headers():
    preview, report, content = mapped_upload(
        TestClient(main.app), synthetic_source(), six_column_rules(formula=True))
    assert preview.status_code == 400
    assert "目标" in preview.json()["detail"]
    assert "B2:B20" in preview.json()["detail"]
    assert "七列表头" in preview.json()["detail"]
    assert report is content is None


def test_six_column_freeze_and_print_repetition_follow_country():
    preview, report, content = mapped_upload(
        TestClient(main.app), synthetic_source(), six_column_rules(frozen=True))
    assert preview.status_code == 200 and report["cells_written"] == 7
    revised = load_workbook(BytesIO(content))
    try:
        sheet = revised["目标"]
        assert sheet["B1"].value == "国家/地区"
        assert sheet.freeze_panes == "C2"
        assert sheet.print_title_cols == "$B:$B"
        assert sheet["C2"].value == "亚洲"
        assert sheet.auto_filter.ref == "A1:G20"
    finally:
        revised.close()


def test_six_column_unsafe_print_repetition_is_located_in_preview():
    preview, report, content = mapped_upload(
        TestClient(main.app), synthetic_source(), six_column_rules(frozen=True, unsafe_print=True))
    assert preview.status_code == 400 and report is content is None
    assert "目标" in preview.json()["detail"] and "打印重复列" in preview.json()["detail"]
    assert "$G:$G" in preview.json()["detail"] and "七列表头" in preview.json()["detail"]


@pytest.mark.parametrize("values,code,expected", [
    ((None, "未统计"), "incompatible_value", [(3, "未统计")]),
    ((None, None), "source_blank", []),
    (("待核", "未统计"), "incompatible_value", [(2, "待核"), (3, "未统计")]),
    ((None, 0), None, []),
])
def test_refusal_provenance_points_to_offending_cell_in_both_entries(tmp_path, values, code, expected):
    from filler import fill_workbook
    from parsers import parse_document
    from store import Store
    from fill_spec import COVID_OUTPUT

    source_book = Workbook()
    sheet = source_book.active
    sheet.title = "Data"
    sheet.append([LABELS[key] for key in ("date", "country", "continent", "gdp_per_capita",
                                          "population", "每日检测数", "病例数")])
    for value in values:
        sheet.append([date(2020, 7, 1), "甲国", "亚洲", 123, 1000, 7, value])
    source_path = tmp_path / "源.xlsx"
    source_book.save(source_path)
    source_book.close()
    source = source_path.read_bytes()

    preview, report, content = mapped_upload(TestClient(main.app), source, seven_column_template())
    assert preview.status_code == 200
    item = next(cell for cell in report["cells"] if cell["field"] == "病例数")
    assert item.get("reason_code") == code
    if code is None:
        assert item["value"] == 0
    source_row = expected[0][0] if expected else (3 if code is None else 2)
    assert item["source_row"] == source_row
    if expected:
        assert item["source_sheet"] == "Data" and item["source_column"] == "G"
        assert item["source_value"] == expected[0][1]
        assert [(entry["source_row"], entry["source_value"]) for entry in item["conflicts"]] == expected
    target = load_workbook(BytesIO(content))
    assert target["目标"]["G2"].value == (0 if code is None else None)
    target.close()

    fixed_template = tmp_path / "模板.xlsx"
    fixed_book = Workbook()
    fixed_book.active.append([LABELS[key] for key in ("date", "country", "continent", "gdp_per_capita",
                                                        "population", "每日检测数", "病例数")])
    fixed_book.save(fixed_template)
    fixed_book.close()
    saved = Store(tmp_path / "source.db")
    saved.save(parse_document(source, source_path.name), source)
    fill_workbook(saved, source_path.name, fixed_template, tmp_path)
    fixed_report = json.loads((tmp_path / "fill_report_covid19.json").read_text())
    fixed_item = next(cell for cell in fixed_report["cells"] if cell["column"] == 7)
    assert fixed_item["source"]["locator"] == {"sheet": "Data", "row": source_row}
    assert fixed_item.get("reason_code") == code
    fixed_output = load_workbook(tmp_path / COVID_OUTPUT)
    assert fixed_output.active["G2"].value == (0 if code is None else None)
    fixed_output.close()
