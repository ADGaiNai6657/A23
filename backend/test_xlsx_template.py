"""Uploaded 03-style XLSX target mapping and downloaded-copy checks."""

import json
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook
from openpyxl.utils import column_index_from_string

import main
from unpack_dataset import unpack_dataset

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
