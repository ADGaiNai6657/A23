"""Public air-quality DOCX template, checked against an independent XLSX read."""

from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
from io import BytesIO
from pathlib import Path
from threading import Event

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

import main
import air_quality
import operations
from air_quality import CITIES, HEADERS, PERIOD
from parser_xlsx import compatible_workbook
from unpack_dataset import unpack_dataset

ARCHIVE = Path(__file__).resolve().parents[1] / "dataset/测试集.zip"


@pytest.fixture(scope="module")
def public_files(tmp_path_factory):
    root = tmp_path_factory.mktemp("air-quality-source")
    unpack_dataset(ARCHIVE, root)
    matches = list(root.rglob("2025山东省环境空气质量监测数据信息-模板.docx"))
    assert len(matches) == 1
    directory = matches[0].parent
    return next(directory.glob("*.xlsx")), matches[0]


def independent_values(content):
    workbook = load_workbook(compatible_workbook(content), read_only=True, data_only=True)
    expected = defaultdict(list)
    try:
        for sheet in workbook:
            sheet.reset_dimensions()
            rows = sheet.iter_rows(values_only=True)
            header = next(rows)
            positions = {name: header.index(name) for name in ("监测时间", *HEADERS)}
            for row_number, row in enumerate(rows, 2):
                if row[positions["监测时间"]] == PERIOD and row[positions["城市"]] in CITIES:
                    expected[row[positions["城市"]]].append((sheet.title, row_number,
                        [row[positions[name]] for name in HEADERS]))
    finally:
        workbook.close()
    return expected


def test_public_air_quality_template_upload_fill_download(public_files):
    source_path, template_path = public_files
    source, template = source_path.read_bytes(), template_path.read_bytes()
    client = TestClient(main.app)
    upload = {"source": (source_path.name, source), "template": (template_path.name, template)}
    preview = client.post("/templates/inspect", files=upload)
    assert preview.status_code == 200, preview.text
    assert [item["source_rows"] for item in preview.json()["tables"]] == [25, 29, 25]
    assert [item["added_rows"] for item in preview.json()["tables"]] == [1, 1, 0]
    filled = client.post("/templates/fill", files=upload, data={"plan_id": preview.json()["plan_id"]})
    assert filled.status_code == 200, filled.text
    report = filled.json()
    assert len(report["cells"]) == 79 * 8
    download = client.get(f"/operations/files/{report['result_id']}")
    assert download.status_code == 200
    original = Document(BytesIO(template))
    revised = Document(BytesIO(download.content))
    assert [len(table.rows) for table in revised.tables] == [26, 30, 26]
    assert [paragraph.text for paragraph in revised.paragraphs] == [p.text for p in original.paragraphs]
    expected = independent_values(source)
    for table_number, city in enumerate(CITIES, 1):
        table = revised.tables[table_number - 1]
        assert tuple(cell.text for cell in table.rows[0].cells) == HEADERS
        for target_row, (sheet, source_row, values) in enumerate(expected[city], 2):
            assert [cell.text for cell in table.rows[target_row - 1].cells] == [
                "—" if value is None or value == "" else str(value) for value in values
            ]
            evidence = [cell for cell in report["cells"] if cell["table"] == table_number and cell["row"] == target_row]
            assert [(item["source_sheet"], item["source_row"], item["source_value"])
                    for item in evidence] == [(sheet, source_row, value) for value in values]
    assert any(item["source_missing"] for item in report["cells"])
    assert template_path.read_bytes() == template


def test_air_quality_confirmed_plan_rejects_swapped_template(public_files):
    source_path, template_path = public_files
    source, template = source_path.read_bytes(), template_path.read_bytes()
    client = TestClient(main.app)
    preview = client.post("/templates/inspect", files={
        "source": (source_path.name, source), "template": (template_path.name, template),
    })
    assert preview.status_code == 200
    changed = Document(BytesIO(template))
    changed.paragraphs[0].text += " 修改"
    output = BytesIO()
    changed.save(output)
    refused = client.post("/templates/fill", files={
        "source": (source_path.name, source), "template": (template_path.name, output.getvalue()),
    }, data={"plan_id": preview.json()["plan_id"]})
    assert refused.status_code == 409
    assert "版本" in refused.json()["detail"]


def varied_source():
    workbook = Workbook()
    first = workbook.active
    first.title = "甲表"
    order_a = ("监测时间", *HEADERS)
    first.append(order_a)
    first.append((PERIOD, "德州市", "德城区", "甲站", "0", "11", "12", "—", "优"))
    second = workbook.create_sheet("乙表")
    order_b = (*reversed(HEADERS), "监测时间")
    second.append(order_b)
    for city, station in (("潍坊市", "乙站"), ("临沂市", "丙站")):
        values = dict(zip(HEADERS, (city, "辖区", station, "31", "32", "33", "—", "优")))
        second.append(tuple(values[name] for name in reversed(HEADERS)) + (PERIOD,))
    last = workbook.create_sheet("仅表头")
    last.append(("空列", "监测时间", *HEADERS))
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def test_multiple_sheet_column_coordinates_and_word_style(public_files):
    _, template_path = public_files
    source = varied_source()
    document = Document(template_path)
    paragraph = document.tables[0].cell(1, 0).paragraphs[0]
    paragraph.style = "Title"
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    paragraph.add_run("").bold = True
    original_paragraphs = [(p.text, p.style.name, p.alignment) for p in document.paragraphs]
    output = BytesIO()
    document.save(output)
    template = output.getvalue()
    client = TestClient(main.app)
    upload = {"source": ("多表.xlsx", source), "template": ("空气模板.docx", template)}
    preview = client.post("/templates/inspect", files=upload)
    assert preview.status_code == 200, preview.text
    assert preview.json()["source_columns"]["甲表"]["城市"] == "B"
    assert preview.json()["source_columns"]["乙表"]["城市"] == "H"
    assert preview.json()["source_columns"]["仅表头"]["城市"] == "C"
    filled = client.post("/templates/fill", files=upload, data={"plan_id": preview.json()["plan_id"]})
    assert filled.status_code == 200, filled.text
    report = filled.json()
    downloaded = client.get(f"/operations/files/{report['result_id']}")
    revised = Document(BytesIO(downloaded.content))
    source_book = load_workbook(BytesIO(source), read_only=True, data_only=True)
    try:
        for item in report["cells"]:
            address = f"{item['source_column']}{item['source_row']}"
            assert source_book[item["source_sheet"]][address].value == item["source_value"]
            assert revised.tables[item["table"] - 1].cell(item["row"] - 1, item["column"] - 1).text == item["value"]
        assert len(report["cells"]) == 24
    finally:
        source_book.close()
    target = revised.tables[0].cell(1, 0).paragraphs[0]
    assert target.style.name == "Title"
    assert target.alignment == WD_ALIGN_PARAGRAPH.RIGHT
    assert target.runs[0].bold is True
    assert [(p.text, p.style.name, p.alignment) for p in revised.paragraphs] == original_paragraphs
    assert revised.tables[0].cell(2, 0).text == ""


def test_merged_target_is_refused_during_preview(public_files):
    _, template_path = public_files
    document = Document(template_path)
    document.tables[0].cell(1, 0).merge(document.tables[0].cell(1, 1))
    output = BytesIO()
    document.save(output)
    client = TestClient(main.app)
    before = set(operations.outputs)
    refused = client.post("/templates/inspect", files={
        "source": ("多表.xlsx", varied_source()),
        "template": ("合并模板.docx", output.getvalue()),
    })
    assert refused.status_code == 400
    assert "第 1 张表第 2 行第 1 列" in refused.json()["detail"]
    assert "合并单元格" in refused.json()["detail"]
    assert set(operations.outputs) == before


@pytest.mark.parametrize("shape,column", [("leading", 1), ("trailing", 8), ("short", 8)])
def test_omitted_or_short_physical_word_row_is_refused(public_files, shape, column):
    _, template_path = public_files
    document = Document(template_path)
    row = document.tables[0].rows[1]
    if shape == "leading":
        row._tr.remove(row._tr.tc_lst[0])
        row._tr.get_or_add_trPr().get_or_add_gridBefore().val = 1
    else:
        row._tr.remove(row._tr.tc_lst[-1])
        if shape == "trailing":
            row._tr.get_or_add_trPr().get_or_add_gridAfter().val = 1
    output = BytesIO()
    document.save(output)
    client = TestClient(main.app)
    before = set(operations.outputs)
    refused = client.post("/templates/inspect", files={
        "source": ("多表.xlsx", varied_source()),
        "template": ("省格模板.docx", output.getvalue()),
    })
    assert refused.status_code == 400
    assert f"第 1 张表第 2 行第 {column} 列" in refused.json()["detail"]
    assert set(operations.outputs) == before
    assert len(document.tables[0].rows[2]._tr.tc_lst) == 8
    assert all(not cell.text for cell in document.tables[0].rows[2].cells)


def test_duplicate_required_source_header_is_refused(public_files):
    _, template_path = public_files
    workbook = load_workbook(BytesIO(varied_source()))
    sheet = workbook["甲表"]
    sheet["J1"] = "空气质量指数"
    sheet["J2"] = 99
    output = BytesIO()
    workbook.save(output)
    client = TestClient(main.app)
    refused = client.post("/templates/inspect", files={
        "source": ("重复列.xlsx", output.getvalue()),
        "template": ("空气模板.docx", template_path.read_bytes()),
    })
    assert refused.status_code == 400
    assert all(term in refused.json()["detail"] for term in ("甲表", "空气质量指数", "E", "J"))
    assert "plan_id" not in refused.json()


def test_template_plan_and_download_expiry_and_capacity(public_files, monkeypatch):
    _, template_path = public_files
    source, template = varied_source(), template_path.read_bytes()
    client = TestClient(main.app)
    upload = {"source": ("多表.xlsx", source), "template": ("空气模板.docx", template)}
    clock = [1000.0]
    monkeypatch.setattr(air_quality, "monotonic", lambda: clock[0])
    monkeypatch.setattr(operations, "monotonic", lambda: clock[0])
    monkeypatch.setattr(air_quality, "MAX_PLANS", 2)
    monkeypatch.setattr(operations, "MAX_OUTPUTS", 2)
    old_plans, old_outputs = air_quality.plans.copy(), operations.outputs.copy()
    try:
        air_quality.plans.clear()
        operations.outputs.clear()
        plan_ids = [client.post("/templates/inspect", files=upload).json()["plan_id"] for _ in range(3)]
        assert len(air_quality.plans) == 2 and plan_ids[0] not in air_quality.plans
        expired_plan = plan_ids[1]
        clock[0] += operations.PLAN_LIFETIME + 1
        refused = client.post("/templates/fill", files=upload, data={"plan_id": expired_plan})
        assert refused.status_code == 409 and expired_plan not in air_quality.plans
        assert not air_quality.plans

        fresh = client.post("/templates/inspect", files=upload).json()["plan_id"]
        filled = client.post("/templates/fill", files=upload, data={"plan_id": fresh})
        assert filled.status_code == 200
        result_id = filled.json()["result_id"]
        assert client.get(f"/operations/files/{result_id}").status_code == 200
        other = [operations.store_output(f"副本{i}.docx", b"test") for i in range(2)]
        assert len(operations.outputs) == 2 and result_id not in operations.outputs
        assert client.get(f"/operations/files/{other[-1]}").status_code == 200
        clock[0] += operations.PLAN_LIFETIME + 1
        assert client.get(f"/operations/files/{other[-1]}").status_code == 404
        assert not operations.outputs
    finally:
        air_quality.plans.clear()
        air_quality.plans.update(old_plans)
        operations.outputs.clear()
        operations.outputs.update(old_outputs)


def test_template_plan_concurrent_confirmation_and_failure_retry(public_files, monkeypatch):
    _, template_path = public_files
    source, template = varied_source(), template_path.read_bytes()
    client = TestClient(main.app)
    upload = {"source": ("多表.xlsx", source), "template": ("空气模板.docx", template)}
    preview = client.post("/templates/inspect", files=upload)
    plan_id = preview.json()["plan_id"]
    entered, release = Event(), Event()
    actual = air_quality._fill_claimed

    def held(*args):
        entered.set()
        assert release.wait(10)
        return actual(*args)

    monkeypatch.setattr(air_quality, "_fill_claimed", held)
    before = set(operations.outputs)
    with ThreadPoolExecutor(max_workers=2) as pool:
        first = pool.submit(client.post, "/templates/fill", files=upload, data={"plan_id": plan_id})
        assert entered.wait(5)
        second = pool.submit(client.post, "/templates/fill", files=upload, data={"plan_id": plan_id})
        assert second.result(timeout=10).status_code == 409
        release.set()
        accepted = first.result(timeout=10)
    assert accepted.status_code == 200, accepted.text
    assert len(set(operations.outputs) - before) == 1
    assert client.post("/templates/fill", files=upload, data={"plan_id": plan_id}).status_code == 409

    monkeypatch.setattr(air_quality, "_fill_claimed", actual)
    retry_id = client.post("/templates/inspect", files=upload).json()["plan_id"]
    calls = [0]

    def fail_once(*args):
        calls[0] += 1
        if calls[0] == 1:
            raise RuntimeError("受控写入失败")
        return actual(*args)

    monkeypatch.setattr(air_quality, "_fill_claimed", fail_once)
    assert client.post("/templates/fill", files=upload, data={"plan_id": retry_id}).status_code == 400
    assert retry_id in air_quality.plans
    assert client.post("/templates/fill", files=upload, data={"plan_id": retry_id}).status_code == 200
