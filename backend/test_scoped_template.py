"""Empty-database cross-document Word and Excel workflow with a model substitute."""

import json
import time
from io import BytesIO

from docx import Document
from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

import extractor
import main
import task_queue
from store import Store


def word_source() -> bytes:
    document = Document()
    document.add_paragraph("2025年甲城GDP总量100亿元。")
    stream = BytesIO()
    document.save(stream)
    return stream.getvalue()


def word_template() -> bytes:
    document = Document()
    document.add_paragraph("跨文档统计表")
    table = document.add_table(rows=3, cols=3)
    for column, label in enumerate(("城市", "GDP总量", "常住人口")):
        table.cell(0, column).text = label
    table.cell(1, 0).text = "甲城"
    table.cell(2, 0).text = "乙城"
    stream = BytesIO()
    document.save(stream)
    return stream.getvalue()


def excel_template() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "目标"
    sheet.append(["城市", "GDP总量", "常住人口"])
    sheet["A2"] = "甲城"
    sheet["A3"] = "乙城"
    sheet["E5"] = "原备注"
    other = workbook.create_sheet("保留")
    other["B2"] = "原值"
    stream = BytesIO()
    workbook.save(stream)
    workbook.close()
    return stream.getvalue()


def wait_tasks(client, tasks):
    complete = []
    for task in tasks:
        for _ in range(300):
            response = client.get(f"/tasks/{task['task_id']}")
            assert response.status_code == 200
            item = response.json()
            if item["status"] in {"succeeded", "failed", "superseded"}:
                break
            time.sleep(.01)
        assert item["status"] == "succeeded", item
        complete.append(item)
    return complete


def test_scoped_word_excel_flow_from_empty_database(tmp_path, monkeypatch):
    database = Store(tmp_path / "empty.db")
    monkeypatch.setattr(main, "store", database)
    assert not database.sources()
    snippets = {
        "经济.docx": ("GDP总量", "100", "亿元", "2025年甲城GDP总量100亿元"),
        "人口.md": ("常住人口", "20", "万人", "2025年甲城常住人口20万人"),
        "范围外.txt": ("GDP总量", "999", "亿元", "2025年甲城GDP总量999亿元"),
    }

    def model_substitute(content, name, metrics, store, *, audit=None, publish=None):
        metric, value, unit, snippet = snippets[name]
        fact = {"block_no": 1, "entity": "甲城", "metric": metric,
                "value": value, "unit": unit, "snippet": snippet}
        return extractor.extract_document(content, name, metrics, store, audit=audit, publish=publish,
                                          respond=lambda _: {"facts": [fact]})

    monkeypatch.setattr(task_queue, "extract_document", model_substitute)
    client = TestClient(main.app)
    uploaded = client.post("/workspace/import", files=[
        ("files", ("经济.docx", word_source())),
        ("files", ("人口.md", "2025年甲城常住人口20万人。\n".encode())),
    ])
    assert uploaded.status_code == 202, uploaded.text
    tasks = wait_tasks(client, uploaded.json()["tasks"])
    assert {item["doc"] for item in tasks} == {"经济.docx", "人口.md"}
    outside = client.post("/workspace/import", files=[
        ("files", ("范围外.txt", "2025年甲城GDP总量999亿元。\n".encode())),
    ])
    assert outside.status_code == 202
    wait_tasks(client, outside.json()["tasks"])
    sources = client.get("/workspace/sources").json()["sources"]
    assert len(sources) == 3 and all(item["fact_count"] == 1 for item in sources)

    docs = ["经济.docx", "人口.md"]
    mapping = {"2": {"metric": "GDP总量", "period": "2025", "scope": "", "unit": "亿元"},
               "3": {"metric": "常住人口", "period": "2025", "scope": "", "unit": "万人"}}
    reference_file = tmp_path / "独立参考.xlsx"
    reference_book = Workbook()
    reference_sheet = reference_book.active
    reference_sheet.append(["城市", "GDP总量", "常住人口"])
    reference_sheet.append(["甲城", 100, 20])
    reference_sheet.append(["乙城", None, None])
    reference_book.save(reference_file)
    reference_book.close()
    reference_book = load_workbook(reference_file, read_only=True, data_only=True)
    reference = {(reference_book.active.cell(row, 1).value, column):
                 reference_book.active.cell(row, column).value
                 for row in (2, 3) for column in (2, 3)}
    reference_book.close()
    for name, content, target in (("未预置.docx", word_template(), "table:1"),
                                  ("未预置.xlsx", excel_template(), "目标")):
        inspect = client.post("/workspace/templates/inspect", files={"template": (name, content)})
        assert inspect.status_code == 200, inspect.text
        assert any(item["id"] == target for item in inspect.json()["targets"])
        fields = {"docs": json.dumps(docs, ensure_ascii=False), "target": target,
                  "entity_column": "1", "mapping": json.dumps(mapping, ensure_ascii=False)}
        preview = client.post("/workspace/templates/preview", files={"template": (name, content)}, data=fields)
        assert preview.status_code == 200, preview.text
        plan = preview.json()
        assert plan["cells_written"] == 2 and plan["cells_empty"] == 2
        filled = client.post("/workspace/templates/fill", files={"template": (name, content)},
                             data={"plan_id": plan["plan_id"]})
        assert filled.status_code == 200, filled.text
        report = filled.json()
        assert report["cells_written"] == 2 and report["cells_empty"] == 2
        assert report["fill_seconds"] >= 0 and report["versions"] == plan["versions"]
        assert {item["source"]["doc"] for item in report["cells"] if item["source"]} == set(docs)
        assert all(item["source"]["snippet"] and item["source"]["locator"]
                   for item in report["cells"] if item["source"])
        assert all(all(candidate["doc"] != "范围外.txt" for candidate in item["candidates"])
                   for item in report["cells"])
        assert all(item["reason_code"] == "missing_fact" for item in report["cells"] if item["entity"] == "乙城")
        downloaded = client.get(f"/operations/files/{report['result_id']}")
        assert downloaded.status_code == 200
        if name.endswith(".docx"):
            revised = Document(BytesIO(downloaded.content))
            actual = {(revised.tables[0].cell(row, 0).text, column + 1):
                      (int(revised.tables[0].cell(row, column).text)
                       if revised.tables[0].cell(row, column).text else None)
                      for row in (1, 2) for column in (1, 2)}
            assert revised.paragraphs[0].text == "跨文档统计表"
            assert Document(BytesIO(content)).tables[0].cell(1, 1).text == ""
        else:
            revised = load_workbook(BytesIO(downloaded.content))
            actual = {(revised["目标"].cell(row, 1).value, column): revised["目标"].cell(row, column).value
                      for row in (2, 3) for column in (2, 3)}
            assert revised["目标"]["E5"].value == "原备注"
            assert revised["保留"]["B2"].value == "原值"
            revised.close()
        differences = [{"target": key, "expected": expected, "actual": actual[key]}
                       for key, expected in reference.items() if actual[key] != expected]
        assert differences == []

    conflict = client.post("/workspace/templates/preview", files={"template": ("冲突.xlsx", excel_template())},
                           data={"docs": json.dumps([*docs, "范围外.txt"], ensure_ascii=False),
                                 "target": "目标", "entity_column": "1",
                                 "mapping": json.dumps(mapping, ensure_ascii=False)})
    assert conflict.status_code == 200, conflict.text
    gdp = next(item for item in conflict.json()["cells"] if item["entity"] == "甲城" and item["column"] == 2)
    assert gdp["decision"] == "refused" and gdp["reason_code"] == "conflict"
    assert {item["doc"] for item in gdp["candidates"]} == {"经济.docx", "范围外.txt"}
    conflicting_result = client.post("/workspace/templates/fill",
                                     files={"template": ("冲突.xlsx", excel_template())},
                                     data={"plan_id": conflict.json()["plan_id"]})
    assert conflicting_result.status_code == 200
    conflicting_copy = client.get(f"/operations/files/{conflicting_result.json()['result_id']}")
    conflicting_book = load_workbook(BytesIO(conflicting_copy.content))
    try:
        assert conflicting_book["目标"]["B2"].value is None
        assert conflicting_book["目标"]["C2"].value == 20
    finally:
        conflicting_book.close()

    incompatible_mapping = {**mapping, "2": {**mapping["2"], "unit": "万元"}}
    incompatible = client.post("/workspace/templates/preview",
                               files={"template": ("单位.xlsx", excel_template())},
                               data={"docs": json.dumps(docs, ensure_ascii=False), "target": "目标",
                                     "entity_column": "1", "mapping": json.dumps(incompatible_mapping, ensure_ascii=False)})
    assert incompatible.status_code == 200
    mismatched = next(item for item in incompatible.json()["cells"] if item["entity"] == "甲城"
                      and item["column"] == 2)
    assert mismatched["reason_code"] == "incompatible_value" and mismatched["value"] is None
    assert mismatched["candidates"][0]["doc"] == "经济.docx"

    old_template = excel_template()
    old_plan = client.post("/workspace/templates/preview", files={"template": ("旧计划.xlsx", old_template)},
                           data={"docs": json.dumps(docs, ensure_ascii=False), "target": "目标",
                                 "entity_column": "1", "mapping": json.dumps(mapping, ensure_ascii=False)})
    assert old_plan.status_code == 200
    newer_document = Document(BytesIO(word_source()))
    newer_document.add_paragraph("同名资料的新版附注。")
    newer_stream = BytesIO()
    newer_document.save(newer_stream)
    revised = client.post("/workspace/import", files=[("files", ("经济.docx", newer_stream.getvalue()))])
    assert revised.status_code == 202
    wait_tasks(client, revised.json()["tasks"])
    stale = client.post("/workspace/templates/fill", files={"template": ("旧计划.xlsx", old_template)},
                        data={"plan_id": old_plan.json()["plan_id"]})
    assert stale.status_code == 409 and "事实版本" in stale.json()["detail"]


def test_batch_preflight_and_template_write_area_reject_without_partial_result(tmp_path, monkeypatch):
    database = Store(tmp_path / "empty.db")
    monkeypatch.setattr(main, "store", database)
    client = TestClient(main.app)
    invalid_batch = client.post("/workspace/import", files=[
        ("files", ("正确.md", "2025年甲城GDP总量100亿元。\n".encode())),
        ("files", ("损坏.docx", b"not a docx")),
    ])
    assert invalid_batch.status_code == 400
    assert database.sources() == []

    book = load_workbook(BytesIO(excel_template()))
    book["目标"]["B2"] = 7
    stream = BytesIO()
    book.save(stream)
    book.close()
    mapping = {"2": {"metric": "GDP总量", "period": "2025"},
               "3": {"metric": "常住人口", "period": "2025"}}
    blocked = client.post("/workspace/templates/preview", files={"template": ("有数据.xlsx", stream.getvalue())},
                          data={"docs": '["正确.md"]', "target": "目标", "entity_column": "1",
                                "mapping": json.dumps(mapping, ensure_ascii=False)})
    assert blocked.status_code == 400 and "B2" in blocked.json()["detail"]
