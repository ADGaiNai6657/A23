"""End-to-end checks for confirmed natural-language document operations."""

from io import BytesIO

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from base64 import b64decode
from copy import copy
from concurrent.futures import ThreadPoolExecutor
from threading import Event

import main
import operations


def word_sample() -> bytes:
    document = Document()
    first = document.add_paragraph()
    first.add_run("本通知自")
    first.add_run("试行")
    first.add_run("日起执行。")
    document.add_paragraph("报名截止日期为2026年12月31日。")
    document.add_paragraph("其余内容保持原样。")
    output = BytesIO()
    document.save(output)
    return output.getvalue()


def excel_sample() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "汇总"
    for column, title in enumerate(("项目", "数量", "单位", "备注"), 1):
        sheet.cell(1, column, title)
    sheet["A2"] = "报名"
    sheet["B2"] = 5
    sheet["C2"] = "人"
    sheet["D2"] = "保留"
    sheet["D2"].font = Font(color="FF008800")
    other = workbook.create_sheet("明细")
    other["A1"] = "原始"
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def post_preview(client, name, content, instruction):
    return client.post("/operations/preview", files={"file": (name, content)},
                       data={"instruction": instruction})


def post_execute(client, name, content, plan_id):
    return client.post("/operations", files={"file": (name, content)},
                       data={"plan_id": plan_id})


@pytest.mark.parametrize("instruction", [
    "把第一段里的‘试行’改成‘正式’，然后把第一段加粗并居中。",
    "第一段将试行替换为正式，并设置粗体、居中对齐。",
])
def test_word_edit_and_format_from_different_phrasings(monkeypatch, instruction):
    source = word_sample()
    replies = {"operations": [
        {"type": "docx_replace", "paragraph": 1, "old": "试行", "new": "正式"},
        {"type": "docx_format", "paragraph": 1, "bold": True, "align": "center"},
    ]}
    monkeypatch.setattr(operations, "model_response", lambda prompt: replies)
    client = TestClient(main.app)
    preview = post_preview(client, "通知.docx", source, instruction)
    assert preview.status_code == 200, preview.text
    plan = preview.json()
    assert [item["location"] for item in plan["operations"]] == ["第 1 段", "第 1 段"]
    assert "正式" in plan["operations"][0]["after"]
    applied = post_execute(client, "通知.docx", source, plan["plan_id"])
    assert applied.status_code == 200, applied.text
    result = applied.json()
    download = client.get(f"/operations/files/{result['result_id']}")
    assert download.status_code == 200
    revised = Document(BytesIO(download.content))
    original = Document(BytesIO(source))
    assert revised.paragraphs[0].text == "本通知自正式日起执行。"
    assert all(run.bold for run in revised.paragraphs[0].runs)
    assert revised.paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert [paragraph.text for paragraph in revised.paragraphs[1:]] == [
        paragraph.text for paragraph in original.paragraphs[1:]
    ]
    assert original.paragraphs[0].text == "本通知自试行日起执行。"


@pytest.mark.parametrize("instruction", [
    "把‘汇总’工作表的 B2 改成 120，把 A1:D1 的表头加粗并居中。",
    "在汇总页将 B2 设为 120，同时使 A1 到 D1 粗体且水平居中。",
])
def test_excel_edit_and_format_from_different_phrasings(monkeypatch, instruction):
    source = excel_sample()
    monkeypatch.setattr(operations, "model_response", lambda prompt: {"operations": [
        {"type": "xlsx_set", "sheet": "汇总", "cell": "B2", "value": 120},
        {"type": "xlsx_format", "sheet": "汇总", "range": "A1:D1", "bold": True, "align": "center"},
    ]})
    client = TestClient(main.app)
    preview = post_preview(client, "报名.xlsx", source, instruction)
    assert preview.status_code == 200, preview.text
    assert [item["location"] for item in preview.json()["operations"]] == ["汇总!B2", "汇总!A1:D1"]
    applied = post_execute(client, "报名.xlsx", source, preview.json()["plan_id"])
    assert applied.status_code == 200, applied.text
    download = client.get(f"/operations/files/{applied.json()['result_id']}")
    assert download.status_code == 200
    revised = load_workbook(BytesIO(download.content))
    original = load_workbook(BytesIO(source))
    try:
        assert revised["汇总"]["B2"].value == 120
        assert all(revised["汇总"].cell(1, column).font.bold for column in range(1, 5))
        assert all(revised["汇总"].cell(1, column).alignment.horizontal == "center"
                   for column in range(1, 5))
        assert revised["汇总"]["D2"].value == "保留"
        assert revised["汇总"]["D2"].font.color.rgb == original["汇总"]["D2"].font.color.rgb
        assert revised["明细"]["A1"].value == original["明细"]["A1"].value
        assert original["汇总"]["B2"].value == 5
    finally:
        revised.close()
        original.close()


def test_extract_displays_original_quote_and_paragraph(monkeypatch):
    source = word_sample()
    quote = "报名截止日期为2026年12月31日"
    monkeypatch.setattr(operations, "model_response", lambda prompt: {"operations": [
        {"type": "docx_extract", "paragraph": 2, "quote": quote},
    ]})
    client = TestClient(main.app)
    preview = post_preview(client, "通知.docx", source, "找出报名截止日期的原文和段落")
    assert preview.status_code == 200
    result = post_execute(client, "通知.docx", source, preview.json()["plan_id"])
    assert result.status_code == 200
    item, = result.json()["operations"]
    assert item["location"] == "第 2 段"
    assert item["quote"] == quote
    assert result.json()["result_id"] is None


def test_excel_extract_returns_original_cell_and_sheet(monkeypatch):
    source = excel_sample()
    monkeypatch.setattr(operations, "model_response", lambda prompt: {"operations": [
        {"type": "xlsx_extract", "sheet": "汇总", "cell": "A2", "quote": "报名"},
    ]})
    client = TestClient(main.app)
    preview = post_preview(client, "报名.xlsx", source, "提取汇总工作表 A2 的原文")
    assert preview.status_code == 200, preview.text
    result = post_execute(client, "报名.xlsx", source, preview.json()["plan_id"])
    assert result.status_code == 200
    assert result.json()["operations"][0]["location"] == "汇总!A2"
    assert result.json()["operations"][0]["quote"] == "报名"
    assert result.json()["result_id"] is None


def test_ambiguous_sheet_and_unbounded_references(monkeypatch):
    source = excel_sample()
    client = TestClient(main.app)
    monkeypatch.setattr(operations, "model_response", lambda prompt: {"operations": [
        {"type": "xlsx_set", "cell": "B2", "value": 120},
    ]})
    ambiguous = post_preview(client, "报名.xlsx", source, "把 B2 改成 120")
    assert ambiguous.status_code == 400 and "请补充工作表名称" in ambiguous.json()["detail"]
    monkeypatch.setattr(operations, "model_response", lambda prompt: {"operations": [
        {"type": "xlsx_set", "sheet": "汇总", "cell": "Z999", "value": 120},
    ]})
    outside = post_preview(client, "报名.xlsx", source, "把汇总工作表的 Z999 改成 120")
    assert outside.status_code == 400 and "第 1 项" in outside.json()["detail"]


def test_missing_paragraph_or_cell_requests_precise_target(monkeypatch):
    client = TestClient(main.app)
    monkeypatch.setattr(operations, "model_response", lambda prompt: {"operations": [
        {"type": "docx_replace", "old": "试行", "new": "正式"},
    ]})
    paragraph = post_preview(client, "通知.docx", word_sample(), "把试行改成正式")
    assert paragraph.status_code == 400 and "补充明确的段落号" in paragraph.json()["detail"]
    monkeypatch.setattr(operations, "model_response", lambda prompt: {"operations": [
        {"type": "xlsx_format", "sheet": "汇总", "bold": True, "align": "center"},
    ]})
    cells = post_preview(client, "报名.xlsx", excel_sample(), "把汇总工作表表头加粗居中")
    assert cells.status_code == 400 and "补充明确的单元格或范围" in cells.json()["detail"]


def test_file_version_mismatch_and_failed_compound_publish_no_copy(monkeypatch):
    source = word_sample()
    monkeypatch.setattr(operations, "model_response", lambda prompt: {"operations": [
        {"type": "docx_replace", "paragraph": 1, "old": "试行", "new": "正式"},
    ]})
    client = TestClient(main.app)
    plan = post_preview(client, "通知.docx", source, "替换文字").json()
    changed = Document(BytesIO(source))
    changed.paragraphs[2].text = "另一份文件"
    output = BytesIO()
    changed.save(output)
    mismatch = post_execute(client, "通知.docx", output.getvalue(), plan["plan_id"])
    assert mismatch.status_code == 409 and "版本不符" in mismatch.json()["detail"]
    renamed = post_execute(client, "另一份通知.docx", source, plan["plan_id"])
    assert renamed.status_code == 409 and "版本不符" in renamed.json()["detail"]

    monkeypatch.setattr(operations, "model_response", lambda prompt: {"operations": [
        {"type": "docx_replace", "paragraph": 1, "old": "试行", "new": "正式"},
        {"type": "docx_format", "paragraph": 99, "bold": True, "align": "center"},
    ]})
    before_outputs = set(operations.outputs)
    failed = post_preview(client, "通知.docx", source, "替换文字并排版第99段")
    assert failed.status_code == 400 and "第 2 项" in failed.json()["detail"]
    assert failed.json().get("result_id") is None
    assert set(operations.outputs) == before_outputs
    assert Document(BytesIO(source)).paragraphs[0].text == "本通知自试行日起执行。"


def test_unsupported_and_model_failure(monkeypatch):
    source = word_sample()
    client = TestClient(main.app)
    monkeypatch.setattr(operations, "model_response", lambda prompt: {"operations": [
        {"type": "delete_document"},
    ]})
    unsupported = post_preview(client, "通知.docx", source, "删除文件")
    assert unsupported.status_code == 400 and "不受支持" in unsupported.json()["detail"]
    def fail_model(prompt):
        raise ValueError("DeepSeek 请求超时")
    monkeypatch.setattr(operations, "model_response", fail_model)
    failure = post_preview(client, "通知.docx", source, "修改第一段")
    assert failure.status_code == 400 and "指令解析失败" in failure.json()["detail"]
    corrupt = post_preview(client, "通知.docx", b"broken", "修改第一段")
    assert corrupt.status_code == 400 and "无法解析 .docx 文件" in corrupt.json()["detail"]


def test_replace_keeps_picture_in_same_run_and_other_paragraphs(monkeypatch):
    document = Document()
    run = document.add_paragraph().add_run("试行")
    picture = b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jR1cAAAAASUVORK5CYII=")
    run.add_picture(BytesIO(picture))
    run.add_text("版本")
    document.add_paragraph("其他段落")
    source_file = BytesIO()
    document.save(source_file)
    source = source_file.getvalue()
    monkeypatch.setattr(operations, "model_response", lambda _: {"operations": [
        {"type": "docx_replace", "paragraph": 1, "old": "试行", "new": "正式"}
    ]})
    client = TestClient(main.app)
    preview = post_preview(client, "图片.docx", source, "把第一段试行改为正式")
    assert preview.status_code == 200, preview.text
    applied = post_execute(client, "图片.docx", source, preview.json()["plan_id"])
    downloaded = client.get(f"/operations/files/{applied.json()['result_id']}")
    revised = Document(BytesIO(downloaded.content))
    assert [p.text for p in revised.paragraphs] == ["正式版本", "其他段落"]
    assert len(revised.inline_shapes) == len(document.inline_shapes) == 1
    assert Document(BytesIO(source)).paragraphs[0].text == "试行版本"
    monkeypatch.setattr(operations, "model_response", lambda _: {"operations": [
        {"type": "docx_replace", "paragraph": 1, "old": "试行版本", "new": "正式版本"}
    ]})
    refused = post_preview(client, "图片.docx", source, "把试行版本改为正式版本")
    assert refused.status_code == 400 and "跨越内嵌图片" in refused.json()["detail"]


def test_excel_missing_value_zero_bool_blank_and_clear(monkeypatch):
    source = excel_sample()
    client = TestClient(main.app)
    monkeypatch.setattr(operations, "model_response", lambda _: {"operations": [
        {"type": "xlsx_set", "sheet": "汇总", "cell": "B2"}
    ]})
    refused = post_preview(client, "记录.xlsx", source, "修改汇总工作表 B2")
    assert refused.status_code == 400 and "第 1 项缺少写入值" in refused.json()["detail"]
    assert load_workbook(BytesIO(source))["汇总"]["B2"].value == 5

    workbook = load_workbook(BytesIO(source))
    workbook["汇总"]["A2"] = 0
    workbook["汇总"]["C2"] = True
    workbook["汇总"]["D2"] = None
    output = BytesIO()
    workbook.save(output)
    data = output.getvalue()
    monkeypatch.setattr(operations, "model_response", lambda _: {"operations": [
        {"type": "xlsx_extract", "sheet": "汇总", "cell": "A2", "quote": "0"},
        {"type": "xlsx_extract", "sheet": "汇总", "cell": "C2", "quote": "True"},
        {"type": "xlsx_set", "sheet": "汇总", "cell": "B2", "value": None},
    ]})
    preview = post_preview(client, "记录.xlsx", data, "提取汇总工作表 A2 和 C2，清空 B2")
    assert preview.status_code == 200, preview.text
    assert preview.json()["operations"][0]["quote"] == "0"
    assert preview.json()["operations"][0]["location"] == "汇总!A2"
    applied = post_execute(client, "记录.xlsx", data, preview.json()["plan_id"])
    revised = load_workbook(BytesIO(client.get(f"/operations/files/{applied.json()['result_id']}").content))
    assert [revised["汇总"][cell].value for cell in ("A2", "B2", "C2", "D2")] == [0, None, True, None]
    monkeypatch.setattr(operations, "model_response", lambda _: {"operations": [
        {"type": "xlsx_extract", "sheet": "汇总", "cell": "D2", "quote": "0"}
    ]})
    assert post_preview(client, "记录.xlsx", data, "提取汇总工作表 D2").status_code == 400


@pytest.mark.parametrize("kind", ["docx", "xlsx"])
@pytest.mark.parametrize("fields", [{"bold": True}, {"align": "center"}, {"bold": True, "align": "center"}])
def test_format_changes_only_requested_attributes(monkeypatch, kind, fields):
    if kind == "docx":
        document = Document()
        paragraph = document.add_paragraph("目标")
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        document.add_paragraph("其他")
        output = BytesIO()
        document.save(output)
        source, name = output.getvalue(), "格式.docx"
        operation = {"type": "docx_format", "paragraph": 1, **fields}
    else:
        workbook = Workbook()
        sheet = workbook.active
        sheet["A1"] = "目标"
        aligned = copy(sheet["A1"].alignment)
        aligned.horizontal = "right"
        sheet["A1"].alignment = aligned
        sheet["B1"] = "其他"
        output = BytesIO()
        workbook.save(output)
        source, name = output.getvalue(), "格式.xlsx"
        operation = {"type": "xlsx_format", "sheet": sheet.title, "range": "A1", **fields}
    monkeypatch.setattr(operations, "model_response", lambda _: {"operations": [operation]})
    client = TestClient(main.app)
    preview = post_preview(client, name, source, "设置格式")
    assert preview.status_code == 200, preview.text
    result = post_execute(client, name, source, preview.json()["plan_id"])
    download = client.get(f"/operations/files/{result.json()['result_id']}").content
    if kind == "docx":
        revised = Document(BytesIO(download))
        assert revised.paragraphs[0].runs[0].bold is (True if "bold" in fields else None)
        assert revised.paragraphs[0].alignment == (WD_ALIGN_PARAGRAPH.CENTER if "align" in fields else WD_ALIGN_PARAGRAPH.RIGHT)
        assert revised.paragraphs[1].text == "其他"
    else:
        revised = load_workbook(BytesIO(download))
        assert revised.active["A1"].font.bold is (True if "bold" in fields else False)
        assert revised.active["A1"].alignment.horizontal == ("center" if "align" in fields else "right")
        assert revised.active["B1"].value == "其他"


def test_operation_plan_concurrent_confirmation_and_failure_retry(monkeypatch):
    source = word_sample()
    client = TestClient(main.app)
    monkeypatch.setattr(operations, "model_response", lambda _: {"operations": [
        {"type": "docx_replace", "paragraph": 1, "old": "试行", "new": "正式"},
    ]})
    plan_id = post_preview(client, "通知.docx", source, "替换第一段").json()["plan_id"]
    actual = operations.run_operations
    entered, release = Event(), Event()

    def held(*args, **kwargs):
        if kwargs.get("save"):
            entered.set()
            assert release.wait(10)
        return actual(*args, **kwargs)

    monkeypatch.setattr(operations, "run_operations", held)
    before = set(operations.outputs)
    with ThreadPoolExecutor(max_workers=2) as pool:
        first = pool.submit(post_execute, client, "通知.docx", source, plan_id)
        assert entered.wait(5)
        second = pool.submit(post_execute, client, "通知.docx", source, plan_id)
        assert second.result(timeout=10).status_code == 409
        release.set()
        accepted = first.result(timeout=10)
    assert accepted.status_code == 200, accepted.text
    assert len(set(operations.outputs) - before) == 1
    assert post_execute(client, "通知.docx", source, plan_id).status_code == 409

    monkeypatch.setattr(operations, "run_operations", actual)
    retry_id = post_preview(client, "通知.docx", source, "替换第一段").json()["plan_id"]
    attempts = [0]

    def fail_once(*args, **kwargs):
        if kwargs.get("save"):
            attempts[0] += 1
            if attempts[0] == 1:
                raise RuntimeError("受控处理失败")
        return actual(*args, **kwargs)

    monkeypatch.setattr(operations, "run_operations", fail_once)
    assert post_execute(client, "通知.docx", source, retry_id).status_code == 400
    assert retry_id in operations.plans
    assert post_execute(client, "通知.docx", source, retry_id).status_code == 200
