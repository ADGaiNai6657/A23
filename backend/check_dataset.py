import argparse
import hashlib
import json
import re
from collections import Counter
from datetime import date, datetime, timezone
from itertools import zip_longest
from pathlib import Path
from time import perf_counter

from docx import Document
from openpyxl import load_workbook

from acceptance import check
from config import BACKEND, ROOT
from criteria import checked_rules, load_expectations
from extractor import evidence_error
from evaluate import matches
from check_extract import recorded_extract
from fill_spec import CITY_OUTPUT, COVID_END, COVID_OUTPUT, COVID_START, REPORT_FILES
from filler import atomic_file, block_digest, fill_workbook, write_json
from parsers import parse_document
from rebuild_evidence import HISTORY_INPUT, RUN_EVIDENCE, digest
from schema import RawFact
from store import Store

DATASET = ROOT / "dataset" / "测试集"
OUTPUT = BACKEND / "data" / "actual"
CITY_FOLDER = "2025年中国城市经济百强全景报告"
CITY_SOURCE = f"{CITY_FOLDER}.docx"
CITY_TEMPLATE = f"{CITY_FOLDER}-模板.xlsx"
CITY_METRICS = [
    "GDP总量", "常住人口", "人均GDP", "一般公共预算收入",
    "GDP", "地区生产总值", "人口基数", "人口严控至", "万人口", "人口",
]
DAILY_FOLDER = "COVID-19数据集"
DAILY_SOURCE = "COVID-19全球数据集（节选）.xlsx"
DAILY_TEMPLATE = "COVID-19 模板.xlsx"
DAILY_CONTEXT = "中国COVID-19新冠疫情情况.docx"
DAILY_FIELDS = ("国家/地区", "大洲", "人均GDP", "人口", "每日检测数", "病例数")
VALUE = r"(\d[\d,]*(?:\.\d+)?)"
CORPUS_ANCHORS = {
    "2024年药品流通行业运行统计分析报告.txt": {"text_line": 70},
    "2023年文化和旅游发展统计公报.md": {"text_line": 68},
    "南京市2021~2024年国民经济和社会发展统计公报汇总.docx": {"paragraph": 271, "table_cell": 946},
}


def city_expected(path: Path) -> tuple[list[tuple], list[dict]]:
    paragraphs = [
        (number, paragraph.text) for number, paragraph in enumerate(Document(path).paragraphs, 1)
        if "一般公共预算收入" in paragraph.text
    ]
    expected = []
    locators = []
    for number, text in paragraphs:
        name = re.split(r" GDP 总量|以|紧随|凭借|在转型|作为", text, maxsplit=1)[0]
        if not re.fullmatch(r"[\u4e00-\u9fff]{2,12}", name):
            raise ValueError(f"01 原文核对无法确定城市名，原文段落 {number}")
        # This audit reads the source quantities directly instead of using model facts.
        revenue = text.split("一般公共预算收入", 1)[1]
        numbers = []
        for section, unit in ((text, "亿元"), (text, "万(?!元|亿)"), (text, "元"), (revenue, "亿元")):
            found = re.search(VALUE + r"\s*" + unit, section)
            if not found:
                raise ValueError(f"01 原文核对缺少指标数量，城市 {name}")
            numbers.append(float(found[1].replace(",", "")))
        expected.append((name, *numbers))
        locators.append({"paragraph": number})
    return expected, locators


def daily_expected(path: Path) -> tuple[list[tuple], list[dict]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        rows = workbook.active.values
        headers = next(rows)
        indices = [headers.index(field) for field in ("日期", *DAILY_FIELDS)]
        selected = {}
        locators = {}
        for number, row in enumerate(rows, 2):
            if not isinstance(row[indices[0]], datetime):
                continue
            day = row[indices[0]].date()
            if COVID_START <= day <= COVID_END:
                key = (day, row[indices[1]])
                if key in selected:
                    raise ValueError(f"03 原文核对发现重复日期和国家键 {key}")
                selected[key] = (day, *(row[index] for index in indices[1:]))
                locators[key] = {"sheet": workbook.active.title, "row": number}
        ordered = sorted(selected)
        return [selected[key] for key in ordered], [locators[key] for key in ordered]
    finally:
        workbook.close()


def verify_workbook(path: Path, expected: list[tuple], width: int) -> dict:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.active
        rows = list(sheet.iter_rows(min_row=2, max_col=width, values_only=True))
        diffs = []
        complete = city_order_matches = 0
        nonempty = correct = false_fills = missed = 0
        for row, (actual, target) in enumerate(
            zip_longest(rows, expected, fillvalue=(None,) * width), 2,
        ):
            complete += all(value is not None for value in actual)
            city_order_matches += matches(actual[0], target[0])
            for column, (value, truth) in enumerate(zip(actual, target), 1):
                nonempty += truth is not None
                correct += truth is not None and matches(value, truth)
                false_fills += truth is None and value is not None
                missed += truth is not None and value is None
                if not matches(value, truth):
                    decision = "误填" if truth is None else "漏填" if value is None else "值不符"
                    diffs.append({
                        "row": row, "column": column, "actual": value, "source_value": truth, "decision": decision,
                    })
        return {
            "rows": len(rows), "columns": sheet.max_column, "source_nonempty": nonempty,
            "source_empty": len(expected) * width - nonempty, "correct_nonempty": correct,
            "false_fills": false_fills, "missed": missed, "differences": diffs,
            "complete_rows": complete, "city_order_matches": city_order_matches,
        }
    finally:
        workbook.close()


def verify_trace(store: Store, path: Path, locators: list[dict], columns: int) -> dict:
    report = json.loads(path.read_text(encoding="utf-8"))
    by_number = {block["number"]: block for block in store.blocks(report["doc"])}
    coordinates = {(item["row"], item["column"]) for item in report["cells"]}
    expected = {(row, column) for row in range(2, len(locators) + 2) for column in range(1, columns + 1)}
    if coordinates != expected or len(report["cells"]) != len(expected):
        raise ValueError(f"逐格记录数量或坐标不符：{path.name}")
    for item in report["cells"]:
        source = item["source"]
        alternatives = [option["source"] for option in item.get("conflicts", ())]
        sources = [source] if source is not None else alternatives if item.get("reason_code") == "conflict" else []
        if not sources or not item["reason"]:
            raise ValueError(f"逐格记录缺少原文位置或原因：{path.name} {item['row']},{item['column']}")
        anchor = locators[item["row"] - 2]
        if not any(candidate["locator"] == anchor for candidate in sources):
            raise ValueError(f"逐格记录缺少对应输出行的源定位：{path.name} {item['row']},{item['column']}")
        if any(
            (block := by_number.get(candidate["block_no"])) is None or candidate["doc"] != report["doc"]
            or candidate["locator"] != block["locator"] or candidate.get("block_sha256") != block_digest(block)
            for candidate in sources + alternatives
        ):
            raise ValueError(f"逐格记录源定位无法回查：{path.name} {item['row']},{item['column']}")
        if columns == 7 and item["column"] >= 3 and item["decision"] == "filled":
            if source.get("field") != DAILY_FIELDS[item["column"] - 2]:
                raise ValueError(f"03 逐格记录的源字段不符：{item['row']},{item['column']}")
    return report


def source_score(trace: dict, audit: dict, daily: bool, database_sha256: str | None,
                 *, database_stage: str = "offline_post_extraction_pre_fill",
                 database_observed_at: str | None = None) -> dict:
    score = {
        key: trace[key] for key in (
            "doc", "file", "result_id", "source_sha256", "template_sha256",
            "run_id", "workbook_sha256", "rows", "cells_written", "cells_empty", "block_count",
            "fill_seconds", "refusal_counts", "publication_id",
        ) if key in trace
    }
    score.update({
        "basis": "independent-source-check", "diff": audit["differences"],
        "expected_empty": audit["source_empty"], "false_fills": audit["false_fills"],
        "missed": audit["missed"],
        "wrong_values": sum(item["decision"] == "值不符" for item in audit["differences"]),
    })
    if database_sha256 is not None:
        score["database_sha256"] = database_sha256
        score["database_sha256_stage"] = database_stage
        if database_observed_at is not None:
            score["database_sha256_observed_at"] = database_observed_at
    else:
        score["database_status"] = "changed_after_fill"
    if daily:
        required, empty = audit["source_nonempty"], audit["source_empty"]
        score.update({
            "required": required, "expected_empty": empty, "correct": audit["correct_nonempty"],
            "correct_rate": audit["correct_nonempty"] / required,
            "false_fill_rate": audit["false_fills"] / empty, "missed_rate": audit["missed"] / required,
        })
    else:
        score.update({
            "cell_accuracy": audit["correct_nonempty"] / audit["source_nonempty"],
            "row_completeness": audit["complete_rows"] / score["rows"],
            "city_order_matches": audit["city_order_matches"],
        })
    return score


def score_published_fill(root: Path, directory: Path, store: Store, file: str, doc: str) -> None:
    daily = file == COVID_OUTPUT
    folder, source, template = (DAILY_FOLDER, DAILY_SOURCE, DAILY_TEMPLATE) if daily else (
        CITY_FOLDER, CITY_SOURCE, CITY_TEMPLATE,
    )
    if doc != source:
        return
    source_path = root / "包含模板文件" / folder / source
    trace_name, _, report_name = REPORT_FILES[file]
    trace = json.loads((directory / trace_name).read_text(encoding="utf-8"))
    if trace["source_sha256"] != hashlib.sha256(source_path.read_bytes()).hexdigest() or (
        trace["template_sha256"] != hashlib.sha256(source_path.with_name(template).read_bytes()).hexdigest()
    ):
        return
    expected, _ = daily_expected(source_path) if daily else city_expected(source_path)
    audit = verify_workbook(directory / file, expected, 7 if daily else 5)
    database_sha256 = hashlib.sha256(store.path.read_bytes()).hexdigest()
    database_observed_at = datetime.now(timezone.utc).isoformat(timespec="milliseconds")
    current = store.snapshot(doc, with_facts=False)
    if current is None or (hashlib.sha256(current[0]).hexdigest(), current[3]) != (
        trace["source_sha256"], trace.get("publication_id")
    ):
        database_sha256 = None
    score = source_score(trace, audit, daily, database_sha256,
                         database_stage="web_scoring_after_fill",
                         database_observed_at=database_observed_at if database_sha256 else None)
    write_json(directory / report_name, score)


def parse_corpus(root: Path, store: Store) -> list[dict]:
    inventory = []
    for category, suffix, count in (("Excel", ".xlsx", 5), ("md", ".md", 3), ("txt", ".txt", 3), ("word", ".docx", 5)):
        paths = sorted((root / category).glob(f"*{suffix}"))
        if len(paths) != count:
            raise ValueError(f"{category} 预期 {count} 份样本，实际 {len(paths)}")
        for path in paths:
            start = perf_counter()
            data = path.read_bytes()
            parsed = parse_document(data, path.name)
            blocks_by_type = dict(Counter(block.type for block in parsed.blocks))
            expected = CORPUS_ANCHORS.get(path.name)
            if not blocks_by_type or (expected is not None and blocks_by_type != expected):
                raise ValueError(f"测试语料的块数与预期不符：{path.name} {blocks_by_type}")
            store.save(parsed, data)
            inventory.append({
                "file": path.name, "blocks": len(parsed.blocks), "seconds": round(perf_counter() - start, 3),
                "types": blocks_by_type,
            })
    return inventory


def restore_city_facts(store: Store, source: Path, archive: Path) -> None:
    if store.path.exists():
        raise ValueError("重放历史事实需要空目录，请指定新的 --data-dir")
    content = source.read_bytes()
    saved = json.loads(archive.read_text(encoding="utf-8"))
    if saved["source_sha256"] != hashlib.sha256(content).hexdigest():
        raise ValueError("历史事实与本次城市原文的 SHA256 不符")
    doc = parse_document(content, source.name)
    blocks = dict(doc.numbered_blocks())
    records = []
    for row in saved["facts"]:
        if row["doc"] != doc.file or row["mode"] != "list" or evidence_error(RawFact.model_validate(row), blocks):
            raise ValueError(f"历史事实无法回指本次原文，块号 {row['block_no']}")
        records.append(tuple(row[key] for key in (
            "block_no", "entity", "metric", "metric_key", "value", "value_num", "unit",
            "snippet", "period", "period_block_no", "scope",
        )))
    store.save(doc, content, "list", records)


def check_dataset(root: Path, output: Path, extract: bool, archived_facts: Path | None = None) -> dict:
    templates = root / "包含模板文件"
    city_source = templates / CITY_FOLDER / CITY_SOURCE
    city_template = city_source.with_name(CITY_TEMPLATE)
    daily_source = templates / DAILY_FOLDER / DAILY_SOURCE
    daily_template = daily_source.with_name(DAILY_TEMPLATE)
    store = Store(output / "a23.db")
    mode = "rebuilt" if extract or archived_facts is not None else "same-db"
    initial_hash = digest(store.path) if store.path.is_file() else None
    expectations = checked_rules(load_expectations(), mode)
    evidence = {"mode": mode, "initial_database_sha256": initial_hash}
    if mode == "rebuilt" and initial_hash is not None:
        raise ValueError("重建验收要求初始空库，请指定新的 --data-dir")
    if mode == "same-db" and initial_hash is None:
        raise ValueError("同库验收需要已有 a23.db，请先用 --extract 或 --archived-facts 建库")
    if archived_facts is not None:
        restore_city_facts(store, city_source, archived_facts)
        archive = output / HISTORY_INPUT
        atomic_file(archive, lambda target: target.write_bytes(archived_facts.read_bytes()))
        evidence.update({"extraction_kind": "history", "extraction_sha256": digest(archive),
                         "extraction_id": json.loads(archive.read_bytes())["extraction_id"]})
    if extract:
        print("正在调用 DeepSeek 抽取 01 城市原文", flush=True)
        run = recorded_extract(city_source, city_source.read_bytes(), CITY_METRICS, store, 1)
        extraction = {**run["result"], **{key: run[key] for key in ("run_id", "model", "seconds")}}
        write_json(output / "city_extraction.json", extraction)
        evidence.update({"extraction_kind": "online", "extraction_id": run["run_id"],
                         "extraction_sha256": digest(output / "extraction-runs" / f"{run['run_id']}.json")})
    if not store.facts(CITY_SOURCE):
        raise ValueError("还没有 01 事实，请加 --extract 从城市报告抽取")
    inventory = parse_corpus(root, store)
    daily_content = daily_source.read_bytes()
    parsed = parse_document(daily_content, DAILY_SOURCE)
    store.save(parsed, daily_content)
    context_content = daily_source.with_name(DAILY_CONTEXT).read_bytes()
    context = parse_document(context_content, DAILY_CONTEXT)
    store.save(context, context_content)
    originals = {
        path: hashlib.sha256(path.read_bytes()).hexdigest() for path in (city_template, daily_template)
    }
    expected_city, city_locators = city_expected(city_source)
    expected_daily, daily_locators = daily_expected(daily_source)
    expected_rules = checked_rules(load_expectations(), "same-db")
    city_rules, daily_rules = expected_rules["01"], expected_rules["03"]
    if len(expected_city) != city_rules["rows"] or len(expected_daily) != daily_rules["rows"]:
        raise ValueError("独立原文核对得出的预期行数不符合任务书")
    if sum(value is None for row in expected_daily for value in row) != daily_rules["expected_empty"]:
        raise ValueError("逐日原文的空格数与验收配置不一致")

    results = {}
    scores = {}
    database_sha256 = digest(store.path)
    database_observed_at = datetime.now(timezone.utc).isoformat(timespec="milliseconds")
    for name, source, template, workbook, expected, locators, width in (
        ("01", CITY_SOURCE, city_template, CITY_OUTPUT, expected_city, city_locators, 5),
        ("03", DAILY_SOURCE, daily_template, COVID_OUTPUT, expected_daily, daily_locators, 7),
    ):
        trace_file, _, score_file = REPORT_FILES[workbook]
        previous = None
        timings = []
        for _ in range(2):
            outcome = fill_workbook(store, source, template, output)
            timings.append(outcome.pop("fill_seconds"))
            audit = verify_workbook(output / workbook, expected, width)
            trace = verify_trace(store, output / trace_file, locators, width)
            current = {**outcome, **audit}
            if previous is not None and previous != current:
                raise ValueError(f"{name} 同库两次填写结果不一致")
            previous = current
        results[name] = {**current, "fill_seconds": timings, "block_count": len(store.blocks(source))}
        scores[name] = source_score(trace, audit, name == "03", database_sha256,
                                    database_observed_at=database_observed_at)
        write_json(output / score_file, scores[name])
        if hashlib.sha256(template.read_bytes()).hexdigest() != originals[template]:
            raise ValueError(f"模板原件出现变更：{template.name}")

    report = {
        "basis": "independent-source-check",
        "mode": mode,
        "corpus": inventory,
        "scenarios": results,
        "daily_context": {
            "file": DAILY_CONTEXT, "blocks": len(context.blocks),
            "used_for_fill": False,
            "reason": "单日省级记述与模板需要的国家或地区逐日统计口径不同",
        },
    }
    write_json(output / "source_audit.json", report)
    write_json(output / RUN_EVIDENCE, {
        **evidence, "database_sha256": database_sha256,
        "outputs": {case: {key: score[key] for key in ("run_id", "source_sha256", "template_sha256", "workbook_sha256")}
                    for case, score in scores.items()},
    })
    failures = check(scores, expectations, "source")
    for name, width in (("01", 5), ("03", 7)):
        if results[name]["columns"] != width:
            failures.append(f"{name} 输出列数预期 {width}，实际 {results[name]['columns']}")
    if failures:
        raise ValueError("\n".join(failures) + f"\n完整差异见 {output / 'source_audit.json'}")
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description="用测试集原文独立核对 01 和 03，输出本地审计记录")
    parser.add_argument("--root", type=Path, default=DATASET)
    parser.add_argument("--data-dir", type=Path, default=OUTPUT)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--extract", action="store_true", help="调用 DeepSeek 重新抽取 01 城市报告")
    mode.add_argument("--archived-facts", type=Path, help="在空库重建历史事实，无网络调用")
    args = parser.parse_args()
    try:
        report = check_dataset(args.root, args.data_dir, args.extract, args.archived_facts)
    except (OSError, ValueError) as error:
        parser.exit(1, f"核对失败：{error}\n")
    print(json.dumps(report["scenarios"], ensure_ascii=False, indent=2))
    print(f"{len(report['corpus'])} 份语料解析完成，报告：{args.data_dir / 'source_audit.json'}")


if __name__ == "__main__":
    main()
