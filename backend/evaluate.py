import json
import math
from io import BytesIO
from datetime import date, datetime
from hashlib import sha256
from itertools import zip_longest
from numbers import Real
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

from criteria import checked_rules, load_expectations
from fill_spec import CITY_HEADERS, DATE_KEYS, REPORT_FILES, normalize
from filler import write_json


def matches(actual: object, expected: object) -> bool:
    if isinstance(expected, datetime):
        expected = expected.date()
    if isinstance(actual, datetime):
        actual = actual.date()
    if isinstance(expected, date):
        return actual == expected
    if isinstance(expected, Real) and not isinstance(expected, bool):
        return isinstance(actual, Real) and not isinstance(actual, bool) and math.isclose(
            actual, expected, rel_tol=1e-12, abs_tol=1e-6,
        )
    if isinstance(expected, str):
        return isinstance(actual, str) and actual.strip() == expected.strip()
    return actual == expected


def excel_day(value: object, epoch: datetime) -> object:
    if isinstance(value, Real) and not isinstance(value, bool):
        try:
            value = from_excel(value, epoch)
        except (OverflowError, ValueError):
            return value
    return value.date() if isinstance(value, datetime) else value


def scoring_origin(filled: Path, content: bytes, doc: str) -> dict:
    filenames = REPORT_FILES.get(filled.name)
    if filenames is None:
        return {}
    trace = json.loads(filled.with_name(filenames[0]).read_text(encoding="utf-8"))
    if trace.get("doc") != doc:
        raise ValueError("当前填写工作簿属于另一份原文")
    if trace.get("workbook_sha256") != sha256(content).hexdigest() or not trace.get("result_id"):
        raise ValueError("工作簿与本次逐格结果版本不一致，请重新填写后评分")
    return {
        key: trace[key] for key in (
            "result_id", "run_id", "source_sha256", "template_sha256", "workbook_sha256",
            "fill_seconds", "refusal_counts",
        ) if key in trace
    }


def evaluate(
    filled: Path, reference: Path, report: Path, doc: str, covid: bool,
    expected_rows: int | None = None, block_count: int | None = None,
    database: Path | None = None,
) -> dict:
    filled_content = filled.read_bytes()
    identity = scoring_origin(filled, filled_content, doc)
    reference_content = reference.read_bytes()
    result_book = load_workbook(BytesIO(filled_content), read_only=True, data_only=True)
    try:
        reference_book = load_workbook(BytesIO(reference_content), read_only=True, data_only=True)
        try:
            actual = result_book.active
            expected = reference_book.active
            width = 7 if covid else len(CITY_HEADERS)
            count = expected_rows if expected_rows is not None else expected.max_row - 1
            if count < 1 or expected.max_row != count + 1:
                raise ValueError("参考结果行数与验收范围不符")
            if expected.max_column != width:
                raise ValueError("参考结果列数与验收范围不符")
            if actual.max_column != width or actual.max_row != count + 1:
                raise ValueError("填写结果的行列数与验收范围不符")
            columns = next(actual.iter_rows(max_row=1, values_only=True))
            reference_columns = next(expected.iter_rows(max_row=1, values_only=True))
            header_keys = tuple(normalize(str(title or "")) for title in columns)
            if header_keys != tuple(normalize(str(title or "")) for title in reference_columns):
                raise ValueError("参考结果列名或列顺序与填写工作簿不一致")
            if not covid and header_keys != tuple(map(normalize, CITY_HEADERS)):
                raise ValueError("01 参考结果表头与五列字段定义不一致")
            date_columns = {
                column for column, title in enumerate(columns, 1)
                if covid and normalize(str(title)) in DATE_KEYS
            }
            actual_rows = actual.iter_rows(min_row=2, max_row=count + 1, max_col=width, values_only=True)
            expected_rows_iter = expected.iter_rows(min_row=2, max_row=count + 1, max_col=width, values_only=True)
            differences = []
            city_names = set()
            correct = complete = allowed_empty = false_fills = missed = same_city = written = wrong = 0
            pairs = zip_longest(actual_rows, expected_rows_iter, fillvalue=(None,) * width)
            for row, (seen, target) in enumerate(pairs, 2):
                complete += all(value is not None for value in seen)
                same_city += matches(seen[0], target[0])
                for column, (value, truth) in enumerate(zip(seen, target), 1):
                    if not covid:
                        valid = (
                            isinstance(truth, str) and bool(truth.strip()) if column == 1
                            else isinstance(truth, Real) and not isinstance(truth, bool) and math.isfinite(truth)
                        )
                        if not valid:
                            raise ValueError(f"01 参考结果第 {row} 行第 {column} 列缺少有效应填值")
                        if column == 1:
                            city_name = normalize(truth).removesuffix("市")
                            if city_name in city_names:
                                raise ValueError(f"01 参考结果第 {row} 行重复出现城市 {truth}")
                            city_names.add(city_name)
                    if column in date_columns:
                        value = excel_day(value, result_book.epoch)
                        truth = excel_day(truth, reference_book.epoch)
                    allowed_empty += truth is None
                    written += value is not None
                    matched = matches(value, truth)
                    correct += matched and (truth is not None or not covid)
                    if matched:
                        continue
                    false_fills += truth is None and value is not None
                    missed += truth is not None and value is None
                    wrong += truth is not None and value is not None
                    decision = "误填" if truth is None else "漏填" if value is None else "值不符"
                    differences.append({"row": row, "column": column, "actual": value, "expected": truth, "decision": decision})
            total = width * count
            output = {
                "doc": doc, "file": filled.name, "rows": actual.max_row - 1,
                "cells_written": written, "cells_empty": total - written, "diff": differences,
                "expected_empty": allowed_empty, "false_fills": false_fills, "missed": missed,
                "wrong_values": wrong,
                "basis": "reference-workbook",
                "reference_sha256": sha256(reference_content).hexdigest(),
                **identity,
            }
            if block_count is not None:
                output["block_count"] = block_count
            if database is not None:
                output["database_sha256"] = sha256(database.read_bytes()).hexdigest()
            if covid:
                required = total - allowed_empty
                if expected_rows is not None:
                    rules = checked_rules(load_expectations(), "same-db")["03"]
                    if (required, allowed_empty) != (rules["required"], rules["expected_empty"]):
                        raise ValueError("03 正式参考结果的应填、应空格数与判据不一致")
                output.update({
                    "required": required, "expected_empty": allowed_empty, "correct": correct,
                    "false_fills": false_fills, "missed": missed,
                    "correct_rate": correct / required if required else None,
                    "false_fill_rate": false_fills / allowed_empty if allowed_empty else None,
                    "missed_rate": missed / required if required else None,
                })
            else:
                output.update({
                    "cell_accuracy": correct / total, "row_completeness": complete / count,
                    "city_order_matches": same_city,
                })
            write_json(report, output)
            return output
        finally:
            reference_book.close()
    finally:
        result_book.close()
