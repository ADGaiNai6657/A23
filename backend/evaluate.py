import math
from datetime import date, datetime
from itertools import zip_longest
from numbers import Real
from pathlib import Path

from openpyxl import load_workbook

from fill_spec import CITY_HEADERS
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


def evaluate(
    filled: Path, reference: Path, report: Path, doc: str, covid: bool,
    expected_rows: int | None = None, block_count: int | None = None,
) -> dict:
    result_book = load_workbook(filled, read_only=True, data_only=True)
    try:
        reference_book = load_workbook(reference, read_only=True, data_only=True)
        try:
            actual = result_book.active
            expected = reference_book.active
            width = 7 if covid else len(CITY_HEADERS)
            count = expected_rows if expected_rows is not None else expected.max_row - 1
            if count < 1 or expected.max_row != count + 1:
                raise ValueError("参考结果行数与验收范围不符")
            if expected.max_column != width:
                raise ValueError("参考结果列数与验收范围不符")
            actual_rows = actual.iter_rows(min_row=2, max_row=count + 1, max_col=width, values_only=True)
            expected_rows_iter = expected.iter_rows(min_row=2, max_row=count + 1, max_col=width, values_only=True)
            differences = []
            correct = complete = allowed_empty = false_fills = missed = same_city = written = 0
            pairs = zip_longest(actual_rows, expected_rows_iter, fillvalue=(None,) * width)
            for row, (seen, target) in enumerate(pairs, 2):
                complete += all(value is not None for value in seen)
                same_city += matches(seen[0], target[0])
                for column, (value, truth) in enumerate(zip(seen, target), 1):
                    allowed_empty += truth is None
                    written += value is not None
                    matched = matches(value, truth)
                    correct += matched and (truth is not None or not covid)
                    if matched:
                        continue
                    false_fills += truth is None and value is not None
                    missed += truth is not None and value is None
                    decision = "误填" if truth is None else "漏填" if value is None else "值不符"
                    differences.append({"row": row, "column": column, "actual": value, "expected": truth, "decision": decision})
            total = width * count
            output = {
                "doc": doc, "file": filled.name, "rows": actual.max_row - 1,
                "cells_written": written, "cells_empty": total - written, "diff": differences,
            }
            if block_count is not None:
                output["block_count"] = block_count
            if covid:
                required = total - allowed_empty
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
