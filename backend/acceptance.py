import argparse
import json
import math
import sqlite3
from contextlib import closing
from hashlib import sha256
from numbers import Real
from pathlib import Path

from config import DATA
from demo import build_demo
from fill_spec import CITY_OUTPUT, COVID_OUTPUT, REPORT_FILES
from filler import json_digest

DEMO_EXPECTATIONS = {
    "01": {"block_count": 1, "rows": 1, "cells_written": 5, "cells_empty": 0,
           "cell_accuracy": 1, "row_completeness": 1, "city_order_matches": 1},
    "03": {"rows": 3, "required": 19, "expected_empty": 2, "correct": 19,
           "false_fills": 0, "missed": 0},
}
REPORTS = {"01": "eval_report.json", "03": "eval_report_covid19.json"}
OUTPUTS = {"01": CITY_OUTPUT, "03": COVID_OUTPUT}
FIXED = {
    "01": {"block_count": 113, "rows": 100, "city_order_matches": 100},
    "03": {"block_count": 15088, "rows": 2542, "required": 17464, "expected_empty": 330},
}
MODEL_FIELDS = {"cells_written", "cells_empty", "cell_accuracy", "row_completeness"}
REQUIRED = {
    "01": set(FIXED["01"]) | MODEL_FIELDS,
    "03": set(FIXED["03"]) | {
        "cells_written", "cells_empty", "correct", "false_fills", "missed",
        "correct_rate", "false_fill_rate", "missed_rate",
    },
}
BASES = {"source": "independent-source-check", "reference": "reference-workbook"}


def is_number(value: object) -> bool:
    return isinstance(value, Real) and not isinstance(value, bool) and math.isfinite(value)


def checked_rules(config: dict, mode: str) -> dict:
    rules = config[mode]
    if set(rules) != set(REQUIRED):
        raise ValueError("验收配置必须同时包含场景 01 和 03")
    for scenario, expected in rules.items():
        missing = REQUIRED[scenario] - expected.keys()
        if missing:
            raise ValueError(f"{scenario} 缺少必验字段：{', '.join(sorted(missing))}")
        for key, target in expected.items():
            if key in FIXED[scenario] and target != FIXED[scenario][key]:
                raise ValueError(f"{scenario}.{key} 必须精确等于 {FIXED[scenario][key]}，配置为 {target!r}")
            if is_number(target):
                continue
            if mode != "rebuilt" or scenario != "01" or key not in MODEL_FIELDS or not isinstance(target, list):
                raise ValueError(f"{scenario}.{key} 必须使用精确数值，配置为 {target!r}")
            points = config.get("observations", [])
            valid = {
                point["run_id"]: point["01"][key] for point in points
                if point.get("run_id") and is_number(point.get("01", {}).get(key))
            }
            if len(valid) < 2 or target != [min(valid.values()), max(valid.values())]:
                raise ValueError(f"01.{key} 区间需由至少两个独立重抽的实测值确定，配置为 {target!r}")
    return rules


def checked_outputs(directory: Path, basis: str) -> dict:
    database = directory / "a23.db"
    database_hash = sha256(database.read_bytes()).hexdigest()
    reports = {}
    for case, file in OUTPUTS.items():
        trace_name, reference_name, source_name = REPORT_FILES[file]
        report_name = source_name if basis == "source" else reference_name
        report_path = directory / report_name
        if not report_path.is_file():
            raise ValueError(f"{case} 缺少本次报告 {report_path}，预期文件存在，实际缺失")
        report = json.loads(report_path.read_text(encoding="utf-8"))
        trace_path = directory / trace_name
        if not trace_path.is_file():
            raise ValueError(f"{case} 缺少逐格记录 {trace_path}，预期文件存在，实际缺失")
        trace = json.loads(trace_path.read_text(encoding="utf-8"))
        expected_basis = BASES[basis]
        if report.get("basis") != expected_basis:
            raise ValueError(f"{case}.basis 预期 {expected_basis!r}，实际 {report.get('basis')!r}")
        identities = {
            "result_id": json_digest({
                key: value for key, value in trace.items()
                if key not in {"run_id", "result_id", "workbook_sha256"}
            }),
            "run_id": trace["run_id"],
            "workbook_sha256": sha256((directory / file).read_bytes()).hexdigest(),
            "source_sha256": trace["source_sha256"],
            "template_sha256": trace["template_sha256"],
            "database_sha256": database_hash,
            "doc": trace["doc"],
            "file": file,
        }
        for key, expected in identities.items():
            if report.get(key) != expected:
                raise ValueError(f"{case}.{key} 预期 {expected!r}，实际 {report.get(key)!r}，检查 {report_name}")
        with closing(sqlite3.connect(database)) as connection:
            row = connection.execute("SELECT content FROM documents WHERE name = ?", (trace["doc"],)).fetchone()
        if row is None or sha256(row[0]).hexdigest() != report["source_sha256"]:
            raise ValueError(f"{case} 数据库中的原文与本次报告不一致")
        reports[case] = report
    return reports


def check(scenarios: dict, expectations: dict, basis: str = "reference") -> list[str]:
    failures = []
    for scenario, rules in expectations.items():
        report = scenarios.get(scenario)
        if report is None:
            failures.append(f"{scenario} 缺少报告")
            continue
        for key, expected in rules.items():
            actual = report.get(key)
            try:
                if isinstance(expected, list):
                    low, high = expected
                    passed = is_number(actual) and low <= actual <= high
                else:
                    passed = is_number(actual) and actual == expected
            except (ValueError, TypeError):
                passed = False
            if not passed:
                report = REPORT_FILES[OUTPUTS[scenario]][2] if basis == "source" else REPORTS[scenario]
                failures.append(f"{scenario}.{key} 预期 {expected!r}，实际 {actual!r}，检查 {report}")
    return failures


def main() -> None:
    parser = argparse.ArgumentParser(description="按指定判据复核两份自检报告，失败时返回非零")
    parser.add_argument("--data-dir", type=Path, default=DATA)
    parser.add_argument("--expectations", type=Path, help="正式样本的实测判据 JSON")
    parser.add_argument("--mode", choices=("same-db", "rebuilt"), default="same-db")
    parser.add_argument("--basis", choices=BASES, default="reference", help="原文核对或正式参考表评分")
    args = parser.parse_args()
    try:
        if args.expectations is None:
            _, _, city, daily = build_demo(args.data_dir / "demo")
            reports, expectations = {"01": city, "03": daily}, DEMO_EXPECTATIONS
        else:
            expectations = checked_rules(json.loads(args.expectations.read_text(encoding="utf-8")), args.mode)
            reports = checked_outputs(args.data_dir, args.basis)
        failures = check(reports, expectations, args.basis)
    except (OSError, ValueError, KeyError, TypeError, sqlite3.Error) as error:
        parser.exit(1, f"验收失败：{error}\n")
    if failures:
        parser.exit(1, "\n".join(failures) + "\n")
    print(f"验收通过，共 {sum(len(rules) for rules in expectations.values())} 条数值判据")


if __name__ == "__main__":
    main()
