import argparse
import json
from pathlib import Path

from config import DATA
from demo import build_demo

DEMO_EXPECTATIONS = {
    "01": {"block_count": 1, "rows": 1, "cells_written": 5, "cells_empty": 0,
           "cell_accuracy": 1, "row_completeness": 1, "city_order_matches": 1},
    "03": {"rows": 3, "required": 19, "expected_empty": 2, "correct": 19,
           "false_fills": 0, "missed": 0},
}
REPORTS = {"01": "eval_report.json", "03": "eval_report_covid19.json"}


def check(scenarios: dict, expectations: dict) -> list[str]:
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
                    passed = low <= actual <= high
                else:
                    passed = actual == expected
            except (ValueError, TypeError):
                passed = False
            if not passed:
                failures.append(f"{scenario}.{key} 预期 {expected!r}，实际 {actual!r}，检查 {REPORTS[scenario]}")
    return failures


def main() -> None:
    parser = argparse.ArgumentParser(description="按指定判据复核两份自检报告，失败时返回非零")
    parser.add_argument("--data-dir", type=Path, default=DATA)
    parser.add_argument("--expectations", type=Path, help="正式样本的实测判据 JSON")
    parser.add_argument("--mode", choices=("same-db", "rebuilt"), default="same-db")
    args = parser.parse_args()
    try:
        if args.expectations is None:
            _, _, city, daily = build_demo(args.data_dir / "demo")
            reports, expectations = {"01": city, "03": daily}, DEMO_EXPECTATIONS
        else:
            expectations = json.loads(args.expectations.read_text(encoding="utf-8"))[args.mode]
            if not expectations or any(name not in REPORTS for name in expectations):
                raise ValueError("判据须指定至少一个已知场景 01 或 03")
            if args.mode == "same-db" and any(
                isinstance(value, list) for rules in expectations.values() for value in rules.values()
            ):
                raise ValueError("同库复验须使用逐项精确值")
            reports = {}
            for scenario in expectations:
                path = args.data_dir / REPORTS[scenario]
                if path.is_file():
                    reports[scenario] = json.loads(path.read_text(encoding="utf-8"))
        failures = check(reports, expectations)
    except (OSError, ValueError, KeyError) as error:
        parser.exit(1, f"验收失败：{error}\n")
    if failures:
        parser.exit(1, "\n".join(failures) + "\n")
    print(f"验收通过，共 {sum(len(rules) for rules in expectations.values())} 条数值判据")


if __name__ == "__main__":
    main()
