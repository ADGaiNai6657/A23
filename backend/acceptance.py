import argparse
import json
import re
import sqlite3
from contextlib import closing
from hashlib import sha256
from pathlib import Path

from config import BACKEND, DATA
from criteria import FIELDS, MODEL_FIELDS, checked_rules, is_number, load_expectations
from demo import build_demo
from fill_spec import CITY_OUTPUT, COVID_OUTPUT, REPORT_FILES
from filler import json_digest
from rebuild_evidence import check_execution

OUTPUTS = {"01": CITY_OUTPUT, "03": COVID_OUTPUT}
BASES = {"source": "independent-source-check", "reference": "reference-workbook"}
DELIVERABLES = BACKEND / "data" / "deliverables"
DEMO_DELIVERY = DELIVERABLES / "demo"


def required_file(directory: Path, filename: str, case: str) -> bytes:
    path = directory / filename
    try:
        return path.read_bytes()
    except FileNotFoundError as error:
        raise ValueError(f"{case} 缺少 {filename}，预期文件存在，实际缺失（{path}）") from error


def checked_outputs(directory: Path, basis: str, *, require_database: bool = True) -> dict:
    database = directory / "a23.db"
    database_hash = sha256(required_file(directory, database.name, "数据库")).hexdigest() if require_database else None
    reports = {}
    for case, file in OUTPUTS.items():
        trace_name, reference_name, source_name = REPORT_FILES[file]
        report_name = source_name if basis == "source" else reference_name
        try:
            report = json.loads(required_file(directory, report_name, case))
            trace = json.loads(required_file(directory, trace_name, case))
        except json.JSONDecodeError as error:
            raise ValueError(f"{case} JSON 无法读取，检查 {report_name} 和 {trace_name}：{error}") from error
        expected_basis = BASES[basis]
        if report.get("basis") != expected_basis:
            raise ValueError(f"{case}.basis 预期 {expected_basis!r}，实际 {report.get('basis')!r}")
        if basis == "reference" and not re.fullmatch(r"[0-9a-f]{64}", report.get("reference_sha256") or ""):
            raise ValueError(f"{case} 缺少有效的参考表 SHA256")
        identities = {
            "result_id": json_digest({
                key: value for key, value in trace.items()
                if key not in {"run_id", "result_id", "workbook_sha256", "fill_seconds"}
            }),
            "run_id": trace.get("run_id"),
            "workbook_sha256": sha256(required_file(directory, file, case)).hexdigest(),
            "source_sha256": trace.get("source_sha256"),
            "template_sha256": trace.get("template_sha256"),
            "doc": trace.get("doc"),
            "file": file,
        }
        for key, expected in identities.items():
            if not expected or report.get(key) != expected or trace.get(key) != expected:
                raise ValueError(
                    f"{case}.{key} 预期 {expected!r}，实际报告 {report.get(key)!r}、逐格记录 {trace.get(key)!r}"
                    f"，检查 {report_name} 和 {trace_name}"
                )
        if database_hash is not None:
            if report.get("database_sha256") != database_hash:
                raise ValueError(f"{case}.database_sha256 与数据库不符，检查 {report_name}")
            with closing(sqlite3.connect(database.resolve().as_uri() + "?mode=ro", uri=True)) as connection:
                row = connection.execute("SELECT content FROM documents WHERE name = ?", (trace["doc"],)).fetchone()
            if row is None or sha256(row[0]).hexdigest() != report["source_sha256"]:
                raise ValueError(f"{case} 数据库中的原文与本次报告不一致")
        reports[case] = report
    return reports


def validate_profile(config: dict, basis: str, reports: dict | None = None) -> None:
    if config.get("basis") != BASES[basis]:
        raise ValueError(f"判据的 basis 必须为 {BASES[basis]}")
    if basis == "reference":
        hashes = config.get("reference_sha256", {})
        if set(hashes) != set(OUTPUTS) or not all(re.fullmatch(r"[0-9a-f]{64}", value) for value in hashes.values()):
            raise ValueError("参考判据需分别为 01 和 03 指定参考表 SHA256")
        if reports is not None:
            for case, report in reports.items():
                if report.get("reference_sha256") != hashes[case]:
                    raise ValueError(f"{case} 评分参考表 SHA256 与独立判据不一致")


def sample_rules(config: dict, mode: str, basis: str, config_dir: Path) -> dict:
    validate_profile(config, basis)
    points = None
    if basis == "reference" and mode == "rebuilt":
        points = []
        for entry in config.get("observations", []):
            directory = config_dir / entry["data_dir"]
            point_mode = entry["mode"]
            rebuilt_kind = entry.get("rebuilt_kind", "online")
            if point_mode not in {"same-db", "rebuilt"} or rebuilt_kind not in {"online", "history"}:
                raise ValueError("参考判据观察点的模式无效")
            scores = checked_outputs(directory, basis)
            validate_profile(config, basis, scores)
            check_execution(directory, point_mode, scores, rebuilt_kind=rebuilt_kind)
            points.append(scores["01"])
        if (len({point["database_sha256"] for point in points}) != len(points)
                or len({point["run_id"] for point in points}) != len(points)
                or len({(point["source_sha256"], point["template_sha256"]) for point in points}) != 1
                or not any(entry["mode"] == "rebuilt" for entry in config.get("observations", []))):
            raise ValueError("参考观察点需要同一原文与模板、独立数据库及填写批次，且含重建记录")
    return checked_rules(config, mode, reference_points=points)


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
                filename = REPORT_FILES[OUTPUTS[scenario]][2 if basis == "source" else 1]
                failures.append(f"{scenario}.{key} 预期 {expected!r}，实际 {actual!r}，检查 {filename}")
    return failures


def print_criteria(config: dict, basis: str, config_dir: Path) -> None:
    exact, rebuilt = (sample_rules(config, mode, basis, config_dir) for mode in ("same-db", "rebuilt"))
    print("| 名称 | 键名 | 判据 | 判据类型 | 适用场景 | 失败时检查的位置 |")
    print("| --- | --- | --- | --- | --- | --- |")
    for case, fields in FIELDS.items():
        report_name = REPORT_FILES[OUTPUTS[case]][2 if basis == "source" else 1]
        for key, label in fields.items():
            values = {"same-db": exact[case][key], "rebuilt": rebuilt[case][key]}
            for mode, value in values.items():
                kind = "区间" if mode == "rebuilt" and case == "01" and key in MODEL_FIELDS else "逐位"
                print(f"| {label} | `{key}` | {value} | {kind} | {case} {mode} | `{report_name}` |")


def main() -> None:
    parser = argparse.ArgumentParser(description="按指定判据复核两份自检报告，失败时返回非零")
    parser.add_argument("--data-dir", type=Path, default=DATA)
    parser.add_argument("--expectations", type=Path, help="正式样本的实测判据 JSON")
    parser.add_argument("--mode", choices=("same-db", "rebuilt"), default="same-db")
    parser.add_argument("--rebuilt-kind", choices=("online", "history"), default="online")
    parser.add_argument("--basis", choices=BASES, help="原文核对或正式参考表评分")
    parser.add_argument("--print-criteria", action="store_true", help="从权威配置生成六字段判据表")
    parser.add_argument("--check-existing", action="store_true", help="只检查指定目录的合成报告，不重新生成")
    args = parser.parse_args()
    try:
        config = load_expectations(args.expectations) if args.expectations else load_expectations()
        basis = args.basis or ("source" if args.expectations or args.print_criteria else "reference")
        config_dir = args.expectations.resolve().parent if args.expectations else BACKEND
        if args.print_criteria:
            print_criteria(config, basis, config_dir)
            return
        if args.expectations is None:
            directory = args.data_dir
            if not args.check_existing:
                _, directory, _, _ = build_demo(directory / "demo")
            reports, expectations = checked_outputs(directory, basis, require_database=not args.check_existing), config["demo"]
        else:
            expectations = sample_rules(config, args.mode, basis, config_dir)
            reports = checked_outputs(args.data_dir, basis)
            validate_profile(config, basis, reports)
            check_execution(args.data_dir, args.mode, reports, rebuilt_kind=args.rebuilt_kind)
        failures = check(reports, expectations, basis)
    except (OSError, ValueError, KeyError, TypeError, sqlite3.Error) as error:
        parser.exit(1, f"验收失败：{error}\n")
    if failures:
        parser.exit(1, "\n".join(failures) + "\n")
    print(f"验收通过，共 {sum(len(rules) for rules in expectations.values())} 条数值判据")


if __name__ == "__main__":
    main()
