import argparse
import json
import os
import shutil
import sqlite3
import subprocess
import sys
from pathlib import Path
from tempfile import TemporaryDirectory
from time import perf_counter
from zipfile import BadZipFile, ZipFile

from acceptance import DELIVERABLES, DEMO_DELIVERY, check, checked_outputs
from config import BACKEND, DIST, ROOT
from criteria import EXPECTATIONS, checked_rules, load_expectations
from demo import build_demo, fill_demo
from fill_spec import CITY_OUTPUT, REPORT_FILES
from filler import write_json
from rebuild_evidence import check_distinct_runs, check_execution


def verify_reports(directory: Path, basis: str, rules: dict, *, require_database: bool = True) -> dict:
    reports = checked_outputs(directory, basis, require_database=require_database)
    failures = check(reports, rules, basis)
    if failures:
        raise ValueError("\n".join(failures))
    return reports


def verify_damage(source: Path, root: Path, rules: dict) -> None:
    directory = root / "damage"
    shutil.copytree(source, directory)
    report = directory / REPORT_FILES[CITY_OUTPUT][1]
    original = report.read_bytes()
    for code, field in (("改数", "cell_accuracy"), ("改键", "rows"), ("删报告", None), ("恢复", None)):
        report.write_bytes(original)
        if code == "删报告":
            report.unlink()
        elif field is not None:
            modified = json.loads(original)
            if code == "改键":
                modified[f"{field}_changed"] = modified.pop(field)
            else:
                modified[field] /= 2
            write_json(report, modified)
        run = subprocess.run(
            [sys.executable, str(BACKEND / "acceptance.py"), "--check-existing", "--data-dir", str(directory)],
            capture_output=True, text=True,
        )
        print(f"{code} | exit={run.returncode} | {(run.stderr or run.stdout).strip()}", flush=True)
        expected_location = f"01.{field}" if field else f"缺少 {report.name}"
        if (code == "恢复" and run.returncode != 0) or (
            code != "恢复" and (run.returncode == 0 or expected_location not in run.stderr)
        ):
            raise ValueError(f"破坏性验收未正确报告：{code}")
        if code in {"改数", "恢复"}:
            pytest_run = subprocess.run(
                [sys.executable, "-m", "pytest", "-q", "backend/test_a23.py",
                 "-k", "delivered_report_predicate and cell_accuracy"],
                cwd=ROOT, capture_output=True, text=True,
                env={**os.environ, "A23_TEST_REPORT_DIR": str(directory)},
            )
            print(f"{code} pytest | exit={pytest_run.returncode} | {pytest_run.stdout.strip()}", flush=True)
            if (code == "恢复" and pytest_run.returncode != 0) or (
                code == "改数" and (pytest_run.returncode != 1 or "01.cell_accuracy" not in pytest_run.stdout)
            ):
                raise ValueError(f"pytest 未正确报告：{code}")
    verify_reports(directory, "reference", rules, require_database=False)


def verify_demo(directory: Path, root: Path, rules: dict) -> None:
    verify_reports(directory, "reference", rules, require_database=False)
    print(f"合成交付文件通过：{directory}", flush=True)
    inputs, output, _, _ = build_demo(root / "same")
    before = verify_reports(output, "reference", rules)
    fill_demo(inputs, output)
    after = verify_reports(output, "reference", rules)
    if any(before[case]["result_id"] != after[case]["result_id"] for case in before):
        raise ValueError("合成同库两次填写的内容标识不同")
    print("合成 same-db 通过（同库重新填写并逐项比较）", flush=True)
    _, rebuilt, _, _ = build_demo(root / "rebuilt")
    verify_reports(rebuilt, "reference", rules)
    print("合成 rebuilt 通过（空库执行固定模型替身并重新入库）", flush=True)
    verify_damage(directory, root, rules)


def verify_published_archives(root: Path, config: dict) -> None:
    archives = (
        ("source/same-db.zip", "same-db"), ("source/rebuilt-history.zip", "rebuilt"),
        ("follow-up/same-db.zip", "same-db"), ("follow-up/city-online.zip", "rebuilt"),
    )
    for archive, mode in archives:
        with ZipFile(DELIVERABLES / archive) as bundle:
            directory = root / archive
            bundle.extractall(directory)
        verify_reports(directory, "source", checked_rules(config, mode), require_database=False)
        print(f"公开样本交付文件通过：{archive}（核对报告与工作簿）")


def main() -> None:
    parser = argparse.ArgumentParser(description="运行 pytest、交付文件、双模式和临时破坏性验收")
    parser.add_argument("--demo-reports", type=Path, default=DEMO_DELIVERY)
    parser.add_argument("--same-db-dir", type=Path, help="本轮同库填写及评分目录")
    parser.add_argument("--rebuilt-data-dir", type=Path, help="独立重建并评分的目录")
    parser.add_argument("--rebuilt-kind", choices=("online", "history"), default="online")
    parser.add_argument("--basis", choices=("source", "reference"), default="source")
    args = parser.parse_args()
    extra_runs = {"same-db": args.same_db_dir, "rebuilt": args.rebuilt_data_dir}
    if any(extra_runs.values()) and not all(extra_runs.values()):
        parser.error("样本验收同时需要 --same-db-dir 和 --rebuilt-data-dir")
    if args.same_db_dir and args.same_db_dir.resolve() == args.rebuilt_data_dir.resolve():
        parser.error("同库和重建库必须使用不同目录")
    if not (DIST / "index.html").is_file():
        parser.error("缺少 frontend/dist/index.html，请先在 frontend 执行 npm run build")
    started = perf_counter()
    try:
        with TemporaryDirectory(prefix="a23-verify-") as temporary:
            root = Path(temporary)
            subprocess.run(
                [sys.executable, "-m", "pytest", "-q", "backend"], cwd=ROOT, check=True,
                env={**os.environ, "A23_TEST_REPORT_DIR": str(args.demo_reports.resolve()),
                     "A23_DATA_DIR": str(root / "pytest"), "A23_REQUIRE_FRONTEND": "1"},
            )
            config = load_expectations(EXPECTATIONS)
            verify_demo(args.demo_reports, root, config["demo"])
            verify_published_archives(root, config)
            checked = {}
            for mode, directory in extra_runs.items():
                if directory is not None:
                    rules = checked_rules(config, mode)
                    checked[mode] = verify_reports(directory, args.basis, rules)
                    check_execution(directory, mode, checked[mode], rebuilt_kind=args.rebuilt_kind)
            if checked:
                check_distinct_runs(checked["same-db"], checked["rebuilt"])
                for mode, directory in extra_runs.items():
                    label = f"{mode}/{args.rebuilt_kind}" if mode == "rebuilt" else mode
                    print(f"{args.basis} {label} 通过（{sum(map(len, checked_rules(config, mode).values()))} 条判据）：{directory}")
    except (OSError, ValueError, KeyError, TypeError, sqlite3.Error,
            BadZipFile, subprocess.CalledProcessError) as error:
        parser.exit(1, f"综合验收失败：{error}\n")
    print(f"综合验收通过，耗时 {perf_counter() - started:.2f} 秒")


if __name__ == "__main__":
    main()
