import argparse
import json
from pathlib import Path

from config import DATA
from criteria import checked_rules, load_expectations
from evaluate import evaluate
from fill_spec import CITY_OUTPUT, COVID_OUTPUT
from filler import fill_workbook
from store import Store


def main() -> None:
    parser = argparse.ArgumentParser(description="填写模板并生成独立评估报告")
    parser.add_argument("doc", help="已入库的文档名")
    parser.add_argument("template", type=Path)
    parser.add_argument("reference", type=Path)
    parser.add_argument("--scenario", choices=("01", "03"), required=True)
    parser.add_argument("--data-dir", type=Path, default=DATA)
    parser.add_argument("--synthetic", action="store_true", help="按合成参考工作簿的实际数据行数评估")
    parser.add_argument("--target-period", help="未写明年份的 01 标题可通过此项指定目标年份")
    parser.add_argument("--score-existing", action="store_true", help="评分当前工作簿，保留本轮填写批次")
    args = parser.parse_args()
    covid = args.scenario == "03"
    if not args.doc.lower().endswith(".xlsx" if covid else ".docx"):
        parser.error("文档后缀与所选场景不匹配")
    output, report = (COVID_OUTPUT, "eval_report_covid19.json") if covid else (CITY_OUTPUT, "eval_report.json")
    store = Store(args.data_dir / "a23.db")
    filled = {}
    if not args.score_existing:
        filled = fill_workbook(store, args.doc, args.template, args.data_dir,
                               target_period=args.target_period, formal_city=not args.synthetic and not covid)
    rows = None if args.synthetic else checked_rules(load_expectations(), "same-db")[args.scenario]["rows"]
    grading = evaluate(
        args.data_dir / output, args.reference, args.data_dir / report, args.doc, covid,
        expected_rows=rows, block_count=len(store.blocks(args.doc)), database=store.path,
    )
    print(json.dumps({**filled, **{key: value for key, value in grading.items() if key != "diff"}},
                     ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
