from pathlib import Path
import shutil

import pandas as pd

from .config import DEFAULT_RULE_RENAME_MAP
from .reports import write_rule_report
from .utils import ensure_dir, prefix_number, sanitize_filename_component, split_filename, timestamp
from .zip_utils import make_zip


def build_rule_plan(paths, rules=None):
    rules = rules or DEFAULT_RULE_RENAME_MAP
    rows = []
    for raw_path in paths:
        path = Path(raw_path)
        stem, extension, _ = split_filename(path.name)
        number = prefix_number(stem)
        hit = number in rules
        target = rules[number] if hit else sanitize_filename_component(stem)
        output_name = f"{sanitize_filename_component(target)}{extension}"
        rows.append({
            "原文件名": path.name,
            "输出文件名": output_name,
            "原扩展名": extension,
            "提取编号": number,
            "是否命中规则": hit,
            "规则目标主体": target,
            "处理方式": "命中规则" if hit else "未命中，清洗原主体",
            "状态": "待执行",
            "source_path": str(path.resolve()),
        })
    return pd.DataFrame(rows)


def validate_rule_plan(plan_df):
    duplicated = plan_df["输出文件名"].str.lower().duplicated(keep=False)
    if duplicated.any():
        return {"ok": False, "errors": plan_df.loc[duplicated, ["原文件名", "输出文件名"]]}
    return {"ok": True, "errors": pd.DataFrame()}


def execute_rule_rename(paths, output_dir, zip_dir, rules=None):
    plan = build_rule_plan(paths, rules)
    result = validate_rule_plan(plan)
    if not result["ok"]:
        return {"ok": False, "error": "规则输出文件名冲突", "errors": result["errors"]}
    output_dir = ensure_dir(output_dir)
    copied = []
    for index, row in plan.iterrows():
        src = Path(row["source_path"])
        dst = output_dir / row["输出文件名"]
        shutil.copy2(src, dst)
        copied.append(dst)
        plan.loc[index, "状态"] = "成功"
    report = write_rule_report(plan, output_dir / "rule_rename_report.xlsx")
    copied.append(report)
    zip_path = make_zip(copied, Path(zip_dir) / f"rule_renamed_files_{timestamp()}.zip")
    return {"ok": True, "plan_df": plan, "report_path": report, "zip_path": zip_path}
