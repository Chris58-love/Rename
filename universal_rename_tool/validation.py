from pathlib import Path

import pandas as pd

from .utils import sanitize_filename_component


def build_rename_plan(file_inventory, rename_state):
    rows = []
    for item in file_inventory:
        state = rename_state[item["original_name"]]
        clean_stem = sanitize_filename_component(state.get("new_stem", ""))
        extension = item["extension"] if state.get("keep_extension", True) else ""
        final_name = f"{clean_stem}{extension}"
        rows.append({
            "原文件名": item["original_name"],
            "输出文件名": final_name,
            "原扩展名": item["extension"],
            "新文件名主体": state.get("new_stem", ""),
            "清洗后的新文件名主体": clean_stem,
            "是否保留扩展名": bool(state.get("keep_extension", True)),
            "是否实际改名": item["original_name"] != final_name,
            "AI建议名": state.get("ai_suggested_stem", ""),
            "AI建议原因": state.get("ai_reason", ""),
            "AI置信度": state.get("ai_confidence", ""),
            "AI审核结果": state.get("ai_review", ""),
            "AI审核风险": state.get("ai_review_risk_level", ""),
            "AI审核说明": state.get("ai_review_comment", ""),
            "处理说明": "",
            "状态": "待执行",
            "source_path": item["source_path"],
        })
    return pd.DataFrame(rows)


def validate_rename_plan(plan_df, case_insensitive=True):
    errors = []
    seen = {}
    for _, row in plan_df.iterrows():
        output_name = str(row["输出文件名"]).strip()
        if not output_name:
            errors.append({"问题类型": "输出名为空", "原文件名": row["原文件名"], "输出文件名": output_name})
            continue
        if Path(output_name).name != output_name or "/" in output_name or "\\" in output_name:
            errors.append({"问题类型": "路径不安全", "原文件名": row["原文件名"], "输出文件名": output_name})
        key = output_name.lower() if case_insensitive else output_name
        if key in seen:
            errors.append({"问题类型": "最终文件名冲突", "原文件名": row["原文件名"], "输出文件名": output_name, "冲突对象": seen[key]})
        else:
            seen[key] = row["原文件名"]
    return {"ok": not errors, "errors": pd.DataFrame(errors), "plan_df": plan_df}
