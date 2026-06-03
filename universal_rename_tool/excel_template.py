from pathlib import Path

import pandas as pd

from .utils import parse_bool, strip_same_extension


def export_template(file_inventory, rename_state, output_path):
    rows = []
    for item in file_inventory:
        state = rename_state[item["original_name"]]
        rows.append({
            "原文件名": item["original_name"],
            "原扩展名": item["extension"],
            "新文件名（不含扩展名）": state.get("new_stem", ""),
            "是否保留原扩展名": state.get("keep_extension", True),
            "采用AI建议": state.get("use_ai_suggestion", False),
        })
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(rows).to_excel(output_path, index=False)
    return output_path


def import_template(file_inventory, rename_state, template_path):
    df = pd.read_excel(template_path)
    by_name = {item["original_name"]: item for item in file_inventory}
    updated = skipped = 0
    for _, row in df.iterrows():
        original = str(row.get("原文件名", "")).strip()
        if original not in by_name:
            skipped += 1
            continue
        new_value = row.get("新文件名（不含扩展名）", "")
        if pd.isna(new_value) or not str(new_value).strip():
            skipped += 1
            continue
        item = by_name[original]
        state = rename_state[original]
        state["new_stem"] = strip_same_extension(str(new_value), item["extension"])
        state["keep_extension"] = parse_bool(row.get("是否保留原扩展名", True))
        state["use_ai_suggestion"] = parse_bool(row.get("采用AI建议", False))
        updated += 1
    return {"updated": updated, "skipped": skipped}
