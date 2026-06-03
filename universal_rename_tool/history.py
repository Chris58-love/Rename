from pathlib import Path

import pandas as pd

from .config import HISTORY_COLUMNS
from .utils import prefix_number, timestamp


def _compatible(df):
    for column in HISTORY_COLUMNS:
        if column not in df.columns:
            df[column] = ""
    return df[HISTORY_COLUMNS]


def load_history(path):
    path = Path(path)
    if not path.exists():
        return pd.DataFrame(columns=HISTORY_COLUMNS)
    try:
        if path.suffix.lower() == ".csv":
            return _compatible(pd.read_csv(path))
        return _compatible(pd.read_excel(path))
    except Exception:
        return pd.DataFrame(columns=HISTORY_COLUMNS)


def append_history(path, plan_df):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    history = load_history(path)
    rows = []
    for _, row in plan_df.iterrows():
        rows.append({
            "时间戳": timestamp(),
            "原文件名": row.get("原文件名", ""),
            "输出文件名": row.get("输出文件名", ""),
            "原扩展名": row.get("原扩展名", ""),
            "新文件名主体": row.get("新文件名主体", ""),
            "清洗后的新文件名主体": row.get("清洗后的新文件名主体", ""),
            "是否保留扩展名": row.get("是否保留扩展名", True),
            "是否实际改名": row.get("是否实际改名", False),
            "处理说明": row.get("处理说明", ""),
            "状态": row.get("状态", ""),
            "original_prefix_number": prefix_number(row.get("原文件名", "")),
            "final_prefix_number": prefix_number(row.get("输出文件名", "")),
        })
    out = _compatible(pd.concat([history, pd.DataFrame(rows)], ignore_index=True))
    out.to_excel(path, index=False)
    return out


def clear_history(path):
    path = Path(path)
    if path.exists():
        path.unlink()


def suggest_from_history(file_inventory, history_df):
    suggestions = {}
    if history_df.empty:
        return suggestions
    by_original = {str(row["原文件名"]): row for _, row in history_df.iterrows()}
    by_prefix = {}
    for _, row in history_df.iterrows():
        number = str(row.get("original_prefix_number", "") or "")
        if number and number not in by_prefix:
            by_prefix[number] = row
    for item in file_inventory:
        row = by_original.get(item["original_name"])
        if row is None:
            row = by_prefix.get(prefix_number(item["original_name"]))
        if row is not None:
            suggestions[item["original_name"]] = str(row.get("清洗后的新文件名主体", "") or "")
    return suggestions
