from pathlib import Path

import pandas as pd

from .config import REPORT_COLUMNS, RULE_REPORT_COLUMNS
from .utils import ensure_dir


def _with_columns(df, columns):
    out = df.copy()
    for column in columns:
        if column not in out.columns:
            out[column] = ""
    return out[columns]


def write_rename_report(plan_df, output_path):
    output_path = Path(output_path)
    ensure_dir(output_path.parent)
    _with_columns(plan_df, REPORT_COLUMNS).to_excel(output_path, index=False)
    return output_path


def write_rule_report(plan_df, output_path):
    output_path = Path(output_path)
    ensure_dir(output_path.parent)
    _with_columns(plan_df, RULE_REPORT_COLUMNS).to_excel(output_path, index=False)
    return output_path
