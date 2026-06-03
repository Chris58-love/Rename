import shutil
from pathlib import Path

from .reports import write_rename_report
from .utils import ensure_dir, timestamp
from .validation import validate_rename_plan
from .zip_utils import make_zip


def execute_rename(plan_df, output_dir, zip_dir):
    result = validate_rename_plan(plan_df)
    if not result["ok"]:
        return {"ok": False, "error": "校验未通过", "errors": result["errors"]}
    output_dir = ensure_dir(output_dir)
    zip_dir = ensure_dir(zip_dir)
    copied = []
    executed = plan_df.copy()
    for index, row in executed.iterrows():
        src = Path(row["source_path"])
        if not src.exists():
            return {"ok": False, "error": f"源文件缺失：{row['原文件名']}"}
        dst = output_dir / row["输出文件名"]
        shutil.copy2(src, dst)
        copied.append(dst)
        executed.loc[index, "状态"] = "成功"
    report_path = write_rename_report(executed, output_dir / "rename_report.xlsx")
    copied.append(report_path)
    zip_path = make_zip(copied, zip_dir / f"renamed_files_{timestamp()}.zip")
    return {"ok": True, "plan_df": executed, "report_path": report_path, "zip_path": zip_path}
