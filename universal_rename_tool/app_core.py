import shutil
from pathlib import Path

from .config import DEFAULT_WORKSPACE
from .executor import execute_rename
from .file_inventory import build_file_inventory
from .history import append_history
from .rename_state import apply_paste_name_list, create_initial_rename_state
from .utils import ensure_dir, safe_remove
from .validation import build_rename_plan, validate_rename_plan


class UniversalRenameCore:
    def __init__(self, workspace=DEFAULT_WORKSPACE):
        self.workspace = Path(workspace)
        self.upload_dir = self.workspace / "uploads"
        self.output_dir = self.workspace / "outputs"
        self.export_dir = self.workspace / "exports"
        self.history_path = self.workspace / "history" / "rename_history.xlsx"
        self.file_inventory = []
        self.rename_state = {}
        self.latest_plan_df = None
        self.reset_workspace()

    def reset_workspace(self):
        ensure_dir(self.workspace)
        for path in [self.upload_dir, self.output_dir, self.export_dir, self.history_path.parent]:
            ensure_dir(path)

    def new_task(self):
        safe_remove(self.upload_dir)
        safe_remove(self.output_dir)
        ensure_dir(self.upload_dir)
        ensure_dir(self.output_dir)
        self.file_inventory = []
        self.rename_state = {}
        self.latest_plan_df = None

    def add_uploaded_files(self, file_paths):
        ensure_dir(self.upload_dir)
        stored = []
        for raw in file_paths:
            src = Path(raw)
            dst = self.upload_dir / src.name
            if src.resolve() != dst.resolve():
                shutil.copy2(src, dst)
            stored.append(dst)
        self.file_inventory = build_file_inventory(stored)
        self.rename_state = create_initial_rename_state(self.file_inventory)
        return self.file_inventory

    def apply_paste_list(self, text):
        return apply_paste_name_list(self.file_inventory, self.rename_state, text)

    def validate(self):
        plan = build_rename_plan(self.file_inventory, self.rename_state)
        result = validate_rename_plan(plan)
        self.latest_plan_df = result["plan_df"] if result["ok"] else None
        return result

    def execute(self):
        result = self.validate()
        if not result["ok"]:
            return result
        executed = execute_rename(result["plan_df"], self.output_dir, self.export_dir)
        if executed.get("ok"):
            append_history(self.history_path, executed["plan_df"])
        return executed
