from pathlib import Path

import pandas as pd

from .ai_features import build_ai_config, test_ai_connection
from .app_core import UniversalRenameCore
from .rule_dictionary import execute_rule_rename


def _format_inventory(inventory):
    if not inventory:
        return pd.DataFrame(columns=["original_name", "extension", "size_bytes"])
    return pd.DataFrame(inventory)[["original_name", "extension", "size_bytes"]]


def launch_local_app():
    try:
        import gradio as gr
    except Exception as exc:
        raise RuntimeError("本地 UI 需要安装 gradio：pip install gradio") from exc

    core = UniversalRenameCore()
    state = {"zip_path": None, "rule_zip_path": None}

    def upload_files(files):
        core.new_task()
        paths = [file.name for file in files or []]
        if not paths:
            return "未选择文件。", _format_inventory([])
        inventory = core.add_uploaded_files(paths)
        return f"已上传 {len(inventory)} 个文件。", _format_inventory(inventory)

    def apply_paste(text):
        if not core.file_inventory:
            return "请先上传文件。"
        result = core.apply_paste_list(text)
        return f"已更新 {result['updated']} 行，跳过空行 {result['skipped_empty']} 行，多余行 {result['extra']} 行。"

    def validate():
        if not core.file_inventory:
            return "请先上传文件。", pd.DataFrame()
        result = core.validate()
        if result["ok"]:
            preview = result["plan_df"].drop(columns=["source_path"], errors="ignore")
            return "校验通过，可以执行改名。", preview
        return "校验失败，请修改后重试。", result["errors"]

    def execute():
        if not core.file_inventory:
            return "请先上传文件。", None
        result = core.execute()
        if not result.get("ok"):
            return result.get("error", "执行失败。"), None
        state["zip_path"] = str(result["zip_path"])
        return f"执行完成：{result['zip_path'].name}", str(result["zip_path"])

    def test_ai(enabled, api_key, base_url, model):
        config = build_ai_config(enabled=enabled, api_key=api_key, base_url=base_url, model=model)
        result = test_ai_connection(config)
        return "连接成功。" if result["ok"] else result["error"]

    def process_rules(files):
        paths = [file.name for file in files or []]
        if not paths:
            return "未选择规则模块文件。", None
        result = execute_rule_rename(paths, core.workspace / "rule_outputs", core.export_dir)
        if not result.get("ok"):
            return result.get("error", "规则处理失败。"), None
        state["rule_zip_path"] = str(result["zip_path"])
        return f"规则处理完成：{result['zip_path'].name}", str(result["zip_path"])

    with gr.Blocks(title="Universal Rename Tool") as demo:
        gr.Markdown("# Universal Rename Tool\n本地版核心流程。只处理文件名，不读取文件内容，不修改原始上传文件。")
        with gr.Row():
            enable_ai = gr.Checkbox(label="启用 DSV4Flash AI 辅助模式", value=False)
            model = gr.Textbox(label="模型名称", value="dsv4flash")
        api_key = gr.Textbox(label="API Key", type="password")
        base_url = gr.Textbox(label="Base URL", value="https://api.deepseek.com/v1")
        ai_status = gr.Textbox(label="AI 状态", interactive=False)
        btn_test_ai = gr.Button("测试连接")
        btn_test_ai.click(test_ai, [enable_ai, api_key, base_url, model], ai_status)

        files = gr.File(label="上传待改名文件", file_count="multiple")
        upload_status = gr.Textbox(label="上传状态", interactive=False)
        inventory_table = gr.Dataframe(label="文件索引预览", interactive=False)
        files.upload(upload_files, files, [upload_status, inventory_table])

        paste_text = gr.Textbox(label="批量粘贴新文件名列表", lines=10)
        paste_status = gr.Textbox(label="粘贴状态", interactive=False)
        gr.Button("应用粘贴的新文件名列表").click(apply_paste, paste_text, paste_status)

        validate_status = gr.Textbox(label="校验状态", interactive=False)
        plan_table = gr.Dataframe(label="校验/计划预览", interactive=False)
        gr.Button("校验当前命名").click(validate, None, [validate_status, plan_table])

        execute_status = gr.Textbox(label="执行状态", interactive=False)
        zip_file = gr.File(label="下载 ZIP")
        gr.Button("执行批量改名").click(execute, None, [execute_status, zip_file])

        with gr.Accordion("规则字典自动改名模块", open=False):
            rule_files = gr.File(label="上传规则模块文件", file_count="multiple")
            rule_status = gr.Textbox(label="规则状态", interactive=False)
            rule_zip = gr.File(label="下载规则 ZIP")
            gr.Button("上传并处理规则字典模块").click(process_rules, rule_files, [rule_status, rule_zip])

    demo.launch()


if __name__ == "__main__":
    launch_local_app()
