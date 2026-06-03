# Colab Universal Rename Tool

通用批量改名工具，支持 Google Colab 交互界面和本地 Python Web UI。工具只处理文件名，不读取文件内容，不修改原始上传文件；执行时会复制到输出目录并生成报告与 ZIP。

## 功能列表

- 任意文件类型上传，支持无扩展名、中文、英文、数字和符号混合文件名。
- `file_inventory` 源文件索引和 `rename_state` 编辑状态。
- Colab 在线表格逐行编辑。
- 批量粘贴一列新文件名。
- Excel 模板导出与上传覆盖当前表格。
- 本地硬规则校验，拦截空名、路径不安全和最终文件名冲突。
- 执行批量改名、生成 `rename_report.xlsx`、输出 ZIP。
- 历史记录保存、回顾、导出、清空和历史推荐。
- 规则字典自动改名模块。
- 可选 DSV4Flash / OpenAI 兼容接口：AI 建议名、AI 审核、当前指引、新用户问答。
- 全流程自检面板。

## Colab 使用方法

在 Colab 中打开 `run_colab.ipynb` 或 `notebooks/colab_launcher.ipynb`，运行安装和启动单元：

```python
!pip install -r requirements.txt

from universal_rename_tool.ui_colab import launch_colab_app
launch_colab_app()
```

如果从 GitHub 使用：

```python
!git clone https://github.com/your-name/colab-universal-rename-tool.git
%cd colab-universal-rename-tool
!pip install -r requirements.txt

from universal_rename_tool.ui_colab import launch_colab_app
launch_colab_app()
```

Colab 版保留完整 ipywidgets 交互、`google.colab.files` 上传下载、在线表格、历史、Excel 模板、规则字典、AI 和自检功能。

## 本地运行方法

```bash
pip install -r requirements.txt
python app.py
```

本地环境会自动启动 Gradio Web UI，不依赖 `google.colab`。当前本地 UI 覆盖核心流程：上传文件、批量粘贴、校验、执行、下载 ZIP、AI 配置测试和规则字典处理；项目结构已为完整功能继续扩展预留模块入口。

## AI 配置说明

AI 是可选功能。未启用 AI 时，上传、编辑、校验、执行、历史、Excel 模板和规则字典模块都可正常运行。

启用 DSV4Flash AI 辅助模式后，需要填写：

- API Key：密码框，只保存在当前运行时内存或页面控件中。
- Base URL：默认 `https://api.deepseek.com/v1`，可替换为其他 OpenAI 兼容接口地址。
- 模型名称：默认 `dsv4flash`。
- 测试连接：只发送最小 JSON 测试请求，不发送文件内容。

安全规则：

- 不保存 API Key 到文件。
- 不把 API Key 写入日志、报告、历史记录或 ZIP。
- 不向 AI 发送文件内容，只发送文件名和用户输入的改名上下文。
- AI 建议和 AI 审核不自动执行改名，也不能绕过本地校验。

## 批量粘贴说明

可从 Excel/WPS/飞书复制一列新文件名并粘贴。系统按当前文件顺序覆盖 `new_stem`，空行跳过，多余行忽略；如果粘贴值带有与原文件相同的扩展名，会自动去掉，避免 `.xlsx.xlsx`。

## Excel 模板模式

导出模板后可在表格软件中编辑，再上传覆盖当前在线表格。模板只包含文件名和改名字段，不包含文件内容。

## 历史推荐

执行成功后会追加历史记录。历史推荐会优先匹配原文件名，也可按开头编号进行推荐；旧版本历史字段会自动补齐。

## 规则字典自动改名

规则字典模块独立于主流程，可单独上传任意文件，按文件名主体开头编号匹配规则。命中规则则替换为标准主体，未命中则清洗原主体；冲突会被拦截。

## 输出文件说明

- `renamed_files_*.zip`：改名后的复制文件和报告。
- `rename_report.xlsx`：主流程改名报告。
- `rule_renamed_files_*.zip`：规则字典模块输出。
- `rule_rename_report.xlsx`：规则字典模块报告。
- `rename_history.xlsx`：历史记录。

报告、历史记录和 ZIP 不包含绝对路径、上传目录路径、API Key 或文件内容。

## 常见问题

**启用 AI 后看不到配置框？**  
Colab 版的 AI 设置区已挂载在“AI 辅助模式选择”主区域下方。选择 DSV4Flash 模式后会立即显示 API Key、Base URL、模型名称和测试连接按钮。

**执行按钮为什么不可用？**  
表格修改后必须重新校验。只有校验通过且当前状态未修改时，执行按钮才会启用；执行前也会再次校验。

**会修改原文件吗？**  
不会。工具只复制文件到输出目录，并对复制件使用新文件名。

**会读取文件内容吗？**  
不会。工具只处理文件名和文件大小等元数据。
