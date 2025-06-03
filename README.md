# Markdown Batch Processing / Markdown 批量处理

**Only tested on win10, there are many bugs in the function, it is not yet finished, please back up before use**

**只在win10上测试通过，功能存在许多bug，还没做完，使用前先一定备份**

A powerful tool for processing and checking YAML frontmatter in Markdown files, with support for multi-language interfaces and advanced configuration options.

一款功能强大的工具，用于处理和检查 Markdown 文件中的 YAML frontmatter，支持多语言界面和高级配置选项。

## A Frontmatter Processing & Checking Tool / Frontmatter 处理与检查工具

This tool allows users to analyze, modify, and standardize frontmatter in multiple Markdown files efficiently. It provides a user-friendly PyQt6-based GUI, supports command-line operations, and generates detailed Excel reports for analysis.

此工具允许用户高效地分析、修改和标准化多个 Markdown 文件中的 frontmatter。它提供基于 PyQt6 的用户友好 GUI，支持命令行操作，并生成详细的 Excel 分析报告。

### Features / 功能

- **Parse and Serialize YAML Frontmatter**: Extract and update frontmatter with robust YAML parsing.
- **Type Detection and Conflict Analysis**: Automatically detect field types (e.g., string, list, date) and identify type conflicts, with enhanced string-to-list conversion.
- **Data Type Conversion**: Convert field values to specified types (e.g., string to list, date to datetime).
- **Field Merging and Default Values**: Merge multiple fields into one and apply default values for missing fields.
- **Interactive Change Confirmation**: Display changes in a dialog for user confirmation before applying.
- **GUI Results Display**: Show field types and conflicts in a tabular format within the GUI.
- **Excel Report Generation**: Generate detailed reports with valid files, type conflicts, and field statistics, saved to a user-specified output directory.
- **PyQt6 GUI and CLI Support**: Operate via an intuitive GUI or command-line for batch processing.
- **Configuration Persistence**: Load and save settings (language, field types, merge rules) in YAML config files.
- **Multi-Language Support**: Interface available in UN official languages (English, Chinese, French, Spanish, Arabic, Russian) with dynamic switching.

- **解析和序列化 YAML Frontmatter**：使用强大的 YAML 解析提取和更新 frontmatter。
- **类型检测与冲突分析**：自动检测字段类型（例如字符串、列表、日期）并识别类型冲突，增强字符串到列表的转换。
- **数据类型转换**：将字段值转换为指定类型（例如字符串到列表，日期到日期时间）。
- **字段合并与默认值**：将多个字段合并为一个，并为缺失字段应用默认值。
- **交互式变更确认**：在应用更改前通过对话框显示变更供用户确认。
- **GUI 结果显示**：在 GUI 中以表格格式显示字段类型和冲突。
- **Excel 报告生成**：生成详细报告，包括有效文件、类型冲突和字段统计，保存到用户指定的输出目录。
- **PyQt6 GUI 和 CLI 支持**：通过直观的 GUI 或命令行进行批量操作。
- **配置持久化**：在 YAML 配置文件中加载和保存设置（语言、字段类型、合并规则）。
- **多语言支持**：界面支持联合国官方语言（英语、中文、法语、西班牙语、阿拉伯语、俄语），并支持动态切换。

### Download / 下载

Download the latest version of the tool:

下载工具的最新版本：

- [Markdown Frontmatter Editor - v1.4](https://raw.githubusercontent.com/strangeZombies/markdown-batch-processing/refs/heads/main/xds_md_frontmatter_tool_gui_v2.py)

### Installation / 安装

1. **Prerequisites / 前提条件**:
   - Python 3.8+ / Python 3.8 或更高版本
   - Install dependencies / 安装依赖：
     ```bash
     pip install PyQt6 pyyaml pandas xlsxwriter
     ```

2. **Download the Script / 下载脚本**:
   - Download `xds_md_frontmatter_tool_gui_v2.py` from the link above / 从上方链接下载 `xds_md_frontmatter_tool_gui_v2.py`。

3. **Run the Tool / 运行工具**:
   - GUI mode / GUI 模式：
     ```bash
     python xds_md_frontmatter_tool_gui_v2.py
     ```
   - Command-line mode (batch processing) / 命令行模式（批量处理）：
     ```bash
     python xds_md_frontmatter_tool_gui_v2.py zh /path/to/input /path/to/output --batch
     ```
     Replace `zh` with desired language code (e.g., `en`, `fr`, `es`, `ar`, `ru`) / 将 `zh` 替换为所需的语言代码（例如 `en`、`fr`、`es`、`ar`、`ru`）。

### Usage / 使用方法

1. **GUI Mode / GUI 模式**:
   - Launch the tool and select the input directory containing Markdown files / 启动工具并选择包含 Markdown 文件的输入目录。
   - Optionally specify an output directory or enable overwrite mode / 可选择指定输出目录或启用覆盖模式。
   - Configure field types, merge rules, and default values in the respective tabs / 在相应选项卡中配置字段类型、合并规则和默认值。
   - Click "Analyze" to detect types and conflicts, or "Process" to apply changes / 点击“分析”检测类型和冲突，或点击“处理”应用更改。
   - View the generated Excel report by clicking "View Report" / 点击“查看报告”查看生成的 Excel 报告。

2. **Command-Line Mode / 命令行模式**:
   - Specify language, input directory, output directory, and optional config file / 指定语言、输入目录、输出目录和可选配置文件。
   - Example / 示例：
     ```bash
     python xds_md_frontmatter_tool_gui_v2.py en /path/to/input /path/to/output /path/to/config.yaml --batch
     ```

### Configuration / 配置

The tool supports a YAML configuration file to persist settings. Example:

该工具支持 YAML 配置文件以持久化设置。示例：

```yaml
language: zh
list_separators: [",", ";", "|"]
field_types:
  tags: list
  title: str
merge_rules:
  tags: [keywords, categories]
default_values:
  author:
    type: str
    value: "Anonymous"
```

- Save the config file (e.g., `frontmatter_config.yaml`) and load it via the GUI or command-line / 保存配置文件（例如 `frontmatter_config.yaml`）并通过 GUI 或命令行加载。

### Contributing / 贡献

Contributions are welcome! Please submit issues or pull requests to the [GitHub repository](https://github.com/strangeZombies/markdown-batch-processing).

欢迎贡献！请在 [GitHub 仓库](https://github.com/strangeZombies/markdown-batch-processing) 提交问题或拉取请求。
