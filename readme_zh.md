# Excel 智能助手

一个简洁、强大且可扩展的智能助手，通过网页对话界面来分析和处理Excel文件。

本项目采用模块化架构，将工具逻辑、Agent智能和Web界面分离，以确保代码的清晰性和可维护性。

## 功能特性

- **对话式交互**: 使用自然语言与您的Excel文件进行交互。
- **多步任务规划**: Agent可以理解复杂请求，并将其分解为多个步骤来执行。
- **可扩展的工具箱**: 通过轻松添加新工具来扩展Agent的能力。
- **文件生成**: 从筛选、排序或数据透视表的结果中创建新的Excel文件。
- **图表生成**: 根据您的数据创建图表，并直接在界面中查看。
- **安全设计**: 从不修改原始上传文件；所有写操作都会创建新文件。

## 核心工具

本助手内置了一套强大的数据分析和处理工具：

- **工作簿与工作表管理**
  - `list_sheets`: 列出工作簿中所有工作表的名称。
  - `create_sheet`: 创建一个新的、空白的工作表。
  - `delete_sheet`: 按名称删除一个工作表。
  - `duplicate_sheet`: 创建一个现有工作表的副本。
- **数据摘要与读取**
  - `get_data_summary`: 快速获取工作表的概览，包括表头和总行数。
  - `read_rows`: 从工作表中读取指定范围的行。
- **数据清洗与转换**
  - `delete_columns`: 删除一个或多个列。
  - `rename_column`: 重命名指定的列。
  - `handle_duplicates`: 基于指定列查找或移除重复行。
  - `fill_missing_values`: 用指定值填充列中的空白单元格。
  - `string_manipulation_in_column`: 对指定列进行字符串操作（如 `uppercase`, `lowercase`, `trim`）。
  - `conditional_value_column`: 基于一个现有列的 if/else 条件创建新列。
  - `lookup_and_merge_columns`: 执行类似 VLOOKUP 的操作，以合并另一个文件/工作表中的列。
- **分析与重塑**
  - `column_aggregate`: 对列执行聚合计算（如 `sum`, `mean`, `min`, `max`）。
  - `get_unique_values`: 获取某列的所有唯一值。
  - `filter_rows`: 根据条件筛选数据（例如 `Price > 100`）。
  - `sort_data`: 按指定列对整个工作表进行排序。
  - `add_column_from_formula`: 基于两个现有列之间的简单公式创建一个新列。
  - `create_pivot_table`: 生成数据透视表以汇总数据。
  - `group_by_and_aggregate`: 按列分组，并对另一列执行聚合运算（sum, mean 等）。
- **格式化**
  - `apply_conditional_formatting`: 对列应用条件格式（例如，高亮显示 > 100 的单元格）。
- **可视化**
  - `create_chart`: 从数据创建 `bar`（柱状图）、`line`（折线图）或 `scatter`（散点图）。

## 安装与运行

请遵循以下步骤来运行本应用：

1.  **安装依赖**: 推荐使用虚拟环境。从`requirements.txt`文件中安装所有必需的包。
    ```bash
    pip install -r requirements.txt
    ```

2.  **配置API密钥**: 打开`config.py`文件，输入您的LLM API凭据。它已预先配置为火山引擎方舟平台，但可以适配任何兼容OpenAI的API。

3.  **运行Web服务器**: 启动Flask web服务器。
    ```bash
    python web_server.py
    ```

4.  **访问UI**: 打开您的网络浏览器，访问 `http://127.0.0.1:5001`。

5.  **开始处理**: 在界面上上传一个Excel文件以开始使用。

---

## 如何添加新工具

本项目的威力源于其工具箱。新的架构设计使得添加工具变得简单而安全，无需触及Agent的核心逻辑。

### 第一步：在 `processor.py` 中实现工具逻辑

打开`processor.py`文件，在末尾添加一个新的、独立的Python函数。这个函数就是你的新工具的实现。

**一个优秀工具函数的要求：**

- 它应该执行一个单一、清晰的任务。
- 如果它需要读取文件，应该接受 `file_path` 参数。
- 如果它需要写入文件，应该接受 `file_output_dir` 和 `output_filename` 参数。
- 所有其他参数都应该有明确的类型提示。

**示例：**
```python
# 在 processor.py 的末尾添加
from typing import Dict, Any

def my_new_tool(file_path: str, some_parameter: str) -> Dict[str, Any]:
    """一个清晰的文档字符串，解释这个工具做什么。"""
    # 在这里实现你的逻辑
    print(f"使用文件 {file_path} 和参数 {some_parameter} 执行 my_new_tool")
    return {"success": True, "message": "新工具成功执行！"}
```

### 第二步：在 `tools.py` 中定义并注册工具

打开`tools.py`文件。这是所有可用工具的唯一真实来源。在`TOOLS`列表中添加一个新字典。

这个字典有两个键：
- `schema`: 一个JSON schema，用于向LLM描述你的工具。这对Agent理解如何使用你的工具至关重要。
- `implementation`: 一个指向你在`processor.py`中创建的函数的引用。

**示例：**

```python
# 在 tools.py 中，将下面的字典追加到 TOOLS 列表
{
    "schema": {
        "name": "my_new_tool",
        "description": "一句话清晰描述这个工具的功能。",
        "parameters": {
            "type": "object",
            "properties": {
                "some_parameter": {
                    "type": "string",
                    "description": "关于这个参数的描述。",
                    "enum": ["option1", "option2"] # 可选：如果参数有固定选项
                }
            },
            "required": ["some_parameter"] # 列出所有必需的参数
        }
    },
    "implementation": processor.my_new_tool
}
```

完成了。一旦你重启`web_server.py`，Agent就会自动感知到你的新工具，并能立刻开始使用它。旧的修改`agent.py`的方式已不再需要。