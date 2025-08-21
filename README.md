# Excel Agent

A simple, powerful, and extensible agent for analyzing and processing Excel files through a web-based conversational interface.

This project is built with a modular architecture, separating the tool logic, agent intelligence, and web interface for clarity and maintainability.

## Features

- **Conversational Interface**: Interact with your Excel files using natural language.
- **Multi-Step Task Planning**: The agent can reason about complex requests and break them down into multiple steps.
- **Extensible Toolbox**: Easily add new tools to expand the agent's capabilities.
- **File Generation**: Create new Excel files from filtered data, sorted data, or pivot tables.
- **Chart Generation**: Create charts from your data and view them directly in the UI.
- **Safe by Design**: Never modifies the original uploaded file; all write operations create new files.

## Core Tools

The agent comes with a powerful set of built-in tools for data analysis and manipulation:

- **Workbook & Sheet Management**
  - `list_sheets`: List the names of all sheets in the workbook.
  - `create_sheet`: Create a new, empty sheet.
  - `delete_sheet`: Delete a sheet by name.
  - `duplicate_sheet`: Create a copy of an existing sheet.
- **Data Summary & Reading**
  - `get_data_summary`: Get a quick overview of a sheet, including headers and row count.
  - `read_rows`: Read a specific range of rows from a sheet.
- **Data Cleaning & Transformation**
  - `delete_columns`: Remove one or more columns.
  - `rename_column`: Rename a specific column.
  - `handle_duplicates`: Find or remove duplicate rows based on specific columns.
  - `fill_missing_values`: Fill empty cells in a column with a specific value.
  - `string_manipulation_in_column`: Apply string operations like `uppercase`, `lowercase`, or `trim` to a column.
  - `conditional_value_column`: Create a new column with values based on an if/else condition on another column.
  - `lookup_and_merge_columns`: Perform a VLOOKUP-like operation to merge columns from a second file/sheet.
- **Analysis & Reshaping**
  - `column_aggregate`: Perform aggregations like `sum`, `mean`, `min`, `max` on a column.
  - `get_unique_values`: Get all unique values from a column.
  - `filter_rows`: Filter data based on a condition (e.g., `Price > 100`).
  - `sort_data`: Sort an entire sheet by a specific column.
  - `add_column_from_formula`: Create a new column based on a simple formula between two existing columns.
  - `create_pivot_table`: Generate a pivot table to summarize data.
  - `group_by_and_aggregate`: Group data by a column and perform aggregations (sum, mean, etc.) on another column.
- **Formatting**
  - `apply_conditional_formatting`: Apply conditional formatting (e.g., highlight cells > 100) to a column.
- **Visualization**
  - `create_chart`: Create `bar`, `line`, or `scatter` charts from your data.

## Setup and Running

Follow these steps to get the application running:

1.  **Install Dependencies**: It's recommended to use a virtual environment. Install all required packages from `requirements.txt`.
    ```bash
    pip install -r requirements.txt
    ```

2.  **Configure API Key**: Open the `config.py` file and enter your LLM API credentials. It is pre-configured for VolcEngine Ark, but can be adapted for any OpenAI-compatible API.

3.  **Run the Web Server**: Start the Flask web server.
    ```bash
    python web_server.py
    ```

4.  **Access the UI**: Open your web browser and navigate to `http://127.0.0.1:5001`.

5.  **Start Processing**: Upload an Excel file using the UI to begin.

---

## How to Add New Tools

The power of this agent comes from its toolbox. The architecture is designed to make adding new tools simple and safe, without needing to touch the core agent logic.

### Step 1: Add the Tool's Logic in `processor.py`

Open the `processor.py` file and add a new, self-contained Python function. This function is your new tool's implementation.

**Requirements for a good tool function:**

- It should perform a single, clear task.
- It should accept `file_path` if it reads a file.
- If the tool writes a file, it should accept `file_output_dir` and `output_filename`.
- All other parameters should be clearly typed.

**Example:**
```python
# In processor.py
from typing import Dict, Any

def my_new_tool(file_path: str, some_parameter: str) -> Dict[str, Any]:
    """A clear docstring explaining what the tool does."""
    # Your logic here
    print(f"Executing my_new_tool with file {file_path} and parameter {some_parameter}")
    return {"success": True, "message": "New tool executed successfully!"}
```

### Step 2: Define and Register the Tool in `tools.py`

Open `tools.py`. This file is the single source of truth for all available tools. Add a new dictionary entry to the `TOOLS` list.

This dictionary has two keys:
- `schema`: A JSON schema that describes the tool to the LLM. This is critical for the agent to understand how to use your tool.
- `implementation`: A reference to the function you just created in `processor.py`.

**Example:**

```python
# In tools.py, append to the TOOLS list:
{
    "schema": {
        "name": "my_new_tool",
        "description": "A clear, one-sentence description of what this tool does.",
        "parameters": {
            "type": "object",
            "properties": {
                "some_parameter": {
                    "type": "string",
                    "description": "A description of this parameter.",
                    "enum": ["option1", "option2"] # Optional: if the parameter has fixed choices
                }
            },
            "required": ["some_parameter"] # List all required parameters
        }
    },
    "implementation": processor.my_new_tool
}
```

That's it. Once you restart `web_server.py`, the agent will automatically be aware of your new tool and can start using it immediately. The old method of modifying `agent.py` is no longer necessary.