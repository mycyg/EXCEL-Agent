"""
This module is the single source of truth for all tools available to the agent.
It defines the schema for each tool and maps it to its implementation in the processor module.
This approach avoids complex runtime introspection and keeps tool definitions clear and maintainable.
"""

import processor

# A list of all available tools. Each tool is a dictionary containing:
# - schema: A JSON-serializable dictionary describing the tool for the LLM.
# - implementation: The actual Python function to call.

TOOLS = [
    {
        "schema": {
            "name": "get_data_summary",
            "description": "Reads an Excel file and returns a summary of its contents, including sheet name, total rows, and header columns.",
            "parameters": {
                "type": "object",
                "properties": {
                    "sheet_name": {
                        "type": "string",
                        "description": "The name of the sheet to summarize. Defaults to the active sheet if not provided."
                    }
                },
                "required": []
            }
        },
        "implementation": processor.get_data_summary
    },
    {
        "schema": {
            "name": "read_rows",
            "description": "Reads a specific range of rows from an Excel sheet.",
            "parameters": {
                "type": "object",
                "properties": {
                    "sheet_name": {
                        "type": "string",
                        "description": "The name of the sheet to read from. Defaults to the active sheet."
                    },
                    "offset": {
                        "type": "integer",
                        "description": "The 0-based row number to start reading from. Defaults to 1."
                    },
                    "limit": {
                        "type": "integer",
                        "description": "The maximum number of rows to read. Defaults to 5."
                    }
                },
                "required": []
            }
        },
        "implementation": processor.read_rows
    },
    {
        "schema": {
            "name": "get_unique_values",
            "description": "Gets a list of unique values from a specified column.",
            "parameters": {
                "type": "object",
                "properties": {
                    "column_name": {
                        "type": "string",
                        "description": "The name of the column to get unique values from."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "The name of the sheet to use. Defaults to the active sheet."
                    }
                },
                "required": ["column_name"]
            }
        },
        "implementation": processor.get_unique_values
    },
    {
        "schema": {
            "name": "column_aggregate",
            "description": "Performs a simple aggregation (sum, mean, min, max) on a single column.",
            "parameters": {
                "type": "object",
                "properties": {
                    "column_name": {
                        "type": "string",
                        "description": "The name of the column to aggregate."
                    },
                    "aggregate_function": {
                        "type": "string",
                        "description": "The aggregation function to apply.",
                        "enum": ["sum", "mean", "min", "max"]
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "The name of the sheet to use. Defaults to the active sheet."
                    }
                },
                "required": ["column_name", "aggregate_function"]
            }
        },
        "implementation": processor.column_aggregate
    },
    {
        "schema": {
            "name": "filter_rows",
            "description": "Filters rows based on a condition and saves the result to a new file.",
            "parameters": {
                "type": "object",
                "properties": {
                    "output_filename": {
                        "type": "string",
                        "description": "The name for the new output Excel file (e.g., 'filtered_data.xlsx')."
                    },
                    "column_name": {
                        "type": "string",
                        "description": "The column to filter on."
                    },
                    "operator": {
                        "type": "string",
                        "description": "The comparison operator.",
                        "enum": ["==", "!=", ">", "<", ">=", "<=", "contains"]
                    },
                    "value": {
                        "type": "string",
                        "description": "The value to compare against. For numeric comparisons, provide a number as a string."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "The name of the sheet to use. Defaults to the active sheet."
                    }
                },
                "required": ["output_filename", "column_name", "operator", "value"]
            }
        },
        "implementation": processor.filter_rows
    },
    {
        "schema": {
            "name": "sort_data",
            "description": "Sorts the data by a specific column and saves to a new file.",
            "parameters": {
                "type": "object",
                "properties": {
                    "output_filename": {
                        "type": "string",
                        "description": "The name for the new output Excel file (e.g., 'sorted_data.xlsx')."
                    },
                    "sort_by_column": {
                        "type": "string",
                        "description": "The column to sort the data by."
                    },
                    "ascending": {
                        "type": "boolean",
                        "description": "Whether to sort in ascending (True) or descending (False) order. Defaults to True."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "The name of the sheet to use. Defaults to the active sheet."
                    }
                },
                "required": ["output_filename", "sort_by_column"]
            }
        },
        "implementation": processor.sort_data
    },
    {
        "schema": {
            "name": "add_column_from_formula",
            "description": "Adds a new column based on a formula involving two existing columns and saves to a new file.",
            "parameters": {
                "type": "object",
                "properties": {
                    "output_filename": {
                        "type": "string",
                        "description": "The name for the new output Excel file (e.g., 'calculated_data.xlsx')."
                    },
                    "new_column_name": {
                        "type": "string",
                        "description": "The name for the new column being created."
                    },
                    "col1": {
                        "type": "string",
                        "description": "The name of the first column in the formula."
                    },
                    "operator": {
                        "type": "string",
                        "description": "The mathematical operator to use.",
                        "enum": ["+", "-", "*", "/"]
                    },
                    "col2": {
                        "type": "string",
                        "description": "The name of the second column in the formula."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "The name of the sheet to use. Defaults to the active sheet."
                    }
                },
                "required": ["output_filename", "new_column_name", "col1", "operator", "col2"]
            }
        },
        "implementation": processor.add_column_from_formula
    },
    {
        "schema": {
            "name": "create_pivot_table",
            "description": "Creates a pivot table and saves it to a new file.",
            "parameters": {
                "type": "object",
                "properties": {
                    "output_filename": {
                        "type": "string",
                        "description": "The name for the new output Excel file (e.g., 'pivot_table.xlsx')."
                    },
                    "index_column": {
                        "type": "string",
                        "description": "The column to use as the pivot table's index (rows)."
                    },
                    "columns_column": {
                        "type": "string",
                        "description": "The column to use for the pivot table's columns."
                    },
                    "values_column": {
                        "type": "string",
                        "description": "The column to aggregate."
                    },
                    "agg_func": {
                        "type": "string",
                        "description": "The aggregation function to use.",
                        "enum": ["sum", "mean", "count"]
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "The name of the sheet to use. Defaults to the active sheet."
                    }
                },
                "required": ["output_filename", "index_column", "columns_column", "values_column"]
            }
        },
        "implementation": processor.create_pivot_table
    },
    {
        "schema": {
            "name": "create_chart",
            "description": "Creates a chart from data and saves it as an image file.",
            "parameters": {
                "type": "object",
                "properties": {
                    "chart_type": {
                        "type": "string",
                        "description": "The type of chart to create.",
                        "enum": ["bar", "line", "scatter"]
                    },
                    "x_column": {
                        "type": "string",
                        "description": "The column to use for the x-axis."
                    },
                    "y_column": {
                        "type": "string",
                        "description": "The column to use for the y-axis."
                    },
                    "output_filename": {
                        "type": "string",
                        "description": "The name for the output image file (e.g., 'my_chart.png'). If not provided, a name will be generated."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "The name of the sheet to use. Defaults to the active sheet."
                    }
                },
                "required": ["chart_type", "x_column", "y_column"]
            }
        },
        "implementation": processor.create_chart
    },
    {
        "schema": {
            "name": "delete_columns",
            "description": "Deletes one or more specified columns from a sheet and saves the result to a new file.",
            "parameters": {
                "type": "object",
                "properties": {
                    "output_filename": {
                        "type": "string",
                        "description": "The name for the new output Excel file (e.g., 'data_with_deleted_columns.xlsx')."
                    },
                    "columns_to_delete": {
                        "type": "array",
                        "items": {
                            "type": "string"
                        },
                        "description": "A list of column names to be deleted."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "The name of the sheet to use. Defaults to the active sheet."
                    }
                },
                "required": ["output_filename", "columns_to_delete"]
            }
        },
        "implementation": processor.delete_columns
    },
    {
        "schema": {
            "name": "rename_column",
            "description": "Renames a single existing column to a new name.",
            "parameters": {
                "type": "object",
                "properties": {
                    "output_filename": {
                        "type": "string",
                        "description": "The name for the new output Excel file (e.g., 'data_renamed.xlsx')."
                    },
                    "old_column_name": {
                        "type": "string",
                        "description": "The current name of the column you want to rename."
                    },
                    "new_column_name": {
                        "type": "string",
                        "description": "The new name for the column."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "The name of the sheet to use. Defaults to the active sheet."
                    }
                },
                "required": ["output_filename", "old_column_name", "new_column_name"]
            }
        },
        "implementation": processor.rename_column
    },
    {
        "schema": {
            "name": "handle_duplicates",
            "description": "Finds or removes duplicate rows based on specified columns.",
            "parameters": {
                "type": "object",
                "properties": {
                    "output_filename": {
                        "type": "string",
                        "description": "The name for the new output Excel file (e.g., 'data_no_duplicates.xlsx'). Required only for 'remove' action."
                    },
                    "subset_columns": {
                        "type": "array",
                        "items": {
                            "type": "string"
                        },
                        "description": "A list of column names to check for duplicates. If empty, all columns are used."
                    },
                    "action": {
                        "type": "string",
                        "description": "The action to perform: 'find' to report duplicates, 'remove' to delete them.",
                        "enum": ["find", "remove"]
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "The name of the sheet to use. Defaults to the active sheet."
                    }
                },
                "required": ["action"]
            }
        },
        "implementation": processor.handle_duplicates
    },
    {
        "schema": {
            "name": "fill_missing_values",
            "description": "Fills missing values (NaN) in a specified column with a given value.",
            "parameters": {
                "type": "object",
                "properties": {
                    "output_filename": {
                        "type": "string",
                        "description": "The name for the new output Excel file (e.g., 'data_filled.xlsx')."
                    },
                    "column_name": {
                        "type": "string",
                        "description": "The name of the column with missing values to fill."
                    },
                    "fill_value": {
                        "type": "string",
                        "description": "The value to use for filling missing cells. It will be converted to numeric if the column is numeric."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "The name of the sheet to use. Defaults to the active sheet."
                    }
                },
                "required": ["output_filename", "column_name", "fill_value"]
            }
        },
        "implementation": processor.fill_missing_values
    },
    {
        "schema": {
            "name": "string_manipulation_in_column",
            "description": "Performs a string operation (uppercase, lowercase, trim) on all values in a column.",
            "parameters": {
                "type": "object",
                "properties": {
                    "output_filename": {
                        "type": "string",
                        "description": "The name for the new output Excel file (e.g., 'data_string_op.xlsx')."
                    },
                    "column_name": {
                        "type": "string",
                        "description": "The name of the column to perform the string operation on."
                    },
                    "operation": {
                        "type": "string",
                        "description": "The string operation to perform.",
                        "enum": ["uppercase", "lowercase", "trim"]
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "The name of the sheet to use. Defaults to the active sheet."
                    }
                },
                "required": ["output_filename", "column_name", "operation"]
            }
        },
        "implementation": processor.string_manipulation_in_column
    },
    {
        "schema": {
            "name": "list_sheets",
            "description": "Lists all the names of the sheets in the Excel workbook.",
            "parameters": {
                "type": "object",
                "properties": {},
                "required": []
            }
        },
        "implementation": processor.list_sheets
    },
    {
        "schema": {
            "name": "lookup_and_merge_columns",
            "description": "Merges columns from a secondary file/sheet into the primary file, similar to VLOOKUP.",
            "parameters": {
                "type": "object",
                "properties": {
                    "output_filename": {
                        "type": "string",
                        "description": "The name for the new output Excel file (e.g., 'merged_data.xlsx')."
                    },
                    "left_on_column": {
                        "type": "string",
                        "description": "The key column in the primary (left) file to join on."
                    },
                    "right_file_path": {
                        "type": "string",
                        "description": "The absolute path to the secondary (right) file to merge from."
                    },
                    "right_on_column": {
                        "type": "string",
                        "description": "The key column in the secondary (right) file to join on."
                    },
                    "columns_to_merge": {
                        "type": "array",
                        "items": {
                            "type": "string"
                        },
                        "description": "A list of column names from the secondary file to add to the primary file."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "The sheet in the primary file to use. Defaults to the active sheet."
                    },
                    "right_sheet_name": {
                        "type": "string",
                        "description": "The sheet in the secondary file to use. Defaults to the active sheet."
                    }
                },
                "required": ["output_filename", "left_on_column", "right_file_path", "right_on_column", "columns_to_merge"]
            }
        },
        "implementation": processor.lookup_and_merge_columns
    },
    {
        "schema": {
            "name": "group_by_and_aggregate",
            "description": "Groups data by a column and calculates aggregations (e.g., sum, mean) for another column.",
            "parameters": {
                "type": "object",
                "properties": {
                    "output_filename": {
                        "type": "string",
                        "description": "The name for the new output Excel file (e.g., 'grouped_data.xlsx')."
                    },
                    "group_by_column": {
                        "type": "string",
                        "description": "The column to group the data by."
                    },
                    "agg_column": {
                        "type": "string",
                        "description": "The numeric column to perform the aggregations on."
                    },
                    "agg_functions": {
                        "type": "array",
                        "items": {
                            "type": "string",
                            "enum": ["sum", "mean", "min", "max", "count"]
                        },
                        "description": "A list of aggregation functions to apply."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "The sheet to use. Defaults to the active sheet."
                    }
                },
                "required": ["output_filename", "group_by_column", "agg_column", "agg_functions"]
            }
        },
        "implementation": processor.group_by_and_aggregate
    },
    {
        "schema": {
            "name": "conditional_value_column",
            "description": "Creates a new column with values based on an if/else condition on a source column.",
            "parameters": {
                "type": "object",
                "properties": {
                    "output_filename": {
                        "type": "string",
                        "description": "The name for the new output Excel file (e.g., 'conditional_data.xlsx')."
                    },
                    "new_column_name": {
                        "type": "string",
                        "description": "The name of the new column to create."
                    },
                    "source_column": {
                        "type": "string",
                        "description": "The name of the column to check the condition against."
                    },
                    "operator": {
                        "type": "string",
                        "description": "The comparison operator for the condition.",
                        "enum": ["==", "!=", ">", "<", ">=", "<=", "contains"]
                    },
                    "value": {
                        "type": "string",
                        "description": "The value to compare against in the condition."
                    },
                    "true_value": {
                        "type": "string",
                        "description": "The value to set in the new column if the condition is true."
                    },
                    "false_value": {
                        "type": "string",
                        "description": "The value to set in the new column if the condition is false."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "The sheet to use. Defaults to the active sheet."
                    }
                },
                "required": ["output_filename", "new_column_name", "source_column", "operator", "value", "true_value", "false_value"]
            }
        },
        "implementation": processor.conditional_value_column
    },
    {
        "schema": {
            "name": "create_sheet",
            "description": "Creates a new, empty sheet with a specified name in the workbook.",
            "parameters": {
                "type": "object",
                "properties": {
                    "output_filename": {
                        "type": "string",
                        "description": "The name for the new output Excel file (e.g., 'workbook_new_sheet.xlsx')."
                    },
                    "new_sheet_name": {
                        "type": "string",
                        "description": "The name for the new sheet to be created."
                    }
                },
                "required": ["output_filename", "new_sheet_name"]
            }
        },
        "implementation": processor.create_sheet
    },
    {
        "schema": {
            "name": "delete_sheet",
            "description": "Deletes a specific sheet from the workbook.",
            "parameters": {
                "type": "object",
                "properties": {
                    "output_filename": {
                        "type": "string",
                        "description": "The name for the new output Excel file (e.g., 'workbook_deleted_sheet.xlsx')."
                    },
                    "sheet_to_delete": {
                        "type": "string",
                        "description": "The name of the sheet to be deleted."
                    }
                },
                "required": ["output_filename", "sheet_to_delete"]
            }
        },
        "implementation": processor.delete_sheet
    },
    {
        "schema": {
            "name": "duplicate_sheet",
            "description": "Duplicates an existing sheet and saves the workbook with the new sheet.",
            "parameters": {
                "type": "object",
                "properties": {
                    "output_filename": {
                        "type": "string",
                        "description": "The name for the new output Excel file (e.g., 'workbook_duplicated.xlsx')."
                    },
                    "source_sheet": {
                        "type": "string",
                        "description": "The name of the existing sheet to duplicate."
                    },
                    "new_sheet_name": {
                        "type": "string",
                        "description": "The name for the newly created duplicate sheet."
                    }
                },
                "required": ["output_filename", "source_sheet", "new_sheet_name"]
            }
        },
        "implementation": processor.duplicate_sheet
    },
    {
        "schema": {
            "name": "apply_conditional_formatting",
            "description": "Applies conditional formatting to a column based on a cell value condition.",
            "parameters": {
                "type": "object",
                "properties": {
                    "output_filename": {
                        "type": "string",
                        "description": "The name for the new output Excel file (e.g., 'formatted_data.xlsx')."
                    },
                    "column_name": {
                        "type": "string",
                        "description": "The name of the column to apply the formatting to."
                    },
                    "operator": {
                        "type": "string",
                        "description": "The comparison operator for the rule.",
                        "enum": ["greaterThan", "lessThan", "equal", "notEqual"]
                    },
                    "value": {
                        "type": "string",
                        "description": "The value to compare against. Must be a number for numeric comparisons."
                    },
                    "color": {
                        "type": "string",
                        "description": "The color to apply if the rule is met.",
                        "enum": ["red", "green", "yellow"]
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "The sheet to use. Defaults to the active sheet."
                    }
                },
                "required": ["output_filename", "column_name", "operator", "value", "color"]
            }
        },
        "implementation": processor.apply_conditional_formatting
    }
]

# --- Tool Access and Execution ---

import inspect

# A dictionary mapping tool names to their implementation for quick access.
_TOOL_IMPLEMENTATIONS = {tool["schema"]["name"]: tool["implementation"] for tool in TOOLS}

def get_tool_schemas() -> list:
    """Returns the list of all tool schemas."""
    return [tool["schema"] for tool in TOOLS]

def execute_tool(tool_name: str, parameters: dict, context: dict) -> dict:
    """
    Executes a tool by its name with the given parameters and context.
    It intelligently filters the combined parameters to only pass what the
    target function actually accepts.
    """
    if tool_name not in _TOOL_IMPLEMENTATIONS:
        return {"error": f"Tool '{tool_name}' does not exist."}

    tool_function = _TOOL_IMPLEMENTATIONS[tool_name]
    
    # Get the signature of the tool function to find out what parameters it accepts.
    sig = inspect.signature(tool_function)
    accepted_params = sig.parameters.keys()

    # Combine the LLM-provided parameters with the system-provided context.
    full_params = {**context, **parameters}

    # Filter the combined parameters to only include those accepted by the function.
    params_to_pass = {k: v for k, v in full_params.items() if k in accepted_params}

    return tool_function(**params_to_pass)
