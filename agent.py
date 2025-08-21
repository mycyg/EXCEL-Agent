"""
The final, most advanced version of the agent logic.
It uses a ReAct-style loop and provides a structured JSON schema of the tools
to the LLM, ensuring reliable tool calls.
"""

import json
import os
from openai import OpenAI
import config
import tools # Import the new tools module

# --- LLM Client Initialization ---
ark_api_key = config.ARK_API_KEY or os.environ.get("ARK_API_KEY")
if not ark_api_key:
    print("Error: ARK_API_KEY not found.")
    exit(1)

client = OpenAI(
    base_url=config.ARK_BASE_URL,
    api_key=ark_api_key,
)

# --- Core Agent Logic ---

def run_agent_task(user_input: str, file_path: str, chart_output_dir: str, file_output_dir: str) -> dict:
    """
    Runs a full agent task, passing safe output directories to the tools.
    """
    print(f"--- New Task ---\nUser Input: {user_input}")

    system_prompt = _create_system_prompt()
    conversation_history = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": f"User Request: \"{user_input}\"\nFile to be used: \"{file_path}\""}
    ]
    
    steps = []
    observations = []

    # Define the context dictionary that will be passed to the tool executor
    tool_context = {
        "file_path": file_path,
        "chart_output_dir": chart_output_dir,
        "file_output_dir": file_output_dir
    }

    max_steps = 5
    for i in range(max_steps):
        print(f"--- Agent Step {i+1} ---")
        
        try:
            llm_response_str = _call_llm(conversation_history)
            conversation_history.append({"role": "assistant", "content": llm_response_str})
            llm_response = json.loads(llm_response_str)
            
            thought = llm_response.get('thought', '(No thought provided)')
            print(f"LLM Thought: {thought}")
            current_step = {"thought": thought}

            if "final_answer" in llm_response:
                print("LLM provided final answer.")
                steps.append(current_step)
                return {"answer": llm_response["final_answer"], "steps": steps, "observations": observations}

            tool_call = llm_response.get("tool_call")
            if not tool_call:
                return {"answer": "I am not sure how to proceed.", "steps": steps, "observations": observations}
            
            current_step["tool_call"] = tool_call
            tool_name = tool_call.get("tool_name")
            parameters = tool_call.get("parameters", {})
            
            print(f"Executing Tool: {tool_name} with params {parameters}")
            try:
                # All tool execution is now handled by the tools module
                observation_dict = tools.execute_tool(tool_name, parameters, tool_context)
            except Exception as e:
                observation_dict = {"error": f"Tool '{tool_name}' failed with error: {e}"}
            
            observation = json.dumps(observation_dict)
            observations.append(observation_dict)
            print(f"Tool Observation: {observation}")
            current_step["observation"] = observation
            steps.append(current_step)
            conversation_history.append({"role": "user", "content": f"Observation: {observation}"})

        except Exception as e:
            print(f"An error occurred in the agent loop: {e}")
            return {"answer": "I encountered an unexpected error.", "steps": steps, "observations": observations}

    return {"answer": "I seem to be stuck in a loop.", "steps": steps, "observations": observations}


# --- Prompt and Schema Generation ---

def _create_system_prompt() -> str:
    # Get schemas directly from the tools module
    tool_schemas = tools.get_tool_schemas()
    tools_schema_str = json.dumps(tool_schemas, indent=2)
    
    return f"""You are a smart agent that can solve user requests by breaking them down into a series of steps using the available tools.

Here is the list of tools available to you in a JSON schema format:
{tools_schema_str}

You must respond in a specific JSON format with two possible keys:
1. `thought`: Your reasoning and plan for the next step.
2. `tool_call`: The specific tool to execute for this step. The `tool_name` must be one of the names from the schema. The `parameters` must adhere to the schema for that tool.

OR

1. `thought`: Your reasoning for why you are finished.
2. `final_answer`: A concise, natural language response to the user's original request, based on your observations.

A key first step for many tasks is to use `get_data_summary` to understand the file's structure before trying to access columns.
After receiving an "Observation", you must decide on the next step: either another `tool_call` or a `final_answer`.
"""

# --- LLM Communication ---

def _call_llm(history: list) -> str:
    """A wrapper for calling the LLM API with conversation history."""
    response = client.chat.completions.create(
        model=config.ARK_MODEL_ID,
        messages=history,
        response_format={"type": "json_object"} 
    )
    return response.choices[0].message.content
