import streamlit as st

st.set_page_config(page_title="StudyBuddy AI Agent", page_icon="🤖", layout="wide")

from dotenv import load_dotenv
load_dotenv()

import anthropic
import subprocess
import sys
import io
import contextlib
import pandas as pd
import sqlite3
import json
import os
import time
from pypdf import PdfReader
from docx import Document

client = anthropic.Anthropic(
    default_headers={"anthropic-workspace-id": os.getenv("ANTHROPIC_WORKSPACE_ID")}  
)

CHATS_DIR = "chats"

# ---------- Memory: multiple saved chats, one JSON file each ----------

def ensure_chats_dir():
    if not os.path.exists(CHATS_DIR):
        os.makedirs(CHATS_DIR)

def list_chats():
    """Return all saved chat IDs, newest first."""
    ensure_chats_dir()
    files = [f for f in os.listdir(CHATS_DIR) if f.endswith(".json")]
    files.sort(reverse=True)  # newest timestamp first
    return [f.replace(".json", "") for f in files]

def load_chat(chat_id):
    path = os.path.join(CHATS_DIR, f"{chat_id}.json")
    if os.path.exists(path):
        with open(path, "r") as f:
            return json.load(f)
    return []

def save_chat(chat_id, messages):
    ensure_chats_dir()
    path = os.path.join(CHATS_DIR, f"{chat_id}.json")
    with open(path, "w") as f:
        json.dump(messages, f, indent=2)

def delete_chat(chat_id):
    path = os.path.join(CHATS_DIR, f"{chat_id}.json")
    if os.path.exists(path):
        os.remove(path)

def new_chat_id():
    return str(int(time.time()))

def chat_label(chat_id):
    """Build a readable label from the chat's first user message, if any."""
    messages = load_chat(chat_id)
    for msg in messages:
        if msg["role"] == "user" and isinstance(msg["content"], str):
            text = msg["content"].strip()
            if text:
                return text[:30] + ("..." if len(text) > 30 else "")
    return "New chat"

def serialize_content(content):
    if isinstance(content, str):
        return content
    return [block.model_dump() if hasattr(block, "model_dump") else block for block in content]

# ---------- Tool schemas ----------

tools = [
    {
        "name": "calculator",
        "description": "Performs basic arithmetic: addition, subtraction, multiplication, or division of two numbers.",
        "input_schema": {
            "type": "object",
            "properties": {
                "operation": {"type": "string", "enum": ["add", "subtract", "multiply", "divide"], "description": "The arithmetic operation to perform"},
                "a": {"type": "number", "description": "The first number"},
                "b": {"type": "number", "description": "The second number"}
            },
            "required": ["operation", "a", "b"]
        }
    },
    {
        "name": "word_counter",
        "description": "Counts the number of words in a given piece of text.",
        "input_schema": {
            "type": "object",
            "properties": {"text": {"type": "string", "description": "The text to count words in"}},
            "required": ["text"]
        }
    },
    {
        "name": "execute_python",
        "description": "Executes a snippet of Python code and returns its printed output. Use this for calculations, data analysis, pandas operations, or any task Python can solve. The code MUST use print() to output whatever result should be returned.",
        "input_schema": {
            "type": "object",
            "properties": {"code": {"type": "string", "description": "Valid Python code to execute. Must call print() on the final result."}},
            "required": ["code"]
        }
    },
    {
        "name": "inspect_file",
        "description": "Reads a CSV or Excel file and returns its column names, data types, number of rows, and the first 5 rows as a preview. Use this BEFORE writing analysis code, to understand a file's structure.",
        "input_schema": {
            "type": "object",
            "properties": {"filepath": {"type": "string", "description": "Path to the CSV or Excel file to inspect"}},
            "required": ["filepath"]
        }
    },
    {
        "name": "run_sql_query",
        "description": (
            "Executes a SQL query against the company.db SQLite database and returns the results. "
            "The database has a table called 'employees' with columns: "
            "Employee_Id, Employee_Name, company, job, degree, salary_more_then_100k, salary, pay_category. "
            "Use this for any question that can be answered with SQL (filtering, grouping, aggregating, sorting). "
            "Only SELECT queries are permitted."
        ),
        "input_schema": {
            "type": "object",
            "properties": {"query": {"type": "string", "description": "A valid SELECT SQL query to run against the employees table"}},
            "required": ["query"]
        }
    },
    {"type": "web_search_20250305", "name": "web_search"},
    {
        "name": "read_pdf",
        "description": "Extracts and returns all text content from a PDF file. Use this to read reports, resumes, assignments, or any PDF document before answering questions about it.",
        "input_schema": {
            "type": "object",
            "properties": {"filepath": {"type": "string", "description": "Path to the PDF file to read"}},
            "required": ["filepath"]
        }
    },
    {
        "name": "generate_report",
        "description": "Creates a Word document (.docx) report with a title and body content. Use this when the user asks for a report, summary document, or something to save/download after an analysis.",
        "input_schema": {
            "type": "object",
            "properties": {
                "title": {"type": "string", "description": "The report's title, shown as a heading at the top"},
                "content": {"type": "string", "description": "The full body text of the report. Use double newlines (\\n\\n) to separate paragraphs."},
                "filename": {"type": "string", "description": "Filename to save as, e.g. 'sales_report.docx'"}
            },
            "required": ["title", "content", "filename"]
        }
    }
]

# ---------- Real Python functions ----------

def calculator(operation, a, b):
    try:
        if operation == "add": return a + b
        elif operation == "subtract": return a - b
        elif operation == "multiply": return a * b
        elif operation == "divide":
            if b == 0:
                return "Error: cannot divide by zero"
            return a / b
        else:
            return f"Error: unknown operation '{operation}'"
    except Exception as e:
        return f"Error: {e}"

def word_counter(text):
    try:
        return len(text.split())
    except Exception as e:
        return f"Error: {e}"

def execute_python(code, timeout=30):
    safe_env = {
        key: os.environ[key]
        for key in ("PATH", "SYSTEMROOT", "TEMP", "TMP")
        if key in os.environ
    }
    safe_env["PYTHONIOENCODING"] = "utf-8"

    try:
        result = subprocess.run(
            [sys.executable, "-c", code],
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=timeout,
            env=safe_env,
        )
    except subprocess.TimeoutExpired:
        return f"Error: code took longer than {timeout} seconds and was stopped."
    except Exception as e:
        return f"Error: {e}"

    if result.returncode != 0:
        return f"Error: {result.stderr.strip()[-2000:]}"

    output = result.stdout
    if len(output) > 10000:
        output = output[:10000] + "\n...[output truncated]"
    return output

def is_safe_path(filepath):
    """Only allow access to files inside the current project folder."""
    project_dir = os.path.abspath(os.getcwd())
    target_path = os.path.abspath(filepath)
    return target_path.startswith(project_dir)

def inspect_file(filepath):
    if not is_safe_path(filepath):
        return "Error: access denied. Only files in the project folder can be read."
    try:
        if filepath.endswith(".csv"):
            df = pd.read_csv(filepath)
        elif filepath.endswith((".xlsx", ".xls")):
            df = pd.read_excel(filepath)
        else:
            return "Error: unsupported file type. Use .csv or .xlsx"
        info = f"Columns: {list(df.columns)}\n"
        info += f"Shape: {df.shape[0]} rows, {df.shape[1]} columns\n"
        info += f"Data types:\n{df.dtypes.to_string()}\n"
        info += f"First 5 rows:\n{df.head().to_string()}"
        return info
    except Exception as e:
        return f"Error: {e}"

def run_sql_query(query):
    try:
        if not query.strip():
            return "Error: empty SQL query provided"

        if not query.strip().lower().startswith("select"):
            return "Error: only SELECT queries are allowed. This tool is read-only."

        conn = sqlite3.connect("company.db")
        cursor = conn.cursor()
        cursor.execute(query)

        rows = cursor.fetchall()
        columns = [description[0] for description in cursor.description]
        conn.close()
        if not rows:
            return "Query ran successfully but returned no rows."
        result = f"Columns: {columns}\n"
        result += "\n".join(str(row) for row in rows)
        return result
    except Exception as e:
        return f"Error: {e}"

def read_pdf(filepath):
    if not is_safe_path(filepath):
        return "Error: access denied. Only files in the project folder can be read."
    try:
        reader = PdfReader(filepath)
        text = ""
        for page_num, page in enumerate(reader.pages):
            page_text = page.extract_text()
            text += f"--- Page {page_num + 1} ---\n{page_text}\n"

        if not text.strip():
            return "Warning: no extractable text found. This PDF may be a scanned image or password-protected."

        return text
    except Exception as e:
        return f"Error: {e}"

def generate_report(title, content, filename):
    try:
        if not filename.endswith(".docx"):
            filename += ".docx"

        doc = Document()
        doc.add_heading(title, level=1)
        paragraphs = content.split("\n\n")
        for para in paragraphs:
            if para.strip():
                doc.add_paragraph(para.strip())
        doc.save(filename)
        return f"Report saved successfully as {filename}"
    except Exception as e:
        return f"Error: {e}"

tool_functions = {
    "calculator": calculator,
    "word_counter": word_counter,
    "execute_python": execute_python,
    "inspect_file": inspect_file,
    "run_sql_query": run_sql_query,
    "read_pdf": read_pdf,
    "generate_report": generate_report
}

# ---------- The agent loop - with rollback on error ----------

def run_agent(user_input, messages):
    """
    Runs the full tool-calling loop. If ANYTHING fails partway through
    (API error, tool crash, etc.), rolls back messages to exactly the
    state it was in before this call started - preventing corrupted
    tool_use/tool_result pairs from poisoning future requests.
    """
    checkpoint = len(messages)
    messages.append({"role": "user", "content": user_input})

    try:
        while True:
            response = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=4096,
                tools=tools,
                messages=messages
            )

            messages.append({"role": "assistant", "content": serialize_content(response.content)})

            if response.stop_reason != "tool_use":
                text_parts = [block.text for block in response.content if block.type == "text"]
                return "\n".join(text_parts)

            tool_results = []
            for block in response.content:
                if block.type == "tool_use":
                    func = tool_functions.get(block.name)
                    if func:
                        try:
                            result = func(**block.input)
                        except Exception as tool_error:
                            result = f"Error running tool: {tool_error}"
                    else:
                        result = f"Error: unknown tool {block.name}"
                    tool_results.append({
                        "type": "tool_result",
                        "tool_use_id": block.id,
                        "content": str(result)
                    })

            if tool_results:
                messages.append({"role": "user", "content": tool_results})

    except Exception as e:
        del messages[checkpoint:]
        return (
            f"⚠️ Something went wrong while processing that request: {e}\n\n"
            f"Your conversation history has been preserved from before this attempt - "
            f"you can try again."
        )


# ============================================================
# STREAMLIT UI
# ============================================================

# ---------- Initialize current chat (must happen before sidebar renders) ----------

if "current_chat_id" not in st.session_state:
    existing = list_chats()
    if existing:
        st.session_state.current_chat_id = existing[0]
        st.session_state.messages = load_chat(existing[0])
    else:
        st.session_state.current_chat_id = new_chat_id()
        st.session_state.messages = []

if "processing" not in st.session_state:
    st.session_state.processing = False

# ---------- Sidebar ----------

with st.sidebar:
    st.header("⚙️ Settings")

    if st.button("➕ New Chat", use_container_width=True, disabled=st.session_state.processing):
        st.session_state.current_chat_id = new_chat_id()
        st.session_state.messages = []
        st.rerun()

    st.divider()
    st.subheader("💬 Past Chats")

    chat_ids = list_chats()
    if not chat_ids:
        st.caption("No past chats yet.")
    else:
        for chat_id in chat_ids:
            is_active = (chat_id == st.session_state.current_chat_id)
            label = ("🟢 " if is_active else "") + chat_label(chat_id)

            col1, col2 = st.columns([4, 1])
            with col1:
                if st.button(label, key=f"chat_{chat_id}", use_container_width=True,
                             disabled=st.session_state.processing):
                    st.session_state.current_chat_id = chat_id
                    st.session_state.messages = load_chat(chat_id)
                    st.rerun()
            with col2:
                if st.button("🗑️", key=f"del_{chat_id}", disabled=st.session_state.processing):
                    delete_chat(chat_id)
                    if chat_id == st.session_state.current_chat_id:
                        remaining = list_chats()
                        if remaining:
                            st.session_state.current_chat_id = remaining[0]
                            st.session_state.messages = load_chat(remaining[0])
                        else:
                            st.session_state.current_chat_id = new_chat_id()
                            st.session_state.messages = []
                    st.rerun()

    st.divider()

    st.subheader("📁 Upload a data file")
    uploaded_file = st.file_uploader(
        "Upload a CSV or Excel file to analyze",
        type=["csv", "xlsx", "xls"]
    )
    if uploaded_file is not None:
        try:
            with open(uploaded_file.name, "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.success(f"Saved as `{uploaded_file.name}` — ask the agent to inspect or analyze it!")
        except Exception as e:
            st.error(f"Couldn't save the uploaded file: {e}")

    st.divider()
    st.caption("Built with Claude + Streamlit")

# ---------- Main chat area ----------

st.title("🤖 StudyBuddy AI Agent")
st.caption("Your personal data-analysis and productivity agent")

for msg in st.session_state.messages:
    role = msg["role"]
    content = msg["content"]

    if role == "user" and isinstance(content, str):
        with st.chat_message("user", avatar="🧑‍💻"):
            st.write(content)
    elif role == "assistant" and isinstance(content, list):
        text_parts = [b.get("text") for b in content if isinstance(b, dict) and b.get("type") == "text"]
        if text_parts:
            with st.chat_message("assistant", avatar="🤖"):
                st.write("\n".join(text_parts))

user_input = st.chat_input(
    "Ask your agent something..." if not st.session_state.processing else "Waiting for the agent to finish...",
    disabled=st.session_state.processing
)

if user_input:
    st.session_state.processing = True

    with st.chat_message("user", avatar="🧑‍💻"):
        st.write(user_input)

    with st.chat_message("assistant", avatar="🤖"):
        with st.spinner("Thinking..."):
            answer = run_agent(user_input, st.session_state.messages)
        st.write(answer)

    save_chat(st.session_state.current_chat_id, st.session_state.messages)
    st.session_state.processing = False
    st.rerun()  # refresh so the sidebar's chat label updates with the new first message





# import streamlit as st

# st.set_page_config(page_title="MyAI Agent", page_icon="🤖", layout="wide")

# from dotenv import load_dotenv
# load_dotenv()

# import anthropic
# import subprocess
# import sys
# import io
# import contextlib
# import pandas as pd
# import sqlite3
# import json
# import os
# from pypdf import PdfReader
# from docx import Document

# client = anthropic.Anthropic(
#     default_headers={"anthropic-workspace-id": "wrkspc_01Dc5LkRcCPmdo9sehrtkX3X"} 
# )

# MEMORY_FILE = "conversation_history.json"

# print("SCRIPT STARTED — IF YOU SEE THIS, PRINTS ARE WORKING", flush=True)

# # ---------- Memory: load/save conversation history to disk ----------

# def load_history():
#     if os.path.exists(MEMORY_FILE):
#         with open(MEMORY_FILE, "r") as f:
#             return json.load(f)
#     return []

# def save_history(messages):
#     with open(MEMORY_FILE, "w") as f:
#         json.dump(messages, f, indent=2)

# def serialize_content(content):
#     if isinstance(content, str):
#         return content
#     return [block.model_dump() if hasattr(block, "model_dump") else block for block in content]

# # ---------- Tool schemas ----------

# tools = [
#     {
#         "name": "calculator",
#         "description": "Performs basic arithmetic: addition, subtraction, multiplication, or division of two numbers.",
#         "input_schema": {
#             "type": "object",
#             "properties": {
#                 "operation": {"type": "string", "enum": ["add", "subtract", "multiply", "divide"], "description": "The arithmetic operation to perform"},
#                 "a": {"type": "number", "description": "The first number"},
#                 "b": {"type": "number", "description": "The second number"}
#             },
#             "required": ["operation", "a", "b"]
#         }
#     },
#     {
#         "name": "word_counter",
#         "description": "Counts the number of words in a given piece of text.",
#         "input_schema": {
#             "type": "object",
#             "properties": {"text": {"type": "string", "description": "The text to count words in"}},
#             "required": ["text"]
#         }
#     },
#     {
#         "name": "execute_python",
#         "description": "Executes a snippet of Python code and returns its printed output. Use this for calculations, data analysis, pandas operations, or any task Python can solve. The code MUST use print() to output whatever result should be returned.",
#         "input_schema": {
#             "type": "object",
#             "properties": {"code": {"type": "string", "description": "Valid Python code to execute. Must call print() on the final result."}},
#             "required": ["code"]
#         }
#     },
#     {
#         "name": "inspect_file",
#         "description": "Reads a CSV or Excel file and returns its column names, data types, number of rows, and the first 5 rows as a preview. Use this BEFORE writing analysis code, to understand a file's structure.",
#         "input_schema": {
#             "type": "object",
#             "properties": {"filepath": {"type": "string", "description": "Path to the CSV or Excel file to inspect"}},
#             "required": ["filepath"]
#         }
#     },
#     {
#         "name": "run_sql_query",
#         "description": (
#             "Executes a SQL query against the company.db SQLite database and returns the results. "
#             "The database has a table called 'employees' with columns: "
#             "Employee_Id, Employee_Name, company, job, degree, salary_more_then_100k, salary, pay_category. "
#             "Use this for any question that can be answered with SQL (filtering, grouping, aggregating, sorting)."
#         ),
#         "input_schema": {
#             "type": "object",
#             "properties": {"query": {"type": "string", "description": "A valid SQL query to run against the employees table"}},
#             "required": ["query"]
#         }
#     },
#     {"type": "web_search_20250305", "name": "web_search"},
#     {
#         "name": "read_pdf",
#         "description": "Extracts and returns all text content from a PDF file. Use this to read reports, resumes, assignments, or any PDF document before answering questions about it.",
#         "input_schema": {
#             "type": "object",
#             "properties": {"filepath": {"type": "string", "description": "Path to the PDF file to read"}},
#             "required": ["filepath"]
#         }
#     },
#     {
#         "name": "generate_report",
#         "description": "Creates a Word document (.docx) report with a title and body content. Use this when the user asks for a report, summary document, or something to save/download after an analysis.",
#         "input_schema": {
#             "type": "object",
#             "properties": {
#                 "title": {"type": "string", "description": "The report's title, shown as a heading at the top"},
#                 "content": {"type": "string", "description": "The full body text of the report. Use double newlines (\\n\\n) to separate paragraphs."},
#                 "filename": {"type": "string", "description": "Filename to save as, e.g. 'sales_report.docx'"}
#             },
#             "required": ["title", "content", "filename"]
#         }
#     }
# ]

# # ---------- Real Python functions ----------

# def calculator(operation, a, b):
#     try:
#         if operation == "add": return a + b
#         elif operation == "subtract": return a - b
#         elif operation == "multiply": return a * b
#         elif operation == "divide":
#             if b == 0:
#                 return "Error: cannot divide by zero"
#             return a / b
#         else:
#             return f"Error: unknown operation '{operation}'"
#     except Exception as e:
#         return f"Error: {e}"

# def word_counter(text):
#     try:
#         return len(text.split())
#     except Exception as e:
#         return f"Error: {e}"

# def execute_python(code, timeout=30):
#     safe_env = {
#         key: os.environ[key]
#         for key in ("PATH", "SYSTEMROOT", "TEMP", "TMP")
#         if key in os.environ
#     }
#     safe_env["PYTHONIOENCODING"] = "utf-8"
    

#     try:
#         result = subprocess.run(
#             [sys.executable, "-c", code],
#             capture_output=True,
#             text=True,
#             encoding="utf-8",
#             timeout=timeout,
#             env=safe_env,
#         )
#     except subprocess.TimeoutExpired:
#         return f"Error: code took longer than {timeout} seconds and was stopped."
#     except Exception as e:
#         return f"Error: {e}"

#     if result.returncode != 0:
#         return f"Error: {result.stderr.strip()[-2000:]}"

#     output = result.stdout
#     if len(output) > 10000:
#         output = output[:10000] + "\n...[output truncated]"
#     return output

# def is_safe_path(filepath):
#     """Only allow access to files inside the current project folder."""
#     project_dir = os.path.abspath(os.getcwd())
#     target_path = os.path.abspath(filepath)
#     return target_path.startswith(project_dir)

# def inspect_file(filepath):
#     if not is_safe_path(filepath):
#         return "Error: access denied. Only files in the project folder can be read."
#     try:
#         if filepath.endswith(".csv"):
#             df = pd.read_csv(filepath)
#         elif filepath.endswith((".xlsx", ".xls")):
#             df = pd.read_excel(filepath)
#         else:
#             return "Error: unsupported file type. Use .csv or .xlsx"
#         info = f"Columns: {list(df.columns)}\n"
#         info += f"Shape: {df.shape[0]} rows, {df.shape[1]} columns\n"
#         info += f"Data types:\n{df.dtypes.to_string()}\n"
#         info += f"First 5 rows:\n{df.head().to_string()}"
#         return info
#     except Exception as e:
#         return f"Error: {e}"

# def run_sql_query(query):
#     try:
#         if not query.strip():
#             return "Error: empty SQL query provided"

#         # Only allow read-only queries - block anything that could modify
#         # or destroy data (DROP, DELETE, UPDATE, INSERT, ALTER, etc.)
#         if not query.strip().lower().startswith("select"):
#             return "Error: only SELECT queries are allowed. This tool is read-only."

#         conn = sqlite3.connect("company.db")
#         cursor = conn.cursor()
#         cursor.execute(query)

#         rows = cursor.fetchall()
#         columns = [description[0] for description in cursor.description]
#         conn.close()
#         if not rows:
#             return "Query ran successfully but returned no rows."
#         result = f"Columns: {columns}\n"
#         result += "\n".join(str(row) for row in rows)
#         return result
#     except Exception as e:
#         return f"Error: {e}"

# def read_pdf(filepath):
#     if not is_safe_path(filepath):
#         return "Error: access denied. Only files in the project folder can be read."
#     try:
#         reader = PdfReader(filepath)
#         text = ""
#         for page_num, page in enumerate(reader.pages):
#             page_text = page.extract_text()
#             text += f"--- Page {page_num + 1} ---\n{page_text}\n"

#         if not text.strip():
#             return "Warning: no extractable text found. This PDF may be a scanned image or password-protected."

#         return text
#     except Exception as e:
#         return f"Error: {e}"

# def generate_report(title, content, filename):
#     try:
#         if not filename.endswith(".docx"):
#             filename += ".docx"

#         doc = Document()
#         doc.add_heading(title, level=1)
#         paragraphs = content.split("\n\n")
#         for para in paragraphs:
#             if para.strip():
#                 doc.add_paragraph(para.strip())
#         doc.save(filename)
#         return f"Report saved successfully as {filename}"
#     except Exception as e:
#         return f"Error: {e}"

# tool_functions = {
#     "calculator": calculator,
#     "word_counter": word_counter,
#     "execute_python": execute_python,
#     "inspect_file": inspect_file,
#     "run_sql_query": run_sql_query,
#     "read_pdf": read_pdf,
#     "generate_report": generate_report
# }

# # ---------- The agent loop - NOW WITH ROLLBACK ON ERROR ----------

# def run_agent(user_input, messages):
#     """
#     Runs the full tool-calling loop. If ANYTHING fails partway through
#     (API error, tool crash, etc.), rolls back messages to exactly the
#     state it was in before this call started - preventing corrupted
#     tool_use/tool_result pairs from poisoning future requests.
#     """
#     checkpoint = len(messages)  # remember where we started
#     messages.append({"role": "user", "content": user_input})

#     try:
#         while True:
#             response = client.messages.create(
#                 model="claude-sonnet-4-6",
#                 max_tokens=4096,
#                 tools=tools,
#                 messages=messages
#             )
#             print(f"[DEBUG] stop_reason={response.stop_reason}, blocks={[b.type for b in response.content]}", flush=True)
#             messages.append({"role": "assistant", "content": serialize_content(response.content)})

#             if response.stop_reason != "tool_use":
#                 text_parts = [block.text for block in response.content if block.type == "text"]
#                 return "\n".join(text_parts)

#             tool_results = []
#             for block in response.content:
#                 if block.type == "tool_use":
#                     func = tool_functions.get(block.name)
#                     if func:
#                         try:
#                             result = func(**block.input)
#                         except Exception as tool_error:
#                             result = f"Error running tool: {tool_error}"
#                     else:
#                         result = f"Error: unknown tool {block.name}"
#                     tool_results.append({
#                         "type": "tool_result",
#                         "tool_use_id": block.id,
#                         "content": str(result)
#                     })

#             if tool_results:
#                 messages.append({"role": "user", "content": tool_results})

#     except Exception as e:
#         # Roll back to before this turn started. This is the key fix:
#         # without this, a failed request could leave a tool_use block
#         # with no matching tool_result, which permanently breaks every
#         # future API call using this history.
#         del messages[checkpoint:]
#         return (
#             f"⚠️ Something went wrong while processing that request: {e}\n\n"
#             f"Your conversation history has been preserved from before this attempt - "
#             f"you can try again."
#         )


# # ============================================================
# # STREAMLIT UI
# # ============================================================

# with st.sidebar:
#     st.header("⚙️ Settings")

#     st.subheader("📁 Upload a data file")
#     uploaded_file = st.file_uploader(
#         "Upload a CSV or Excel file to analyze",
#         type=["csv", "xlsx", "xls"]
#     )
#     if uploaded_file is not None:
#         try:
#             with open(uploaded_file.name, "wb") as f:
#                 f.write(uploaded_file.getbuffer())
#             st.success(f"Saved as `{uploaded_file.name}` — ask the agent to inspect or analyze it!")
#         except Exception as e:
#             st.error(f"Couldn't save the uploaded file: {e}")

#     st.divider()

#     if st.button("🗑️ Clear conversation", use_container_width=True):
#         st.session_state.messages = []
#         save_history([])
#         st.rerun()

#     st.divider()
#     st.caption("Built with Claude + Streamlit")

# st.title("🤖 MyAI Agent")
# st.caption("Your personal data-analysis and productivity agent")

# if "messages" not in st.session_state:
#     st.session_state.messages = load_history()

# if "processing" not in st.session_state:
#     st.session_state.processing = False

# for msg in st.session_state.messages:
#     role = msg["role"]
#     content = msg["content"]

#     if role == "user" and isinstance(content, str):
#         with st.chat_message("user", avatar="🧑‍💻"):
#             st.write(content)
#     elif role == "assistant" and isinstance(content, list):
#         text_parts = [b.get("text") for b in content if isinstance(b, dict) and b.get("type") == "text"]
#         if text_parts:
#             with st.chat_message("assistant", avatar="🤖"):
#                 st.write("\n".join(text_parts))

# # The input box is DISABLED while a request is processing - this is
# # what actually prevents the corruption bug: you physically cannot
# # send a second message until the first one finishes.
# user_input = st.chat_input(
#     "Ask your agent something..." if not st.session_state.processing else "Waiting for the agent to finish...",
#     disabled=st.session_state.processing
# )

# if user_input:
#     st.session_state.processing = True

#     with st.chat_message("user", avatar="🧑‍💻"):
#         st.write(user_input)

#     with st.chat_message("assistant", avatar="🤖"):
#         with st.spinner("Thinking..."):
#             answer = run_agent(user_input, st.session_state.messages)
#         st.write(answer)

#     save_history(st.session_state.messages)
#     st.session_state.processing = False