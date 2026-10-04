# StudyBuddy AI — Personal AI Agent

A general-purpose AI agent built from first principles in Python: no agent framework, just the Anthropic API, a hand-written tool-calling loop, and a set of real tools for data analysis, file handling, and document generation. Built as a learning project and portfolio piece, with a focus on understanding — not hiding behind — the mechanics of how LLM agents actually work.

## Why this project

Most "AI agent" tutorials wrap everything in a framework (LangChain, LangGraph) from line one, which makes it easy to use an agent without understanding what an agent *is*. This project was built the opposite way: implementing the tool-calling loop, memory, and error handling by hand first, so every piece is something I can explain, not just something that works.

## What it does

StudyBuddy AI is a chat-based agent that can reason about a request, decide whether it needs a tool, call that tool, read the result, and continue — looping until it has a real answer. It supports multiple, independently saved conversations (a "New Chat" button plus a switchable history list, like ChatGPT/Claude's own interface), a file upload sidebar, and a persistent memory system so conversations survive restarts.

### Multi-chat memory

Each conversation is its own file on disk (`chats/<timestamp>.json`), not one giant history. The sidebar lists past chats (labeled from their first message), lets you switch between them or delete one, and "New Chat" simply starts a fresh, empty message list under a new id. This mirrors how the API itself has no built-in concept of "conversations" — it's stateless per call — so anything resembling separate, resumable chats is this project's own storage logic, not an API feature.

### Tools

| Tool | What it does |
|---|---|
| `calculator` | Basic arithmetic, with explicit divide-by-zero handling |
| `execute_python` | Runs arbitrary Python (pandas, etc.) in an isolated subprocess — see **Security** below |
| `inspect_file` | Reads a CSV/Excel file's structure (columns, types, preview) before analysis |
| `run_sql_query` | Read-only SQL access to a SQLite database |
| `read_pdf` | Extracts text from PDF documents |
| `generate_report` | Produces a formatted Word (.docx) report |
| `web_search` | Live web search (Anthropic-hosted tool) |
| `word_counter` | Simple text utility |

## Architecture

The core loop (the actual design I planned before writing a line of code):

```
User input
   ↓
Call Claude with full conversation history + tool definitions
   ↓
Does the response need a tool? ──No──→ Return final answer
   ↓ Yes
Run the real Python function for that tool
   ↓
Append the tool's result to the conversation history
   ↓
Loop back to "Call Claude" — repeat until no more tools are needed
```

Every reasoning step and every tool result lives in one growing `messages` list, which is resent to the API in full on each call. The model has no memory of its own between API calls — any appearance of "memory" (within a conversation, or across restarts) is this project's own code rebuilding and resending that list.

## Tech stack

- **Python 3.13**
- **Anthropic API** (Claude Sonnet) — tool use / function calling
- **Streamlit** — UI
- **pandas** — data analysis
- **SQLite** — database querying
- **pypdf**, **python-docx** — document handling
- **pytest** — automated tests

## Setup

```bash
pip install anthropic python-dotenv streamlit pandas pypdf python-docx pytest
```

Create a `.env` file in the project root:
```
ANTHROPIC_API_KEY=your-key-here
```

Run the app:
```bash
streamlit run app.py
```

## Security

Running LLM-generated Python is a real risk, not a theoretical one, so this isn't an afterthought:

- **Sandboxed code execution** — `execute_python` runs in a separate OS subprocess, not in-process via `exec()`. The subprocess gets a scrubbed environment (no API keys or other secrets), a 30-second timeout, and an output size cap.
- **Read-only SQL** — `run_sql_query` only permits `SELECT` statements; destructive SQL (`DROP`, `DELETE`, `UPDATE`) is rejected before it reaches the database.
- **Path containment** — `inspect_file` and `read_pdf` can only read files inside the project directory; absolute paths pointing elsewhere are rejected.
- **Defense in depth** — these are enforced in code, not left to the model's judgment, though in practice the model also declines destructive requests on its own.

## Error handling

Every tool fails safely and returns a readable error string rather than crashing. The agent loop itself is wrapped so that if any step fails mid-request (a bad API response, a tool exception), the conversation history is rolled back to its last valid state — this prevents one failed request from corrupting the whole conversation for every future message.

## Testing

```bash
pytest test_tools.py -v
```

Tests cover the deterministic logic directly (arithmetic, path-safety checks) rather than the live API, since API calls are costly and non-deterministic to test against on every run.

## What I'd add next

- A proper sandbox (container-level isolation) for `execute_python`, rather than process-level
- Retrieval-augmented generation for large document sets, instead of resending full conversation history every call
- Automated chart/visualization generation surfaced directly in the Streamlit UI

## Built by

Megha Gandhi — built incrementally, milestone by milestone, as a learning project.
