# MS Excel MCP

A simple Model Context Protocol (MCP) server for working with an Excel workbook.

## What it does

This project allows the user to work on MS Excel workbooks by connecting this MCP server to their preferred AI chat(Claude, ChatGPT, etc). Start working on your excel sheet by just talking to AI.
The MCP server exposes a small set of tools for a local Excel file:

- `read_excel` - read the used range from the workbook as a 2D matrix
- `write_excel` - overwrite the used range with a new matrix
- `save_excel` - save and close the workbook
- `discard_changes` - close the workbook without saving
- `apply_formula` - write a formula into a specific cell

## Requirements

- Windows
- Microsoft Excel installed
- Any LLM chat interface that allows MCP plugins

## Setup

1. Activate the virtual environment:

```powershell
.\myenv\Scripts\Activate.ps1
```

2. Install dependencies if needed:

```powershell
pip install mcp xlwings
```

## Run

Start the MCP server with:

```powershell
python server.py
```
