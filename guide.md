# Google Sheets MCP API Guide

This guide documents how to call every MCP tool in this project, including accepted input data and current mode behavior.

## Base Endpoint

- JSON-RPC endpoint: `http://127.0.0.1:8000/mcp`
- JSON-RPC method: `tools/call`
- Tool name is passed in `params.name`
- Tool arguments are passed in `params.arguments`

### Generic JSON-RPC Envelope

```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "tools/call",
  "params": {
    "name": "<tool_name>",
    "arguments": {
      "...": "..."
    }
  }
}
```

---

## Required Common Inputs

Most tools require:

- `spreadsheet_id` (string): Spreadsheet identifier used by the server (`gspread.open(...)` in current implementation)
- `sheet` (string): Sheet/tab name
- `user_id` (string): Used to fetch OAuth access token

---

## 1) `get_sheet_data`

Reads data from a sheet or a range.

### Arguments

- `spreadsheet_id` (string, required)
- `sheet` (string, required)
- `user_id` (string, required)
- `range` (string, optional): A1-style range like `"A1:D50"`
- `include_grid_data` (bool, optional, default `false`)

### Example

```json
{
  "spreadsheet_id": "your_sheet_id",
  "sheet": "Ledger",
  "range": "A1:C20",
  "include_grid_data": true,
  "user_id": "user-uuid"
}
```

### Success output (typical)

- `spreadsheetId`
- `sheet`
- `range`
- `values` (2D list)
- `rowCount`
- `columnCount`
- optional `extra` when `include_grid_data=true`

---

## 2) `get_sheet_formulas`

Reads formulas (not computed values).

### Arguments

- `spreadsheet_id` (string, required)
- `sheet` (string, required)
- `user_id` (string, required)
- `range` (string, optional)

### Example

```json
{
  "spreadsheet_id": "your_sheet_id",
  "sheet": "Ledger",
  "range": "B2:D20",
  "user_id": "user-uuid"
}
```

### Success output (typical)

- `formulas` (2D list)
- `range_used`
- `rowCount`
- `columnCount`
- `has_formulas`

---

## 3) `create_new_sheet`

Creates a new tab.

### Arguments

- `spreadsheet_id` (string, required)
- `title` (string, required)
- `user_id` (string, required)
- `rows` (int, optional, default `100`)
- `cols` (int, optional, default `26`)

### Example

```json
{
  "spreadsheet_id": "your_sheet_id",
  "title": "MyNewTab",
  "rows": 200,
  "cols": 20,
  "user_id": "user-uuid"
}
```

---

## 4) `update_single_cell`

Updates one cell. Supports **two modes** with `cell_name`.

### Arguments

- `spreadsheet_id` (string, required)
- `sheet` (string, required)
- `row` (int, required by signature; ignored in A1 mode)
- `col` (int|string, required)
- `value` (any JSON value, required)
- `user_id` (string, required)
- `header_row` (int, optional, default `1`)
- `cell_name` (bool, optional, default `false`)

### Mode A: `cell_name = false` (classic mode)

- `row` must be `>= 1`
- `col` can be:
  - integer column index (1-based)
  - column letters like `"A"`, `"AA"`
  - header name (resolved from `header_row`)

### Mode B: `cell_name = true`

- `col` is interpreted in this order:
  1. **A1 reference** (e.g., `"B5"`) → row+col parsed from A1
  2. if not A1, treated as **header name** and paired with provided `row`
- In header fallback path, `row` must be `>= 1`

### Examples

Classic:

```json
{
  "spreadsheet_id": "your_sheet_id",
  "sheet": "Ledger",
  "row": 4,
  "col": "B",
  "value": 1200,
  "user_id": "user-uuid",
  "cell_name": false
}
```

A1 mode:

```json
{
  "spreadsheet_id": "your_sheet_id",
  "sheet": "Ledger",
  "row": 999,
  "col": "C7",
  "value": "Utilities",
  "user_id": "user-uuid",
  "cell_name": true
}
```

Header+row fallback mode:

```json
{
  "spreadsheet_id": "your_sheet_id",
  "sheet": "Ledger",
  "row": 7,
  "col": "Description",
  "value": "Utilities",
  "user_id": "user-uuid",
  "header_row": 1,
  "cell_name": true
}
```

---

## 5) `update_single_row`

Updates one row using list or header-mapped dict.

### Arguments

- `spreadsheet_id` (string, required)
- `sheet` (string, required)
- `row` (int, required)
- `values_json` (string, required): JSON-encoded list or dict
- `user_id` (string, required)
- `header_row` (int, optional, default `1`)

### Accepted `values_json`

- list form: `"[\"2026-02-26\", \"Rent\", 5000]"`
- dict form: `"{\"Date\":\"2026-02-26\",\"Amount\":5000}"`

### Example

```json
{
  "spreadsheet_id": "your_sheet_id",
  "sheet": "Ledger",
  "row": 10,
  "values_json": "{\"Date\":\"2026-02-26\",\"Description\":\"Rent\",\"Amount\":5000}",
  "user_id": "user-uuid",
  "header_row": 1
}
```

---

## 6) `append_row`

Appends a row at the bottom.

### Arguments

- `spreadsheet_id` (string, required)
- `sheet` (string, required)
- `values_json` (string, required): JSON-encoded list or dict
- `user_id` (string, required)
- `header_row` (int, optional, default `1`)

### Accepted `values_json`

- list form
- dict form (mapped by header names)

### Example

```json
{
  "spreadsheet_id": "your_sheet_id",
  "sheet": "Ledger",
  "values_json": "[\"2026-02-26\",\"Salary\",120000]",
  "user_id": "user-uuid"
}
```

---

## 7) `batch_update_cells`

Batch updates multiple cells. Supports two modes via `cell_name`.

### Arguments

- `spreadsheet_id` (string, required)
- `sheet` (string, required)
- `updates_json` (string, required): JSON-encoded list of update objects
- `user_id` (string, required)
- `header_row` (int, optional, default `1`)
- `cell_name` (bool, optional, default `false`)

### Mode A: `cell_name = false` (classic)

Each update item:

```json
{"row": 6, "col": "B", "value": 45000}
```

- `col` supports int or string
- string is interpreted as column letters first, then header fallback

### Mode B: `cell_name = true` (current implementation)

**Important:** current code expects:

```json
{"row": 6, "col": "HeaderName", "value": "..."}
```

- It does **not** currently parse A1 reference for batch mode.
- It resolves `col` as a header key from `header_row`.

### Example

```json
{
  "spreadsheet_id": "your_sheet_id",
  "sheet": "Ledger",
  "updates_json": "[{\"row\":6,\"col\":\"Description\",\"value\":\"Electricity\"}]",
  "user_id": "user-uuid",
  "cell_name": true,
  "header_row": 1
}
```

---

## 8) `batch_update_rows`

Batch updates multiple rows.

### Arguments

- `spreadsheet_id` (string, required)
- `sheet` (string, required)
- `updates_json` (string, required): JSON-encoded list
- `user_id` (string, required)
- `header_row` (int, optional, default `1`)

### Accepted row item formats

- list values:

```json
{"row": 12, "values": ["2026-02-26", "Fuel", 900]}
```

- dict values:

```json
{"row": 13, "values": {"Date": "2026-02-27", "Description": "Bonus", "Amount": 15000}}
```

- `values` may also be a JSON string that decodes to list/dict.

---

## Error Shape

Typical errors from tools are returned as dictionaries such as:

```json
{
  "error": "Failed to ...",
  "detail": "..."
}
```

Some tools also return `status` for not-found cases.

---

## How to Run Comprehensive Tests

Use the included integration suite:

```bash
uv run test.py
```

Optional environment variables:

- `MCP_SERVER_URL`
- `MCP_USER_ID`
- `MCP_SPREADSHEET_ID`
- `MCP_BASE_SHEET`

The test script creates a temporary sheet and validates all tools including mode-specific paths.
