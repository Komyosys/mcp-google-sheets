#!/usr/bin/env python3
"""Comprehensive integration tests for all Google Spreadsheet MCP tools.

Usage:
  uv run test.py

Optional environment overrides:
  MCP_SERVER_URL
  MCP_USER_ID
  MCP_SPREADSHEET_ID
  MCP_BASE_SHEET
"""

import json
import os
import time
from datetime import datetime
from typing import Any, Callable, Dict, Optional, Tuple

import requests

try:
    from termcolor import colored
except ImportError:
    def colored(text, color=None, attrs=None):
        return text


SERVER_URL = os.getenv("MCP_SERVER_URL", "http://mcp-google-sheets.komyosys.ai/mcp")
USER_ID = os.getenv("MCP_USER_ID", "")
SPREADSHEET_ID = os.getenv("MCP_SPREADSHEET_ID", "muhammad_ahmad_accountio_journal_sheet")
BASE_SHEET = os.getenv("MCP_BASE_SHEET", "Ledger")

TIMEOUT_SEC = 20
DELAY_BETWEEN_TESTS = 0.5

HEADERS = {
    "Content-Type": "application/json",
    "Accept": "application/json, text/event-stream",
    "User-Agent": "MCP-Comprehensive-Test/2.0",
}


def _extract_tool_result(json_rpc_response: Dict[str, Any]) -> Any:
    result = json_rpc_response.get("result", {})
    if isinstance(result, dict) and "content" in result and isinstance(result["content"], list):
        for item in result["content"]:
            if isinstance(item, dict) and item.get("type") == "text":
                text = item.get("text", "")
                if isinstance(text, str):
                    try:
                        return json.loads(text)
                    except Exception:
                        return {"raw_text": text}
    return result


def send_jsonrpc(method: str, args: Dict[str, Any], rid: int) -> Tuple[bool, Dict[str, Any], Any, str]:
    payload = {
        "jsonrpc": "2.0",
        "method": "tools/call",
        "params": {"name": method, "arguments": args},
        "id": rid,
    }

    print(colored(f"\n→ TEST {rid} : {method}", "cyan", attrs=["bold"]))
    print(json.dumps(payload, indent=2, ensure_ascii=False))

    try:
        response = requests.post(SERVER_URL, headers=HEADERS, json=payload, timeout=TIMEOUT_SEC)
    except requests.Timeout:
        return False, {}, {}, "Request timed out"
    except requests.RequestException as exc:
        return False, {}, {}, f"Request failed: {exc}"

    print(f"← Status code : {response.status_code}")
    if response.status_code != 200:
        body = response.text[:800]
        return False, {}, {}, f"HTTP {response.status_code} - {response.reason} - {body}"

    try:
        data = response.json()
    except Exception:
        return False, {}, {}, "Response is not valid JSON"

    if "error" in data:
        return False, data, {}, f"JSON-RPC error: {json.dumps(data['error'], ensure_ascii=False)}"

    tool_result = _extract_tool_result(data)
    if isinstance(tool_result, dict) and tool_result.get("error"):
        return False, data, tool_result, f"Tool error: {tool_result.get('error')}"

    summary = "OK"
    if isinstance(tool_result, dict):
        if "success" in tool_result:
            summary = f"success={tool_result.get('success')}"
        elif "values" in tool_result:
            values = tool_result.get("values", [])
            summary = f"rows={len(values)}"
        elif "formulas" in tool_result:
            formulas = tool_result.get("formulas", [])
            summary = f"formula_rows={len(formulas)}"
        else:
            summary = f"keys={list(tool_result.keys())[:5]}"

    return True, data, tool_result, summary


def expect_true(condition: bool, message: str) -> Tuple[bool, str]:
    return condition, message


def run_test(
    rid: int,
    method: str,
    args: Dict[str, Any],
    expectation: str,
    validator: Optional[Callable[[Any], Tuple[bool, str]]] = None,
) -> bool:
    success, _, tool_result, summary = send_jsonrpc(method, args, rid)
    print(colored("Expectation:", "yellow"), expectation)
    print(colored("Result:", "yellow"), summary)

    if success and validator:
        valid, msg = validator(tool_result)
        success = success and valid
        if not valid:
            print(colored("Validator failed:", "red"), msg)

    if success:
        print(colored("→ PASSED", "green", attrs=["bold"]))
    else:
        print(colored("→ FAILED", "red", attrs=["bold"]))
        print("Detail:", tool_result)

    print("─" * 80)
    return success


def main() -> None:
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    test_sheet = f"MCP_Test_{now}"

    print(colored("Google Spreadsheet MCP - Comprehensive Test Suite", "cyan", attrs=["bold", "underline"]))
    print(f"Server        : {SERVER_URL}")
    print(f"User          : {USER_ID}")
    print(f"Spreadsheet   : {SPREADSHEET_ID}")
    print(f"Base Sheet    : {BASE_SHEET}")
    print(f"Test Sheet    : {test_sheet}")
    print("=" * 90)

    rid = 2000
    total = 0
    passed = 0

    tests = [
        {
            "name": "get_sheet_data",
            "args": {
                "spreadsheet_id": SPREADSHEET_ID,
                "sheet": BASE_SHEET,
                "range": "A1:C10",
                "include_grid_data": False,
                "user_id": USER_ID,
            },
            "expectation": "Returns values list with rowCount/columnCount",
            "validator": lambda r: expect_true(isinstance(r, dict) and "values" in r, "Missing values key"),
        },
        {
            "name": "get_sheet_formulas",
            "args": {
                "spreadsheet_id": SPREADSHEET_ID,
                "sheet": BASE_SHEET,
                "range": "A1:C10",
                "user_id": USER_ID,
            },
            "expectation": "Returns formulas and has_formulas flag",
            "validator": lambda r: expect_true(isinstance(r, dict) and "formulas" in r, "Missing formulas key"),
        },
        {
            "name": "create_new_sheet",
            "args": {
                "spreadsheet_id": SPREADSHEET_ID,
                "title": test_sheet,
                "user_id": USER_ID,
                "rows": 120,
                "cols": 12,
            },
            "expectation": "Creates a sheet and returns success=True",
            "validator": lambda r: expect_true(isinstance(r, dict) and r.get("success") is True, "Sheet creation failed"),
        },
        {
            "name": "update_single_cell",
            "args": {
                "spreadsheet_id": SPREADSHEET_ID,
                "sheet": test_sheet,
                "row": 1,
                "col": "A",
                "value": "Date",
                "user_id": USER_ID,
                "cell_name": False,
            },
            "expectation": "Classic mode updates a cell using row+column",
            "validator": lambda r: expect_true(isinstance(r, dict) and r.get("success") is True, "update_single_cell classic mode failed"),
        },
        {
            "name": "update_single_cell",
            "args": {
                "spreadsheet_id": SPREADSHEET_ID,
                "sheet": test_sheet,
                "row": 999,
                "col": "B1",
                "value": "Description",
                "user_id": USER_ID,
                "cell_name": True,
            },
            "expectation": "cell_name mode updates via A1 reference (B1)",
            "validator": lambda r: expect_true(
                isinstance(r, dict) and r.get("success") is True and r.get("updated", {}).get("row") == 1,
                "update_single_cell cell_name A1 mode failed",
            ),
        },
        {
            "name": "append_row",
            "args": {
                "spreadsheet_id": SPREADSHEET_ID,
                "sheet": test_sheet,
                "values_json": json.dumps(["2026-02-26", "Salary", "Income", 120000]),
                "user_id": USER_ID,
            },
            "expectation": "Appends a list row",
            "validator": lambda r: expect_true(isinstance(r, dict) and r.get("success") is True, "append_row list mode failed"),
        },
        {
            "name": "append_row",
            "args": {
                "spreadsheet_id": SPREADSHEET_ID,
                "sheet": test_sheet,
                "values_json": json.dumps({"Date": "2026-02-27", "Description": "Rent"}),
                "user_id": USER_ID,
                "header_row": 1,
            },
            "expectation": "Appends a dict row using header mapping",
            "validator": lambda r: expect_true(isinstance(r, dict) and r.get("success") is True, "append_row dict mode failed"),
        },
        {
            "name": "update_single_row",
            "args": {
                "spreadsheet_id": SPREADSHEET_ID,
                "sheet": test_sheet,
                "row": 3,
                "values_json": json.dumps(["2026-02-28", "Utilities", "Expense", 2500]),
                "user_id": USER_ID,
            },
            "expectation": "Updates row using list payload",
            "validator": lambda r: expect_true(isinstance(r, dict) and r.get("success") is True, "update_single_row list mode failed"),
        },
        {
            "name": "update_single_row",
            "args": {
                "spreadsheet_id": SPREADSHEET_ID,
                "sheet": test_sheet,
                "row": 4,
                "values_json": json.dumps({"Date": "2026-03-01", "Description": "Internet"}),
                "user_id": USER_ID,
                "header_row": 1,
            },
            "expectation": "Updates row using dict payload with headers",
            "validator": lambda r: expect_true(isinstance(r, dict) and r.get("success") is True, "update_single_row dict mode failed"),
        },
        {
            "name": "batch_update_cells",
            "args": {
                "spreadsheet_id": SPREADSHEET_ID,
                "sheet": test_sheet,
                "updates_json": json.dumps([
                    {"row": 5, "col": "A", "value": "2026-03-02"},
                    {"row": 5, "col": "B", "value": "Transport"},
                    {"row": 5, "col": "D", "value": 900},
                ]),
                "user_id": USER_ID,
                "cell_name": False,
            },
            "expectation": "Batch updates cells in classic mode",
            "validator": lambda r: expect_true(isinstance(r, dict) and r.get("success") is True, "batch_update_cells classic mode failed"),
        },
        {
            "name": "batch_update_cells",
            "args": {
                "spreadsheet_id": SPREADSHEET_ID,
                "sheet": test_sheet,
                "updates_json": json.dumps([
                    {"row": 6, "col": "Date", "value": "2026-03-03"},
                    {"row": 6, "col": "Description", "value": "Fuel"},
                ]),
                "user_id": USER_ID,
                "cell_name": True,
                "header_row": 1,
            },
            "expectation": "Batch updates cells in header-based cell_name mode",
            "validator": lambda r: expect_true(isinstance(r, dict) and r.get("success") is True, "batch_update_cells cell_name mode failed"),
        },
        {
            "name": "batch_update_rows",
            "args": {
                "spreadsheet_id": SPREADSHEET_ID,
                "sheet": test_sheet,
                "updates_json": json.dumps([
                    {"row": 7, "values": ["2026-03-04", "Groceries", "Expense", 1800]},
                    {"row": 8, "values": {"Date": "2026-03-05", "Description": "Bonus"}},
                ]),
                "user_id": USER_ID,
                "header_row": 1,
            },
            "expectation": "Batch updates rows with list and dict value forms",
            "validator": lambda r: expect_true(isinstance(r, dict) and r.get("success") is True, "batch_update_rows failed"),
        },
        {
            "name": "get_sheet_data",
            "args": {
                "spreadsheet_id": SPREADSHEET_ID,
                "sheet": test_sheet,
                "range": "A1:D10",
                "include_grid_data": True,
                "user_id": USER_ID,
            },
            "expectation": "Final read should include data and optional extra metadata",
            "validator": lambda r: expect_true(
                isinstance(r, dict) and "values" in r and "extra" in r,
                "Final get_sheet_data did not return expected keys",
            ),
        },
    ]

    for case in tests:
        total += 1
        ok = run_test(
            rid=rid,
            method=case["name"],
            args=case["args"],
            expectation=case["expectation"],
            validator=case.get("validator"),
        )
        if ok:
            passed += 1
        rid += 1
        time.sleep(DELAY_BETWEEN_TESTS)

    print("\n" + "═" * 90)
    print(colored("TEST SUMMARY", "magenta", attrs=["bold"]))
    print(f"Total tests : {total}")
    print(f"Passed      : {passed}")
    print(f"Failed      : {total - passed}")
    if passed == total:
        print(colored("ALL TESTS PASSED", "green", attrs=["bold", "underline"]))
    elif passed >= max(1, total // 2):
        print(colored("PARTIAL PASS - inspect failed cases", "yellow", attrs=["bold"]))
    else:
        print(colored("MOST TESTS FAILED - inspect server logs and credentials", "red", attrs=["bold"]))
    print("═" * 90)


if __name__ == "__main__":
    main()