"""
Google Spreadsheet MCP Server
A Model Context Protocol (MCP) server built with FastMCP for interacting with Google Sheets.
"""

import os
import httpx
from typing import Dict, Any, Optional, Union, List
import gspread
from termcolor import colored
from gspread.exceptions import SpreadsheetNotFound, WorksheetNotFound
from gspread.cell import Cell
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import json

# MCP imports
from fastmcp import FastMCP
from mcp.types import ToolAnnotations

# Environment / configuration
TOKEN_ENDPOINT = "https://api.accountio.ai/google/get_token"

# These should come from environment variables in production
GOOGLE_CLIENT_ID = os.environ.get("GOOGLE_CLIENT_ID")
GOOGLE_CLIENT_SECRET = os.environ.get("GOOGLE_CLIENT_SECRET")
HOST = os.environ.get("HOST") or "127.0.0.1"
PORT_STR = os.environ.get("PORT") or "8000"

if not GOOGLE_CLIENT_ID or not GOOGLE_CLIENT_SECRET:
    raise ValueError("GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET must be set in environment variables")

# Resolve host/port
# _resolved_host = os.environ.get('HOST') or os.environ.get('FASTMCP_HOST') or "0.0.0.0"
# _resolved_port_str = os.environ.get('PORT') or os.environ.get('FASTMCP_PORT') or "8000"
# try:
#     _resolved_port = int(_resolved_port_str)
# except ValueError:
#     _resolved_port = 8000

mcp = FastMCP(
    "Google Spreadsheet",
    # dependencies=["google-auth", "google-auth-oauthlib", "google-api-python-client", "httpx"],
    # host=_resolved_host,
    # port=_resolved_port
    
)


def _log_function_call(function_name: str, **kwargs: Any) -> None:
    if kwargs:
        args_preview = ", ".join(f"{k}={v}" for k, v in kwargs.items())
        print(colored(f"[CALL] {function_name}({args_preview})", "cyan"))
    else:
        print(colored(f"[CALL] {function_name}()", "cyan"))


def _log_function_status(function_name: str, success: bool, detail: Optional[str] = None) -> None:
    status = "SUCCESS" if success else "ERROR"
    color = "green" if success else "red"
    message = f"[{status}] {function_name}"
    if detail:
        message = f"{message} - {detail}"
    print(colored(message, color))


async def get_fresh_access_token(user_id: str) -> Dict[str, Any]:
    """
    Calls the internal /get-token endpoint to retrieve fresh Google credentials.
    """
    _log_function_call("get_fresh_access_token", user_id=user_id)
    print(colored(f"[get_fresh_access_token] Requesting token for user_id={user_id}", "yellow"))

    async with httpx.AsyncClient(timeout=15.0) as client:
        try:
            response = await client.get(f"{TOKEN_ENDPOINT}/{user_id}")
            response.raise_for_status()
            print(colored(f"[get_fresh_access_token] Token endpoint success for user_id={user_id}", "green"))
            _log_function_status("get_fresh_access_token", True)
            return response.json()
        except httpx.HTTPStatusError as e:
            print(colored(
                f"[get_fresh_access_token] HTTP error for user_id={user_id}: "
                f"status={e.response.status_code} body={e.response.text}"
            , "red"))
            _log_function_status("get_fresh_access_token", False, f"HTTP {e.response.status_code}")
            raise RuntimeError(
                f"Failed to get token for user {user_id} - status {e.response.status_code} - {e.response.text}"
            ) from e
        except Exception as e:
            print(colored(f"[get_fresh_access_token] Unexpected error for user_id={user_id}: {str(e)}", "red"))
            _log_function_status("get_fresh_access_token", False, str(e))
            raise RuntimeError(f"Error contacting token endpoint for user {user_id}: {str(e)}") from e


def _get_header_map(worksheet: gspread.Worksheet, header_row: int = 1) -> Dict[str, int]:
    """
    Helper to get mapping from column header (lowercased, stripped) to 1-based column index.
    """
    _log_function_call("_get_header_map", worksheet=worksheet.title, header_row=header_row)
    try:
        headers = worksheet.row_values(header_row)
        header_map = {h.strip().lower(): idx + 1 for idx, h in enumerate(headers) if h.strip()}
        _log_function_status("_get_header_map", True, f"headers={len(header_map)}")
        return header_map
    except Exception as exc:
        _log_function_status("_get_header_map", False, str(exc))
        raise


async def _get_authorized_gspread(user_id: str) -> gspread.Client:
    """
    Helper to get authorized gspread client.
    """
    _log_function_call("_get_authorized_gspread", user_id=user_id)
    print(colored(f"[_get_authorized_gspread] Building gspread client for user_id={user_id}", "yellow"))
    try:
        token_data = await get_fresh_access_token(user_id)
        access_token = token_data.get("access_token")
        if not access_token:
            print(colored(f"[_get_authorized_gspread] Missing access_token for user_id={user_id}", "red"))
            _log_function_status("_get_authorized_gspread", False, "No access token")
            raise ValueError("No access_token returned from token endpoint")

        creds = Credentials(
            token=access_token,
            client_id=GOOGLE_CLIENT_ID,
            client_secret=GOOGLE_CLIENT_SECRET,
            token_uri="https://oauth2.googleapis.com/token",
            scopes=[
                "https://www.googleapis.com/auth/drive.file",
            ],
        )

        if creds.expired:
            print(colored(f"[_get_authorized_gspread] Access token expired; refreshing for user_id={user_id}", "yellow"))
            creds.refresh(Request())

        print(colored(f"[_get_authorized_gspread] gspread client authorized for user_id={user_id}", "green"))
        _log_function_status("_get_authorized_gspread", True)
        return gspread.authorize(creds)
    except Exception as exc:
        _log_function_status("_get_authorized_gspread", False, str(exc))
        raise


async def _get_worksheet(
    user_id: str, spreadsheet_id: str, sheet: str
) -> gspread.Worksheet:
    """
    Helper to get spreadsheet and worksheet.
    """
    _log_function_call("_get_worksheet", user_id=user_id, spreadsheet_id=spreadsheet_id, sheet=sheet)
    print(colored(f"[_get_worksheet] Loading worksheet '{sheet}' from spreadsheet_id={spreadsheet_id}", "yellow"))
    try:
        gc = await _get_authorized_gspread(user_id)
        try:
            ss = gc.open(spreadsheet_id)
        except SpreadsheetNotFound:
            print(colored(f"[_get_worksheet] Spreadsheet not found: spreadsheet_id={spreadsheet_id}", "red"))
            _log_function_status("_get_worksheet", False, "Spreadsheet not found")
            raise ValueError("Spreadsheet not found or no access")
        try:
            worksheet = ss.worksheet(sheet)
        except WorksheetNotFound:
            print(colored(f"[_get_worksheet] Worksheet not found: sheet='{sheet}'", "red"))
            _log_function_status("_get_worksheet", False, "Worksheet not found")
            raise ValueError(f"Sheet '{sheet}' not found")
        print(colored(f"[_get_worksheet] Worksheet resolved successfully: sheet='{sheet}'", "green"))
        _log_function_status("_get_worksheet", True)
        return worksheet
    except Exception as exc:
        _log_function_status("_get_worksheet", False, str(exc))
        raise


@mcp.tool(
    annotations=ToolAnnotations(
        title="Get Sheet Data",
        readOnlyHint=True,
    ),
)
async def get_sheet_data(
    spreadsheet_id: str,
    sheet: str,
    user_id: str,
    range: Optional[str] = None,
    include_grid_data: bool = False,
) -> Dict[str, Any]:
    """
    Get data from a specific sheet in a Google Spreadsheet using user-specific OAuth credentials.
    Now powered by gspread for simpler range handling.

    Args:
        spreadsheet_id: The ID of the spreadsheet
        sheet: The name of the sheet/tab
        range: Optional cell range in A1 notation (e.g. 'A1:C10'). If omitted, fetches the entire sheet.
        include_grid_data: If True, attempts to return more metadata (but gspread is limited here)
        user_id: The identifier of the user whose Google credentials should be used

    Returns:
        Dictionary with spreadsheet data (values or more detailed if possible)
    """
    _log_function_call(
        "get_sheet_data",
        spreadsheet_id=spreadsheet_id,
        sheet=sheet,
        user_id=user_id,
        range=range,
        include_grid_data=include_grid_data,
    )
    # 1. Get fresh token from your backend
    try:
        token_data = await get_fresh_access_token(user_id)
    except Exception as exc:
        _log_function_status("get_sheet_data", False, "Failed to obtain Google access token")
        return {"error": "Failed to obtain Google access token", "detail": str(exc)}

    access_token = token_data.get("access_token")
    if not access_token:
        _log_function_status("get_sheet_data", False, "No access_token returned")
        return {"error": "No access_token returned from token endpoint"}

    # 2. Build credentials (same as before)
    creds = Credentials(
        token=access_token,
        client_id=GOOGLE_CLIENT_ID,
        client_secret=GOOGLE_CLIENT_SECRET,
        token_uri="https://oauth2.googleapis.com/token",
        scopes=[
            "https://www.googleapis.com/auth/drive.file",  # often enough for reading
        ],
    )

    if creds.expired:
        try:
            creds.refresh(Request())
        except Exception as exc:
            _log_function_status("get_sheet_data", False, "Token refresh failed")
            return {"error": "Token expired and refresh failed", "detail": str(exc)}

    # 3. Authorize gspread
    try:
        gc = gspread.authorize(creds)
        ss = gc.open(spreadsheet_id)
    except SpreadsheetNotFound:
        _log_function_status("get_sheet_data", False, "Spreadsheet not found")
        return {"error": "Spreadsheet not found or no access", "status": 404}
    except Exception as exc:
        _log_function_status("get_sheet_data", False, "Failed to open spreadsheet")
        return {"error": "Failed to open spreadsheet", "detail": str(exc)}

    # 4. Get the worksheet
    try:
        worksheet = ss.worksheet(sheet)
    except WorksheetNotFound:
        _log_function_status("get_sheet_data", False, "Sheet not found")
        return {"error": f"Sheet/tab '{sheet}' not found in the spreadsheet", "status": 404}

    # 5. Fetch the data
    try:
        if range:
            # Specific range
            values = worksheet.get(range, value_render_option="FORMATTED_VALUE")
        else:
            # Whole sheet
            values = worksheet.get_all_values()

        result = {
            "spreadsheetId": spreadsheet_id,
            "sheet": sheet,
            "range": range or "full sheet",
            "values": values,
            "rowCount": len(values),
            "columnCount": len(values[0]) if values else 0,
        }

        # Optional: if include_grid_data=True, we can try to get more info
        # (gspread doesn't give full grid data like the native API, but we can add basics)
        if include_grid_data:
            result["extra"] = {
                "title": worksheet.title,
                "url": ss.url,
                "updated": worksheet.updated if hasattr(worksheet, 'updated') else None,
                "row_count": worksheet.row_count,
                "col_count": worksheet.col_count,
            }

        _log_function_status("get_sheet_data", True)
        return result

    except Exception as exc:
        _log_function_status("get_sheet_data", False, "Failed to read sheet data")
        return {"error": "Failed to read sheet data", "detail": str(exc)}

@mcp.tool(
    annotations=ToolAnnotations(
        title="Get Sheet Formulas",
        readOnlyHint=True,
        description="Retrieve the raw formulas (not calculated values) from a range or entire sheet."
    ),
)
async def get_sheet_formulas(
    spreadsheet_id: str,
    sheet: str,
    user_id: str,
    range: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Get formulas from a specific sheet in a Google Spreadsheet using gspread.
    
    Args:
        spreadsheet_id: The ID of the spreadsheet
        sheet: The name of the sheet/tab
        user_id: The user identifier for token retrieval
        range: Optional cell range in A1 notation (e.g. 'B2:D100'). 
               If omitted, returns formulas from the entire sheet.
    
    Returns:
        Dictionary containing:
        - formulas: 2D list of formulas (empty string if cell has no formula)
        - range_used: the actual range that was queried
        - row_count, column_count
        - error (if any)
    """
    _log_function_call(
        "get_sheet_formulas",
        spreadsheet_id=spreadsheet_id,
        sheet=sheet,
        user_id=user_id,
        range=range,
    )
    # 1. Get fresh token
    try:
        token_data = await get_fresh_access_token(user_id)
    except Exception as exc:
        _log_function_status("get_sheet_formulas", False, "Failed to obtain Google access token")
        return {"error": "Failed to obtain Google access token", "detail": str(exc)}

    access_token = token_data.get("access_token")
    if not access_token:
        _log_function_status("get_sheet_formulas", False, "No access_token returned")
        return {"error": "No access_token returned from token endpoint"}

    # 2. Build credentials
    creds = Credentials(
        token=access_token,
        client_id=GOOGLE_CLIENT_ID,
        client_secret=GOOGLE_CLIENT_SECRET,
        token_uri="https://oauth2.googleapis.com/token",
        scopes=[
            "https://www.googleapis.com/auth/drive.file",
        ],
    )

    if creds.expired:
        try:
            creds.refresh(Request())
        except Exception as exc:
            _log_function_status("get_sheet_formulas", False, "Token refresh failed")
            return {"error": "Token expired and refresh failed", "detail": str(exc)}

    # 3. Authorize gspread
    try:
        gc = gspread.authorize(creds)
        ss = gc.open(spreadsheet_id)
    except SpreadsheetNotFound:
        _log_function_status("get_sheet_formulas", False, "Spreadsheet not found")
        return {"error": "Spreadsheet not found or no access", "status": 404}
    except Exception as exc:
        _log_function_status("get_sheet_formulas", False, "Failed to open spreadsheet")
        return {"error": "Failed to open spreadsheet", "detail": str(exc)}

    # 4. Get worksheet
    try:
        worksheet = ss.worksheet(sheet)
    except WorksheetNotFound:
        _log_function_status("get_sheet_formulas", False, "Sheet not found")
        return {"error": f"Sheet '{sheet}' not found in the spreadsheet", "status": 404}

    # 5. Determine range to query
    if range:
        target_range = f"{sheet}!{range}"
    else:
        # whole sheet — gspread uses 'A1' to last cell with data + some buffer
        # but to be safe we use worksheet.get_all_values() style range
        target_range = None  # we'll use get_all_values() + formula fetch separately

    try:
        if range:
            # Get formulas for specific range
            formulas = worksheet.get(
                range,
                value_render_option="FORMULA",          # ← this is the key
                date_time_render_option="FORMATTED_STRING"
            )
            used_range = range
        else:
            # Get formulas for entire sheet
            # gspread .get() with no range → whole sheet
            formulas = worksheet.get(
                value_render_option="FORMULA"
            )
            used_range = worksheet.get_all_values()  # just for info
            used_range = f"A1:{gspread.utils.rowcol_to_a1(worksheet.row_count, worksheet.col_count)}"

        result = {
            "spreadsheetId": spreadsheet_id,
            "sheet": sheet,
            "range_used": used_range,
            "formulas": formulas,
            "rowCount": len(formulas),
            "columnCount": len(formulas[0]) if formulas and formulas[0] else 0,
            "has_formulas": any(
                any(cell.strip().startswith('=') for cell in row)
                for row in formulas
            ) if formulas else False
        }

        _log_function_status("get_sheet_formulas", True)
        return result

    except Exception as exc:
        _log_function_status("get_sheet_formulas", False, "Failed to retrieve formulas")
        return {
            "error": "Failed to retrieve formulas",
            "detail": str(exc),
            "status": getattr(exc, 'resp', {}).get('status', None) if hasattr(exc, 'resp') else None
        }


@mcp.tool(
    annotations=ToolAnnotations(
        title="Create New Sheet",
        readOnlyHint=False,
        description="Create a new sheet (tab) in the spreadsheet."
    ),
)
async def create_new_sheet(
    spreadsheet_id: str,
    title: str,
    user_id: str,
    rows: int = 100,
    cols: int = 26,
) -> Dict[str, Any]:
    """
    Create a new worksheet in the spreadsheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet
        title: The name for the new sheet
        user_id: The user identifier
        rows: Number of rows for the new sheet (default 100)
        cols: Number of columns for the new sheet (default 26)

    Returns:
        Dictionary with new sheet info
    """
    _log_function_call(
        "create_new_sheet",
        spreadsheet_id=spreadsheet_id,
        title=title,
        user_id=user_id,
        rows=rows,
        cols=cols,
    )
    try:
        gc = await _get_authorized_gspread(user_id)
        ss = gc.open(spreadsheet_id)
        worksheet = ss.add_worksheet(title=title, rows=rows, cols=cols)
        _log_function_status("create_new_sheet", True)
        return {
            "success": True,
            "sheet": worksheet.title,
            "id": worksheet.id,
            "url": worksheet.url,
            "row_count": worksheet.row_count,
            "col_count": worksheet.col_count,
        }
    except Exception as exc:
        _log_function_status("create_new_sheet", False, "Failed to create new sheet")
        return {"error": "Failed to create new sheet", "detail": str(exc)}


@mcp.tool(
    annotations=ToolAnnotations(
        title="Update Single Cell",
        readOnlyHint=False,
        description="Update a single cell by row number and column (number or header name)."
    ),
)
async def update_single_cell(
    spreadsheet_id: str,
    sheet: str,
    row: int,
    col: Union[int, str],
    value: Any,
    user_id: str,
    header_row: int = 1,
    cell_name:bool=False
) -> Dict[str, Any]:
    """
    Update a single cell in the sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet
        sheet: The name of the sheet
        row: 1-based row number
        col: Column as 1-based int or str header name
        value: The value to set (str, int, float, etc.)
        user_id: The user identifier
        header_row: Row number with headers (if using col as str)

    Returns:
        Success dict or error
    """
    _log_function_call(
        "update_single_cell",
        spreadsheet_id=spreadsheet_id,
        sheet=sheet,
        row=row,
        col=col,
        user_id=user_id,
        header_row=header_row,
        cell_name=cell_name,
    )
    try:
        worksheet = await _get_worksheet(user_id, spreadsheet_id, sheet)
        mode = "row + column/header"

        if cell_name:
            mode = "cell_name"
            if not isinstance(col, str) or not col.strip():
                _log_function_status("update_single_cell", False, "In cell_name mode, col must be a non-empty string")
                return {"error": "In cell_name mode, col must be a non-empty string"}

            col_input = col.strip()

            try:
                parsed_row, parsed_col = gspread.utils.a1_to_rowcol(col_input)
                row_num = parsed_row
                col_num = parsed_col
                mode = "A1 cell reference"
            except Exception:
                if row is None or row < 1:
                    _log_function_status("update_single_cell", False, "Row must be >= 1 when using header-based cell_name mode")
                    return {"error": "Row must be >= 1 when using header-based cell_name mode"}
                header_map = _get_header_map(worksheet, header_row)
                col_num = header_map.get(col_input.lower())
                if col_num is None:
                    _log_function_status("update_single_cell", False, f"Column header '{col_input}' not found")
                    return {"error": f"Column header '{col_input}' not found"}
                row_num = row
                mode = "header + row"
        else:
            if row is None or row < 1:
                _log_function_status("update_single_cell", False, "Row must be >= 1")
                return {"error": "Row must be >= 1"}

            row_num = row

            if isinstance(col, int):
                col_num = col
            elif isinstance(col, str):
                col_str = col.strip().upper()
                if not col_str:
                    _log_function_status("update_single_cell", False, "Column string cannot be empty")
                    return {"error": "Column string cannot be empty"}

                if col_str.isalpha():
                    col_num = 0
                    for c in col_str:
                        col_num = col_num * 26 + (ord(c) - ord('A') + 1)
                else:
                    header_map = _get_header_map(worksheet, header_row)
                    col_num = header_map.get(col.strip().lower())
                    if col_num is None:
                        _log_function_status("update_single_cell", False, f"Column header '{col}' not found")
                        return {"error": f"Column header '{col}' not found"}
            else:
                _log_function_status("update_single_cell", False, "col must be int or str")
                return {"error": "col must be int or str"}

        if col_num < 1:
            _log_function_status("update_single_cell", False, "Column must be >= 1")
            return {"error": "Column must be >= 1"}

        worksheet.update_cell(row_num, col_num, value)
        _log_function_status("update_single_cell", True)
        return {
            "success": True,
            "updated": {"row": row_num, "col": col_num, "value": value},
            "mode": mode,
        }
    except Exception as exc:
        _log_function_status("update_single_cell", False, "Failed to update cell")
        return {"error": "Failed to update cell", "detail": str(exc)}


@mcp.tool(
    annotations=ToolAnnotations(
        title="Update Single Row",
        readOnlyHint=False,
        description="Update a single row by row number and values as list or dict (header to value)."
    ),
)
async def update_single_row(
    spreadsheet_id: str,
    sheet: str,
    row: int,
    values_json: str,  # Take as JSON string to handle list/dict
    user_id: str,
    header_row: int = 1,
) -> Dict[str, Any]:
    """
    Update a single row in the sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet
        sheet: The name of the sheet
        row: 1-based row number to update
        values_json: JSON string of list of values or dict of {header: value}
        user_id: The user identifier
        header_row: Row number with headers (if using dict)

    Returns:
        Success dict or error
    """
    _log_function_call(
        "update_single_row",
        spreadsheet_id=spreadsheet_id,
        sheet=sheet,
        row=row,
        user_id=user_id,
        header_row=header_row,
    )
    try:
        values = json.loads(values_json)
    except json.JSONDecodeError:
        _log_function_status("update_single_row", False, "Invalid JSON for values")
        return {"error": "Invalid JSON for values"}

    try:
        worksheet = await _get_worksheet(user_id, spreadsheet_id, sheet)
        cells: List[Cell] = []
        if isinstance(values, list):
            for idx, val in enumerate(values):
                cells.append(Cell(row, idx + 1, val))
        elif isinstance(values, dict):
            header_map = _get_header_map(worksheet, header_row)
            for header, val in values.items():
                col = header_map.get(header.strip().lower())
                if col:
                    cells.append(Cell(row, col, val))
                else:
                    _log_function_status("update_single_row", False, f"Column header '{header}' not found")
                    return {"error": f"Column header '{header}' not found"}
        else:
            _log_function_status("update_single_row", False, "values must be list or dict")
            return {"error": "values must be list or dict"}

        if cells:
            worksheet.update_cells(cells)
        _log_function_status("update_single_row", True)
        return {"success": True, "updated_row": row, "cell_count": len(cells)}
    except Exception as exc:
        _log_function_status("update_single_row", False, "Failed to update row")
        return {"error": "Failed to update row", "detail": str(exc)}


@mcp.tool(
    annotations=ToolAnnotations(
        title="Append Row",
        readOnlyHint=False,
        description="Append a new row at the end of the sheet with values as list or dict."
    ),
)
async def append_row(
    spreadsheet_id: str,
    sheet: str,
    values_json: str,  # JSON string for list/dict
    user_id: str,
    header_row: int = 1,
) -> Dict[str, Any]:
    """
    Append a row to the sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet
        sheet: The name of the sheet
        values_json: JSON string of list of values or dict of {header: value}
        user_id: The user identifier
        header_row: Row number with headers (if using dict)

    Returns:
        Success dict with new row number or error
    """
    _log_function_call(
        "append_row",
        spreadsheet_id=spreadsheet_id,
        sheet=sheet,
        user_id=user_id,
        header_row=header_row,
    )
    try:
        values = json.loads(values_json)
    except json.JSONDecodeError:
        _log_function_status("append_row", False, "Invalid JSON for values")
        return {"error": "Invalid JSON for values"}

    try:
        worksheet = await _get_worksheet(user_id, spreadsheet_id, sheet)
        if isinstance(values, dict):
            header_map = _get_header_map(worksheet, header_row)
            if not header_map:
                _log_function_status("append_row", False, "No headers found")
                return {"error": "No headers found"}
            max_col = max(header_map.values())
            row_list = [''] * max_col
            for header, val in values.items():
                col = header_map.get(header.strip().lower())
                if col:
                    row_list[col - 1] = val
            values = row_list
        elif not isinstance(values, list):
            _log_function_status("append_row", False, "values must be list or dict")
            return {"error": "values must be list or dict"}

        new_row = worksheet.append_row(values)
        _log_function_status("append_row", True)
        return {"success": True, "appended_row": new_row.get('updates', {}).get('updatedRange', '')}
    except Exception as exc:
        _log_function_status("append_row", False, "Failed to append row")
        return {"error": "Failed to append row", "detail": str(exc)}


@mcp.tool(
  annotations=ToolAnnotations(
        title="Batch Update Cells",
        readOnlyHint=False,
        description="Batch update multiple cells. Supports both header-based and A1-style cell references."
    ),
)
async def batch_update_cells(
    spreadsheet_id: str,
    sheet: str,
    updates_json: str,
    user_id: str,
    header_row: int = 1,
    cell_name: bool = False,          # ← new flag
) -> Dict[str, Any]:
    """
    Batch update individual cells.

    Args:
        spreadsheet_id: The ID of the spreadsheet
        sheet: The name of the sheet/tab
        updates_json: JSON list of dicts, each with:
            - When cell_name=False: {'row': int, 'col': int|str, 'value': Any}
            - When cell_name=True:  {'cell': str (A1-style), 'value': Any}
        user_id: The user identifier
        header_row: Row number containing headers (used when cell_name=False and col is string header)
        cell_name: If True, expect 'cell' key with A1 notation (e.g. "B5") instead of separate row+col

    Returns:
        {"success": bool, "updated_cells": int} or {"error": str, ...}
    """
    _log_function_call(
        "batch_update_cells",
        spreadsheet_id=spreadsheet_id,
        sheet=sheet,
        user_id=user_id,
        header_row=header_row,
        cell_name=cell_name,
    )

    try:
        updates = json.loads(updates_json)
        if not isinstance(updates, list):
            _log_function_status("batch_update_cells", False, "updates must be a list")
            return {"error": "updates must be a list"}
    except json.JSONDecodeError:
        _log_function_status("batch_update_cells", False, "Invalid JSON for updates")
        return {"error": "Invalid JSON for updates"}

    try:
        worksheet = await _get_worksheet(user_id, spreadsheet_id, sheet)

        header_map = {}
        if cell_name:
            # Only build header map if we're using header names
            header_map = _get_header_map(worksheet, header_row)

        print(colored(f"[batch_update_cells] cell_name mode: {cell_name}", "yellow"))
        print(colored(f"[batch_update_cells] Header map: {header_map}", "yellow"))

        cells: List[Cell] = []
        skipped = 0

        for update in updates:
            value = update.get('value')
            row=update.get('row')
            col_input=update.get('col')
            if value is None:
                skipped += 1
                continue

            if cell_name:
                
                if not isinstance(col_input, str) or not col_input.strip():
                  skipped += 1
                  continue
                col_num = header_map.get(col_input.strip().lower())
                if col_num is None:
                    _log_function_status("batch_update_cells", False, f"Column '{col_input}' not found")
                    return {"error": f"Column '{col_input}' not found"}

                try:
                    cells.append(Cell(row, col_num, value))
                except Exception as e:
                    print(colored(f"[batch_update_cells] Invalid cell ref '{col_input}': {e}", "red"))
                    skipped += 1
                    continue

            else:
                # ── Classic row + col mode ──
                col_input = update.get('col')

                if row is None or col_input is None:
                    skipped += 1
                    continue

                # Parse column
                if isinstance(col_input, int):
                    col_num = col_input
                elif isinstance(col_input, str):
                    col_str = col_input.strip().upper()

                    # Try as column letter (A → 1, B → 2, AA → 27, ...)
                    try:
                        col_num = 0
                        for c in col_str:
                            if not 'A' <= c <= 'Z':
                                raise ValueError
                            col_num = col_num * 26 + (ord(c) - ord('A') + 1)
                    except:
                        # Then try as header name
                        col_num = header_map.get(col_str.lower())
                        if col_num is None:
                            print(colored(f"[batch_update_cells] Column '{col_input}' not found", "red"))
                            return {"error": f"Column '{col_input}' not found (not a valid header or letter)"}
                else:
                    skipped += 1
                    continue

                cells.append(Cell(row, col_num, value))

        if not cells:
            msg = f"No valid cells to update (skipped {skipped})"
            _log_function_status("batch_update_cells", False, msg)
            return {"success": False, "message": msg, "updated_cells": 0}

        worksheet.update_cells(cells)

        msg = f"Updated {len(cells)} cells (skipped {skipped})"
        _log_function_status("batch_update_cells", True, msg)
        return {
            "success": True,
            "updated_cells": len(cells),
            "skipped": skipped,
            "mode": "Header Row" if cell_name else "row + column/header"
        }

    except gspread.exceptions.APIError as api_err:
        detail = api_err.response.text if hasattr(api_err, 'response') else str(api_err)
        _log_function_status("batch_update_cells", False, "Google API error")
        return {"error": "Google Sheets API error", "detail": detail}
    except Exception as exc:
        _log_function_status("batch_update_cells", False, str(exc))
        return {"error": "Failed to batch update cells", "detail": str(exc)}

@mcp.tool(
    annotations=ToolAnnotations(
        title="Batch Update Rows",
        readOnlyHint=False,
        description="Batch update multiple rows."
    ),
)
async def batch_update_rows(
    spreadsheet_id: str,
    sheet: str,
    updates_json: str,  # JSON list of dicts: [{'row':int, 'values': list or dict}]
    user_id: str,
    header_row: int = 1,
) -> Dict[str, Any]:
    """
    Batch update multiple rows.

    Args:
        spreadsheet_id: The ID of the spreadsheet
        sheet: The name of the sheet
        updates_json: JSON list of row updates, each {'row': int, 'values': JSON list or dict}
        user_id: The user identifier
        header_row: Row with headers if values is dict

    Returns:
        Success dict or error
    """
    _log_function_call(
        "batch_update_rows",
        spreadsheet_id=spreadsheet_id,
        sheet=sheet,
        user_id=user_id,
        header_row=header_row,
    )
    try:
        updates = json.loads(updates_json)
        if not isinstance(updates, list):
            _log_function_status("batch_update_rows", False, "updates must be a list")
            return {"error": "updates must be a list"}
    except json.JSONDecodeError:
        _log_function_status("batch_update_rows", False, "Invalid JSON for updates")
        return {"error": "Invalid JSON for updates"}

    try:
        worksheet = await _get_worksheet(user_id, spreadsheet_id, sheet)
        cells: List[Cell] = []
        header_map = _get_header_map(worksheet, header_row)
        for update in updates:
            row = update.get('row')
            values = update.get('values')
            if not row or not values:
                continue
            if isinstance(values, str):
                try:
                    values = json.loads(values)
                except:
                    continue
            if isinstance(values, list):
                for idx, val in enumerate(values):
                    cells.append(Cell(row, idx + 1, val))
            elif isinstance(values, dict):
                for header, val in values.items():
                    col = header_map.get(header.strip().lower())
                    if col:
                        cells.append(Cell(row, col, val))
            else:
                continue
        if cells:
            worksheet.update_cells(cells)
        _log_function_status("batch_update_rows", True)
        return {"success": True, "updated_cells": len(cells)}
    except Exception as exc:
        _log_function_status("batch_update_rows", False, "Failed to batch update rows")
        return {"error": "Failed to batch update rows", "detail": str(exc)}

if __name__ == "__main__":
    import sys
    transport = "stdio"
    for i, arg in enumerate(sys.argv):
        if arg == "--transport" and i + 1 < len(sys.argv):
            transport = sys.argv[i + 1]
            break
    print(colored(f"[startup] Running transport={transport}", "magenta"))
    print(colored("[startup] Launching FastMCP server on http://127.0.0.1:8000", "magenta"))
    mcp.run(transport="http", host=HOST, port=int(PORT_STR), json_response=True,   stateless_http=True
)
