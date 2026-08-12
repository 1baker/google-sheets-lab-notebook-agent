from __future__ import annotations

import json
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from .agent import suggestion_to_workbook_row
from .material_scaffold import formulation_key
from .planning import result_row_key
from .schema import (
    CONTROLLED_VOCAB_VALIDATIONS,
    RUN_CONSOLE_SHEET,
    SHEETS,
    WORKBOOK_CONTRACT_VERSION,
    column_data_type,
    column_number_format,
    sheet_by_name,
)
from .sheets import rows_from_values

GENERATED_SHEET_ID_START = 900_000_000
DEFAULT_VALIDATION_END_ROW = 1000
DEFAULT_WORKBOOK_TIMEZONE = "America/Chicago"
PLOT_DASHBOARD_MIN_COLUMNS = 28
RUN_CONSOLE_COLUMN_COUNT = 12
RUN_CONSOLE_ROW_COUNT = 100
TECHNICAL_SHEETS = frozenset(
    {
        "Daily Reviews",
        "Project Notebook Records",
        "Source Sync",
        "Plot Data",
        "Plot Definitions",
        "Process Knowledge",
        "Controlled Vocab",
        "Agent Config",
        "Workbook Metadata",
        "Audit Log",
    }
)
CORE_ENTRY_SHEETS = frozenset(
    {
        "Experiments",
        "Daily Log",
        "Formulations",
        "Results",
        "Run Capture Plan",
        "Samples",
        "Deviations",
        "Raw Data Files",
    }
)
REFERENCE_SHEETS = frozenset(
    {"Master Reagents", "Equipment", "Protocols", "Specifications"}
)
OPTIONAL_EXTENSION_SHEETS = {
    "Project Notebook Records",
    "Source Sync",
    "Plot Data",
    "Plot Definitions",
    "Plot Dashboard",
    "Workbook Metadata",
    "Run Capture Plan",
    "Samples",
    "Equipment",
    "Protocols",
    "Specifications",
    "Deviations",
    "Raw Data Files",
    "Audit Log",
}


def contract_sheet_names() -> tuple[str, ...]:
    """Return every physical contract sheet, including interface-only views."""

    return (RUN_CONSOLE_SHEET, *(spec.name for spec in SHEETS))


def load_agent_report(path: str | Path) -> dict[str, Any]:
    with Path(path).expanduser().open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, dict):
        raise ValueError("Agent report must be a JSON object.")
    return data


def load_sheet_snapshot(path: str | Path) -> dict[str, Any]:
    with Path(path).expanduser().open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, dict):
        raise ValueError("Google Sheets snapshot must be a JSON object.")
    return data


def snapshot_to_tables(snapshot: dict[str, Any]) -> dict[str, list[dict[str, Any]]]:
    sheets = snapshot.get("sheets")
    if not isinstance(sheets, dict):
        raise ValueError("Snapshot must include a sheets object.")
    tables: dict[str, list[dict[str, Any]]] = {}
    for sheet_name, payload in sheets.items():
        values = sheet_values(payload)
        tables[sheet_name] = rows_from_values(values)
    return tables


def sheet_ids_from_snapshot(snapshot: dict[str, Any]) -> dict[str, int]:
    sheets = snapshot.get("sheets")
    if not isinstance(sheets, dict):
        raise ValueError("Snapshot must include a sheets object.")
    ids: dict[str, int] = {}
    for sheet_name, payload in sheets.items():
        if isinstance(payload, dict) and payload.get("sheet_id") is not None:
            ids[str(sheet_name)] = int(payload["sheet_id"])
    return ids


def sheet_ids_from_metadata_payload(metadata: dict[str, Any]) -> dict[str, int]:
    ids: dict[str, int] = {}
    for sheet in metadata.get("sheets", []) or []:
        if not isinstance(sheet, dict):
            continue
        properties = sheet.get("properties", {})
        if not isinstance(properties, dict):
            continue
        title = properties.get("title")
        sheet_id = properties.get("sheetId")
        if title is not None and sheet_id is not None:
            ids[str(title)] = int(sheet_id)
    return ids


def sheet_properties_from_metadata_payload(metadata: dict[str, Any]) -> dict[str, dict[str, Any]]:
    properties_by_title: dict[str, dict[str, Any]] = {}
    for sheet in metadata.get("sheets", []) or []:
        if not isinstance(sheet, dict):
            continue
        properties = sheet.get("properties", {})
        if not isinstance(properties, dict):
            continue
        title = properties.get("title")
        if title is not None:
            properties_by_title[str(title)] = properties
    return properties_by_title


def sheet_values(payload: Any) -> list[list[Any]]:
    if isinstance(payload, dict):
        values = payload.get("values", [])
    else:
        values = payload
    if not isinstance(values, list):
        raise ValueError("Snapshot sheet values must be a list of rows.")
    return [list(row) if isinstance(row, list) else [row] for row in values]


def snapshot_from_tables(
    tables: dict[str, list[dict[str, Any]]],
    sheet_ids: dict[str, int] | None = None,
) -> dict[str, Any]:
    sheet_ids = sheet_ids or {}
    sheets: dict[str, Any] = {}
    for sheet_name, rows in tables.items():
        headers = list(sheet_by_name(sheet_name).headers)
        values = [headers]
        values.extend([[row.get(header, "") for header in headers] for row in rows])
        sheets[sheet_name] = {
            "sheet_id": sheet_ids.get(sheet_name),
            "values": values,
        }
    return {
        "schema": "lab-notebook-agent-google-sheets-snapshot.v1",
        "sheets": sheets,
    }


def snapshot_capture_plan(
    spreadsheet_id: str = "",
    value_range: str = "A1:Z1000",
) -> dict[str, Any]:
    return {
        "schema": "lab-notebook-agent-google-sheets-capture-plan.v1",
        "spreadsheet_id": spreadsheet_id,
        "value_render_option": "FORMATTED_VALUE",
        "sheets": [
            {
                "sheet_name": spec.name,
                "range": value_range,
                "headers": list(spec.headers),
                "used_for_apply": spec.name in {
                    "Master Reagents",
                    "Experiments",
                    "Formulations",
                    "Results",
                    "Daily Log",
                    "Literature Evidence",
                    "Agent Suggestions",
                    "Daily Reviews",
                },
            }
            for spec in SHEETS
        ],
    }


def google_setup_audit_from_metadata(
    metadata: dict[str, Any],
    include_validations: bool = True,
    validation_end_row: int = DEFAULT_VALIDATION_END_ROW,
) -> dict[str, Any]:
    existing_sheet_ids = sheet_ids_from_metadata_payload(metadata)
    generated_sheet_ids = generated_sheet_ids_for_missing(existing_sheet_ids)
    requests = google_setup_requests_from_metadata(
        metadata,
        include_validations=include_validations,
        validation_end_row=validation_end_row,
    )
    expected_sheet_names = contract_sheet_names()
    contract_sheet_name_set = set(expected_sheet_names)
    existing_contract_sheets = [
        name for name in expected_sheet_names if name in existing_sheet_ids
    ]
    missing_sheets = [
        name for name in expected_sheet_names if name not in existing_sheet_ids
    ]
    unknown_sheets = sorted(set(existing_sheet_ids) - contract_sheet_name_set)
    validation_rule_count = (
        sum(len(fields) for fields in CONTROLLED_VOCAB_VALIDATIONS.values()) if include_validations else 0
    )
    return {
        "schema": "lab-notebook-agent-google-setup-audit.v1",
        "ready_after_apply": True,
        "existing_contract_sheets": existing_contract_sheets,
        "missing_sheets_to_create": missing_sheets,
        "unknown_sheets": unknown_sheets,
        "generated_sheet_ids": generated_sheet_ids,
        "summary": {
            "contract_sheet_count": len(expected_sheet_names),
            "existing_contract_sheet_count": len(existing_contract_sheets),
            "missing_sheet_count": len(missing_sheets),
            "validation_rule_count": validation_rule_count,
            "request_count": len(requests),
        },
    }


def google_setup_requests_from_metadata(
    metadata: dict[str, Any],
    include_validations: bool = True,
    validation_end_row: int = DEFAULT_VALIDATION_END_ROW,
    time_zone: str = DEFAULT_WORKBOOK_TIMEZONE,
) -> list[dict[str, Any]]:
    validation_end_row = max(2, validation_end_row)
    existing_sheet_ids = sheet_ids_from_metadata_payload(metadata)
    properties_by_title = sheet_properties_from_metadata_payload(metadata)
    generated_sheet_ids = generated_sheet_ids_for_missing(existing_sheet_ids)
    sheet_ids = {**existing_sheet_ids, **generated_sheet_ids}
    requests: list[dict[str, Any]] = []
    current_time_zone = str((metadata.get("properties") or {}).get("timeZone", ""))
    if time_zone and current_time_zone != time_zone:
        requests.append(
            {
                "updateSpreadsheetProperties": {
                    "properties": {"timeZone": time_zone},
                    "fields": "timeZone",
                }
            }
        )
    run_console_id = sheet_ids[RUN_CONSOLE_SHEET]
    run_console_is_new = RUN_CONSOLE_SHEET not in existing_sheet_ids
    if run_console_is_new:
        requests.append(
            add_sheet_request(
                RUN_CONSOLE_SHEET,
                run_console_id,
                RUN_CONSOLE_COLUMN_COUNT,
                RUN_CONSOLE_ROW_COUNT,
                index=0,
            )
        )
    requests.extend(
        run_console_setup_requests(
            run_console_id,
            sheet_ids,
            is_new=run_console_is_new,
        )
    )
    for spec in SHEETS:
        sheet_id = sheet_ids[spec.name]
        grid_column_count = minimum_grid_column_count(spec.name)
        if spec.name not in existing_sheet_ids:
            requests.append(
                add_sheet_request(
                    spec.name,
                    sheet_id,
                    grid_column_count,
                    validation_end_row,
                )
            )
        requests.append(
            sheet_grid_setup_request(
                spec.name,
                sheet_id,
                grid_column_count,
                properties_by_title.get(spec.name, {}),
                validation_end_row,
            )
        )
        requests.append(header_update_request(spec.name, sheet_ids))
        requests.append(header_format_request(spec.name, sheet_ids))
        requests.extend(column_number_format_requests(spec.name, sheet_ids, validation_end_row))
        requests.append(basic_filter_request(spec.name, sheet_ids, validation_end_row))
        requests.extend(column_width_requests(spec.name, sheet_ids))
        requests.append(header_row_height_request(spec.name, sheet_ids))
        if include_validations:
            for field, allowed_values in CONTROLLED_VOCAB_VALIDATIONS.get(spec.name, {}).items():
                requests.append(
                    data_validation_request(
                        spec.name,
                        field,
                        allowed_values,
                        sheet_ids,
                        validation_end_row=validation_end_row,
                    )
                )
    return requests


def generated_sheet_ids_for_missing(existing_sheet_ids: dict[str, int]) -> dict[str, int]:
    used_ids = set(existing_sheet_ids.values())
    generated: dict[str, int] = {}
    next_sheet_id = GENERATED_SHEET_ID_START
    for sheet_name in contract_sheet_names():
        if sheet_name in existing_sheet_ids:
            continue
        while next_sheet_id in used_ids:
            next_sheet_id += 1
        generated[sheet_name] = next_sheet_id
        used_ids.add(next_sheet_id)
        next_sheet_id += 1
    return generated


def minimum_grid_column_count(sheet_name: str) -> int:
    if sheet_name == RUN_CONSOLE_SHEET:
        return RUN_CONSOLE_COLUMN_COUNT
    header_count = len(sheet_by_name(sheet_name).headers)
    if sheet_name == "Plot Dashboard":
        return max(header_count, PLOT_DASHBOARD_MIN_COLUMNS)
    return header_count


def add_sheet_request(
    sheet_name: str,
    sheet_id: int,
    column_count: int,
    row_count: int,
    *,
    index: int | None = None,
) -> dict[str, Any]:
    properties: dict[str, Any] = {
        "sheetId": sheet_id,
        "title": sheet_name,
        "gridProperties": {
            "rowCount": row_count,
            "columnCount": column_count,
            "frozenRowCount": 3 if sheet_name == RUN_CONSOLE_SHEET else 1,
            "hideGridlines": sheet_name in {RUN_CONSOLE_SHEET, "Plot Dashboard"},
        },
    }
    if index is not None:
        properties["index"] = index
    return {
        "addSheet": {
            "properties": properties
        }
    }


def sheet_grid_setup_request(
    sheet_name: str,
    sheet_id: int,
    column_count: int,
    properties: dict[str, Any],
    validation_end_row: int,
) -> dict[str, Any]:
    grid = properties.get("gridProperties", {})
    if not isinstance(grid, dict):
        grid = {}
    grid_properties: dict[str, Any] = {
        "frozenRowCount": 1,
        "frozenColumnCount": frozen_column_count(sheet_name),
        "hideGridlines": sheet_name == "Plot Dashboard",
    }
    fields = [
        "gridProperties.frozenRowCount",
        "gridProperties.frozenColumnCount",
        "gridProperties.hideGridlines",
        "hidden",
        "tabColorStyle",
    ]
    row_count = optional_int(grid.get("rowCount"))
    if row_count is not None and row_count < validation_end_row:
        grid_properties["rowCount"] = validation_end_row
        fields.append("gridProperties.rowCount")
    existing_column_count = optional_int(grid.get("columnCount"))
    if existing_column_count is not None and existing_column_count < column_count:
        grid_properties["columnCount"] = column_count
        fields.append("gridProperties.columnCount")
    return {
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "title": sheet_name,
                "hidden": sheet_name in TECHNICAL_SHEETS,
                "tabColorStyle": {"rgbColor": tab_color(sheet_name)},
                "gridProperties": grid_properties,
            },
            "fields": ",".join(fields),
        }
    }


def header_update_request(sheet_name: str, sheet_ids: dict[str, int]) -> dict[str, Any]:
    if sheet_name not in sheet_ids:
        raise KeyError(f"Missing sheet ID for {sheet_name!r}.")
    spec = sheet_by_name(sheet_name)
    return {
        "updateCells": {
            "start": {
                "sheetId": sheet_ids[sheet_name],
                "rowIndex": 0,
                "columnIndex": 0,
            },
            "rows": [
                {
                    "values": [
                        {
                            "userEnteredValue": {"stringValue": column.name},
                            "note": (
                                f"{column.description}"
                                + (" Required." if column.required else "")
                            ),
                        }
                        for column in spec.columns
                    ]
                }
            ],
            "fields": "userEnteredValue,note",
        }
    }


def header_format_request(sheet_name: str, sheet_ids: dict[str, int]) -> dict[str, Any]:
    if sheet_name not in sheet_ids:
        raise KeyError(f"Missing sheet ID for {sheet_name!r}.")
    headers = list(sheet_by_name(sheet_name).headers)
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_ids[sheet_name],
                "startRowIndex": 0,
                "endRowIndex": 1,
                "startColumnIndex": 0,
                "endColumnIndex": len(headers),
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {"red": 0.12, "green": 0.31, "blue": 0.47},
                    "textFormat": {
                        "bold": True,
                        "foregroundColor": {"red": 1, "green": 1, "blue": 1},
                    },
                    "wrapStrategy": "WRAP",
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,wrapStrategy)",
        }
    }


def frozen_column_count(sheet_name: str) -> int:
    if sheet_name in {
        "Daily Log",
        "Formulations",
        "Results",
        "Run Capture Plan",
        "Samples",
        "Deviations",
        "Raw Data Files",
        "Project Notebook Records",
    }:
        return 2
    return 1


def tab_color(sheet_name: str) -> dict[str, float]:
    if sheet_name == RUN_CONSOLE_SHEET:
        return {"red": 0.07, "green": 0.18, "blue": 0.29}
    if sheet_name in CORE_ENTRY_SHEETS:
        return {"red": 0.10, "green": 0.48, "blue": 0.55}
    if sheet_name in REFERENCE_SHEETS:
        return {"red": 0.31, "green": 0.56, "blue": 0.36}
    if sheet_name in {"Plot Dashboard", "Literature Evidence"}:
        return {"red": 0.40, "green": 0.31, "blue": 0.64}
    if sheet_name == "Agent Suggestions":
        return {"red": 0.82, "green": 0.52, "blue": 0.17}
    return {"red": 0.55, "green": 0.60, "blue": 0.64}


def column_pixel_width(header: str) -> int:
    text = header.lower()
    if any(
        token in text
        for token in (
            "json",
            "notes",
            "description",
            "objective",
            "hypothesis",
            "summary",
            "finding",
            "rationale",
            "proposed_change",
            "expected_effect",
            "interpretation",
            "guidance",
            "immediate_action",
            "impact_assessment",
            "disposition",
            "change_summary",
        )
    ):
        return 280
    if "url" in text or "source_range" in text or "linked_" in text:
        return 200
    if text in {
        "title",
        "name",
        "material_name",
        "reagent_name",
        "measurement_type",
        "parameter",
        "action",
        "method",
        "search_terms",
        "typical_examples",
        "measured_fields",
    }:
        return 185
    if text.endswith("_at") or text.endswith("_date") or text == "timestamp":
        return 155
    if text.endswith("_id") or text in {"record_id", "source_key"}:
        return 145
    if text in {
        "status",
        "qc_status",
        "quality_flag",
        "process_stage",
        "category",
        "target_role",
        "units",
        "confidence",
        "active",
        "required",
    }:
        return 115
    return 125


def column_width_requests(
    sheet_name: str,
    sheet_ids: dict[str, int],
) -> list[dict[str, Any]]:
    """Use bounded, repeatable widths instead of content-driven runaway sizing."""

    if sheet_name not in sheet_ids:
        raise KeyError(f"Missing sheet ID for {sheet_name!r}.")
    widths = [column_pixel_width(header) for header in sheet_by_name(sheet_name).headers]
    requests: list[dict[str, Any]] = []
    start = 0
    while start < len(widths):
        width = widths[start]
        end = start + 1
        while end < len(widths) and widths[end] == width:
            end += 1
        requests.append(
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_ids[sheet_name],
                        "dimension": "COLUMNS",
                        "startIndex": start,
                        "endIndex": end,
                    },
                    "properties": {"pixelSize": width},
                    "fields": "pixelSize",
                }
            }
        )
        start = end
    return requests


def header_row_height_request(
    sheet_name: str,
    sheet_ids: dict[str, int],
) -> dict[str, Any]:
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_ids[sheet_name],
                "dimension": "ROWS",
                "startIndex": 0,
                "endIndex": 1,
            },
            "properties": {"pixelSize": 36},
            "fields": "pixelSize",
        }
    }


def auto_resize_columns_request(sheet_name: str, sheet_ids: dict[str, int]) -> dict[str, Any]:
    if sheet_name not in sheet_ids:
        raise KeyError(f"Missing sheet ID for {sheet_name!r}.")
    headers = list(sheet_by_name(sheet_name).headers)
    return {
        "autoResizeDimensions": {
            "dimensions": {
                "sheetId": sheet_ids[sheet_name],
                "dimension": "COLUMNS",
                "startIndex": 0,
                "endIndex": len(headers),
            }
        }
    }


def run_console_cell_data(value: Any) -> dict[str, Any]:
    if value in (None, ""):
        return {}
    if isinstance(value, bool):
        entered = {"boolValue": value}
    elif isinstance(value, (int, float)):
        entered = {"numberValue": value}
    elif str(value).startswith("="):
        entered = {"formulaValue": str(value)}
    else:
        entered = {"stringValue": str(value)}
    return {"userEnteredValue": entered}


def run_console_lookup_formula(column_letter: str) -> str:
    return (
        '=IF($B$3="","",IFERROR(IF(INDEX(\'Experiments\'!$'
        f'{column_letter}$2:${column_letter}$1000,'
        "MATCH($B$3,'Experiments'!$A$2:$A$1000,0))=\"\",\"Not recorded\","
        "INDEX('Experiments'!$"
        f'{column_letter}$2:${column_letter}$1000,'
        "MATCH($B$3,'Experiments'!$A$2:$A$1000,0))),\"Not recorded\"))"
    )


def run_console_content_rows(sheet_ids: dict[str, int]) -> list[list[Any]]:
    link_rows = (
        ("Experiment record", "Experiments"),
        ("Run plan", "Run Capture Plan"),
        ("Bench observations", "Daily Log"),
        ("Formulation", "Formulations"),
        ("Samples", "Samples"),
        ("Results", "Results"),
        ("Raw files", "Raw Data Files"),
        ("Deviations", "Deviations"),
        ("Plots", "Plot Dashboard"),
    )
    rows: list[list[Any]] = [
        ["EXPERIMENT OVERVIEW", "", "", "", "READINESS", "VALUE", "CHECK", "", "QUICK LINKS", ""],
        ["Status", run_console_lookup_formula("I"), "", "", "Operator", "=$B$7", '=IF($B$3="","",IF(F6<>"Not recorded","✓ Ready","⚠ Missing"))', "", link_rows[0][0], f'=HYPERLINK("#gid={sheet_ids[link_rows[0][1]]}","Open →")'],
        ["Operator", run_console_lookup_formula("H"), "", "", "Protocol", run_console_lookup_formula("Q"), '=IF($B$3="","",IF(F7<>"Not recorded","✓ Ready","⚠ Missing"))', "", link_rows[1][0], f'=HYPERLINK("#gid={sheet_ids[link_rows[1][1]]}","Open →")'],
        ["Process", run_console_lookup_formula("D"), "", "", "Equipment", run_console_lookup_formula("R"), '=IF($B$3="","",IF(F8<>"Not recorded","✓ Ready","⚠ Missing"))', "", link_rows[2][0], f'=HYPERLINK("#gid={sheet_ids[link_rows[2][1]]}","Open →")'],
        ["Date", run_console_lookup_formula("B"), "", "", "Run steps", '=IF($B$3="","",COUNTIF(\'Run Capture Plan\'!$A$2:$A$1000,$B$3))', '=IF($B$3="","",IF(F9>0,"✓ Ready","⚠ Missing"))', "", link_rows[3][0], f'=HYPERLINK("#gid={sheet_ids[link_rows[3][1]]}","Open →")'],
        ["Objective", run_console_lookup_formula("E"), "", "", "Observations", '=IF($B$3="","",COUNTIF(\'Daily Log\'!$A$2:$A$1000,$B$3))', '=IF($B$3="","",IF(F10>0,"✓ Logged","⚠ Missing"))', "", link_rows[4][0], f'=HYPERLINK("#gid={sheet_ids[link_rows[4][1]]}","Open →")'],
        ["Hypothesis", run_console_lookup_formula("F"), "", "", "Samples", '=IF($B$3="","",COUNTIF(\'Samples\'!$B$2:$B$1000,$B$3))', '=IF($B$3="","",IF(F11>0,"✓ Logged","⚠ Missing"))', "", link_rows[5][0], f'=HYPERLINK("#gid={sheet_ids[link_rows[5][1]]}","Open →")'],
        ["Next step", run_console_lookup_formula("J"), "", "", "Results", '=IF($B$3="","",COUNTIF(\'Results\'!$A$2:$A$1000,$B$3))', '=IF($B$3="","",IF(F12>0,"✓ Logged","⚠ Missing"))', "", link_rows[6][0], f'=HYPERLINK("#gid={sheet_ids[link_rows[6][1]]}","Open →")'],
        ["", "", "", "", "Raw files", '=IF($B$3="","",COUNTIF(\'Raw Data Files\'!$B$2:$B$1000,$B$3))', '=IF($B$3="","",IF(F13>0,"✓ Linked","⚠ Missing"))', "", link_rows[7][0], f'=HYPERLINK("#gid={sheet_ids[link_rows[7][1]]}","Open →")'],
        ["", "", "", "", "Open deviations", '=IF($B$3="","",COUNTIFS(\'Deviations\'!$B$2:$B$1000,$B$3,\'Deviations\'!$K$2:$K$1000,"<>closed"))', '=IF($B$3="","",IF(F14=0,"✓ Clear","⚠ Attention"))', "", link_rows[8][0], f'=HYPERLINK("#gid={sheet_ids[link_rows[8][1]]}","Open →")'],
        ["", "", "", "", "Reviewer", run_console_lookup_formula("U"), '=IF($B$3="","",IF(F15<>"Not recorded","✓ Ready","⚠ Missing"))', "", "", ""],
        ["", "", "", "", "Completeness", '=IF($B$3="","",COUNTIF($G$6:$G$15,"✓*")/10)', '=IF($B$3="","",IF(F16=1,"✓ Ready to close",TEXT(F16,"0%")&" complete"))', "", "", ""],
        ["", "", "", "", "", "", "", "", "", ""],
        ["BENCH WORKFLOW", "", "", "", "", "", "", "", "", ""],
        ["1 · PLAN", "Set protocol, equipment, run steps, acceptance criteria, and sample plan before starting.", "", "", "", "", "", "", "", ""],
        ["2 · PREPARE", "Confirm reagent lots, equipment calibration, formulation targets, and safety controls.", "", "", "", "", "", "", "", ""],
        ["3 · RUN", "Record timestamps, actual additions, process conditions, observations, and deviations as they happen.", "", "", "", "", "", "", "", ""],
        ["4 · MEASURE", "Create sample records, link instrument files, capture uncertainty, and evaluate specifications.", "", "", "", "", "", "", "", ""],
        ["5 · REVIEW", "Complete reviewer fields, resolve deviations, document the conclusion, and define the next experiment.", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", ""],
        ["GOOD RECORDS", "Use stable IDs. Record actual values, times, lots, and operators. Link raw evidence; never replace it with a summary.", "", "", "", "", "", "", "", ""],
    ]
    return rows


def run_console_setup_requests(
    sheet_id: int,
    sheet_ids: dict[str, int],
    *,
    is_new: bool,
) -> list[dict[str, Any]]:
    requests: list[dict[str, Any]] = [
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "hidden": False,
                    "tabColorStyle": {"rgbColor": tab_color(RUN_CONSOLE_SHEET)},
                    "gridProperties": {
                        "frozenRowCount": 3,
                        "frozenColumnCount": 0,
                        "hideGridlines": True,
                    },
                },
                "fields": (
                    "hidden,tabColorStyle,gridProperties.frozenRowCount,"
                    "gridProperties.frozenColumnCount,gridProperties.hideGridlines"
                ),
            }
        },
        {
            "updateCells": {
                "start": {"sheetId": sheet_id, "rowIndex": 0, "columnIndex": 0},
                "rows": [
                    {"values": [run_console_cell_data("COCHRAN LAB NOTEBOOK")]},
                    {"values": [run_console_cell_data("Scientist run console · select an experiment to inspect readiness and navigate the record")]} ,
                    {"values": [run_console_cell_data("Active experiment")]},
                ],
                "fields": "userEnteredValue",
            }
        },
        {
            "updateCells": {
                "start": {"sheetId": sheet_id, "rowIndex": 4, "columnIndex": 0},
                "rows": [
                    {"values": [run_console_cell_data(value) for value in row]}
                    for row in run_console_content_rows(sheet_ids)
                ],
                "fields": "userEnteredValue",
            }
        },
        {
            "setDataValidation": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 2,
                    "endRowIndex": 3,
                    "startColumnIndex": 1,
                    "endColumnIndex": 2,
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_RANGE",
                        "values": [
                            {"userEnteredValue": "='Experiments'!$A$2:$A$1000"}
                        ],
                    },
                    "inputMessage": "Choose an experiment ID from the Experiments table.",
                    "strict": True,
                    "showCustomUi": True,
                },
            }
        },
    ]
    if is_new:
        requests.insert(
            2,
            {
                "updateCells": {
                    "start": {"sheetId": sheet_id, "rowIndex": 2, "columnIndex": 1},
                    "rows": [{"values": [run_console_cell_data("")]}],
                    "fields": "userEnteredValue",
                }
            },
        )

    base_format = {
        "backgroundColor": {"red": 1, "green": 1, "blue": 1},
        "textFormat": {
            "fontFamily": "Arial",
            "fontSize": 10,
            "foregroundColor": {"red": 0.12, "green": 0.16, "blue": 0.20},
        },
        "verticalAlignment": "MIDDLE",
        "wrapStrategy": "WRAP",
    }
    requests.extend(
        [
            repeat_cell_format_request(sheet_id, 0, 26, 0, 10, base_format),
            repeat_cell_format_request(
                sheet_id,
                0,
                1,
                0,
                10,
                {
                    "backgroundColor": {"red": 0.07, "green": 0.18, "blue": 0.29},
                    "textFormat": {
                        "fontFamily": "Arial",
                        "fontSize": 20,
                        "bold": True,
                        "foregroundColor": {"red": 1, "green": 1, "blue": 1},
                    },
                    "verticalAlignment": "MIDDLE",
                },
            ),
            repeat_cell_format_request(
                sheet_id,
                1,
                2,
                0,
                10,
                {
                    "backgroundColor": {"red": 0.88, "green": 0.93, "blue": 0.96},
                    "textFormat": {
                        "fontFamily": "Arial",
                        "fontSize": 10,
                        "foregroundColor": {"red": 0.22, "green": 0.31, "blue": 0.38},
                    },
                },
            ),
            repeat_cell_format_request(
                sheet_id,
                2,
                3,
                1,
                2,
                {
                    "backgroundColor": {"red": 1.0, "green": 0.95, "blue": 0.72},
                    "textFormat": {"bold": True, "foregroundColor": {"red": 0.12, "green": 0.16, "blue": 0.20}},
                    "borders": {
                        "top": {"style": "SOLID", "color": {"red": 0.74, "green": 0.59, "blue": 0.16}},
                        "bottom": {"style": "SOLID", "color": {"red": 0.74, "green": 0.59, "blue": 0.16}},
                        "left": {"style": "SOLID", "color": {"red": 0.74, "green": 0.59, "blue": 0.16}},
                        "right": {"style": "SOLID", "color": {"red": 0.74, "green": 0.59, "blue": 0.16}},
                    },
                },
            ),
        ]
    )
    for row_index in (4, 17):
        requests.append(
            repeat_cell_format_request(
                sheet_id,
                row_index,
                row_index + 1,
                0,
                10,
                {
                    "backgroundColor": {"red": 0.10, "green": 0.48, "blue": 0.55},
                    "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                    "verticalAlignment": "MIDDLE",
                },
            )
        )
    requests.append(
        repeat_cell_format_request(
            sheet_id,
            5,
            16,
            0,
            1,
            {"backgroundColor": {"red": 0.92, "green": 0.96, "blue": 0.97}, "textFormat": {"bold": True}},
        )
    )
    requests.append(
        repeat_cell_format_request(
            sheet_id,
            5,
            16,
            4,
            5,
            {"backgroundColor": {"red": 0.92, "green": 0.96, "blue": 0.97}, "textFormat": {"bold": True}},
        )
    )
    requests.append(
        repeat_cell_format_request(
            sheet_id,
            18,
            23,
            0,
            1,
            {"backgroundColor": {"red": 0.88, "green": 0.93, "blue": 0.96}, "textFormat": {"bold": True}},
        )
    )
    requests.append(
        repeat_cell_format_request(
            sheet_id,
            24,
            25,
            0,
            10,
            {"backgroundColor": {"red": 1.0, "green": 0.95, "blue": 0.72}, "textFormat": {"bold": True}},
        )
    )
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 15,
                    "endRowIndex": 16,
                    "startColumnIndex": 5,
                    "endColumnIndex": 6,
                },
                "cell": {"userEnteredFormat": {"numberFormat": {"type": "PERCENT", "pattern": "0%"}}},
                "fields": "userEnteredFormat.numberFormat",
            }
        }
    )
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 8,
                    "endRowIndex": 9,
                    "startColumnIndex": 1,
                    "endColumnIndex": 2,
                },
                "cell": {"userEnteredFormat": {"numberFormat": {"type": "DATE", "pattern": "yyyy-mm-dd"}}},
                "fields": "userEnteredFormat.numberFormat",
            }
        }
    )
    for index, width in enumerate((145, 285, 26, 26, 145, 95, 135, 26, 145, 95)):
        requests.append(
            {
                "updateDimensionProperties": {
                    "range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": index, "endIndex": index + 1},
                    "properties": {"pixelSize": width},
                    "fields": "pixelSize",
                }
            }
        )
    for row_index, height in ((0, 42), (1, 26), (2, 34), (4, 30), (17, 30), (24, 46)):
        requests.append(
            {
                "updateDimensionProperties": {
                    "range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": row_index, "endIndex": row_index + 1},
                    "properties": {"pixelSize": height},
                    "fields": "pixelSize",
                }
            }
        )
    if is_new:
        for index, (prefix, color) in enumerate(
            (
                ("✓", {"red": 0.78, "green": 0.90, "blue": 0.79}),
                ("⚠", {"red": 1.0, "green": 0.92, "blue": 0.68}),
            )
        ):
            requests.append(
                {
                    "addConditionalFormatRule": {
                        "index": index,
                        "rule": {
                            "ranges": [{"sheetId": sheet_id, "startRowIndex": 5, "endRowIndex": 16, "startColumnIndex": 6, "endColumnIndex": 7}],
                            "booleanRule": {
                                "condition": {"type": "TEXT_STARTS_WITH", "values": [{"userEnteredValue": prefix}]},
                                "format": {"backgroundColor": color, "textFormat": {"bold": True}},
                            },
                        },
                    }
                }
            )
    return requests


def repeat_cell_format_request(
    sheet_id: int,
    start_row: int,
    end_row: int,
    start_column: int,
    end_column: int,
    user_entered_format: dict[str, Any],
) -> dict[str, Any]:
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_column,
                "endColumnIndex": end_column,
            },
            "cell": {"userEnteredFormat": user_entered_format},
            "fields": "userEnteredFormat",
        }
    }


def column_number_format_requests(
    sheet_name: str,
    sheet_ids: dict[str, int],
    end_row: int = DEFAULT_VALIDATION_END_ROW,
) -> list[dict[str, Any]]:
    if sheet_name not in sheet_ids:
        raise KeyError(f"Missing sheet ID for {sheet_name!r}.")
    requests: list[dict[str, Any]] = []
    for column_index, header in enumerate(sheet_by_name(sheet_name).headers):
        pattern = column_number_format(sheet_name, header)
        if not pattern:
            continue
        requests.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_ids[sheet_name],
                        "startRowIndex": 1,
                        "endRowIndex": max(2, end_row),
                        "startColumnIndex": column_index,
                        "endColumnIndex": column_index + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": (
                                    "DATE_TIME"
                                    if column_data_type(sheet_name, header) == "datetime"
                                    else "DATE"
                                    if column_data_type(sheet_name, header) == "date"
                                    else "NUMBER"
                                ),
                                "pattern": pattern,
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat",
                }
            }
        )
    return requests


def basic_filter_request(
    sheet_name: str,
    sheet_ids: dict[str, int],
    end_row: int = DEFAULT_VALIDATION_END_ROW,
) -> dict[str, Any]:
    if sheet_name not in sheet_ids:
        raise KeyError(f"Missing sheet ID for {sheet_name!r}.")
    return {
        "setBasicFilter": {
            "filter": {
                "range": {
                    "sheetId": sheet_ids[sheet_name],
                    "startRowIndex": 0,
                    "endRowIndex": max(2, end_row),
                    "startColumnIndex": 0,
                    "endColumnIndex": len(sheet_by_name(sheet_name).headers),
                }
            }
        }
    }


def data_validation_request(
    sheet_name: str,
    field: str,
    allowed_values: tuple[str, ...],
    sheet_ids: dict[str, int],
    validation_end_row: int = DEFAULT_VALIDATION_END_ROW,
) -> dict[str, Any]:
    if sheet_name not in sheet_ids:
        raise KeyError(f"Missing sheet ID for {sheet_name!r}.")
    headers = list(sheet_by_name(sheet_name).headers)
    if field not in headers:
        raise KeyError(f"Unknown field {field!r} for {sheet_name!r}.")
    column_index = headers.index(field)
    return {
        "setDataValidation": {
            "range": {
                "sheetId": sheet_ids[sheet_name],
                "startRowIndex": 1,
                "endRowIndex": validation_end_row,
                "startColumnIndex": column_index,
                "endColumnIndex": column_index + 1,
            },
            "rule": {
                "condition": {
                    "type": "ONE_OF_LIST",
                    "values": [
                        {"userEnteredValue": str(value)}
                        for value in allowed_values
                    ],
                },
                "inputMessage": "Choose a value from the controlled vocabulary.",
                "strict": True,
                "showCustomUi": True,
            },
        }
    }


def optional_int(value: Any) -> int | None:
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def validate_snapshot(snapshot: dict[str, Any], require_sheet_ids: bool = False) -> dict[str, Any]:
    errors: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    sheets = snapshot.get("sheets")
    if not isinstance(sheets, dict):
        return {
            "valid": False,
            "errors": [{"code": "missing_sheets_object", "message": "Snapshot must include a sheets object."}],
            "warnings": [],
            "summary": {},
        }

    for spec in SHEETS:
        payload = sheets.get(spec.name)
        if payload is None:
            if spec.name in OPTIONAL_EXTENSION_SHEETS:
                continue
            errors.append({"code": "missing_sheet", "sheet": spec.name})
            continue
        values = sheet_values(payload)
        if not values and spec.name in OPTIONAL_EXTENSION_SHEETS:
            continue
        actual_headers = [str(value) for value in values[0]] if values else []
        expected_headers = list(spec.headers)
        if actual_headers[: len(expected_headers)] != expected_headers:
            errors.append(
                {
                    "code": "header_mismatch",
                    "sheet": spec.name,
                    "expected": expected_headers,
                    "actual": actual_headers,
                }
            )
        if require_sheet_ids and (not isinstance(payload, dict) or payload.get("sheet_id") is None):
            errors.append({"code": "missing_sheet_id", "sheet": spec.name})
        if len(actual_headers) > len(expected_headers):
            warnings.append(
                {
                    "code": "extra_columns",
                    "sheet": spec.name,
                    "expected_count": len(expected_headers),
                    "actual_count": len(actual_headers),
                }
            )

    unknown_sheets = sorted(set(sheets) - {sheet.name for sheet in SHEETS})
    for sheet_name in unknown_sheets:
        warnings.append({"code": "unknown_sheet", "sheet": sheet_name})

    return {
        "valid": not errors,
        "errors": errors,
        "warnings": warnings,
        "summary": {
            "sheet_count": len(sheets),
            "missing_sheet_count": sum(1 for error in errors if error.get("code") == "missing_sheet"),
            "header_mismatch_count": sum(1 for error in errors if error.get("code") == "header_mismatch"),
            "warning_count": len(warnings),
        },
    }


def audit_report_against_snapshot(
    report: dict[str, Any],
    snapshot: dict[str, Any],
    require_sheet_ids: bool = True,
) -> dict[str, Any]:
    snapshot_audit = validate_snapshot(snapshot, require_sheet_ids=False)
    errors: list[dict[str, Any]] = list(snapshot_audit["errors"])
    warnings: list[dict[str, Any]] = list(snapshot_audit["warnings"])
    tables = snapshot_to_tables(snapshot) if snapshot_audit["valid"] or "sheets" in snapshot else {}
    sheet_ids = sheet_ids_from_snapshot(snapshot) if isinstance(snapshot.get("sheets"), dict) else {}

    evidence_rows = collect_rows(report, "append_literature_evidence")
    suggestion_rows = [
        suggestion_to_workbook_row(suggestion)
        for suggestion in collect_rows(report, "append_agent_suggestions")
    ]
    master_reagent_rows = collect_rows(report, "append_master_reagents")
    master_reagent_updates = collect_rows(report, "update_master_reagents")
    formulation_rows = collect_rows(report, "append_formulations")
    formulation_updates = collect_rows(report, "update_formulations")
    experiment_rows = collect_rows(report, "append_experiments")
    result_rows = collect_rows(report, "append_results")
    daily_log_rows = collect_rows(report, "append_daily_log")
    daily_review_rows = collect_rows(report, "append_daily_reviews")
    project_notebook_rows = collect_rows(report, "append_project_notebook_records")
    project_notebook_updates = collect_rows(report, "update_project_notebook_records")
    source_sync_rows = collect_rows(report, "append_source_sync")
    source_sync_updates = collect_rows(report, "update_source_sync")
    plot_replacements = {
        sheet_name: collect_rows(report, report_key)
        for sheet_name, report_key in (
            ("Plot Data", "replace_plot_data"),
            ("Plot Definitions", "replace_plot_definitions"),
            ("Plot Dashboard", "replace_plot_dashboard"),
        )
        if report_key in report
    }
    experiment_updates = collect_rows(report, "update_experiments")
    suggestion_updates = collect_rows(report, "update_agent_suggestions")
    for row in master_reagent_rows:
        reagent_id = str(row.get("reagent_id", ""))
        if reagent_id and any(str(existing.get("reagent_id", "")) == reagent_id for existing in tables.get("Master Reagents", [])):
            errors.append({"code": "duplicate_append", "sheet": "Master Reagents", "key": "reagent_id", "value": reagent_id})
    for update in master_reagent_updates:
        reagent_id = str(update.get("reagent_id", "") or update.get("key_value", ""))
        if reagent_id and not any(str(existing.get("reagent_id", "")) == reagent_id for existing in tables.get("Master Reagents", [])):
            errors.append({"code": "missing_update_target", "sheet": "Master Reagents", "key": "reagent_id", "value": reagent_id})
    existing_formulation_keys = {formulation_key(row) for row in tables.get("Formulations", [])}
    for row in formulation_rows:
        row_key = formulation_key(row)
        if row_key in existing_formulation_keys:
            errors.append(
                {
                    "code": "duplicate_append",
                    "sheet": "Formulations",
                    "key": "experiment_id,reagent_id,target_role",
                    "value": "|".join(row_key),
                }
            )
    for update in formulation_updates:
        update_key = (
            str(update.get("experiment_id", "")).strip(),
            str(update.get("reagent_id", "")).strip(),
            str(update.get("target_role", "")).strip(),
        )
        if update_key not in existing_formulation_keys:
            errors.append(
                {
                    "code": "missing_update_target",
                    "sheet": "Formulations",
                    "key": "experiment_id,reagent_id,target_role",
                    "value": "|".join(update_key),
                }
            )
    for row in evidence_rows:
        evidence_id = str(row.get("evidence_id", ""))
        if evidence_id and any(str(existing.get("evidence_id", "")) == evidence_id for existing in tables.get("Literature Evidence", [])):
            errors.append({"code": "duplicate_append", "sheet": "Literature Evidence", "key": "evidence_id", "value": evidence_id})
    for row in suggestion_rows:
        suggestion_id = str(row.get("suggestion_id", ""))
        if suggestion_id and any(str(existing.get("suggestion_id", "")) == suggestion_id for existing in tables.get("Agent Suggestions", [])):
            errors.append({"code": "duplicate_append", "sheet": "Agent Suggestions", "key": "suggestion_id", "value": suggestion_id})
    for update in suggestion_updates:
        suggestion_id = str(update.get("suggestion_id", ""))
        if suggestion_id and not any(str(existing.get("suggestion_id", "")) == suggestion_id for existing in tables.get("Agent Suggestions", [])):
            errors.append({"code": "missing_update_target", "sheet": "Agent Suggestions", "key": "suggestion_id", "value": suggestion_id})
    for update in experiment_updates:
        experiment_id = str(update.get("experiment_id", "") or update.get("key_value", ""))
        if experiment_id and not any(str(existing.get("experiment_id", "")) == experiment_id for existing in tables.get("Experiments", [])):
            errors.append({"code": "missing_update_target", "sheet": "Experiments", "key": "experiment_id", "value": experiment_id})
    for row in experiment_rows:
        experiment_id = str(row.get("experiment_id", ""))
        if experiment_id and any(str(existing.get("experiment_id", "")) == experiment_id for existing in tables.get("Experiments", [])):
            errors.append({"code": "duplicate_append", "sheet": "Experiments", "key": "experiment_id", "value": experiment_id})
    existing_result_keys = {result_row_key(row) for row in tables.get("Results", [])}
    for row in result_rows:
        row_key = result_row_key(row)
        if row_key in existing_result_keys:
            errors.append(
                {
                    "code": "duplicate_append",
                    "sheet": "Results",
                    "key": "experiment_id,sample_id,measurement_type",
                    "value": "|".join(row_key),
                }
            )
    existing_daily_log_keys = {daily_log_row_key(row) for row in tables.get("Daily Log", [])}
    for row in daily_log_rows:
        row_key = daily_log_row_key(row)
        if row_key in existing_daily_log_keys:
            errors.append(
                {
                    "code": "duplicate_append",
                    "sheet": "Daily Log",
                    "key": "experiment_id,timestamp,observation",
                    "value": "|".join(row_key),
                }
            )
    for row in daily_review_rows:
        review_id = str(row.get("review_id", ""))
        if review_id and any(str(existing.get("review_id", "")) == review_id for existing in tables.get("Daily Reviews", [])):
            errors.append({"code": "duplicate_append", "sheet": "Daily Reviews", "key": "review_id", "value": review_id})
    audit_keyed_rows(
        errors,
        tables,
        "Project Notebook Records",
        "record_id",
        project_notebook_rows,
        project_notebook_updates,
    )
    audit_keyed_rows(
        errors,
        tables,
        "Source Sync",
        "source_key",
        source_sync_rows,
        source_sync_updates,
    )

    if require_sheet_ids and evidence_rows and "Literature Evidence" not in sheet_ids:
        errors.append({"code": "missing_apply_sheet_id", "sheet": "Literature Evidence"})
    if require_sheet_ids and suggestion_rows and "Agent Suggestions" not in sheet_ids:
        errors.append({"code": "missing_apply_sheet_id", "sheet": "Agent Suggestions"})
    if require_sheet_ids and suggestion_updates and "Agent Suggestions" not in sheet_ids:
        errors.append({"code": "missing_apply_sheet_id", "sheet": "Agent Suggestions"})
    if require_sheet_ids and master_reagent_rows and "Master Reagents" not in sheet_ids:
        errors.append({"code": "missing_apply_sheet_id", "sheet": "Master Reagents"})
    if require_sheet_ids and master_reagent_updates and "Master Reagents" not in sheet_ids:
        errors.append({"code": "missing_apply_sheet_id", "sheet": "Master Reagents"})
    if require_sheet_ids and formulation_rows and "Formulations" not in sheet_ids:
        errors.append({"code": "missing_apply_sheet_id", "sheet": "Formulations"})
    if require_sheet_ids and formulation_updates and "Formulations" not in sheet_ids:
        errors.append({"code": "missing_apply_sheet_id", "sheet": "Formulations"})
    if require_sheet_ids and experiment_rows and "Experiments" not in sheet_ids:
        errors.append({"code": "missing_apply_sheet_id", "sheet": "Experiments"})
    if require_sheet_ids and experiment_updates and "Experiments" not in sheet_ids:
        errors.append({"code": "missing_apply_sheet_id", "sheet": "Experiments"})
    if require_sheet_ids and result_rows and "Results" not in sheet_ids:
        errors.append({"code": "missing_apply_sheet_id", "sheet": "Results"})
    if require_sheet_ids and daily_log_rows and "Daily Log" not in sheet_ids:
        errors.append({"code": "missing_apply_sheet_id", "sheet": "Daily Log"})
    if require_sheet_ids and daily_review_rows and "Daily Reviews" not in sheet_ids:
        errors.append({"code": "missing_apply_sheet_id", "sheet": "Daily Reviews"})
    if (
        require_sheet_ids
        and (project_notebook_rows or project_notebook_updates)
        and "Project Notebook Records" not in sheet_ids
    ):
        errors.append({"code": "missing_apply_sheet_id", "sheet": "Project Notebook Records"})
    if (
        require_sheet_ids
        and (source_sync_rows or source_sync_updates)
        and "Source Sync" not in sheet_ids
    ):
        errors.append({"code": "missing_apply_sheet_id", "sheet": "Source Sync"})
    if require_sheet_ids:
        for sheet_name in plot_replacements:
            if sheet_name not in sheet_ids:
                errors.append(
                    {"code": "missing_apply_sheet_id", "sheet": sheet_name}
                )

    return {
        "valid": not errors,
        "errors": errors,
        "warnings": warnings,
        "summary": {
            "master_reagent_rows_to_append": len(master_reagent_rows),
            "master_reagent_cells_to_update": len(master_reagent_updates),
            "formulation_rows_to_append": len(formulation_rows),
            "formulation_cells_to_update": len(formulation_updates),
            "evidence_rows_to_append": len(evidence_rows),
            "suggestion_rows_to_append": len(suggestion_rows),
            "suggestion_rows_to_update": len(suggestion_updates),
            "experiment_rows_to_append": len(experiment_rows),
            "experiment_cells_to_update": len(experiment_updates),
            "result_rows_to_append": len(result_rows),
            "daily_log_rows_to_append": len(daily_log_rows),
            "daily_review_rows_to_append": len(daily_review_rows),
            "project_notebook_rows_to_append": len(project_notebook_rows),
            "project_notebook_cells_to_update": len(project_notebook_updates),
            "source_sync_rows_to_append": len(source_sync_rows),
            "source_sync_cells_to_update": len(source_sync_updates),
            "plot_sheets_to_replace": len(plot_replacements),
            "plot_rows_to_replace": sum(
                len(rows) for rows in plot_replacements.values()
            ),
            "request_count": len(batch_update_requests_from_report(report, sheet_ids)) if not errors else 0,
        },
    }


def batch_update_requests_from_report(
    report: dict[str, Any],
    sheet_ids: dict[str, int],
) -> list[dict[str, Any]]:
    requests: list[dict[str, Any]] = []
    evidence_rows = collect_rows(report, "append_literature_evidence")
    suggestion_rows = [
        suggestion_to_workbook_row(suggestion)
        for suggestion in collect_rows(report, "append_agent_suggestions")
    ]
    master_reagent_rows = collect_rows(report, "append_master_reagents")
    master_reagent_updates = collect_rows(report, "update_master_reagents")
    formulation_rows = collect_rows(report, "append_formulations")
    formulation_updates = collect_rows(report, "update_formulations")
    experiment_rows = collect_rows(report, "append_experiments")
    result_rows = collect_rows(report, "append_results")
    daily_log_rows = collect_rows(report, "append_daily_log")
    daily_review_rows = collect_rows(report, "append_daily_reviews")
    project_notebook_rows = collect_rows(report, "append_project_notebook_records")
    project_notebook_updates = collect_rows(report, "update_project_notebook_records")
    source_sync_rows = collect_rows(report, "append_source_sync")
    source_sync_updates = collect_rows(report, "update_source_sync")
    plot_replacements = (
        ("Plot Data", "replace_plot_data"),
        ("Plot Definitions", "replace_plot_definitions"),
        ("Plot Dashboard", "replace_plot_dashboard"),
    )
    experiment_updates = collect_rows(report, "update_experiments")
    suggestion_updates = collect_rows(report, "update_agent_suggestions")
    if master_reagent_rows:
        requests.append(append_cells_request("Master Reagents", master_reagent_rows, sheet_ids))
    for update in master_reagent_updates:
        requests.append(update_cell_request("Master Reagents", update, sheet_ids))
    if experiment_rows:
        requests.append(append_cells_request("Experiments", experiment_rows, sheet_ids))
    if formulation_rows:
        requests.append(append_cells_request("Formulations", formulation_rows, sheet_ids))
    for update in formulation_updates:
        requests.append(update_cell_request("Formulations", update, sheet_ids))
    if daily_log_rows:
        requests.append(append_cells_request("Daily Log", daily_log_rows, sheet_ids))
    if result_rows:
        requests.append(append_cells_request("Results", result_rows, sheet_ids))
    if evidence_rows:
        requests.append(append_cells_request("Literature Evidence", evidence_rows, sheet_ids))
    if suggestion_rows:
        requests.append(append_cells_request("Agent Suggestions", suggestion_rows, sheet_ids))
    if daily_review_rows:
        requests.append(append_cells_request("Daily Reviews", daily_review_rows, sheet_ids))
    if project_notebook_rows:
        requests.append(
            append_cells_request(
                "Project Notebook Records",
                project_notebook_rows,
                sheet_ids,
            )
        )
    for update in project_notebook_updates:
        requests.append(
            update_cell_request("Project Notebook Records", update, sheet_ids)
        )
    if source_sync_rows:
        requests.append(append_cells_request("Source Sync", source_sync_rows, sheet_ids))
    for update in source_sync_updates:
        requests.append(update_cell_request("Source Sync", update, sheet_ids))
    for update in experiment_updates:
        requests.append(update_cell_request("Experiments", update, sheet_ids))
    for update in suggestion_updates:
        requests.append(update_cell_request("Agent Suggestions", update, sheet_ids))
    existing_counts = report.get("existing_plot_row_counts", {})
    for sheet_name, report_key in plot_replacements:
        if report_key not in report:
            continue
        rows = collect_rows(report, report_key)
        requests.extend(
            replace_sheet_rows_requests(
                sheet_name,
                rows,
                int(existing_counts.get(sheet_name, 0) or 0),
                sheet_ids,
            )
        )
    return requests


def append_cells_request(
    sheet_name: str,
    rows: list[dict[str, Any]],
    sheet_ids: dict[str, int],
) -> dict[str, Any]:
    if sheet_name not in sheet_ids:
        raise KeyError(f"Missing sheet ID for {sheet_name!r}.")
    columns = list(sheet_by_name(sheet_name).columns)
    return {
        "appendCells": {
            "sheetId": sheet_ids[sheet_name],
            "rows": [
                {
                    "values": [
                        google_cell_data(
                            row.get(column.name, ""),
                            data_type=column_data_type(sheet_name, column.name),
                        )
                        for column in columns
                    ]
                }
                for row in rows
            ],
            "fields": "userEnteredValue",
        }
    }


def replace_sheet_rows_requests(
    sheet_name: str,
    rows: list[dict[str, Any]],
    existing_row_count: int,
    sheet_ids: dict[str, int],
) -> list[dict[str, Any]]:
    """Replace all managed data rows while preserving the contract header row."""

    if sheet_name not in sheet_ids:
        raise KeyError(f"Missing sheet ID for {sheet_name!r}.")
    headers = list(sheet_by_name(sheet_name).headers)
    managed_row_count = max(len(rows), max(0, existing_row_count))
    requests = [
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_ids[sheet_name],
                    "gridProperties": {
                        "rowCount": max(
                            DEFAULT_VALIDATION_END_ROW,
                            managed_row_count + 1,
                        ),
                        "columnCount": minimum_grid_column_count(sheet_name),
                    },
                },
                "fields": "gridProperties(rowCount,columnCount)",
            }
        }
    ]
    if managed_row_count == 0:
        return requests
    update_rows = [
        {
            "values": [
                google_cell_data(
                    row.get(header, ""),
                    data_type=column_data_type(sheet_name, header),
                )
                for header in headers
            ]
        }
        for row in rows
    ]
    update_rows.extend(
        {
            "values": [
                {} for _ in headers
            ]
        }
        for _ in range(managed_row_count - len(rows))
    )
    requests.append(
        {
            "updateCells": {
                "range": {
                    "sheetId": sheet_ids[sheet_name],
                    "startRowIndex": 1,
                    "endRowIndex": managed_row_count + 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": len(headers),
                },
                "rows": update_rows,
                "fields": "userEnteredValue",
            }
        }
    )
    return requests


def google_cell_data(value: Any, data_type: str = "text") -> dict[str, Any]:
    entered_value = google_user_entered_value(value, data_type=data_type)
    return {"userEnteredValue": entered_value} if entered_value else {}


def google_user_entered_value(value: Any, data_type: str = "text") -> dict[str, Any]:
    if value in ("", None):
        return {}
    if isinstance(value, bool):
        return {"boolValue": value}
    if isinstance(value, (int, float)):
        return {"numberValue": value}
    if data_type == "number":
        numeric_value = parse_number(value)
        if numeric_value is not None:
            return {"numberValue": numeric_value}
    if data_type in {"date", "datetime"}:
        serial_value = google_date_serial(value, include_time=data_type == "datetime")
        if serial_value is not None:
            return {"numberValue": serial_value}
    return {"stringValue": stringify_cell(value)}


def update_cell_request(
    sheet_name: str,
    update: dict[str, Any],
    sheet_ids: dict[str, int],
) -> dict[str, Any]:
    if sheet_name not in sheet_ids:
        raise KeyError(f"Missing sheet ID for {sheet_name!r}.")
    headers = list(sheet_by_name(sheet_name).headers)
    field = str(update.get("field", ""))
    if field not in headers:
        raise KeyError(f"Unknown field {field!r} for {sheet_name!r}.")
    row_number = int(update.get("row_number", 0))
    if row_number < 2:
        raise ValueError("Update row_number must be a 1-based sheet row number greater than or equal to 2.")
    return {
        "updateCells": {
            "start": {
                "sheetId": sheet_ids[sheet_name],
                "rowIndex": row_number - 1,
                "columnIndex": headers.index(field),
            },
            "rows": [
                {
                    "values": [
                        {
                            "userEnteredValue": {
                                "stringValue": stringify_cell(update.get("value", ""))
                            }
                        }
                    ]
                }
            ],
            "fields": "userEnteredValue",
        }
    }


def parse_number(value: Any) -> int | float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    if text.endswith("%"):
        text = text[:-1].strip()
    try:
        parsed = float(text)
    except ValueError:
        return None
    return int(parsed) if parsed.is_integer() else parsed


def google_date_serial(value: Any, include_time: bool = False) -> float | None:
    parsed: datetime
    if isinstance(value, datetime):
        parsed = value
    elif isinstance(value, date):
        parsed = datetime(value.year, value.month, value.day)
    else:
        text = str(value).strip()
        if not text:
            return None
        normalized = text[:-1] + "+00:00" if text.endswith("Z") else text
        try:
            parsed = datetime.fromisoformat(normalized)
        except ValueError:
            try:
                parsed_date = date.fromisoformat(text)
            except ValueError:
                return None
            parsed = datetime(parsed_date.year, parsed_date.month, parsed_date.day)
    if parsed.tzinfo is not None:
        parsed = parsed.astimezone(timezone.utc).replace(tzinfo=None)
    if not include_time:
        parsed = datetime(parsed.year, parsed.month, parsed.day)
    epoch = datetime(1899, 12, 30)
    return (parsed - epoch).total_seconds() / 86_400


def google_contract_migration_requests(
    tables: dict[str, list[dict[str, Any]]],
    sheet_ids: dict[str, int],
    *,
    migrated_at: str | None = None,
    time_zone: str = DEFAULT_WORKBOOK_TIMEZONE,
    normalize_existing_types: bool = False,
) -> list[dict[str, Any]]:
    """Build idempotent seed, metadata, lineage, and optional type-repair requests."""

    migrated_at = migrated_at or datetime.now(timezone.utc).isoformat()
    requests: list[dict[str, Any]] = []
    keyed_seed_specs = {
        "Process Knowledge": ("process_type", "material_role"),
        "Controlled Vocab": ("field", "allowed_value"),
        "Agent Config": ("key",),
    }
    for sheet_name, key_fields in keyed_seed_specs.items():
        spec = sheet_by_name(sheet_name)
        existing_keys = {
            tuple(str(row.get(field, "")).strip() for field in key_fields)
            for row in tables.get(sheet_name, [])
        }
        missing_rows: list[dict[str, Any]] = []
        for values in spec.example_rows:
            row = {
                header: values[index] if index < len(values) else ""
                for index, header in enumerate(spec.headers)
            }
            key = tuple(str(row.get(field, "")).strip() for field in key_fields)
            if key not in existing_keys:
                missing_rows.append(row)
                existing_keys.add(key)
        if missing_rows:
            requests.append(append_cells_request(sheet_name, missing_rows, sheet_ids))

    metadata_rows = tables.get("Workbook Metadata", [])
    metadata_by_key = {
        str(row.get("key", "")).strip(): (index + 2, row)
        for index, row in enumerate(metadata_rows)
        if str(row.get("key", "")).strip()
    }
    desired_metadata = (
        (
            "contract_name",
            "lab-notebook-agent-workbook",
            "Machine-readable workbook contract.",
        ),
        (
            "contract_version",
            WORKBOOK_CONTRACT_VERSION,
            "Schema version currently applied to this workbook.",
        ),
        (
            "workbook_timezone",
            time_zone,
            "Timezone used for local laboratory timestamps.",
        ),
        (
            "migration_status",
            "current",
            "Set by a successful contract migration.",
        ),
        (
            "last_migrated_at",
            migrated_at,
            "UTC timestamp of the most recent contract migration.",
        ),
    )
    metadata_appends: list[dict[str, Any]] = []
    for key, value, notes in desired_metadata:
        existing = metadata_by_key.get(key)
        if existing is None:
            metadata_appends.append(
                {
                    "key": key,
                    "value": value,
                    "updated_at": migrated_at,
                    "notes": notes,
                }
            )
            continue
        row_number, row = existing
        for field, desired_value in (
            ("value", value),
            ("updated_at", migrated_at),
            ("notes", notes),
        ):
            if str(row.get(field, "")) != str(desired_value):
                requests.append(
                    update_cell_request(
                        "Workbook Metadata",
                        {
                            "row_number": row_number,
                            "field": field,
                            "value": desired_value,
                        },
                        sheet_ids,
                    )
                )
    if metadata_appends:
        requests.append(
            append_cells_request("Workbook Metadata", metadata_appends, sheet_ids)
        )

    audit_id = f"MIGRATION-{WORKBOOK_CONTRACT_VERSION}"
    existing_audit_ids = {
        str(row.get("audit_id", "")).strip()
        for row in tables.get("Audit Log", [])
    }
    if audit_id not in existing_audit_ids:
        requests.append(
            append_cells_request(
                "Audit Log",
                [
                    {
                        "audit_id": audit_id,
                        "occurred_at": migrated_at,
                        "actor": "lab-notebook-agent",
                        "action": "migrate",
                        "sheet_name": "Workbook Metadata",
                        "row_key": "contract_version",
                        "field_name": "value",
                        "old_value": "",
                        "new_value": WORKBOOK_CONTRACT_VERSION,
                        "reason": "Upgrade workbook contract and scientific record structure.",
                        "source": "google-setup-live",
                    }
                ],
                sheet_ids,
            )
        )

    if normalize_existing_types:
        requests.extend(google_type_normalization_requests(tables, sheet_ids))
    return requests


def google_type_normalization_requests(
    tables: dict[str, list[dict[str, Any]]],
    sheet_ids: dict[str, int],
) -> list[dict[str, Any]]:
    """Rewrite parseable typed cells without changing text fields or blank cells."""

    requests: list[dict[str, Any]] = []
    for sheet_name, rows in tables.items():
        if sheet_name not in sheet_ids:
            continue
        headers = list(sheet_by_name(sheet_name).headers)
        for row_index, row in enumerate(rows, start=1):
            for column_index, header in enumerate(headers):
                data_type = column_data_type(sheet_name, header)
                if data_type == "text":
                    continue
                value = row.get(header, "")
                entered_value = google_user_entered_value(value, data_type=data_type)
                if not entered_value or "stringValue" in entered_value:
                    continue
                requests.append(
                    {
                        "updateCells": {
                            "start": {
                                "sheetId": sheet_ids[sheet_name],
                                "rowIndex": row_index,
                                "columnIndex": column_index,
                            },
                            "rows": [
                                {
                                    "values": [
                                        {"userEnteredValue": entered_value}
                                    ]
                                }
                            ],
                            "fields": "userEnteredValue",
                        }
                    }
                )
    return requests


def quality_conditional_format_requests(
    sheet_ids: dict[str, int],
    end_row: int = DEFAULT_VALIDATION_END_ROW,
) -> list[dict[str, Any]]:
    requests: list[dict[str, Any]] = []
    color_by_status = {
        "pass": {"red": 0.78, "green": 0.90, "blue": 0.79},
        "warn": {"red": 1.0, "green": 0.92, "blue": 0.68},
        "fail": {"red": 0.96, "green": 0.76, "blue": 0.76},
        "not_evaluated": {"red": 0.90, "green": 0.90, "blue": 0.90},
    }
    if "Results" in sheet_ids:
        headers = list(sheet_by_name("Results").headers)
        column_index = headers.index("qc_status")
        for index, (status, color) in enumerate(color_by_status.items()):
            requests.append(
                {
                    "addConditionalFormatRule": {
                        "index": index,
                        "rule": {
                            "ranges": [
                                {
                                    "sheetId": sheet_ids["Results"],
                                    "startRowIndex": 1,
                                    "endRowIndex": max(2, end_row),
                                    "startColumnIndex": column_index,
                                    "endColumnIndex": column_index + 1,
                                }
                            ],
                            "booleanRule": {
                                "condition": {
                                    "type": "TEXT_EQ",
                                    "values": [{"userEnteredValue": status}],
                                },
                                "format": {"backgroundColor": color},
                            },
                        },
                    }
                }
            )
    if "Deviations" in sheet_ids:
        headers = list(sheet_by_name("Deviations").headers)
        column_index = headers.index("status")
        for index, (status, color) in enumerate(
            (
                ("open", {"red": 0.96, "green": 0.76, "blue": 0.76}),
                ("under_review", {"red": 1.0, "green": 0.92, "blue": 0.68}),
                ("closed", {"red": 0.78, "green": 0.90, "blue": 0.79}),
            )
        ):
            requests.append(
                {
                    "addConditionalFormatRule": {
                        "index": index,
                        "rule": {
                            "ranges": [
                                {
                                    "sheetId": sheet_ids["Deviations"],
                                    "startRowIndex": 1,
                                    "endRowIndex": max(2, end_row),
                                    "startColumnIndex": column_index,
                                    "endColumnIndex": column_index + 1,
                                }
                            ],
                            "booleanRule": {
                                "condition": {
                                    "type": "TEXT_EQ",
                                    "values": [{"userEnteredValue": status}],
                                },
                                "format": {"backgroundColor": color},
                            },
                        },
                    }
                }
            )
    return requests


def daily_log_row_key(row: dict[str, Any]) -> tuple[str, str, str]:
    return (
        str(row.get("experiment_id", "")).strip(),
        str(row.get("timestamp", "")).strip(),
        str(row.get("observation", "")).strip(),
    )


def audit_keyed_rows(
    errors: list[dict[str, Any]],
    tables: dict[str, list[dict[str, Any]]],
    sheet_name: str,
    key_field: str,
    append_rows: list[dict[str, Any]],
    updates: list[dict[str, Any]],
) -> None:
    existing_keys = {
        str(row.get(key_field, "")).strip()
        for row in tables.get(sheet_name, [])
        if str(row.get(key_field, "")).strip()
    }
    pending_keys: set[str] = set()
    for row in append_rows:
        key_value = str(row.get(key_field, "")).strip()
        if not key_value:
            errors.append(
                {
                    "code": "missing_append_key",
                    "sheet": sheet_name,
                    "key": key_field,
                }
            )
        elif key_value in existing_keys or key_value in pending_keys:
            errors.append(
                {
                    "code": "duplicate_append",
                    "sheet": sheet_name,
                    "key": key_field,
                    "value": key_value,
                }
            )
        pending_keys.add(key_value)
    for update in updates:
        key_value = str(
            update.get(key_field, "") or update.get("key_value", "")
        ).strip()
        if key_value and key_value not in existing_keys:
            errors.append(
                {
                    "code": "missing_update_target",
                    "sheet": sheet_name,
                    "key": key_field,
                    "value": key_value,
                }
            )


def collect_rows(report: dict[str, Any], key: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for row in report.get(key, []) or []:
        if isinstance(row, dict):
            rows.append(row)
    for run in report.get("runs", []):
        for row in run.get(key, []) or []:
            if isinstance(row, dict):
                rows.append(row)
    return rows


def stringify_cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (list, tuple)):
        return ",".join(str(item) for item in value)
    return str(value)
