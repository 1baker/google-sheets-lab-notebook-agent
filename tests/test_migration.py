from __future__ import annotations

import unittest

from lab_notebook_agent.google_sheets import (
    append_cells_request,
    generated_sheet_ids_for_missing,
    google_contract_migration_requests,
    google_setup_requests_from_metadata,
    replace_sheet_rows_requests,
)
from lab_notebook_agent.schema import RUN_CONSOLE_SHEET, SHEETS, workbook_contract


class WorkbookMigrationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.sheet_ids = {
            spec.name: index
            for index, spec in enumerate(SHEETS, start=100)
        }

    def test_contract_adds_operational_lineage_and_qc_tabs(self) -> None:
        contract = workbook_contract()
        self.assertEqual("0.5.0", contract["version"])
        self.assertEqual(RUN_CONSOLE_SHEET, contract["views"][0]["name"])
        self.assertEqual("B3", contract["views"][0]["active_experiment_cell"])
        self.assertEqual("A30", contract["views"][0]["active_queue_anchor"])
        names = {sheet["name"] for sheet in contract["sheets"]}
        self.assertTrue(
            {
                "Workbook Metadata",
                "Run Capture Plan",
                "Batch Builder",
                "Samples",
                "Equipment",
                "Protocols",
                "Specifications",
                "Deviations",
                "Raw Data Files",
                "Audit Log",
            }.issubset(names)
        )
        results = next(
            sheet for sheet in contract["sheets"] if sheet["name"] == "Results"
        )
        numeric_value = next(
            column
            for column in results["columns"]
            if column["name"] == "numeric_value"
        )
        self.assertEqual("number", numeric_value["data_type"])

    def test_setup_adds_timezone_notes_number_formats_and_filters(self) -> None:
        requests = google_setup_requests_from_metadata(
            {
                "properties": {"timeZone": "America/Los_Angeles"},
                "sheets": [],
            }
        )
        self.assertEqual(
            "America/Chicago",
            requests[0]["updateSpreadsheetProperties"]["properties"]["timeZone"],
        )
        result_id = generated_sheet_ids_for_missing({})["Results"]
        result_header = next(
            request["updateCells"]
            for request in requests
            if request.get("updateCells", {}).get("start", {}).get("sheetId")
            == result_id
        )
        self.assertIn(
            "Machine-readable numeric result",
            result_header["rows"][0]["values"][10]["note"],
        )
        self.assertTrue(
            any(
                request.get("repeatCell", {})
                .get("cell", {})
                .get("userEnteredFormat", {})
                .get("numberFormat", {})
                .get("type")
                == "NUMBER"
                for request in requests
            )
        )
        self.assertTrue(any("setBasicFilter" in request for request in requests))

    def test_setup_builds_formula_driven_run_console_and_bounded_presentation(self) -> None:
        requests = google_setup_requests_from_metadata({"properties": {}, "sheets": []})
        console_id = generated_sheet_ids_for_missing({})[RUN_CONSOLE_SHEET]
        console_formula_cells = [
            value["userEnteredValue"]["formulaValue"]
            for request in requests
            for row in request.get("updateCells", {}).get("rows", [])
            for value in row.get("values", [])
            if "formulaValue" in value.get("userEnteredValue", {})
            and request.get("updateCells", {}).get("start", {}).get("sheetId")
            == console_id
        ]
        self.assertTrue(console_formula_cells)
        self.assertTrue(any("Not recorded" in formula for formula in console_formula_cells))
        self.assertTrue(any("FILTER" in formula for formula in console_formula_cells))
        self.assertTrue(any("MAP(" in formula for formula in console_formula_cells))
        self.assertTrue(any("Assign operator" in formula for formula in console_formula_cells))
        self.assertTrue(any("COUNTIF" in formula for formula in console_formula_cells))
        self.assertTrue(
            all(formula.count("(") == formula.count(")") for formula in console_formula_cells)
        )
        self.assertTrue(any("source_notebook_id" not in formula and "$L$2:$L$1000" in formula for formula in console_formula_cells))
        self.assertTrue(
            any(
                request.get("repeatCell", {}).get("range", {}).get("startColumnIndex") == 6
                and request.get("repeatCell", {}).get("range", {}).get("endColumnIndex") == 7
                and request.get("repeatCell", {})
                .get("cell", {})
                .get("userEnteredFormat", {})
                .get("numberFormat", {})
                .get("type")
                == "PERCENT"
                for request in requests
            )
        )
        self.assertFalse(any("autoResizeDimensions" in request for request in requests))
        self.assertTrue(
            any(
                request.get("updateSheetProperties", {})
                .get("properties", {})
                .get("hidden")
                is True
                for request in requests
            )
        )

    def test_migration_seeds_missing_reference_rows_and_is_idempotent(self) -> None:
        empty_tables = {spec.name: [] for spec in SHEETS}
        requests = google_contract_migration_requests(
            empty_tables,
            self.sheet_ids,
            migrated_at="2026-07-24T12:00:00+00:00",
        )
        appended_sheet_ids = {
            request["appendCells"]["sheetId"]
            for request in requests
            if "appendCells" in request
        }
        self.assertIn(self.sheet_ids["Controlled Vocab"], appended_sheet_ids)
        self.assertIn(self.sheet_ids["Agent Config"], appended_sheet_ids)
        self.assertIn(self.sheet_ids["Workbook Metadata"], appended_sheet_ids)
        self.assertIn(self.sheet_ids["Audit Log"], appended_sheet_ids)

        seeded_tables = {spec.name: [] for spec in SHEETS}
        for name in ("Process Knowledge", "Controlled Vocab", "Agent Config"):
            spec = next(spec for spec in SHEETS if spec.name == name)
            seeded_tables[name] = [
                {
                    header: values[index] if index < len(values) else ""
                    for index, header in enumerate(spec.headers)
                }
                for values in spec.example_rows
            ]
        seeded_tables["Workbook Metadata"] = [
            {
                "key": key,
                "value": value,
                "updated_at": "2026-07-24T12:00:00+00:00",
                "notes": notes,
            }
            for key, value, notes in (
                (
                    "contract_name",
                    "lab-notebook-agent-workbook",
                    "Machine-readable workbook contract.",
                ),
                (
                    "contract_version",
                    "0.5.0",
                    "Schema version currently applied to this workbook.",
                ),
                (
                    "workbook_timezone",
                    "America/Chicago",
                    "Timezone used for local laboratory timestamps.",
                ),
                (
                    "migration_status",
                    "current",
                    "Set by a successful contract migration.",
                ),
                (
                    "last_migrated_at",
                    "2026-07-24T12:00:00+00:00",
                    "UTC timestamp of the most recent contract migration.",
                ),
            )
        ]
        seeded_tables["Audit Log"] = [{"audit_id": "MIGRATION-0.5.0"}]
        rerun = google_contract_migration_requests(
            seeded_tables,
            self.sheet_ids,
            migrated_at="2026-07-24T12:00:00+00:00",
        )
        self.assertFalse(any("appendCells" in request for request in rerun))

    def test_append_cells_writes_contract_numbers_and_dates_as_numbers(self) -> None:
        request = append_cells_request(
            "Experiments",
            [
                {
                    "experiment_id": "EP-TYPED",
                    "date": "2026-07-24",
                    "process_type": "emulsion polymerization",
                    "objective": "Verify typed dates.",
                    "status": "planned",
                }
            ],
            self.sheet_ids,
        )
        values = request["appendCells"]["rows"][0]["values"]
        self.assertIn("numberValue", values[1]["userEnteredValue"])

        request = append_cells_request(
            "Formulations",
            [
                {
                    "experiment_id": "EP-TYPED",
                    "reagent_id": "R-1",
                    "target_role": "surfactant",
                    "mass_g": "12.5",
                }
            ],
            self.sheet_ids,
        )
        values = request["appendCells"]["rows"][0]["values"]
        self.assertEqual(
            12.5,
            values[4]["userEnteredValue"]["numberValue"],
        )

    def test_plot_dashboard_reserves_a_canvas_for_managed_charts(self) -> None:
        requests = replace_sheet_rows_requests(
            "Plot Dashboard",
            [],
            0,
            self.sheet_ids,
        )
        self.assertGreaterEqual(
            requests[0]["updateSheetProperties"]["properties"]["gridProperties"][
                "columnCount"
            ],
            28,
        )


if __name__ == "__main__":
    unittest.main()
