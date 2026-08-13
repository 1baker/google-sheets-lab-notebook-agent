from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from lab_notebook_agent.google_sheets import (
    google_scientist_workspace_upgrade_requests,
)
from lab_notebook_agent.scientist_workspace import (
    daily_log_rows_from_tables,
    google_reaction_master_array_formulas,
    plot_studio_measurement_formula,
    plot_studio_process_formula,
    reaction_master_excel_formula,
    result_rows_from_tables,
)
from lab_notebook_agent.templates import save_workbook


class ScientistWorkspaceTests(unittest.TestCase):
    def test_compact_entry_rows_project_into_canonical_science_tables(self) -> None:
        tables = {
            "Daily Log": [],
            "Bench Log": [
                {
                    "Run ID": "EP-900",
                    "Date & time": "2026-08-12T10:00:00",
                    "Stage": "feed",
                    "Temperature (°C)": 70,
                    "RPM": 300,
                    "Observation / action": "Feed started cleanly.",
                }
            ],
            "Results": [],
            "Measurements": [
                {
                    "Run ID": "EP-900",
                    "Sample ID": "EP-900-S1",
                    "Measurement": "DLS particle size",
                    "Numeric value": 185,
                    "Units": "nm",
                }
            ],
        }
        log = daily_log_rows_from_tables(tables)[0]
        result = result_rows_from_tables(tables)[0]
        self.assertEqual("EP-900", log["experiment_id"])
        self.assertEqual(70, log["temperature_C"])
        self.assertEqual("Feed started cleanly.", log["observation"])
        self.assertEqual("DLS particle size", result["measurement_type"])
        self.assertEqual(185, result["numeric_value"])
        self.assertEqual(185, result["value"])

    def test_plot_studio_formulas_follow_selected_run_and_metric(self) -> None:
        process = plot_studio_process_formula()
        measurement = plot_studio_measurement_formula()
        self.assertIn("'Bench Log'!$A$2:$A$1000=$B$3", process)
        self.assertIn("MATCH($B$4", process)
        self.assertIn("'Measurements'!$A$2:$A$1000=$B$3", measurement)
        self.assertIn("'Measurements'!$C$2:$C$1000=$B$5", measurement)

    def test_reaction_master_rolls_up_entry_tables_without_manual_copying(self) -> None:
        planned = reaction_master_excel_formula("Planned mass (g)", 2)
        activity = reaction_master_excel_formula("Last activity", 2)
        next_action = reaction_master_excel_formula("Next action", 2)
        google = google_reaction_master_array_formulas()
        self.assertIn("SUMIF('Batch Builder'!$A$2:$A$1000,$A2", planned)
        self.assertIn("'Bench Log'!$A$2:$A$1000", activity)
        self.assertIn("Measurements!$A$2:$A$1000", activity)
        self.assertIn("Enter batch quantities", next_action)
        self.assertIn("MAP(A2:A1000", google["Planned mass (g)"])
        self.assertIn("COUNTIF('Bench Log'!A2:A1000", google["Bench entries"])

    def test_generated_workbook_prioritizes_compact_entry_and_plotting(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = save_workbook(Path(tmpdir) / "scientist.xlsx")
            workbook = load_workbook(path, data_only=False)
            self.assertEqual("visible", workbook["Bench Log"].sheet_state)
            self.assertEqual("visible", workbook["Measurements"].sheet_state)
            self.assertEqual("visible", workbook["Plot Studio"].sheet_state)
            self.assertEqual("visible", workbook["Reaction Master"].sheet_state)
            self.assertEqual("hidden", workbook["Daily Log"].sheet_state)
            self.assertEqual("hidden", workbook["Results"].sheet_state)
            self.assertEqual("hidden", workbook["Plot Dashboard"].sheet_state)
            self.assertEqual("Run ID", workbook["Bench Log"]["A1"].value)
            self.assertEqual("Observation / action", workbook["Bench Log"]["K1"].value)
            self.assertEqual("PLOT STUDIO", workbook["Plot Studio"]["A1"].value)
            self.assertEqual(2, len(workbook["Plot Studio"]._charts))
            self.assertIn("SUMIF", workbook["Reaction Master"]["J2"].value)
            self.assertIn("COUNTIF", workbook["Reaction Master"]["N2"].value)
            self.assertIn("Continue run", workbook["Reaction Master"]["R2"].value)
            metadata = {
                row[0].value: row[1].value
                for row in workbook["Workbook Metadata"].iter_rows(min_row=2)
                if row[0].value
            }
            self.assertEqual("0.7.0", metadata["contract_version"])

    def test_live_upgrade_adds_entry_sheets_seeds_history_and_adds_two_charts(self) -> None:
        metadata = {
            "properties": {"timeZone": "America/Chicago"},
            "sheets": [
                {"properties": {"sheetId": index + 10, "title": title, "gridProperties": {"rowCount": 1000, "columnCount": 40}}}
                for index, title in enumerate(
                    (
                        "Run Console",
                        "Master Reagents",
                        "Experiments",
                        "Batch Builder",
                        "Daily Log",
                        "Formulations",
                        "Results",
                        "Plot Dashboard",
                    )
                )
            ],
        }
        tables = {
            "Daily Log": [{"experiment_id": "EP-001", "timestamp": "2026-08-12T10:00:00", "process_stage": "feed", "observation": "Stable."}],
            "Results": [{"experiment_id": "EP-001", "sample_id": "S1", "measurement_type": "DLS particle size", "numeric_value": 200, "units": "nm"}],
        }
        requests = google_scientist_workspace_upgrade_requests(metadata, tables)
        added_titles = {
            request["addSheet"]["properties"]["title"]
            for request in requests
            if "addSheet" in request
        }
        self.assertTrue({"Reaction Master", "Bench Log", "Measurements", "Plot Studio"}.issubset(added_titles))
        self.assertEqual(2, sum("addChart" in request for request in requests))
        self.assertTrue(any("Stable." in str(request) for request in requests))
        self.assertTrue(any("DLS particle size" in str(request) for request in requests))
        self.assertTrue(any("Bench Log" in str(request) for request in requests))
        self.assertTrue(any("Plot Studio" in str(request) for request in requests))
        self.assertTrue(any("Batch Builder" in str(request) and "formulaValue" in str(request) for request in requests))


if __name__ == "__main__":
    unittest.main()
