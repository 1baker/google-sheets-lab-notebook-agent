from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from lab_notebook_agent.batch_builder import (
    batch_builder_to_formulation_row,
    formulation_rows_from_tables,
    google_batch_builder_array_formulas,
)
from lab_notebook_agent.google_sheets import (
    generated_sheet_ids_for_missing,
    google_scientist_batch_builder_upgrade_requests,
    google_setup_requests_from_metadata,
)
from lab_notebook_agent.schema import sheet_by_name
from lab_notebook_agent.templates import save_workbook


class BatchBuilderTests(unittest.TestCase):
    def test_contract_exposes_scientist_quantity_inputs_and_calculations(self) -> None:
        headers = sheet_by_name("Batch Builder").headers
        for header in (
            "stage",
            "charge_type",
            "reagent_id",
            "parts_per_hundred_monomer",
            "stage_monomer_basis_g",
            "direct_mass_g",
            "planned_mass_g",
            "stock_active_fraction",
            "active_mass_g",
            "density_g_mL",
            "planned_volume_mL",
            "actual_mass_g",
            "mass_variance_g",
            "feed_duration_min",
            "feed_rate_mL_min",
            "target_temperature_C",
            "lot",
            "charge_status",
        ):
            self.assertIn(header, headers)

    def test_google_formulas_support_direct_mass_and_phr_scaling(self) -> None:
        formulas = google_batch_builder_array_formulas(1000)
        self.assertIn("I2:I1000*J2:J1000/100", formulas["planned_mass_g"])
        self.assertIn("K2:K1000", formulas["planned_mass_g"])
        self.assertIn("'Master Reagents'!A2:G1000", formulas["density_g_mL"])
        self.assertIn("R2:R1000-L2:L1000", formulas["mass_variance_g"])
        self.assertIn("Q2:Q1000/V2:V1000", formulas["feed_rate_mL_min"])
        self.assertTrue(
            all(formula.count("(") == formula.count(")") for formula in formulas.values())
        )

    def test_generated_excel_batch_builder_has_formula_and_input_zones(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = save_workbook(Path(tmpdir) / "template.xlsx")
            workbook = load_workbook(path, data_only=False)
            worksheet = workbook["Batch Builder"]
            self.assertIn("VLOOKUP", worksheet["H2"].value)
            self.assertIn("$I2*$J2/100", worksheet["L2"].value)
            self.assertIn("$R2-$L2", worksheet["S2"].value)
            self.assertEqual("FFF9E3", worksheet["K2"].fill.fgColor.rgb[-6:])
            self.assertEqual("EAF4F7", worksheet["L2"].fill.fgColor.rgb[-6:])
            self.assertEqual("0.########", worksheet["L2"].number_format)
            self.assertEqual("C2", worksheet.freeze_panes)

    def test_google_setup_installs_batch_builder_formulas_and_id_validations(self) -> None:
        requests = google_setup_requests_from_metadata({"properties": {}, "sheets": []})
        sheet_id = generated_sheet_ids_for_missing({})["Batch Builder"]
        formulas = [
            value["userEnteredValue"]["formulaValue"]
            for request in requests
            for row in request.get("updateCells", {}).get("rows", [])
            for value in row.get("values", [])
            if request.get("updateCells", {}).get("start", {}).get("sheetId") == sheet_id
            and "formulaValue" in value.get("userEnteredValue", {})
        ]
        self.assertEqual(8, len(formulas))
        self.assertTrue(any("ARRAYFORMULA" in formula for formula in formulas))
        id_validations = [
            request["setDataValidation"]
            for request in requests
            if request.get("setDataValidation", {}).get("range", {}).get("sheetId")
            == sheet_id
            and request.get("setDataValidation", {}).get("rule", {}).get("condition", {}).get("type")
            == "ONE_OF_RANGE"
        ]
        self.assertEqual(2, len(id_validations))

    def test_bounded_upgrade_adds_batch_builder_hides_legacy_and_refreshes_console(self) -> None:
        metadata = {
            "properties": {},
            "sheets": [
                {"properties": {"sheetId": 10, "title": "Run Console", "gridProperties": {"rowCount": 100, "columnCount": 12}}},
                {"properties": {"sheetId": 11, "title": "Experiments", "gridProperties": {"rowCount": 1000, "columnCount": 26}}},
                {"properties": {"sheetId": 12, "title": "Master Reagents", "gridProperties": {"rowCount": 1000, "columnCount": 26}}},
                {"properties": {"sheetId": 13, "title": "Formulations", "gridProperties": {"rowCount": 1000, "columnCount": 20}}},
                *[
                    {"properties": {"sheetId": 100 + index, "title": name, "gridProperties": {"rowCount": 1000, "columnCount": 40}}}
                    for index, name in enumerate(
                        ("Run Capture Plan", "Daily Log", "Samples", "Results", "Raw Data Files", "Deviations", "Plot Dashboard")
                    )
                ],
            ],
        }
        requests = google_scientist_batch_builder_upgrade_requests(metadata)
        self.assertTrue(any("addSheet" in request and request["addSheet"]["properties"]["title"] == "Batch Builder" for request in requests))
        self.assertTrue(any(request.get("updateSheetProperties", {}).get("properties", {}).get("sheetId") == 13 and request["updateSheetProperties"]["properties"].get("hidden") is True for request in requests))
        self.assertTrue(any("Enter batch quantities" in str(request) for request in requests))

    def test_batch_builder_rows_feed_existing_scientific_analysis(self) -> None:
        charge = {
            "experiment_id": "EP-900",
            "charge_id": "EP-900-CHG-001",
            "stage": "core",
            "charge_type": "monomer_pre_emulsion",
            "target_role": "core_monomer",
            "reagent_id": "M-BA",
            "material_name": "Butyl Acrylate",
            "planned_mass_g": 99.7,
            "planned_volume_mL": 112.02,
            "actual_mass_g": 99.6,
            "mass_variance_g": -0.1,
            "feed_start_min": 0,
            "feed_duration_min": 180,
            "feed_rate_mL_min": 0.622,
            "lot": "BA-LOT-1",
            "charge_status": "charged",
        }
        projected = batch_builder_to_formulation_row(charge)
        self.assertEqual("core / monomer_pre_emulsion", projected["phase"])
        self.assertEqual(99.7, projected["mass_g"])
        self.assertEqual("EP-900-CHG-001", projected["charge_id"])
        combined = formulation_rows_from_tables(
            {"Formulations": [], "Batch Builder": [charge]}
        )
        self.assertEqual([projected], combined)


if __name__ == "__main__":
    unittest.main()
