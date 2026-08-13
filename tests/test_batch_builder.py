from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from lab_notebook_agent.batch_builder import (
    batch_builder_to_formulation_row,
    calculate_batch_mass_row,
    formulation_rows_from_tables,
    google_batch_builder_array_formulas,
    summarize_experiment_mass_plan,
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
            "molecular_weight_g_mol",
            "planned_moles_mmol",
            "equivalent_basis_mmol",
            "equivalents",
            "actual_moles_mmol",
            "calculation_mode",
            "recipe_wt_percent",
            "experiment_target_mass_g",
            "target_active_mass_g",
            "target_equivalents",
            "mass_tolerance_percent",
            "experiment_default_tolerance_percent",
            "mass_variance_percent",
            "within_tolerance",
            "formula_status",
            "functional_group",
            "functional_equivalent_weight_g_eq",
            "target_functional_equivalents",
            "planned_functional_equivalents",
            "actual_functional_equivalents",
            "weighing_method",
            "source_container_before_g",
            "source_container_after_g",
            "effective_actual_mass_g",
            "planned_carrier_mass_g",
            "actual_carrier_mass_g",
            "weighing_status",
        ):
            self.assertIn(header, headers)

    def test_google_formulas_support_direct_mass_and_phr_scaling(self) -> None:
        formulas = google_batch_builder_array_formulas(1000)
        self.assertIn('"batch_wt_percent"', formulas["planned_mass_g"])
        self.assertIn("wtpct*target/100", formulas["planned_mass_g"])
        self.assertIn('"pphm"', formulas["planned_mass_g"])
        self.assertIn("pphm*basis/100", formulas["planned_mass_g"])
        self.assertIn('"target_active_mass"', formulas["planned_mass_g"])
        self.assertIn('"equivalents"', formulas["planned_mass_g"])
        self.assertIn('"functional_equivalents"', formulas["planned_mass_g"])
        self.assertIn("'Master Reagents'!A2:G1000", formulas["density_g_mL"])
        self.assertIn('(L2:L1000="")', formulas["mass_variance_g"])
        self.assertIn("BB2:BB1000-L2:L1000", formulas["mass_variance_g"])
        self.assertIn("Q2:Q1000/V2:V1000", formulas["feed_rate_mL_min"])
        self.assertIn("'Master Reagents'!A2:F1000", formulas["molecular_weight_g_mol"])
        self.assertIn("N2:N1000/AD2:AD1000*1000", formulas["planned_moles_mmol"])
        self.assertIn("AE2:AE1000/AF2:AF1000", formulas["equivalents"])
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
            self.assertIn('OR($BB2="",$L2="")', worksheet["S2"].value)
            self.assertIn("$BB2-$L2", worksheet["S2"].value)
            self.assertEqual("FFF9E3", worksheet["K2"].fill.fgColor.rgb[-6:])
            self.assertEqual("EAF4F7", worksheet["L2"].fill.fgColor.rgb[-6:])
            self.assertEqual("0.########", worksheet["L2"].number_format)
            self.assertEqual("C2", worksheet.freeze_panes)
            self.assertIn("VLOOKUP", worksheet["AD2"].value)
            self.assertIn("$N2/$AD2*1000", worksheet["AE2"].value)
            self.assertEqual("FFF9E3", worksheet["AF2"].fill.fgColor.rgb[-6:])
            self.assertEqual("EAF4F7", worksheet["AG2"].fill.fgColor.rgb[-6:])
            self.assertEqual("FFF9E3", worksheet["AI2"].fill.fgColor.rgb[-6:])
            self.assertEqual("EAF4F7", worksheet["AK2"].fill.fgColor.rgb[-6:])
            self.assertIn("MISSING_MODE", worksheet["AR2"].value)
            self.assertIn("MISSING_MODE", worksheet["AR1000"].value)
            self.assertIn("Master Reagents", worksheet["AS2"].value)
            self.assertIn("$AY2-$AZ2", worksheet["BA2"].value)
            self.assertIn("by_difference", worksheet["BB2"].value)
            validations = {
                str(validation.sqref): str(validation.formula1)
                for validation in worksheet.data_validations.dataValidation
            }
            self.assertIn("AI2:AI1000", validations)
            self.assertIn("batch_wt_percent", validations["AI2:AI1000"])
            conditional_ranges = {
                str(key.sqref) for key in worksheet.conditional_formatting._cf_rules
            }
            self.assertTrue({"AP2:AP1000", "AQ2:AQ1000", "AR2:AR1000"}.issubset(conditional_ranges))
            variance_rules = next(
                rules
                for key, rules in worksheet.conditional_formatting._cf_rules.items()
                if str(key.sqref) == "AP2:AP1000"
            )
            self.assertIn('IF($AN2<>"",$AN2,$AO2)>0', variance_rules[0].formula[0])

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
        self.assertEqual(26, len(formulas))
        self.assertTrue(any("ARRAYFORMULA" in formula for formula in formulas))
        self.assertTrue(any("MAP(" in formula for formula in formulas))
        id_validations = [
            request["setDataValidation"]
            for request in requests
            if request.get("setDataValidation", {}).get("range", {}).get("sheetId")
            == sheet_id
            and request.get("setDataValidation", {}).get("rule", {}).get("condition", {}).get("type")
            == "ONE_OF_RANGE"
        ]
        self.assertEqual(2, len(id_validations))
        mode_validations = [
            request["setDataValidation"]
            for request in requests
            if request.get("setDataValidation", {}).get("range", {}).get("sheetId")
            == sheet_id
            and request.get("setDataValidation", {}).get("range", {}).get("startColumnIndex")
            == sheet_by_name("Batch Builder").headers.index("calculation_mode")
        ]
        self.assertEqual(1, len(mode_validations))
        self.assertTrue(
            any(
                value["userEnteredValue"] == "batch_wt_percent"
                for value in mode_validations[0]["rule"]["condition"]["values"]
            )
        )

    def test_mass_modes_match_senior_scientist_acceptance_examples(self) -> None:
        cases = (
            (
                {"calculation_mode": "batch_wt_percent", "experiment_target_mass_g": 500, "recipe_wt_percent": 40},
                200,
            ),
            (
                {"calculation_mode": "pphm", "parts_per_hundred_monomer": 2, "stage_monomer_basis_g": 300},
                6,
            ),
            (
                {"calculation_mode": "target_active_mass", "target_active_mass_g": 5, "stock_active_fraction": 0.25},
                20,
            ),
            (
                {
                    "calculation_mode": "equivalents",
                    "target_equivalents": 0.5,
                    "equivalent_basis_mmol": 100,
                    "molecular_weight_g_mol": 200,
                    "stock_active_fraction": 0.5,
                },
                20,
            ),
            (
                {
                    "calculation_mode": "functional_equivalents",
                    "target_functional_equivalents": 0.05,
                    "functional_equivalent_weight_g_eq": 200,
                    "stock_active_fraction": 0.5,
                },
                20,
            ),
        )
        for row, expected_mass in cases:
            with self.subTest(mode=row["calculation_mode"]):
                result = calculate_batch_mass_row(row)
                self.assertEqual("READY", result["formula_status"])
                self.assertAlmostEqual(expected_mass, result["planned_mass_g"])
        equivalent = calculate_batch_mass_row(cases[-2][0])
        self.assertAlmostEqual(10, equivalent["active_mass_g"])
        self.assertAlmostEqual(50, equivalent["planned_moles_mmol"])
        self.assertAlmostEqual(0.5, equivalent["equivalents"])

    def test_by_difference_weighing_and_carrier_accounting(self) -> None:
        result = calculate_batch_mass_row(
            {
                "calculation_mode": "direct_mass",
                "direct_mass_g": 20,
                "stock_active_fraction": 0.25,
                "functional_equivalent_weight_g_eq": 50,
                "weighing_method": "by_difference",
                "source_container_before_g": 100,
                "source_container_after_g": 80.3,
                "experiment_default_tolerance_percent": 2,
            }
        )
        self.assertAlmostEqual(19.7, result["effective_actual_mass_g"])
        self.assertAlmostEqual(-0.3, result["mass_variance_g"])
        self.assertAlmostEqual(15, result["planned_carrier_mass_g"])
        self.assertAlmostEqual(14.775, result["actual_carrier_mass_g"])
        self.assertEqual("READY", result["weighing_status"])

        invalid = calculate_batch_mass_row(
            {
                "calculation_mode": "direct_mass",
                "direct_mass_g": 20,
                "weighing_method": "by_difference",
                "source_container_before_g": 80,
                "source_container_after_g": 100,
            }
        )
        self.assertEqual("INVALID_CONTAINER_DIFFERENCE", invalid["weighing_status"])
        self.assertIsNone(invalid["effective_actual_mass_g"])

    def test_run_summary_audits_functional_equivalent_ratio(self) -> None:
        experiment = {
            "experiment_id": "PU-1",
            "target_batch_mass_g": 63.125,
            "default_mass_tolerance_percent": 1,
            "stoichiometric_numerator_group": "NCO",
            "stoichiometric_denominator_group": "active_H",
            "target_equivalent_ratio": 1.05,
            "equivalent_ratio_tolerance_percent": 0.5,
        }
        charges = [
            {"experiment_id": "PU-1", "calculation_mode": "direct_mass", "direct_mass_g": 13.125, "functional_group": "NCO", "functional_equivalent_weight_g_eq": 125},
            {"experiment_id": "PU-1", "calculation_mode": "direct_mass", "direct_mass_g": 50, "functional_group": "active_H", "functional_equivalent_weight_g_eq": 500},
        ]
        summary = summarize_experiment_mass_plan(experiment, charges)
        self.assertAlmostEqual(1.05, summary["planned_equivalent_ratio"])
        self.assertEqual("READY", summary["stoichiometry_status"])

    def test_actual_mass_error_uses_row_or_experiment_tolerance(self) -> None:
        base = {
            "calculation_mode": "direct_mass",
            "direct_mass_g": 20,
            "actual_mass_g": 19.7,
            "experiment_default_tolerance_percent": 2,
        }
        passing = calculate_batch_mass_row(base)
        failing = calculate_batch_mass_row({**base, "mass_tolerance_percent": 1})
        self.assertAlmostEqual(-0.3, passing["mass_variance_g"])
        self.assertAlmostEqual(-1.5, passing["mass_variance_percent"])
        self.assertEqual("PASS", passing["within_tolerance"])
        self.assertEqual("FAIL", failing["within_tolerance"])

    def test_invalid_or_missing_inputs_never_return_plausible_mass(self) -> None:
        cases = (
            ({}, "MISSING_MODE"),
            ({"calculation_mode": "batch_wt_percent", "recipe_wt_percent": 40}, "MISSING_TARGET_BATCH_MASS"),
            ({"calculation_mode": "pphm", "parts_per_hundred_monomer": 2}, "MISSING_PPHM_BASIS"),
            ({"calculation_mode": "target_active_mass"}, "MISSING_ACTIVE_TARGET"),
            ({"calculation_mode": "equivalents", "target_equivalents": 1, "equivalent_basis_mmol": 100}, "MISSING_MOLECULAR_WEIGHT"),
            ({"calculation_mode": "direct_mass", "direct_mass_g": 10, "stock_active_fraction": 1.2}, "INVALID_ACTIVE_FRACTION"),
        )
        for row, expected_status in cases:
            with self.subTest(expected_status=expected_status):
                result = calculate_batch_mass_row(row)
                self.assertEqual(expected_status, result["formula_status"])
                self.assertIsNone(result["planned_mass_g"])

        invalid_with_actual = calculate_batch_mass_row(
            {"calculation_mode": "direct_mass", "actual_mass_g": 7.5}
        )
        self.assertIsNone(invalid_with_actual["mass_variance_g"])
        self.assertIsNone(invalid_with_actual["mass_variance_percent"])

    def test_nonpositive_physical_properties_do_not_generate_values(self) -> None:
        result = calculate_batch_mass_row(
            {
                "calculation_mode": "direct_mass",
                "direct_mass_g": 10,
                "actual_mass_g": 9.9,
                "density_g_mL": -1,
                "molecular_weight_g_mol": -100,
            }
        )
        self.assertEqual("READY", result["formula_status"])
        self.assertIsNone(result["planned_volume_mL"])
        self.assertIsNone(result["planned_moles_mmol"])
        self.assertIsNone(result["actual_moles_mmol"])

    def test_run_mass_summary_detects_ready_mismatch_and_charge_failure(self) -> None:
        experiment = {
            "experiment_id": "RXN-500",
            "target_batch_mass_g": 500,
            "default_mass_tolerance_percent": 2,
        }
        ready_rows = [
            {
                "experiment_id": "RXN-500",
                "calculation_mode": "batch_wt_percent",
                "recipe_wt_percent": 40,
                "actual_mass_g": 198,
            },
            {
                "experiment_id": "RXN-500",
                "calculation_mode": "batch_wt_percent",
                "recipe_wt_percent": 60,
                "actual_mass_g": 297,
            },
        ]
        ready = summarize_experiment_mass_plan(experiment, ready_rows)
        mismatch = summarize_experiment_mass_plan(
            experiment, ready_rows[:1]
        )
        charge_failure = summarize_experiment_mass_plan(
            experiment,
            [{**ready_rows[0], "actual_mass_g": 180}, ready_rows[1]],
        )
        self.assertEqual(500, ready["planned_mass_g"])
        self.assertEqual(-5, ready["actual_vs_target_g"])
        self.assertEqual("READY", ready["status"])
        self.assertEqual("MASS_MISMATCH", mismatch["status"])
        self.assertEqual("CHARGE_TOLERANCE_FAILURE", charge_failure["status"])

    def test_run_mass_summary_never_implies_unproven_balance(self) -> None:
        direct = {
            "experiment_id": "RXN-GUARD",
            "calculation_mode": "direct_mass",
            "direct_mass_g": 20,
        }
        no_target = summarize_experiment_mass_plan(
            {"experiment_id": "RXN-GUARD", "default_mass_tolerance_percent": 2},
            [direct],
        )
        missing_tolerance = summarize_experiment_mass_plan(
            {"experiment_id": "RXN-GUARD", "target_batch_mass_g": 20},
            [direct],
        )
        calculation_issue = summarize_experiment_mass_plan(
            {
                "experiment_id": "RXN-GUARD",
                "target_batch_mass_g": 20,
                "default_mass_tolerance_percent": 2,
            },
            [{"experiment_id": "RXN-GUARD", "calculation_mode": "direct_mass"}],
        )
        no_charges = summarize_experiment_mass_plan(
            {
                "experiment_id": "RXN-GUARD",
                "target_batch_mass_g": 20,
                "default_mass_tolerance_percent": 2,
            },
            [],
        )
        self.assertEqual("NO_TARGET_MASS", no_target["status"])
        self.assertEqual("MISSING_TOLERANCE", missing_tolerance["status"])
        self.assertEqual("CALCULATION_ISSUES", calculation_issue["status"])
        self.assertEqual("NO_CHARGES", no_charges["status"])

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
            "planned_moles_mmol": 874.6,
            "molecular_weight_g_mol": 128.17,
            "equivalents": 1.0,
        }
        projected = batch_builder_to_formulation_row(charge)
        self.assertEqual("core / monomer_pre_emulsion", projected["phase"])
        self.assertEqual(99.7, projected["mass_g"])
        self.assertEqual("EP-900-CHG-001", projected["charge_id"])
        self.assertEqual(874.6, projected["moles_mmol"])
        combined = formulation_rows_from_tables(
            {"Formulations": [], "Batch Builder": [charge]}
        )
        self.assertEqual([projected], combined)


if __name__ == "__main__":
    unittest.main()
