from __future__ import annotations

import unittest

from lab_notebook_agent.preflight import build_experiment_preflight_report


def complete_record_tables() -> dict[str, list[dict[str, object]]]:
    sections = []
    for index, section_type in enumerate(
        ("objective", "safety", "setup", "procedure", "observations", "workup", "results", "conclusion"),
        start=1,
    ):
        sections.append(
            {
                "experiment_id": "RXN-1",
                "section_id": f"RXN-1-SEC-{index:03d}",
                "section_type": section_type,
                "required": "true",
                "status": "complete",
                "content": f"Completed {section_type} narrative.",
                "authored_by": "Operator",
                "authored_at": "2026-08-13T09:00:00",
                "completed_by": "Operator",
                "completed_at": "2026-08-13T16:00:00",
            }
        )
    return {
        "Experiments": [{
            "experiment_id": "RXN-1", "date": "2026-08-13", "process_type": "small molecule synthesis",
            "objective": "Test the controlled route.", "operator": "Operator", "status": "complete",
            "protocol_id": "P-1", "protocol_version": "2", "equipment_id": "R-1",
            "template_id": "T-1", "template_version": "3", "summary": "Run completed.",
            "completed_at": "2026-08-13T16:00:00", "reviewer": "Reviewer",
            "reviewed_at": "2026-08-13T17:00:00", "record_fingerprint": "sha256:abc",
        }],
        "Experiment Templates": [{
            "template_id": "T-1", "version": "3", "state": "effective", "owner": "Steward",
            "effective_at": "2026-08-01T09:00:00", "required_capture_sections": "all",
            "source_url": "https://example.test/template/3",
        }],
        "Protocols": [{
            "protocol_id": "P-1", "version": "2", "status": "active",
            "source_url": "https://example.test/protocol/2",
        }],
        "Equipment": [{"equipment_id": "R-1", "calibration_status": "current"}],
        "Master Reagents": [{
            "reagent_id": "A", "name": "Reagent A", "hazards": "SDS reviewed",
            "molecular_weight_g_mol": 100, "density_g_mL": 1,
        }],
        "Formulations": [{
            "experiment_id": "RXN-1", "reagent_id": "A", "target_role": "reactant", "mass_g": 10,
        }],
        "Batch Builder": [{
            "experiment_id": "RXN-1", "charge_id": "RXN-1-CHG-001", "stage": "reaction",
            "charge_type": "initial", "target_role": "reactant", "feed_order": 1, "reagent_id": "A",
            "formula_status": "READY", "effective_actual_mass_g": 10, "lot": "LOT-A",
            "recorded_by": "Operator", "recorded_at": "2026-08-13T10:00:00", "charge_status": "charged",
        }],
        "Notebook Sections": sections,
        "Bench Log": [{
            "Run ID": "RXN-1", "Date & time": "2026-08-13T10:05:00",
            "Observation / action": "Charge complete.",
        }],
        "Measurements": [{
            "Run ID": "RXN-1", "Sample ID": "RXN-1-P", "Measurement": "NMR purity",
            "Numeric value": 98, "Units": "%", "Raw file ID": "RAW-1",
        }],
        "Raw Data Files": [{
            "raw_file_id": "RAW-1", "experiment_id": "RXN-1", "instrument_id": "NMR-1",
            "collected_at": "2026-08-13T15:00:00", "file_name": "rxn1.fid",
            "file_url": "https://example.test/raw/rxn1.fid",
        }],
        "Reaction Outcomes": [{
            "experiment_id": "RXN-1", "product_name": "Product", "theoretical_product_mass_g": 10,
            "recovered_product_mass_g": 9, "purity_percent": 98, "appearance": "white solid",
            "outcome_status": "complete", "mass_balance_status": "CLOSED", "completed_by": "Operator",
            "completed_at": "2026-08-13T16:00:00", "conclusion": "Product obtained.",
        }],
        "Deviations": [],
        "Record Signatures": [{
            "experiment_id": "RXN-1", "status": "signed", "signer": "Operator",
            "signed_at": "2026-08-13T17:05:00", "record_fingerprint": "sha256:abc",
            "witnessed_by": "Reviewer", "witnessed_at": "2026-08-13T17:10:00",
        }],
        "Literature Evidence": [],
        "Agent Suggestions": [],
        "Agent Config": [],
        "Project Notebook Records": [],
        "Daily Log": [],
        "Results": [],
    }


class RecordQualityTests(unittest.TestCase):
    def test_complete_archive_record_passes_all_release_gates(self) -> None:
        report = build_experiment_preflight_report(complete_record_tables(), "RXN-1", "archive")
        self.assertTrue(report["ready_to_archive"], report["checks"])
        self.assertEqual(0, report["summary"]["fail_count"])

    def test_review_blocks_missing_raw_data_and_incomplete_required_section(self) -> None:
        tables = complete_record_tables()
        tables["Measurements"][0]["Raw file ID"] = ""
        tables["Notebook Sections"][-1]["status"] = "in_progress"
        report = build_experiment_preflight_report(tables, "RXN-1", "review")
        checks = {row["name"]: row for row in report["checks"]}
        self.assertEqual("fail", checks["raw_data_traceability"]["status"])
        self.assertEqual("fail", checks["notebook_sections"]["status"])
        self.assertFalse(report["ready_to_run"])

    def test_archive_signature_must_match_current_record_fingerprint(self) -> None:
        tables = complete_record_tables()
        tables["Record Signatures"][0]["record_fingerprint"] = "sha256:old"
        report = build_experiment_preflight_report(tables, "RXN-1", "archive")
        checks = {row["name"]: row for row in report["checks"]}
        self.assertEqual("fail", checks["signature_integrity"]["status"])
        self.assertFalse(report["ready_to_archive"])


if __name__ == "__main__":
    unittest.main()
