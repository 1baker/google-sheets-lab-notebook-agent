from __future__ import annotations

from typing import Any

from .agent import agent_config_values, suggestions_for_experiment
from .materials import audit_experiment_materials, nonblank
from .sheets import build_experiment_entry_from_tables


EXPERIMENT_REQUIRED_FIELDS = ("experiment_id", "date", "process_type", "objective", "status")
STAGES = ("planning", "review", "archive")
PRE_RUN_SECTION_TYPES = ("objective", "safety", "procedure")


def build_experiment_preflight_report(
    tables: dict[str, list[dict[str, Any]]],
    experiment_id: str,
    stage: str = "planning",
) -> dict[str, Any]:
    stage = normalize_stage(stage)
    try:
        entry = build_experiment_entry_from_tables(tables, experiment_id)
    except ValueError as exc:
        checks = [
            check(
                "experiment_row",
                "fail",
                str(exc),
                actions=["Add an Experiments row with this experiment_id before scaffolding or review."],
            )
        ]
        return report_payload(experiment_id, stage, {}, {}, checks)

    material_audit = audit_experiment_materials(entry)
    checks = [
        experiment_required_fields_check(entry),
        accountable_run_identity_check(entry),
        governed_template_check(tables, entry),
        controlled_protocol_check(tables, entry),
        equipment_readiness_check(tables, entry),
        material_roles_check(material_audit),
        formulation_quantities_check(material_audit),
        reagent_properties_check(material_audit),
        reagent_safety_check(entry, safety_review_required(tables)),
        placeholder_reagents_check(entry),
        batch_builder_check(tables, experiment_id, stage),
        notebook_sections_check(tables, experiment_id, stage),
        observations_check(entry, stage),
        results_check(entry, stage),
        raw_data_traceability_check(tables, experiment_id, stage),
        reaction_outcome_check(tables, experiment_id, stage),
        deviations_check(tables, experiment_id, stage),
        review_attribution_check(entry, stage),
        signature_integrity_check(tables, entry, stage),
        literature_evidence_check(entry),
        open_suggestions_check(tables, experiment_id),
    ]
    return report_payload(experiment_id, stage, entry, material_audit, checks)


def normalize_stage(stage: str) -> str:
    normalized = str(stage or "").strip().lower()
    if normalized not in STAGES:
        raise ValueError(f"Preflight stage must be one of {', '.join(STAGES)}.")
    return normalized


def experiment_required_fields_check(entry: dict[str, Any]) -> dict[str, Any]:
    missing = [field for field in EXPERIMENT_REQUIRED_FIELDS if not nonblank(entry.get(field))]
    if missing:
        return check(
            "experiment_required_fields",
            "fail",
            "Experiment row is missing required fields.",
            details={"missing_fields": missing},
            actions=[f"Fill Experiments.{field}." for field in missing],
        )
    return check(
        "experiment_required_fields",
        "pass",
        "Experiment row has the required planning fields.",
    )


def rows_for_experiment(
    tables: dict[str, list[dict[str, Any]]],
    sheet_name: str,
    experiment_id: str,
    *id_fields: str,
) -> list[dict[str, Any]]:
    fields = id_fields or ("experiment_id", "Run ID")
    return [
        row
        for row in tables.get(sheet_name, [])
        if isinstance(row, dict)
        and any(str(row.get(field, "")).strip() == experiment_id for field in fields)
    ]


def accountable_run_identity_check(entry: dict[str, Any]) -> dict[str, Any]:
    missing = [field for field in ("operator",) if not nonblank(entry.get(field))]
    if missing:
        return check(
            "accountable_run_identity",
            "fail",
            "The planned run does not identify the accountable operator.",
            details={"missing_fields": missing},
            actions=["Assign Experiments.operator before the run starts."],
        )
    return check("accountable_run_identity", "pass", "The run identifies its accountable operator.")


def governed_template_check(
    tables: dict[str, list[dict[str, Any]]], entry: dict[str, Any]
) -> dict[str, Any]:
    template_id = str(entry.get("template_id", "")).strip()
    template_version = str(entry.get("template_version", "")).strip()
    matches = [
        row
        for row in tables.get("Experiment Templates", [])
        if str(row.get("template_id", "")).strip() == template_id
        and str(row.get("version", "")).strip() == template_version
    ]
    if not template_id or not template_version:
        return check(
            "governed_template",
            "fail",
            "The run is not pinned to a governed template and version.",
            actions=["Select Experiments.template_id and template_version from an effective Experiment Templates row."],
        )
    if not matches or str(matches[-1].get("state", "")).strip().lower() != "effective":
        return check(
            "governed_template",
            "fail",
            "The selected template/version is missing or not effective.",
            details={"template_id": template_id, "template_version": template_version},
            actions=["Publish or select an effective template version before use."],
        )
    template = matches[-1]
    template_gaps = [
        field for field in ("owner", "effective_at", "required_capture_sections", "source_url")
        if not nonblank(template.get(field))
    ]
    if template_gaps:
        return check(
            "governed_template",
            "fail",
            "The effective template lacks publication or capture provenance.",
            details={"template_id": template_id, "template_version": template_version, "missing_fields": template_gaps},
            actions=["Complete the template owner, effective timestamp, required capture sections, and controlled source URL."],
        )
    return check(
        "governed_template",
        "pass",
        "The run is pinned to an effective template version.",
        details={"template_id": template_id, "template_version": template_version},
    )


def controlled_protocol_check(
    tables: dict[str, list[dict[str, Any]]], entry: dict[str, Any]
) -> dict[str, Any]:
    protocol_id = str(entry.get("protocol_id", "")).strip()
    protocol_version = str(entry.get("protocol_version", "")).strip()
    candidates = [
        row for row in tables.get("Protocols", [])
        if str(row.get("protocol_id", "")).strip() == protocol_id
    ]
    if protocol_version:
        candidates = [
            row for row in candidates
            if str(row.get("version", "")).strip() == protocol_version
        ]
    active = [row for row in candidates if str(row.get("status", "")).strip().lower() == "active"]
    if not protocol_id or not protocol_version:
        return check(
            "controlled_protocol",
            "fail",
            "No exact controlled protocol and version is linked to the run.",
            details={"protocol_id": protocol_id, "protocol_version": protocol_version},
            actions=["Select an active Protocols.protocol_id and record its version before execution."],
        )
    if not active:
        return check(
            "controlled_protocol",
            "fail",
            "The linked protocol/version is missing, retired, or still draft.",
            details={"protocol_id": protocol_id, "protocol_version": protocol_version},
            actions=["Link the run to an active Protocols row."],
        )
    protocol = active[-1]
    gaps = [field for field in ("version", "source_url") if not nonblank(protocol.get(field))]
    if gaps:
        return check(
            "controlled_protocol",
            "fail",
            "The active protocol lacks versioned source provenance.",
            details={"protocol_id": protocol_id, "missing_fields": gaps},
            actions=["Complete Protocols.version and Protocols.source_url."],
        )
    if protocol_version and str(protocol.get("version", "")).strip() != protocol_version:
        return check(
            "controlled_protocol",
            "fail",
            "The run's protocol version does not match the active protocol row.",
            actions=["Correct Experiments.protocol_version or select the intended active protocol version."],
        )
    return check(
        "controlled_protocol", "pass", "An active, versioned protocol with a controlled source is linked.",
        details={"protocol_id": protocol_id, "protocol_version": protocol.get("version", "")},
    )


def equipment_readiness_check(
    tables: dict[str, list[dict[str, Any]]], entry: dict[str, Any]
) -> dict[str, Any]:
    equipment_id = str(entry.get("equipment_id", "")).strip()
    matches = [
        row for row in tables.get("Equipment", [])
        if str(row.get("equipment_id", "")).strip() == equipment_id
    ]
    if not equipment_id or not matches:
        return check(
            "equipment_readiness",
            "fail",
            "The run does not identify registered equipment.",
            actions=["Select Experiments.equipment_id from Equipment."],
        )
    state = str(matches[-1].get("calibration_status", "")).strip().lower()
    if state not in {"current", "not_required"}:
        return check(
            "equipment_readiness",
            "fail",
            "The selected equipment is not calibration-ready.",
            details={"equipment_id": equipment_id, "calibration_status": state or "missing"},
            actions=["Resolve equipment calibration status before running."],
        )
    return check(
        "equipment_readiness", "pass", "The selected equipment is registered and calibration-ready.",
        details={"equipment_id": equipment_id, "calibration_status": state},
    )


def material_roles_check(material_audit: dict[str, Any]) -> dict[str, Any]:
    missing = material_audit.get("missing_required_role_groups", [])
    if missing:
        return check(
            "material_roles",
            "fail",
            "Required process material roles are missing.",
            details={"missing_required_role_groups": missing},
            actions=[
                "Run scaffold-materials or add Formulations rows for: " + ", ".join(str(item) for item in missing) + "."
            ],
        )
    return check(
        "material_roles",
        "pass",
        "Required process material roles are present.",
        details={"role_groups": material_audit.get("role_groups", [])},
    )


def formulation_quantities_check(material_audit: dict[str, Any]) -> dict[str, Any]:
    gaps = material_audit.get("quantity_gaps", [])
    if gaps:
        return check(
            "formulation_quantities",
            "fail",
            "One or more formulation rows lacks a quantitative basis.",
            details={"quantity_gaps": gaps},
            actions=[
                "Fill at least one of mass_g, volume_mL, moles_mmol, wt_percent, or concentration for each Formulations row."
            ],
        )
    return check(
        "formulation_quantities",
        "pass",
        "Every formulation row has a quantitative basis.",
    )


def reagent_properties_check(material_audit: dict[str, Any]) -> dict[str, Any]:
    gaps = material_audit.get("reagent_property_gaps", [])
    if gaps:
        return check(
            "reagent_properties",
            "fail",
            "Master Reagents is missing physical-property fields needed for calculations.",
            details={"reagent_property_gaps": gaps},
            actions=[
                "Fill missing Master Reagents fields such as molecular_weight_g_mol, density_g_mL, or concentration."
            ],
        )
    return check(
        "reagent_properties",
        "pass",
        "Master Reagents has the physical-property fields needed for the current roles.",
    )


def reagent_safety_check(entry: dict[str, Any], required: bool = True) -> dict[str, Any]:
    gaps = []
    for row in entry.get("formulation", []) or []:
        if not isinstance(row, dict) or not str(row.get("reagent_id", "")).strip():
            continue
        if reagent_has_safety_notes(row):
            continue
        gaps.append(
            {
                "reagent_id": row.get("reagent_id", ""),
                "target_role": row.get("target_role", ""),
                "missing_fields": ["hazards"],
            }
        )
    if gaps:
        return check(
            "reagent_safety",
            "fail" if required else "warn",
            "Master Reagents is missing hazards/SDS review notes for one or more formulation reagents.",
            details={"reagent_safety_gaps": gaps, "safety_review_required": required},
            actions=["Fill Master Reagents.hazards with SDS-reviewed handling notes before running or accepting the follow-up."],
        )
    return check(
        "reagent_safety",
        "pass",
        "Master Reagents has hazards/SDS review notes for the formulation reagents.",
        details={"safety_review_required": required},
    )


def reagent_has_safety_notes(row: dict[str, Any]) -> bool:
    reagent = row.get("reagent") if isinstance(row.get("reagent"), dict) else {}
    return any(
        nonblank(value)
        for value in (
            row.get("hazards"),
            row.get("reagent_hazards"),
            reagent.get("hazards"),
        )
    )


def safety_review_required(tables: dict[str, list[dict[str, Any]]]) -> bool:
    raw_value = agent_config_values(tables).get("safety_review_required", "true")
    return str(raw_value).strip().lower() not in {"false", "0", "no", "off"}


def placeholder_reagents_check(entry: dict[str, Any]) -> dict[str, Any]:
    placeholders = []
    for row in entry.get("formulation", []) or []:
        if not isinstance(row, dict):
            continue
        reagent = row.get("reagent") if isinstance(row.get("reagent"), dict) else {}
        text = " ".join(
            str(value)
            for value in [
                row.get("reagent_id", ""),
                reagent.get("name", ""),
                reagent.get("common_name", ""),
                reagent.get("notes", ""),
                row.get("notes", ""),
            ]
        ).lower()
        if "todo" in text or "starter placeholder" in text or str(row.get("reagent_id", "")).startswith("AUTO-"):
            placeholders.append(
                {
                    "reagent_id": row.get("reagent_id", ""),
                    "target_role": row.get("target_role", ""),
                    "name": reagent.get("name", ""),
                }
            )
    if placeholders:
        return check(
            "placeholder_reagents",
            "fail",
            "Generated placeholder reagent rows must be replaced before running.",
            details={"placeholder_reagents": placeholders},
            actions=["Replace TODO/AUTO reagent identities with verified Master Reagents rows and SDS-reviewed hazards."],
        )
    return check(
        "placeholder_reagents",
        "pass",
        "No generated placeholder reagents are linked to this formulation.",
    )


def batch_builder_check(
    tables: dict[str, list[dict[str, Any]]], experiment_id: str, stage: str
) -> dict[str, Any]:
    rows = rows_for_experiment(tables, "Batch Builder", experiment_id)
    if not rows:
        return check(
            "batch_plan_and_actuals",
            "fail",
            "No staged Batch Builder plan exists for the run.",
            actions=["Add one Batch Builder row for every charge, feed, hold, quench, workup, and purification step."],
        )
    plan_gaps = []
    actual_gaps = []
    for row in rows:
        charge_id = row.get("charge_id", "")
        missing_plan = [
            field for field in ("charge_id", "stage", "charge_type", "target_role", "reagent_id", "feed_order")
            if not nonblank(row.get(field))
        ]
        formula_status = str(row.get("formula_status", "")).strip()
        if formula_status and formula_status != "READY":
            missing_plan.append("formula_status:" + formula_status)
        if missing_plan:
            plan_gaps.append({"charge_id": charge_id, "missing_or_invalid": missing_plan})
        if stage in {"review", "archive"} and str(row.get("charge_status", "")).strip() == "charged":
            missing_actual = [
                field for field in ("effective_actual_mass_g", "lot", "recorded_by", "recorded_at")
                if not nonblank(row.get(field))
            ]
            if missing_actual:
                actual_gaps.append({"charge_id": charge_id, "missing_fields": missing_actual})
    if plan_gaps:
        return check(
            "batch_plan_and_actuals", "fail", "The staged batch plan is incomplete or not calculation-ready.",
            details={"plan_gaps": plan_gaps},
            actions=["Resolve every Batch Builder plan gap before execution."],
        )
    if actual_gaps:
        return check(
            "batch_plan_and_actuals", "fail", "Charged materials lack attributable actual values or lot provenance.",
            details={"actual_gaps": actual_gaps},
            actions=["Record actual mass, lot, recorder, and timestamp for every charged row."],
        )
    return check(
        "batch_plan_and_actuals", "pass",
        "The batch plan is structured and its required actual charge records are attributable.",
        details={"charge_count": len(rows)},
    )


def notebook_sections_check(
    tables: dict[str, list[dict[str, Any]]], experiment_id: str, stage: str
) -> dict[str, Any]:
    rows = rows_for_experiment(tables, "Notebook Sections", experiment_id)
    by_type = {str(row.get("section_type", "")).strip().lower(): row for row in rows}
    missing_types = [section for section in PRE_RUN_SECTION_TYPES if section not in by_type]
    gaps = []
    for section_type, row in by_type.items():
        required = str(row.get("required", "")).strip().lower() in {"true", "1", "yes"}
        should_be_complete = required if stage in {"review", "archive"} else section_type in PRE_RUN_SECTION_TYPES
        if not should_be_complete:
            continue
        missing = []
        if str(row.get("status", "")).strip().lower() != "complete":
            missing.append("status:complete")
        if not nonblank(row.get("content")):
            missing.append("content")
        if stage in {"review", "archive"}:
            for field in ("authored_by", "authored_at", "completed_by", "completed_at"):
                if not nonblank(row.get(field)):
                    missing.append(field)
        if missing:
            gaps.append({"section_type": section_type, "missing_or_invalid": missing})
    if missing_types or gaps:
        return check(
            "notebook_sections", "fail", "Required scientific narrative sections are missing or incomplete.",
            details={"missing_section_types": missing_types, "section_gaps": gaps},
            actions=["Complete the required objective, safety, procedure, observation, workup, results, and conclusion sections at the appropriate stage."],
        )
    return check(
        "notebook_sections", "pass", "Required narrative sections are present and complete for this stage.",
        details={"section_count": len(rows)},
    )


def raw_data_traceability_check(
    tables: dict[str, list[dict[str, Any]]], experiment_id: str, stage: str
) -> dict[str, Any]:
    measurements = rows_for_experiment(tables, "Measurements", experiment_id)
    raw_files = rows_for_experiment(tables, "Raw Data Files", experiment_id)
    raw_by_id = {str(row.get("raw_file_id", "")).strip(): row for row in raw_files}
    gaps = []
    for row in measurements:
        raw_id = str(row.get("Raw file ID", "")).strip()
        if not raw_id:
            gaps.append({"sample_id": row.get("Sample ID", ""), "measurement": row.get("Measurement", ""), "problem": "missing_raw_file_id"})
            continue
        raw = raw_by_id.get(raw_id)
        if not raw:
            gaps.append({"raw_file_id": raw_id, "problem": "unresolved_raw_file_id"})
            continue
        missing = [field for field in ("file_name", "file_url", "collected_at", "instrument_id") if not nonblank(raw.get(field))]
        if missing:
            gaps.append({"raw_file_id": raw_id, "problem": "incomplete_provenance", "missing_fields": missing})
    if gaps:
        return check(
            "raw_data_traceability",
            "fail" if stage in {"review", "archive"} else "warn",
            "One or more measurements cannot be traced to a provenance-bearing raw file.",
            details={"raw_data_gaps": gaps},
            actions=["Create a Raw Data Files row and link its ID from each Measurements row."],
        )
    return check(
        "raw_data_traceability", "pass", "Measurements are linked to registered raw data files.",
        details={"measurement_count": len(measurements), "raw_file_count": len(raw_files)},
    )


def reaction_outcome_check(
    tables: dict[str, list[dict[str, Any]]], experiment_id: str, stage: str
) -> dict[str, Any]:
    rows = rows_for_experiment(tables, "Reaction Outcomes", experiment_id)
    if not rows:
        return check(
            "reaction_outcome",
            "fail" if stage in {"review", "archive"} else "warn",
            "No product outcome or material-closure record exists.",
            actions=["Add a Reaction Outcomes row with theoretical basis, recovered product, closure, purity, and conclusion."],
        )
    row = rows[-1]
    if stage in {"review", "archive"}:
        gaps = [
            field for field in ("product_name", "theoretical_product_mass_g", "recovered_product_mass_g", "purity_percent", "appearance", "completed_by", "completed_at", "conclusion")
            if not nonblank(row.get(field))
        ]
        outcome_status = str(row.get("outcome_status", "")).strip().lower()
        balance_status = str(row.get("mass_balance_status", "")).strip().upper()
        if outcome_status != "complete":
            gaps.append("outcome_status:complete")
        if balance_status != "CLOSED":
            gaps.append("mass_balance_status:CLOSED")
        if gaps:
            return check(
                "reaction_outcome", "fail", "The reaction outcome is not scientifically or materially closed.",
                details={"missing_or_invalid": gaps},
                actions=["Complete yield, purity, appearance, material balance, attribution, and conclusion before review."],
            )
    return check("reaction_outcome", "pass", "A stage-appropriate reaction outcome record exists.")


def deviations_check(
    tables: dict[str, list[dict[str, Any]]], experiment_id: str, stage: str
) -> dict[str, Any]:
    rows = rows_for_experiment(tables, "Deviations", experiment_id)
    open_rows = [row for row in rows if str(row.get("status", "")).strip().lower() != "closed"]
    if open_rows and stage in {"review", "archive"}:
        return check(
            "deviations", "fail", "Open deviations still require impact assessment and disposition.",
            details={"open_deviation_ids": [row.get("deviation_id", "") for row in open_rows]},
            actions=["Close each deviation with impact assessment, disposition, owner, and reviewer."],
        )
    return check("deviations", "pass", "No unresolved deviation blocks this stage.", details={"deviation_count": len(rows)})


def review_attribution_check(entry: dict[str, Any], stage: str) -> dict[str, Any]:
    if stage == "planning":
        return check("review_attribution", "pass", "Review attribution is not required before execution.")
    missing = [field for field in ("summary", "completed_at", "reviewer", "reviewed_at") if not nonblank(entry.get(field))]
    if missing:
        return check(
            "review_attribution", "fail", "The completed run lacks summary or independent review attribution.",
            details={"missing_fields": missing},
            actions=["Complete the run summary and record reviewer and review timestamps."],
        )
    return check("review_attribution", "pass", "The completed record has summary and review attribution.")


def signature_integrity_check(
    tables: dict[str, list[dict[str, Any]]], entry: dict[str, Any], stage: str
) -> dict[str, Any]:
    if stage != "archive":
        return check("signature_integrity", "pass", "Signature is evaluated only at archive stage.")
    experiment_id = str(entry.get("experiment_id", "")).strip()
    fingerprint = str(entry.get("record_fingerprint", "")).strip()
    signed = [
        row for row in rows_for_experiment(tables, "Record Signatures", experiment_id)
        if str(row.get("status", "")).strip().lower() == "signed"
    ]
    matching = [
        row for row in signed
        if fingerprint and str(row.get("record_fingerprint", "")).strip() == fingerprint
    ]
    if not fingerprint or not matching:
        return check(
            "signature_integrity", "fail", "No active signature matches the current record fingerprint.",
            actions=["Generate the final record fingerprint, then sign that exact reviewed state."],
        )
    if any(not nonblank(row.get("witnessed_by")) or not nonblank(row.get("witnessed_at")) for row in matching):
        return check(
            "signature_integrity", "fail", "The matching signature lacks an independent witness.",
            actions=["Add an independent witness to the matching signed record."],
        )
    return check("signature_integrity", "pass", "A witnessed signature matches the current record fingerprint.")


def observations_check(entry: dict[str, Any], stage: str) -> dict[str, Any]:
    observations = [row for row in entry.get("observations", []) or [] if isinstance(row, dict)]
    if observations:
        return check(
            "daily_log_observations",
            "pass",
            "Daily Log has observations for this experiment.",
            details={"observation_count": len(observations)},
        )
    status = "fail" if stage in {"review", "archive"} else "warn"
    return check(
        "daily_log_observations",
        status,
        "Daily Log has no observations for this experiment.",
        details={"observation_count": 0},
        actions=["Capture timestamped observations in Bench Log during setup, feed, hold, workup, and testing."],
    )


def results_check(entry: dict[str, Any], stage: str) -> dict[str, Any]:
    results = [row for row in entry.get("results", []) or [] if isinstance(row, dict)]
    if results:
        return check(
            "results_measurements",
            "pass",
            "Results has measurements for this experiment.",
            details={"result_count": len(results)},
        )
    status = "fail" if stage in {"review", "archive"} else "warn"
    return check(
        "results_measurements",
        status,
        "Results has no measurements for this experiment.",
        details={"result_count": 0},
        actions=["Capture measurements in Measurements before asking for a result-driven next experiment."],
    )


def literature_evidence_check(entry: dict[str, Any]) -> dict[str, Any]:
    evidence = [row for row in entry.get("literature_evidence", []) or [] if isinstance(row, dict)]
    if evidence:
        return check(
            "literature_evidence",
            "pass",
            "Literature Evidence rows are linked to this experiment.",
            details={"evidence_count": len(evidence)},
        )
    return check(
        "literature_evidence",
        "warn",
        "No Literature Evidence rows are linked yet.",
        details={"evidence_count": 0},
        actions=["Run the agent with --litscout-export or --run-litscout to ground suggestions in literature evidence."],
    )


def open_suggestions_check(tables: dict[str, list[dict[str, Any]]], experiment_id: str) -> dict[str, Any]:
    suggestions = suggestions_for_experiment(tables, experiment_id)
    if not suggestions:
        return check(
            "open_suggestions",
            "pass",
            "No open Agent Suggestions currently block a new suggestion run.",
        )
    return check(
        "open_suggestions",
        "warn",
        "Open Agent Suggestions exist for this experiment.",
        details={
            "suggestion_ids": [
                row.get("suggestion_id", "") for row in suggestions if row.get("suggestion_id")
            ]
        },
        actions=["Review open suggestions and set status to rejected or run_complete before rerunning without --force."],
    )


def report_payload(
    experiment_id: str,
    stage: str,
    entry: dict[str, Any],
    material_audit: dict[str, Any],
    checks: list[dict[str, Any]],
) -> dict[str, Any]:
    fail_count = sum(1 for row in checks if row.get("status") == "fail")
    warn_count = sum(1 for row in checks if row.get("status") == "warn")
    ready_to_run = fail_count == 0
    ready_for_agent_suggestion = ready_to_run and (
        stage != "review"
        or (
            any(row.get("name") == "daily_log_observations" and row.get("status") == "pass" for row in checks)
            and any(row.get("name") == "results_measurements" and row.get("status") == "pass" for row in checks)
        )
    )
    return {
        "schema": "lab-notebook-agent-experiment-preflight.v2",
        "experiment_id": experiment_id,
        "stage": stage,
        "status": "needs_attention" if fail_count else ("ready_with_warnings" if warn_count else "ready"),
        "ready_to_run": ready_to_run,
        "ready_for_quantitative_suggestion": bool(material_audit.get("ready_for_quantitative_suggestion", False)),
        "ready_for_agent_suggestion": ready_for_agent_suggestion,
        "ready_to_archive": stage == "archive" and fail_count == 0,
        "summary": {
            "check_count": len(checks),
            "pass_count": sum(1 for row in checks if row.get("status") == "pass"),
            "warn_count": warn_count,
            "fail_count": fail_count,
            "action_count": sum(len(row.get("actions", [])) for row in checks),
        },
        "experiment": {
            "date": entry.get("date", ""),
            "project": entry.get("project", ""),
            "process_type": entry.get("process_type", ""),
            "objective": entry.get("objective", ""),
            "status": entry.get("status", ""),
        },
        "checks": checks,
        "next_actions": collect_actions(checks),
        "material_audit": material_audit,
    }


def check(
    name: str,
    status: str,
    message: str,
    details: dict[str, Any] | None = None,
    actions: list[str] | None = None,
) -> dict[str, Any]:
    return {
        "name": name,
        "status": status,
        "message": message,
        "details": details or {},
        "actions": actions or [],
    }


def collect_actions(checks: list[dict[str, Any]]) -> list[str]:
    actions: list[str] = []
    for row in checks:
        for action in row.get("actions", []):
            if action and action not in actions:
                actions.append(str(action))
    if not actions:
        actions.append("Experiment notebook row is ready for the selected preflight stage.")
    return actions
