from __future__ import annotations

from dataclasses import dataclass


WORKBOOK_CONTRACT_VERSION = "0.7.0"


@dataclass(frozen=True)
class Column:
    name: str
    description: str
    required: bool = False


@dataclass(frozen=True)
class SheetSpec:
    name: str
    columns: tuple[Column, ...]
    example_rows: tuple[tuple[object, ...], ...] = ()

    @property
    def headers(self) -> tuple[str, ...]:
        return tuple(column.name for column in self.columns)


PROCESS_TYPES = (
    "emulsion polymerization",
    "solution polymerization",
    "bulk polymerization",
    "suspension polymerization",
    "RAFT polymerization",
    "step-growth polymerization",
    "ring-opening polymerization",
    "polyurethane network curing",
    "photopolymerization",
    "polymer functionalization",
    "monomer synthesis",
    "chain-transfer agent synthesis",
    "latex characterization",
    "compounding",
    "hydrolysis study",
)

REAGENT_CATEGORIES = (
    "monomer",
    "initiator",
    "surfactant",
    "solvent",
    "buffer",
    "chain_transfer_agent",
    "crosslinker",
    "inhibitor",
    "additive",
    "catalyst",
    "acid",
    "base",
    "gas",
    "polyol",
    "isocyanate",
    "chain_extender",
    "functional_reagent",
    "bio_based_substrate",
    "matrix_polymer",
    "unknown",
)

FORMULATION_ROLES = (
    "core_monomer",
    "shell_monomer",
    "comonomer",
    "initiator",
    "surfactant",
    "buffer",
    "chain_transfer_agent",
    "crosslinker",
    "solvent",
    "neutralizer",
    "additive",
    "monomer",
    "substrate",
    "catalyst",
    "acid",
    "base",
    "gas",
    "polyol",
    "isocyanate",
    "chain_extender",
    "functional_reagent",
)

BATCH_STAGES = (
    "setup",
    "initial_charge",
    "addition",
    "reaction",
    "hold",
    "quench",
    "seed",
    "core",
    "shell",
    "functional_shell",
    "workup",
    "purification",
    "isolation",
)

BATCH_CHARGE_TYPES = (
    "initial_charge",
    "monomer_charge",
    "solvent_charge",
    "catalyst_addition",
    "dropwise_addition",
    "gas_feed",
    "reflux",
    "distillation",
    "pre_reactor",
    "monomer_pre_emulsion",
    "aqueous_pre_emulsion",
    "initiator_shot",
    "initiator_feed",
    "redox_pair",
    "chase",
    "adjustment",
    "quench",
    "wash",
    "filtration",
    "isolation",
    "other",
)

BATCH_CHARGE_STATUS = (
    "planned",
    "prepared",
    "charged",
    "skipped",
)

EXPERIMENT_STATUS = (
    "planned",
    "running",
    "complete",
    "needs_review",
    "abandoned",
)

PROCESS_STAGES = (
    "setup",
    "initial_charge",
    "addition",
    "reaction",
    "seed",
    "feed",
    "hold",
    "chase",
    "workup",
    "sampling",
    "test",
    "cleanup",
    "quench",
    "purification",
    "isolation",
)

RESULT_QUALITY_FLAGS = (
    "planned",
    "observed",
    "ok",
    "suspect",
    "repeat",
    "failed",
)

SUGGESTION_STATUS = (
    "draft",
    "accepted",
    "rejected",
    "run_planned",
    "run_complete",
)

DAILY_REVIEW_STATUS = (
    "needs_attention",
    "ready_to_apply",
    "ready_with_warnings",
    "no_action",
)

PROJECT_RECORD_TYPES = (
    "component",
    "process_parameter",
    "feed_step",
    "observation",
    "result",
)

SOURCE_SYNC_STATUS = (
    "synced",
    "unchanged",
    "needs_review",
)

PLOT_ROW_TYPES = (
    "header",
    "point",
)

PLOT_TYPES = (
    "LINE",
    "COLUMN",
)

PLOT_STATUS = (
    "ready",
    "insufficient_data",
)

PLOT_X_AXIS_MODES = (
    "numeric",
    "category",
)

RUN_STEP_STATUS = (
    "planned",
    "ready",
    "in_progress",
    "complete",
    "skipped",
    "blocked",
)

SAMPLE_STATUS = (
    "collected",
    "stored",
    "consumed",
    "disposed",
    "missing",
)

CALIBRATION_STATUS = (
    "current",
    "due_soon",
    "expired",
    "not_required",
    "unknown",
)

PROTOCOL_STATUS = (
    "draft",
    "active",
    "retired",
)

SPECIFICATION_STATUS = (
    "draft",
    "active",
    "retired",
)

QC_STATUS = (
    "pass",
    "warn",
    "fail",
    "not_evaluated",
)

DEVIATION_STATUS = (
    "open",
    "under_review",
    "closed",
)

RAW_FILE_PARSER_STATUS = (
    "unprocessed",
    "parsed",
    "needs_review",
    "failed",
)

AUDIT_ACTIONS = (
    "create",
    "update",
    "correct",
    "delete",
    "import",
    "migrate",
)

INVENTORY_STATUS = (
    "available",
    "low",
    "expired",
    "quarantined",
    "unknown",
)

CONTROLLED_VOCAB_VALIDATIONS: dict[str, dict[str, tuple[str, ...]]] = {
    "Master Reagents": {
        "category": REAGENT_CATEGORIES,
        "inventory_status": INVENTORY_STATUS,
    },
    "Experiments": {"process_type": PROCESS_TYPES, "status": EXPERIMENT_STATUS},
    "Batch Builder": {
        "stage": BATCH_STAGES,
        "charge_type": BATCH_CHARGE_TYPES,
        "target_role": FORMULATION_ROLES,
        "charge_status": BATCH_CHARGE_STATUS,
    },
    "Bench Log": {"Stage": PROCESS_STAGES},
    "Measurements": {"Quality": RESULT_QUALITY_FLAGS},
    "Daily Log": {"process_stage": PROCESS_STAGES},
    "Formulations": {"target_role": FORMULATION_ROLES},
    "Results": {
        "quality_flag": RESULT_QUALITY_FLAGS,
        "qc_status": QC_STATUS,
    },
    "Agent Suggestions": {"status": SUGGESTION_STATUS},
    "Daily Reviews": {"status": DAILY_REVIEW_STATUS},
    "Project Notebook Records": {
        "record_type": PROJECT_RECORD_TYPES,
        "active": ("true", "false"),
    },
    "Source Sync": {"status": SOURCE_SYNC_STATUS},
    "Plot Data": {
        "row_type": PLOT_ROW_TYPES,
        "active": ("true", "false"),
    },
    "Plot Definitions": {
        "chart_type": PLOT_TYPES,
        "status": PLOT_STATUS,
        "x_axis_mode": PLOT_X_AXIS_MODES,
    },
    "Run Capture Plan": {
        "required": ("true", "false"),
        "status": RUN_STEP_STATUS,
    },
    "Samples": {"status": SAMPLE_STATUS},
    "Equipment": {"calibration_status": CALIBRATION_STATUS},
    "Protocols": {"status": PROTOCOL_STATUS},
    "Specifications": {"status": SPECIFICATION_STATUS},
    "Deviations": {"status": DEVIATION_STATUS},
    "Raw Data Files": {"parser_status": RAW_FILE_PARSER_STATUS},
    "Audit Log": {"action": AUDIT_ACTIONS},
}

SHEETS: tuple[SheetSpec, ...] = (
    SheetSpec(
        name="Reaction Master",
        columns=(
            Column("Run ID", "Auto-linked experiment identifier."),
            Column("Date", "Experiment date from Experiments."),
            Column("Project", "Project from Experiments."),
            Column("Process", "Polymerization or reaction process."),
            Column("Objective", "Scientific objective for the run."),
            Column("Status", "Current experiment status."),
            Column("Operator", "Scientist running the reaction."),
            Column("Protocol", "Protocol and version used."),
            Column("Equipment", "Primary reactor or equipment."),
            Column("Planned mass (g)", "Sum of planned Batch Builder charges."),
            Column("Actual mass (g)", "Sum of actual Batch Builder charges."),
            Column("Mass variance (g)", "Actual minus planned mass across charges."),
            Column("Charges recorded", "Number of charges marked charged."),
            Column("Bench entries", "Number of linked Bench Log entries."),
            Column("Measurements", "Number of linked Measurements rows."),
            Column("Last activity", "Most recent bench or measurement timestamp."),
            Column("Run summary", "Completion summary from Experiments."),
            Column("Next action", "Formula-derived next notebook action."),
            Column("Reviewer", "Scientist who reviewed the record."),
            Column("Reviewed at", "Record review timestamp."),
        ),
    ),
    SheetSpec(
        name="Master Reagents",
        columns=(
            Column("reagent_id", "Stable short ID used by formulation rows.", True),
            Column("name", "Full reagent name.", True),
            Column("common_name", "Short lab name or abbreviation."),
            Column("category", "Controlled reagent category.", True),
            Column("role", "Typical role in experiments."),
            Column("molecular_weight_g_mol", "Molecular weight in g/mol."),
            Column("density_g_mL", "Density in g/mL when relevant."),
            Column("purity_fraction", "Purity from 0 to 1."),
            Column("concentration", "Stock concentration if this is a solution."),
            Column("concentration_units", "Units for concentration."),
            Column("supplier", "Supplier or source."),
            Column("lot", "Lot or batch identifier."),
            Column("storage_location", "Freezer, cabinet, hood, or shelf."),
            Column("hazards", "Short safety notes."),
            Column("notes", "Additional user notes."),
            Column("cas_number", "CAS Registry Number when verified."),
            Column("expiration_date", "Expiration or retest date."),
            Column("inventory_status", "available, low, expired, quarantined, or unknown."),
            Column("sds_url", "Link to the current safety data sheet."),
        ),
        example_rows=(
            (
                "M-SKA",
                "solketal acrylate",
                "SKA",
                "monomer",
                "rubbery acrylic core monomer",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "combustible/irritant; verify SDS",
                "Seed row; replace values with verified inventory data.",
            ),
            (
                "I-APS",
                "ammonium persulfate",
                "APS",
                "initiator",
                "water-soluble radical initiator",
                "228.20",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "oxidizer; verify SDS",
                "Common emulsion polymerization initiator.",
            ),
            (
                "S-SDS",
                "sodium dodecyl sulfate",
                "SDS",
                "surfactant",
                "anionic surfactant",
                "288.38",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "irritant; verify SDS",
                "Example surfactant row.",
            ),
        ),
    ),
    SheetSpec(
        name="Experiments",
        columns=(
            Column("experiment_id", "Stable experiment ID.", True),
            Column("date", "Experiment date as YYYY-MM-DD.", True),
            Column("project", "Project or product line."),
            Column("process_type", "Controlled process type.", True),
            Column("objective", "What the experiment is trying to learn.", True),
            Column("hypothesis", "Expected outcome or mechanism."),
            Column("linked_literature_ids", "Evidence IDs from Literature Evidence."),
            Column("operator", "Person running the experiment."),
            Column("status", "Experiment state.", True),
            Column("planned_next_step", "Human-planned next action."),
            Column("summary", "Short run summary after completion."),
            Column("source_notebook_id", "Source spreadsheet or workbook ID for imported experiments."),
            Column("source_sheet_name", "Source tab name for imported experiments."),
            Column("source_url", "Link to the read-only source notebook."),
            Column("source_modified_at", "Source modification timestamp when available."),
            Column("source_fingerprint", "SHA-256 fingerprint of the imported source tab."),
            Column("protocol_id", "Protocol and version used for this run."),
            Column("equipment_id", "Primary reactor or equipment ID."),
            Column("started_at", "Actual run start timestamp."),
            Column("completed_at", "Actual run completion timestamp."),
            Column("reviewer", "Person who reviewed the completed record."),
            Column("reviewed_at", "Record review timestamp."),
        ),
        example_rows=(
            (
                "EP-001",
                "2026-06-09",
                "SABER CCSP",
                "emulsion polymerization",
                "Reduce particle size while avoiding coagulum.",
                "A slower monomer feed and balanced surfactant package should narrow PSD.",
                "",
                "",
                "planned",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
            ),
        ),
    ),
    SheetSpec(
        name="Batch Builder",
        columns=(
            Column("experiment_id", "Experiment ID linked to Experiments.", True),
            Column("charge_id", "Stable charge-row ID such as EP-001-CHG-001.", True),
            Column("stage", "Seed, core, shell, functional shell, or workup stage.", True),
            Column("charge_type", "Where or how this material is charged.", True),
            Column("target_role", "Scientific role of the material in this stage.", True),
            Column("feed_order", "Planned order of addition within the experiment."),
            Column("reagent_id", "Reagent ID from Master Reagents.", True),
            Column("material_name", "Calculated material name from Master Reagents."),
            Column("parts_per_hundred_monomer", "Input parts per hundred monomer when scaling by PHR."),
            Column("stage_monomer_basis_g", "Input monomer basis for this stage in grams."),
            Column("direct_mass_g", "Input planned mass directly when PHR scaling is not used."),
            Column("planned_mass_g", "Calculated direct mass or PHR times stage basis divided by 100."),
            Column("stock_active_fraction", "Input active fraction of the supplied stock from 0 to 1."),
            Column("active_mass_g", "Calculated active-material mass."),
            Column("density_override_g_mL", "Optional density override for this charge."),
            Column("density_g_mL", "Calculated override or Master Reagents density."),
            Column("planned_volume_mL", "Calculated planned mass divided by density."),
            Column("actual_mass_g", "Input actual mass charged."),
            Column("mass_variance_g", "Calculated actual mass minus planned mass."),
            Column("actual_volume_mL", "Calculated actual mass divided by density."),
            Column("feed_start_min", "Input elapsed feed start time in minutes."),
            Column("feed_duration_min", "Input feed duration in minutes."),
            Column("feed_rate_mL_min", "Calculated planned volume divided by feed duration."),
            Column("target_temperature_C", "Input target reactor temperature for the charge."),
            Column("lot", "Input reagent lot used."),
            Column("recorded_by", "Scientist who prepared or charged the material."),
            Column("recorded_at", "Timestamp for the charge record."),
            Column("charge_status", "Planned, prepared, charged, or skipped.", True),
            Column("notes", "Charge preparation, addition, or deviation notes."),
            Column("molecular_weight_g_mol", "Calculated molecular weight from Master Reagents."),
            Column("planned_moles_mmol", "Calculated planned active amount in mmol."),
            Column("equivalent_basis_mmol", "Input reference amount used for equivalents."),
            Column("equivalents", "Calculated planned mmol divided by the equivalent basis."),
            Column("actual_moles_mmol", "Calculated actual active amount in mmol."),
        ),
        example_rows=(
            (
                "EP-001", "EP-001-CHG-001", "core", "monomer_pre_emulsion",
                "core_monomer", 1, "M-SKA", "", "", "", "", "", 1,
                "", "", "", "", "", "", "", 0, 180, "", 70, "", "",
                "", "planned", "Enter direct mass, or enter PHR plus stage monomer basis.",
            ),
            (
                "EP-001", "EP-001-CHG-002", "core", "initiator_feed",
                "initiator", 2, "I-APS", "", "", "", "", "", 1,
                "", "", "", "", "", "", "", 0, 210, "", 70, "", "",
                "", "planned", "Record the stock fraction when using a solution.",
            ),
            (
                "EP-001", "EP-001-CHG-003", "core", "aqueous_pre_emulsion",
                "surfactant", 3, "S-SDS", "", "", "", "", "", 1,
                "", "", "", "", "", "", "", "", "", "", 70, "", "",
                "", "planned", "Enter the planned quantity and actual charge at the bench.",
            ),
        ),
    ),
    SheetSpec(
        name="Bench Log",
        columns=(
            Column("Run ID", "Choose the experiment or run being recorded.", True),
            Column("Date & time", "When the observation was made.", True),
            Column("Stage", "Current process stage."),
            Column("Temperature (°C)", "Observed reactor or sample temperature."),
            Column("RPM", "Observed agitation speed."),
            Column("pH", "Observed pH."),
            Column("Solids (%)", "Observed or estimated solids."),
            Column("Particle size (nm)", "Particle size when measured."),
            Column("Conversion (%)", "Conversion when measured."),
            Column("Viscosity (cP)", "Viscosity when measured."),
            Column("Observation / action", "What happened, what was seen, or what was done.", True),
            Column("Issue tags", "Short comma-separated issue tags."),
            Column("Attachment link", "Link to a photo, instrument file, or folder."),
        ),
        example_rows=((
            "EP-001", "2026-06-09T14:35:00", "feed", 70, 250, "", "", 420,
            "", "", "Latex looked bluish but small coagulum appeared on stir shaft.",
            "coagulum,particle_size_high", "",
        ),),
    ),
    SheetSpec(
        name="Daily Log",
        columns=(
            Column("experiment_id", "Experiment ID linked to Experiments.", True),
            Column("timestamp", "Observation timestamp.", True),
            Column("process_stage", "Setup, seed, feed, hold, workup, test, etc."),
            Column("temperature_C", "Observed reactor/sample temperature."),
            Column("rpm", "Agitation speed."),
            Column("pH", "Measured pH."),
            Column("solids_percent", "Measured or estimated solids percent."),
            Column("particle_size_nm", "Particle size if measured."),
            Column("conversion_percent", "Conversion if measured."),
            Column("viscosity_cP", "Viscosity if measured."),
            Column("observation", "Free-text notebook observation.", True),
            Column("issue_tags", "Comma-separated tags such as coagulum or low_conversion."),
            Column("attachments_url", "Link to photos, spectra, files, or Drive folders."),
            Column("residual_monomer_percent", "Residual monomer if measured."),
            Column("polydispersity_index", "DLS polydispersity index if measured."),
            Column("Tg_C", "Glass transition temperature if measured."),
            Column("hold_time_min", "Thermal or reaction hold time."),
        ),
        example_rows=(
            (
                "EP-001",
                "2026-06-09T14:35:00",
                "feed",
                "70",
                "250",
                "",
                "",
                "420",
                "",
                "",
                "Latex looked bluish but small coagulum appeared on stir shaft.",
                "coagulum,particle_size_high",
                "",
                "",
                "",
                "",
                "",
            ),
        ),
    ),
    SheetSpec(
        name="Formulations",
        columns=(
            Column("experiment_id", "Experiment ID linked to Experiments.", True),
            Column("reagent_id", "Reagent ID from Master Reagents.", True),
            Column("phase", "Aqueous, monomer, seed, initiator, chase, etc."),
            Column("target_role", "Controlled formulation role.", True),
            Column("mass_g", "Mass in grams."),
            Column("volume_mL", "Volume in mL."),
            Column("moles_mmol", "Moles in mmol."),
            Column("concentration", "Concentration value."),
            Column("concentration_units", "Concentration units."),
            Column("wt_percent", "Weight percent in formulation."),
            Column("feed_order", "Order added."),
            Column("feed_start_min", "Feed start time in minutes."),
            Column("feed_duration_min", "Feed duration in minutes."),
            Column("notes", "Formulation notes."),
            Column("actual_mass_g", "Actual charged mass in grams."),
            Column("mass_variance_g", "Actual mass minus planned mass in grams."),
            Column("actual_volume_mL", "Actual charged volume in mL."),
            Column("lot", "Reagent lot used for this experiment."),
            Column("recorded_by", "Person who recorded the charge."),
            Column("recorded_at", "Charge recording timestamp."),
        ),
        example_rows=(
            ("EP-001", "M-SKA", "monomer feed", "core_monomer", "", "", "", "", "", "", "1", "0", "180", ""),
            ("EP-001", "I-APS", "initiator feed", "initiator", "", "", "", "", "", "", "2", "0", "210", ""),
            ("EP-001", "S-SDS", "aqueous", "surfactant", "", "", "", "", "", "", "0", "", "", ""),
        ),
    ),
    SheetSpec(
        name="Measurements",
        columns=(
            Column("Run ID", "Choose the experiment or run.", True),
            Column("Sample ID", "Sample or aliquot identifier.", True),
            Column("Measurement", "What was measured.", True),
            Column("Method", "Instrument, method, or calculation."),
            Column("Numeric value", "Machine-readable numeric result."),
            Column("Units", "Units for the numeric value."),
            Column("Condition", "Relevant test or sample condition."),
            Column("Replicate", "Replicate number or label when applicable."),
            Column("Uncertainty", "Measurement uncertainty when available."),
            Column("Quality", "Observed, okay, suspect, repeat, or failed."),
            Column("Raw file ID", "Linked raw-data record identifier."),
            Column("Measured at", "Measurement timestamp."),
            Column("Analyst", "Person who performed the measurement."),
            Column("Interpretation / notes", "Short result interpretation or notes."),
        ),
        example_rows=((
            "EP-001", "EP-001-L1", "DLS particle size", "intensity average",
            420, "nm", "post-feed", "1", "", "suspect", "", "", "",
            "Above target range.",
        ),),
    ),
    SheetSpec(
        name="Results",
        columns=(
            Column("experiment_id", "Experiment ID linked to Experiments.", True),
            Column("sample_id", "Sample or aliquot ID.", True),
            Column("measurement_type", "DLS, GC, NMR, DSC, tensile, impact, etc.", True),
            Column("method", "Instrument or method details."),
            Column("value", "Numeric or text result."),
            Column("units", "Measurement units."),
            Column("condition", "Temperature, matrix, aging, replicate conditions."),
            Column("replicate", "Replicate number."),
            Column("quality_flag", "ok, suspect, repeat, failed."),
            Column("interpretation", "Human interpretation."),
            Column("numeric_value", "Machine-readable numeric result when value is numeric."),
            Column("uncertainty", "Measurement uncertainty."),
            Column("uncertainty_units", "Units for uncertainty."),
            Column("detection_limit", "Method detection or reporting limit."),
            Column("detection_limit_units", "Units for detection_limit."),
            Column("specification_id", "Specification used to assess this result."),
            Column("qc_status", "pass, warn, fail, or not_evaluated."),
            Column("raw_file_id", "Linked Raw Data Files record."),
            Column("measured_at", "Measurement timestamp."),
            Column("analyst", "Person who performed the measurement."),
            Column("reviewer", "Person who reviewed the result."),
            Column("reviewed_at", "Result review timestamp."),
        ),
        example_rows=(
            ("EP-001", "EP-001-L1", "DLS particle size", "intensity average", "420", "nm", "post-feed", "1", "suspect", "Above target range."),
        ),
    ),
    SheetSpec(
        name="Literature Evidence",
        columns=(
            Column("evidence_id", "Stable evidence row ID.", True),
            Column("source", "LitScout backend, paper, patent, local note, or manual."),
            Column("title", "Source title."),
            Column("authors", "Authors or assignees."),
            Column("year", "Publication year."),
            Column("doi_or_url", "DOI, patent URL, or source URL."),
            Column("query", "Search query that found this evidence."),
            Column("finding", "Short finding relevant to experiments.", True),
            Column("relevance_tags", "Tags such as surfactant, particle_size, initiator."),
            Column("confidence", "low, medium, high."),
            Column("notes", "Additional review notes."),
        ),
    ),
    SheetSpec(
        name="Agent Suggestions",
        columns=(
            Column("suggestion_id", "Stable suggestion ID.", True),
            Column("created_at", "UTC timestamp."),
            Column("experiment_id", "Experiment ID the suggestion responds to.", True),
            Column("recommendation_type", "next_experiment, literature_search, data_cleanup, safety_review."),
            Column("rationale", "Why the agent suggested this."),
            Column("proposed_change", "Concrete change or experiment to run."),
            Column("expected_effect", "Expected measurable effect."),
            Column("linked_evidence_ids", "Literature Evidence IDs used."),
            Column("safety_check", "Safety or review reminder."),
            Column("confidence", "low, medium, high."),
            Column("status", "Controlled suggestion status."),
            Column("proposed_experiment_id", "Suggested follow-up experiment ID if accepted."),
            Column("proposed_plan_json", "Structured proposed experiment plan for review/materialization."),
        ),
    ),
    SheetSpec(
        name="Daily Reviews",
        columns=(
            Column("review_id", "Stable daily review ID.", True),
            Column("created_at", "UTC timestamp for the review row."),
            Column("review_date", "Date reviewed as YYYY-MM-DD.", True),
            Column("selected_experiment_ids", "Comma-separated selected experiment IDs."),
            Column("experiment_count", "Number of experiments reviewed."),
            Column("observation_count", "Daily Log observations counted."),
            Column("result_count", "Existing Results rows counted before appending new rows."),
            Column("normalized_result_rows_to_append", "Pending Results rows from Daily Log normalization."),
            Column("evidence_rows_to_append", "Pending Literature Evidence rows."),
            Column("suggestion_rows_to_append", "Pending Agent Suggestions rows."),
            Column("preflight_fail_count", "Total failed preflight checks."),
            Column("preflight_warn_count", "Total warning preflight checks."),
            Column("apply_request_count", "Google Sheets batchUpdate request count when available."),
            Column("status", "Daily review state."),
            Column("summary", "Short human-readable run summary."),
            Column("next_actions_json", "JSON list of recommended next actions."),
        ),
    ),
    SheetSpec(
        name="Project Notebook Records",
        columns=(
            Column("record_id", "Stable ID used to update an imported source row without duplication.", True),
            Column("experiment_id", "Experiment ID linked to Experiments.", True),
            Column("source_key", "Stable source spreadsheet and tab key.", True),
            Column("record_type", "component, process_parameter, feed_step, observation, or result.", True),
            Column("stage", "Seed, core, shell, functional shell, workup, or other source stage."),
            Column("section", "Source subsection such as pre-reactor, pre-emulsion, or initiator feed."),
            Column("label", "Source label or material name."),
            Column("material_name", "Material name for component records."),
            Column("reagent_id", "Optional mapped Master Reagents ID."),
            Column("target_role", "Normalized material role such as surfactant or initiator."),
            Column("planned_value", "Calculated, goal, or target value from the source."),
            Column("actual_value", "Measured or actual value from the source."),
            Column("units", "Units for planned_value and actual_value."),
            Column("parts_per_hundred_monomer", "Source pphm value when provided."),
            Column("mass_g", "Target or calculated component mass in grams."),
            Column("actual_mass_g", "Actual component mass in grams."),
            Column("mass_variance_g", "Actual-minus-target or source mass-left value in grams."),
            Column("volume_mL", "Target or calculated volume in mL."),
            Column("density_g_mL", "Material density in g/mL."),
            Column("concentration", "Concentration value."),
            Column("concentration_units", "Concentration units."),
            Column("elapsed_start_min", "Feed or step start time in minutes."),
            Column("elapsed_end_min", "Feed or step end time in minutes."),
            Column("feed_rate_mL_min", "Feed rate in mL/min."),
            Column("cumulative_percent", "Cumulative feed percentage."),
            Column("radical_flux_mol_min_L", "Calculated radical flux in mol/min/L."),
            Column("particle_size_nm", "Measured or estimated particle size in nm."),
            Column("temperature_C", "Temperature in degrees Celsius."),
            Column("rpm", "Agitation speed."),
            Column("oil_temperature_C", "Oil-bath or controller temperature in degrees Celsius."),
            Column("torque", "Agitator torque or power value as recorded by the source."),
            Column("source_timestamp", "Clock time or timestamp copied from the source log."),
            Column("lot_number", "Input material lot number when the source provides one."),
            Column("product_lot", "Output product lot or batch identifier."),
            Column("notes", "Observation, instruction, or parser note."),
            Column("details_json", "Lossless JSON for additional source fields."),
            Column("source_range", "A1 range or source row reference."),
            Column("source_fingerprint", "Fingerprint of the complete source tab at sync time."),
            Column("synced_at", "UTC timestamp of the last source synchronization."),
            Column("active", "true for current source records; false when retired after a source change."),
            Column("cumulative_addition_g", "Cumulative initiator or other addition mass in grams."),
            Column("torque_units", "Units parsed from the source torque or power column header."),
        ),
    ),
    SheetSpec(
        name="Source Sync",
        columns=(
            Column("source_key", "Stable source spreadsheet and tab key.", True),
            Column("source_spreadsheet_id", "Google spreadsheet ID or local workbook identifier."),
            Column("source_title", "Source workbook title."),
            Column("source_sheet_name", "Source tab name.", True),
            Column("source_sheet_id", "Google sheet ID when available."),
            Column("source_url", "Link to the read-only source notebook."),
            Column("source_modified_at", "Source modification timestamp when available."),
            Column("source_range", "Bounded source range read by the adapter."),
            Column("parser_profile", "Parser profile selected for the source tab."),
            Column("parser_version", "Version of the parser contract."),
            Column("source_fingerprint", "SHA-256 fingerprint used for no-op detection.", True),
            Column("synced_at", "UTC timestamp of the last successful synchronization."),
            Column("record_count", "Number of active normalized records produced."),
            Column("status", "synced, unchanged, or needs_review."),
            Column("warnings_json", "JSON list of parser warnings."),
        ),
    ),
    SheetSpec(
        name="Plot Studio",
        columns=(
            Column("control", "Plot Studio control or helper label."),
            Column("value", "Selected value or chart data."),
            Column("notes", "Instructions or helper context."),
        ),
    ),
    SheetSpec(
        name="Plot Data",
        columns=(
            Column("row_type", "Managed block header or plot point.", True),
            Column("plot_id", "Stable plot definition ID.", True),
            Column("plot_point_id", "Stable point ID for provenance and QA."),
            Column("experiment_id", "Experiment represented by this point."),
            Column("plot_kind", "Temperature, feed rate, charge accuracy, outcome trend, etc."),
            Column("x_order", "Stable numeric ordering value."),
            Column("x_value", "Domain value used by the chart."),
            Column("x_label", "Human-readable domain label."),
            Column("x_units", "Domain units."),
            Column("series_1_name", "First series label."),
            Column("series_1_value", "First numeric series value."),
            Column("series_1_units", "First series units."),
            Column("series_2_name", "Second series label."),
            Column("series_2_value", "Second numeric series value."),
            Column("series_2_units", "Second series units."),
            Column("series_3_name", "Third series label."),
            Column("series_3_value", "Third numeric series value."),
            Column("series_3_units", "Third series units."),
            Column("series_4_name", "Fourth series label."),
            Column("series_4_value", "Fourth numeric series value."),
            Column("series_4_units", "Fourth series units."),
            Column("stage", "Process stage."),
            Column("section", "Process subsection."),
            Column("source_record_ids", "Comma-separated canonical source record IDs."),
            Column("source_ranges", "Comma-separated source ranges."),
            Column("quality_flag", "Observed, planned, estimated, suspect, etc."),
            Column("recorded_at", "Source or refresh timestamp."),
            Column("active", "true for current derived plot points."),
        ),
    ),
    SheetSpec(
        name="Plot Definitions",
        columns=(
            Column("plot_id", "Stable managed plot ID.", True),
            Column("experiment_id", "Experiment scope, or cross_experiment."),
            Column("plot_kind", "Semantic plot family.", True),
            Column("title", "Human-readable chart title.", True),
            Column("chart_type", "LINE or COLUMN.", True),
            Column("x_axis_title", "Domain-axis title."),
            Column("y_axis_title", "Value-axis title."),
            Column("series_1_name", "First series label."),
            Column("series_1_units", "First series units."),
            Column("series_2_name", "Second series label."),
            Column("series_2_units", "Second series units."),
            Column("series_3_name", "Third series label."),
            Column("series_3_units", "Third series units."),
            Column("series_4_name", "Fourth series label."),
            Column("series_4_units", "Fourth series units."),
            Column("data_start_row", "One-based Plot Data block-header row."),
            Column("data_end_row", "One-based final Plot Data point row."),
            Column("source_point_count", "Number of plotted points."),
            Column("status", "ready or insufficient_data."),
            Column("updated_at", "UTC refresh timestamp."),
            Column("notes", "Plot-generation or interpretation notes."),
            Column("x_axis_mode", "numeric for elapsed/continuous axes or category for named samples/materials."),
        ),
    ),
    SheetSpec(
        name="Plot Dashboard",
        columns=(
            Column("plot_id", "Stable managed plot ID.", True),
            Column("title", "Chart title."),
            Column("experiment_id", "Experiment scope."),
            Column("plot_kind", "Semantic plot family."),
            Column("chart_type", "LINE or COLUMN."),
            Column("point_count", "Number of plotted points."),
            Column("updated_at", "UTC refresh timestamp."),
        ),
    ),
    SheetSpec(
        name="Process Knowledge",
        columns=(
            Column("process_type", "Process name.", True),
            Column("material_role", "Role category."),
            Column("typical_examples", "Typical materials for semantic lookup."),
            Column("measured_fields", "Useful measurements to capture."),
            Column("guidance", "Short process guidance."),
            Column("search_terms", "Terms used to generate literature queries."),
        ),
        example_rows=(
            (
                "emulsion polymerization",
                "monomer",
                "acrylate monomers, methacrylate monomers, functional comonomers, crosslinkers",
                "mass_g, moles_mmol, conversion_percent, Tg",
                "Track monomer feed timing, core/shell target, and conversion.",
                "emulsion polymerization acrylic monomer core shell latex conversion",
            ),
            (
                "emulsion polymerization",
                "initiator",
                "ammonium persulfate, potassium persulfate, redox initiator systems",
                "initiator concentration, temperature_C, conversion_percent",
                "Initiator level and temperature affect radical flux, conversion, and nucleation.",
                "emulsion polymerization persulfate initiator radical flux particle nucleation",
            ),
            (
                "emulsion polymerization",
                "surfactant",
                "SDS, nonionic fatty alcohol ethoxylates, mixed ionic/nonionic surfactants",
                "particle_size_nm, solids_percent, coagulum, viscosity_cP",
                "Surfactant package and feed profile often control latex stability and particle size.",
                "emulsion polymerization surfactant particle size coagulum latex stability",
            ),
        ),
    ),
    SheetSpec(
        name="Controlled Vocab",
        columns=(
            Column("field", "Field name."),
            Column("allowed_value", "Allowed value."),
            Column("description", "Meaning."),
        ),
        example_rows=tuple(
            ("process_type", value, "Experiment process type.") for value in PROCESS_TYPES
        )
        + tuple(("reagent_category", value, "Master Reagents category.") for value in REAGENT_CATEGORIES)
        + tuple(("formulation_role", value, "Formulations target role.") for value in FORMULATION_ROLES)
        + tuple(("batch_stage", value, "Batch Builder stage.") for value in BATCH_STAGES)
        + tuple(("batch_charge_type", value, "Batch Builder charge type.") for value in BATCH_CHARGE_TYPES)
        + tuple(("batch_charge_status", value, "Batch Builder charge status.") for value in BATCH_CHARGE_STATUS)
        + tuple(("experiment_status", value, "Experiments status.") for value in EXPERIMENT_STATUS)
        + tuple(("process_stage", value, "Daily Log process stage.") for value in PROCESS_STAGES)
        + tuple(("result_quality_flag", value, "Results quality flag.") for value in RESULT_QUALITY_FLAGS)
        + tuple(("suggestion_status", value, "Agent Suggestions status.") for value in SUGGESTION_STATUS)
        + tuple(("daily_review_status", value, "Daily Reviews status.") for value in DAILY_REVIEW_STATUS)
        + tuple(
            ("project_record_type", value, "Imported Project Notebook Records type.")
            for value in PROJECT_RECORD_TYPES
        )
        + tuple(
            ("source_sync_status", value, "Source Sync status.")
            for value in SOURCE_SYNC_STATUS
        )
        + tuple(
            ("plot_row_type", value, "Managed Plot Data row type.")
            for value in PLOT_ROW_TYPES
        )
        + tuple(("plot_type", value, "Managed plot chart type.") for value in PLOT_TYPES)
        + tuple(("plot_status", value, "Managed plot status.") for value in PLOT_STATUS)
        + tuple(
            ("plot_x_axis_mode", value, "Managed plot domain interpretation.")
            for value in PLOT_X_AXIS_MODES
        )
        + tuple(("run_step_status", value, "Run Capture Plan step status.") for value in RUN_STEP_STATUS)
        + tuple(("sample_status", value, "Sample lifecycle status.") for value in SAMPLE_STATUS)
        + tuple(
            ("calibration_status", value, "Equipment calibration state.")
            for value in CALIBRATION_STATUS
        )
        + tuple(("protocol_status", value, "Protocol lifecycle status.") for value in PROTOCOL_STATUS)
        + tuple(
            ("specification_status", value, "Specification lifecycle status.")
            for value in SPECIFICATION_STATUS
        )
        + tuple(("qc_status", value, "Result assessment against a specification.") for value in QC_STATUS)
        + tuple(("deviation_status", value, "Deviation workflow status.") for value in DEVIATION_STATUS)
        + tuple(
            ("raw_file_parser_status", value, "Raw data ingest or parser state.")
            for value in RAW_FILE_PARSER_STATUS
        )
        + tuple(("audit_action", value, "Audit trail action.") for value in AUDIT_ACTIONS)
        + tuple(("inventory_status", value, "Reagent inventory state.") for value in INVENTORY_STATUS),
    ),
    SheetSpec(
        name="Agent Config",
        columns=(
            Column("key", "Configuration key.", True),
            Column("value", "Configuration value."),
            Column("notes", "Usage notes."),
        ),
        example_rows=(
            ("default_context_limit", "5", "Notebook search matches to include per agent run."),
            ("default_history_limit", "5", "Same-process prior experiments to include as result benchmarks."),
            ("default_evidence_limit", "3", "Literature Evidence rows to append or select per experiment."),
            ("default_litscout_sources", "openalex,crossref,semantic_scholar", "Sources used by generated LitScout commands."),
            ("default_litscout_depth", "light", "Increase to medium/intense when a search is promising."),
            ("default_litscout_limit", "8", "LitScout works to retrieve when the agent runs LitScout live."),
            ("suggestion_confidence_floor", "medium", "Human review required before running suggested experiments."),
            ("require_literature_evidence", "false", "Set true to skip suggestions unless Literature Evidence is linked or generated."),
            ("safety_review_required", "true", "Agent suggestions do not replace SDS, SOP, or PI review."),
        ),
    ),
    SheetSpec(
        name="Workbook Metadata",
        columns=(
            Column("key", "Stable workbook metadata key.", True),
            Column("value", "Metadata value."),
            Column("updated_at", "Timestamp of the last metadata update."),
            Column("notes", "Meaning or migration context."),
        ),
        example_rows=(
            ("contract_name", "lab-notebook-agent-workbook", "", "Machine-readable workbook contract."),
            ("contract_version", WORKBOOK_CONTRACT_VERSION, "", "Schema version currently applied to this workbook."),
            ("workbook_timezone", "America/Chicago", "", "Timezone used for local laboratory timestamps."),
            ("migration_status", "current", "", "Set by a successful contract migration."),
        ),
    ),
    SheetSpec(
        name="Run Capture Plan",
        columns=(
            Column("experiment_id", "Experiment ID linked to Experiments.", True),
            Column("step_id", "Stable step ID within the experiment.", True),
            Column("sequence", "Execution order.", True),
            Column("stage", "Setup, seed, feed, hold, workup, test, or cleanup."),
            Column("planned_time_min", "Planned elapsed start time in minutes."),
            Column("action", "Operator action or checkpoint.", True),
            Column("parameter", "Controlled parameter or observation name."),
            Column("target_value", "Planned parameter value."),
            Column("units", "Units for target and limits."),
            Column("lower_limit", "Lower acceptable operating limit."),
            Column("upper_limit", "Upper acceptable operating limit."),
            Column("required", "true when the step cannot be skipped."),
            Column("status", "planned, ready, in_progress, complete, skipped, or blocked."),
            Column("actual_time_min", "Actual elapsed completion time in minutes."),
            Column("completed_by", "Operator who completed the step."),
            Column("completed_at", "Step completion timestamp."),
            Column("deviation_id", "Linked deviation if the step differed from plan."),
            Column("notes", "Execution notes."),
        ),
    ),
    SheetSpec(
        name="Samples",
        columns=(
            Column("sample_id", "Stable sample or aliquot ID.", True),
            Column("experiment_id", "Experiment that produced the sample.", True),
            Column("parent_sample_id", "Parent sample for an aliquot or derivative."),
            Column("collected_at", "Sample collection timestamp."),
            Column("stage", "Process stage at collection."),
            Column("sample_type", "Latex, serum, residual, film, coupon, or other type."),
            Column("amount", "Collected amount."),
            Column("units", "Units for amount."),
            Column("storage_location", "Physical storage location."),
            Column("status", "collected, stored, consumed, disposed, or missing."),
            Column("raw_data_url", "Link to related raw data or folder."),
            Column("notes", "Sample handling and condition notes."),
        ),
    ),
    SheetSpec(
        name="Equipment",
        columns=(
            Column("equipment_id", "Stable equipment or instrument ID.", True),
            Column("name", "Human-readable equipment name.", True),
            Column("manufacturer", "Equipment manufacturer."),
            Column("model", "Manufacturer model."),
            Column("serial_number", "Serial number."),
            Column("location", "Lab and physical location."),
            Column("calibration_status", "current, due_soon, expired, not_required, or unknown."),
            Column("last_calibration_date", "Date last calibrated or verified."),
            Column("calibration_due_date", "Next calibration or verification due date."),
            Column("protocol_id", "Operating or calibration protocol."),
            Column("notes", "Equipment condition or usage notes."),
        ),
    ),
    SheetSpec(
        name="Protocols",
        columns=(
            Column("protocol_id", "Stable protocol or SOP ID.", True),
            Column("name", "Protocol name.", True),
            Column("version", "Controlled version.", True),
            Column("effective_date", "Date this version became effective."),
            Column("owner", "Protocol owner."),
            Column("status", "draft, active, or retired."),
            Column("source_url", "Link to the controlled source document."),
            Column("change_summary", "Summary of changes from the prior version."),
            Column("notes", "Additional use notes."),
        ),
    ),
    SheetSpec(
        name="Specifications",
        columns=(
            Column("specification_id", "Stable specification ID.", True),
            Column("scope_type", "experiment, project, product, sample_type, or method."),
            Column("scope_id", "ID or label for the specification scope."),
            Column("measurement_type", "Measurement assessed by this specification.", True),
            Column("target_value", "Nominal or desired numeric value."),
            Column("lower_limit", "Lower acceptable result."),
            Column("upper_limit", "Upper acceptable result."),
            Column("units", "Units for target and limits."),
            Column("method", "Required method or condition."),
            Column("effective_date", "Date the specification became effective."),
            Column("status", "draft, active, or retired."),
            Column("notes", "Interpretation and exception notes."),
        ),
    ),
    SheetSpec(
        name="Deviations",
        columns=(
            Column("deviation_id", "Stable deviation ID.", True),
            Column("experiment_id", "Experiment affected by the deviation.", True),
            Column("step_id", "Run Capture Plan step affected."),
            Column("occurred_at", "Deviation timestamp."),
            Column("category", "Process, material, equipment, method, safety, or data category."),
            Column("description", "What differed from the approved plan.", True),
            Column("immediate_action", "Action taken during the run."),
            Column("impact_assessment", "Potential impact on safety, validity, or quality."),
            Column("disposition", "Use, repeat, investigate, reject, or other decision."),
            Column("owner", "Person responsible for follow-up."),
            Column("status", "open, under_review, or closed."),
            Column("closed_at", "Closure timestamp."),
            Column("reviewer", "Person approving the disposition."),
            Column("notes", "Additional investigation notes."),
        ),
    ),
    SheetSpec(
        name="Raw Data Files",
        columns=(
            Column("raw_file_id", "Stable raw file record ID.", True),
            Column("experiment_id", "Experiment linked to the file.", True),
            Column("sample_id", "Sample linked to the file."),
            Column("measurement_type", "Measurement represented by the file."),
            Column("instrument_id", "Equipment ID that produced the file."),
            Column("collected_at", "Acquisition timestamp."),
            Column("file_name", "Original file name.", True),
            Column("file_url", "Immutable or controlled Drive link.", True),
            Column("checksum_sha256", "SHA-256 checksum when available."),
            Column("method", "Acquisition or parsing method."),
            Column("parser_status", "unprocessed, parsed, needs_review, or failed."),
            Column("notes", "File provenance and parser notes."),
        ),
    ),
    SheetSpec(
        name="Audit Log",
        columns=(
            Column("audit_id", "Stable audit event ID.", True),
            Column("occurred_at", "Audit event timestamp.", True),
            Column("actor", "Person or agent responsible for the action."),
            Column("action", "create, update, correct, delete, import, or migrate.", True),
            Column("sheet_name", "Affected sheet."),
            Column("row_key", "Stable key of the affected row."),
            Column("field_name", "Affected field when applicable."),
            Column("old_value", "Value before the change."),
            Column("new_value", "Value after the change."),
            Column("reason", "Reason for the action or correction."),
            Column("source", "CLI command, manual edit, import, or migration source."),
            Column("notes", "Additional audit context."),
        ),
    ),
)

RUN_CONSOLE_SHEET = "Run Console"

NUMBER_COLUMNS: dict[str, frozenset[str]] = {
    "Master Reagents": frozenset(
        {
            "molecular_weight_g_mol",
            "density_g_mL",
            "purity_fraction",
            "concentration",
        }
    ),
    "Daily Log": frozenset(
        {
            "temperature_C",
            "rpm",
            "pH",
            "solids_percent",
            "particle_size_nm",
            "conversion_percent",
            "viscosity_cP",
            "residual_monomer_percent",
            "polydispersity_index",
            "Tg_C",
            "hold_time_min",
        }
    ),
    "Formulations": frozenset(
        {
            "mass_g",
            "volume_mL",
            "moles_mmol",
            "concentration",
            "wt_percent",
            "feed_order",
            "feed_start_min",
            "feed_duration_min",
            "actual_mass_g",
            "mass_variance_g",
            "actual_volume_mL",
        }
    ),
    "Batch Builder": frozenset(
        {
            "feed_order",
            "parts_per_hundred_monomer",
            "stage_monomer_basis_g",
            "direct_mass_g",
            "planned_mass_g",
            "stock_active_fraction",
            "active_mass_g",
            "density_override_g_mL",
            "density_g_mL",
            "planned_volume_mL",
            "actual_mass_g",
            "mass_variance_g",
            "actual_volume_mL",
            "feed_start_min",
            "feed_duration_min",
            "feed_rate_mL_min",
            "target_temperature_C",
            "molecular_weight_g_mol",
            "planned_moles_mmol",
            "equivalent_basis_mmol",
            "equivalents",
            "actual_moles_mmol",
        }
    ),
    "Bench Log": frozenset(
        {
            "Temperature (°C)",
            "RPM",
            "pH",
            "Solids (%)",
            "Particle size (nm)",
            "Conversion (%)",
            "Viscosity (cP)",
        }
    ),
    "Measurements": frozenset({"Numeric value", "Uncertainty"}),
    "Results": frozenset(
        {
            "replicate",
            "numeric_value",
            "uncertainty",
            "detection_limit",
        }
    ),
    "Literature Evidence": frozenset({"year"}),
    "Daily Reviews": frozenset(
        {
            "experiment_count",
            "observation_count",
            "result_count",
            "normalized_result_rows_to_append",
            "evidence_rows_to_append",
            "suggestion_rows_to_append",
            "preflight_fail_count",
            "preflight_warn_count",
            "apply_request_count",
        }
    ),
    "Project Notebook Records": frozenset(
        {
            "planned_value",
            "actual_value",
            "parts_per_hundred_monomer",
            "mass_g",
            "actual_mass_g",
            "mass_variance_g",
            "volume_mL",
            "density_g_mL",
            "concentration",
            "elapsed_start_min",
            "elapsed_end_min",
            "feed_rate_mL_min",
            "cumulative_percent",
            "radical_flux_mol_min_L",
            "particle_size_nm",
            "temperature_C",
            "rpm",
            "oil_temperature_C",
            "torque",
            "cumulative_addition_g",
        }
    ),
    "Source Sync": frozenset({"source_sheet_id", "record_count"}),
    "Plot Data": frozenset(
        {
            "x_order",
            "x_value",
            "series_1_value",
            "series_2_value",
            "series_3_value",
            "series_4_value",
        }
    ),
    "Plot Definitions": frozenset(
        {
            "data_start_row",
            "data_end_row",
            "source_point_count",
        }
    ),
    "Plot Dashboard": frozenset({"point_count"}),
    "Run Capture Plan": frozenset(
        {
            "sequence",
            "planned_time_min",
            "target_value",
            "lower_limit",
            "upper_limit",
            "actual_time_min",
        }
    ),
    "Samples": frozenset({"amount"}),
    "Specifications": frozenset({"target_value", "lower_limit", "upper_limit"}),
}

DATE_COLUMNS: dict[str, frozenset[str]] = {
    "Master Reagents": frozenset({"expiration_date"}),
    "Experiments": frozenset({"date"}),
    "Daily Reviews": frozenset({"review_date"}),
    "Equipment": frozenset({"last_calibration_date", "calibration_due_date"}),
    "Protocols": frozenset({"effective_date"}),
    "Specifications": frozenset({"effective_date"}),
}

DATETIME_COLUMNS: dict[str, frozenset[str]] = {
    "Experiments": frozenset(
        {
            "source_modified_at",
            "started_at",
            "completed_at",
            "reviewed_at",
        }
    ),
    "Daily Log": frozenset({"timestamp"}),
    "Formulations": frozenset({"recorded_at"}),
    "Batch Builder": frozenset({"recorded_at"}),
    "Bench Log": frozenset({"Date & time"}),
    "Measurements": frozenset({"Measured at"}),
    "Results": frozenset({"measured_at", "reviewed_at"}),
    "Agent Suggestions": frozenset({"created_at"}),
    "Daily Reviews": frozenset({"created_at"}),
    "Project Notebook Records": frozenset({"source_timestamp", "synced_at"}),
    "Source Sync": frozenset({"source_modified_at", "synced_at"}),
    "Plot Data": frozenset({"recorded_at"}),
    "Plot Definitions": frozenset({"updated_at"}),
    "Plot Dashboard": frozenset({"updated_at"}),
    "Workbook Metadata": frozenset({"updated_at"}),
    "Run Capture Plan": frozenset({"completed_at"}),
    "Samples": frozenset({"collected_at"}),
    "Deviations": frozenset({"occurred_at", "closed_at"}),
    "Raw Data Files": frozenset({"collected_at"}),
    "Audit Log": frozenset({"occurred_at"}),
}


def column_data_type(sheet_name: str, column_name: str) -> str:
    if column_name in NUMBER_COLUMNS.get(sheet_name, ()):
        return "number"
    if column_name in DATE_COLUMNS.get(sheet_name, ()):
        return "date"
    if column_name in DATETIME_COLUMNS.get(sheet_name, ()):
        return "datetime"
    return "text"


def column_number_format(sheet_name: str, column_name: str) -> str | None:
    data_type = column_data_type(sheet_name, column_name)
    if data_type == "number":
        return "0.########"
    if data_type == "date":
        return "yyyy-mm-dd"
    if data_type == "datetime":
        return "yyyy-mm-dd hh:mm:ss"
    return None


def sheet_by_name(name: str) -> SheetSpec:
    for sheet in SHEETS:
        if sheet.name == name:
            return sheet
    raise KeyError(name)


def workbook_contract() -> dict[str, object]:
    return {
        "name": "lab-notebook-agent-workbook",
        "version": WORKBOOK_CONTRACT_VERSION,
        "sheets": [
            {
                "name": sheet.name,
                "headers": list(sheet.headers),
                "columns": [
                    {
                        "name": column.name,
                        "description": column.description,
                        "required": column.required,
                        "data_type": column_data_type(sheet.name, column.name),
                        "number_format": column_number_format(sheet.name, column.name),
                    }
                    for column in sheet.columns
                ],
            }
            for sheet in SHEETS
        ],
        "views": [
            {
                "name": RUN_CONSOLE_SHEET,
                "kind": "scientist_run_console",
                "purpose": (
                    "Scientist-facing experiment overview, readiness checks, "
                    "bench workflow, and links into normalized record tables."
                ),
                "active_experiment_cell": "B3",
                "preserve_user_cells": ["B3"],
                "active_queue_anchor": "A30",
                "batch_entry_sheet": "Batch Builder",
                "active_queue_rule": (
                    "status is planned or running, or status is needs_review "
                    "and source_notebook_id is blank"
                ),
            }
        ],
        "controlled_vocab": {
            "process_type": list(PROCESS_TYPES),
            "reagent_category": list(REAGENT_CATEGORIES),
            "formulation_role": list(FORMULATION_ROLES),
            "batch_stage": list(BATCH_STAGES),
            "batch_charge_type": list(BATCH_CHARGE_TYPES),
            "batch_charge_status": list(BATCH_CHARGE_STATUS),
            "experiment_status": list(EXPERIMENT_STATUS),
            "process_stage": list(PROCESS_STAGES),
            "result_quality_flag": list(RESULT_QUALITY_FLAGS),
            "suggestion_status": list(SUGGESTION_STATUS),
            "daily_review_status": list(DAILY_REVIEW_STATUS),
            "run_step_status": list(RUN_STEP_STATUS),
            "sample_status": list(SAMPLE_STATUS),
            "calibration_status": list(CALIBRATION_STATUS),
            "protocol_status": list(PROTOCOL_STATUS),
            "specification_status": list(SPECIFICATION_STATUS),
            "qc_status": list(QC_STATUS),
            "deviation_status": list(DEVIATION_STATUS),
            "raw_file_parser_status": list(RAW_FILE_PARSER_STATUS),
            "audit_action": list(AUDIT_ACTIONS),
            "inventory_status": list(INVENTORY_STATUS),
        },
    }
