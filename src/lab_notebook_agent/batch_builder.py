from __future__ import annotations

from typing import Any


BATCH_BUILDER_FORMULA_COLUMNS = (
    "material_name",
    "planned_mass_g",
    "active_mass_g",
    "density_g_mL",
    "planned_volume_mL",
    "mass_variance_g",
    "actual_volume_mL",
    "feed_rate_mL_min",
)


def excel_batch_builder_formula(column_name: str, row_number: int) -> str:
    """Return a row-local Excel formula for a calculated Batch Builder column."""

    row = row_number
    formulas = {
        "material_name": (
            f'=IF($A{row}="","",IFERROR(VLOOKUP($G{row},'
            f"'Master Reagents'!$A$2:$B$1000,2,FALSE),\"Unknown reagent\"))"
        ),
        "planned_mass_g": (
            f'=IF($A{row}="","",IF($K{row}<>"",$K{row},'
            f'IF(AND(ISNUMBER($I{row}),ISNUMBER($J{row})),$I{row}*$J{row}/100,"")))'
        ),
        "active_mass_g": (
            f'=IF($L{row}="","",$L{row}*IF($M{row}="",1,$M{row}))'
        ),
        "density_g_mL": (
            f'=IF($A{row}="","",IF($O{row}<>"",$O{row},IFERROR(VLOOKUP('
            f"$G{row},'Master Reagents'!$A$2:$G$1000,7,FALSE),\"\")))"
        ),
        "planned_volume_mL": (
            f'=IF(OR($L{row}="",$P{row}="",$P{row}=0),"",$L{row}/$P{row})'
        ),
        "mass_variance_g": (
            f'=IF($R{row}="","",$R{row}-$L{row})'
        ),
        "actual_volume_mL": (
            f'=IF(OR($R{row}="",$P{row}="",$P{row}=0),"",$R{row}/$P{row})'
        ),
        "feed_rate_mL_min": (
            f'=IF(OR($Q{row}="",$V{row}="",$V{row}=0),"",$Q{row}/$V{row})'
        ),
    }
    try:
        return formulas[column_name]
    except KeyError as exc:
        raise KeyError(f"Unknown Batch Builder formula column {column_name!r}.") from exc


def google_batch_builder_array_formulas(end_row: int = 1000) -> dict[str, str]:
    """Return spill formulas for the calculated Batch Builder columns."""

    end_row = max(2, int(end_row))
    return {
        "material_name": (
            f'=ARRAYFORMULA(IF(A2:A{end_row}="","",IFNA(VLOOKUP('
            f"G2:G{end_row},'Master Reagents'!A2:B1000,2,FALSE),\"Unknown reagent\")))"
        ),
        "planned_mass_g": (
            f'=ARRAYFORMULA(IF(A2:A{end_row}="","",IF(K2:K{end_row}<>"",'
            f'K2:K{end_row},IF((I2:I{end_row}<>"")*(J2:J{end_row}<>""),'
            f'I2:I{end_row}*J2:J{end_row}/100,""))))'
        ),
        "active_mass_g": (
            f'=ARRAYFORMULA(IF(L2:L{end_row}="","",L2:L{end_row}*'
            f'IF(M2:M{end_row}="",1,M2:M{end_row})))'
        ),
        "density_g_mL": (
            f'=ARRAYFORMULA(IF(A2:A{end_row}="","",IF(O2:O{end_row}<>"",'
            f'O2:O{end_row},IFNA(VLOOKUP(G2:G{end_row},'
            f"'Master Reagents'!A2:G1000,7,FALSE),\"\"))))"
        ),
        "planned_volume_mL": (
            f'=ARRAYFORMULA(IF((L2:L{end_row}="")+(P2:P{end_row}="")+'
            f'(P2:P{end_row}=0),"",L2:L{end_row}/P2:P{end_row}))'
        ),
        "mass_variance_g": (
            f'=ARRAYFORMULA(IF(R2:R{end_row}="","",R2:R{end_row}-L2:L{end_row}))'
        ),
        "actual_volume_mL": (
            f'=ARRAYFORMULA(IF((R2:R{end_row}="")+(P2:P{end_row}="")+'
            f'(P2:P{end_row}=0),"",R2:R{end_row}/P2:P{end_row}))'
        ),
        "feed_rate_mL_min": (
            f'=ARRAYFORMULA(IF((Q2:Q{end_row}="")+(V2:V{end_row}="")+'
            f'(V2:V{end_row}=0),"",Q2:Q{end_row}/V2:V{end_row}))'
        ),
    }


def batch_builder_to_formulation_row(row: dict[str, Any]) -> dict[str, Any]:
    """Project one scientist-facing charge row into the normalized formulation shape."""

    stage = str(row.get("stage", "")).strip()
    charge_type = str(row.get("charge_type", "")).strip()
    charge_id = str(row.get("charge_id", "")).strip()
    notes = str(row.get("notes", "")).strip()
    if charge_id:
        notes = f"{notes} [charge_id={charge_id}]".strip()
    projected = {
        "experiment_id": row.get("experiment_id", ""),
        "reagent_id": row.get("reagent_id", ""),
        "phase": " / ".join(value for value in (stage, charge_type) if value),
        "target_role": row.get("target_role", ""),
        "mass_g": row.get("planned_mass_g", "") or row.get("direct_mass_g", ""),
        "volume_mL": row.get("planned_volume_mL", ""),
        "moles_mmol": "",
        "concentration": "",
        "concentration_units": "",
        "wt_percent": "",
        "feed_order": row.get("feed_order", ""),
        "feed_start_min": row.get("feed_start_min", ""),
        "feed_duration_min": row.get("feed_duration_min", ""),
        "notes": notes,
        "actual_mass_g": row.get("actual_mass_g", ""),
        "mass_variance_g": row.get("mass_variance_g", ""),
        "actual_volume_mL": row.get("actual_volume_mL", ""),
        "lot": row.get("lot", ""),
        "recorded_by": row.get("recorded_by", ""),
        "recorded_at": row.get("recorded_at", ""),
        "charge_id": charge_id,
        "stage": stage,
        "charge_type": charge_type,
        "parts_per_hundred_monomer": row.get("parts_per_hundred_monomer", ""),
        "stage_monomer_basis_g": row.get("stage_monomer_basis_g", ""),
        "active_mass_g": row.get("active_mass_g", ""),
        "feed_rate_mL_min": row.get("feed_rate_mL_min", ""),
        "target_temperature_C": row.get("target_temperature_C", ""),
        "charge_status": row.get("charge_status", ""),
        "stock_active_fraction": row.get("stock_active_fraction", ""),
    }
    material_name = row.get("material_name", "")
    if material_name not in (None, ""):
        projected["reagent_name"] = material_name
    return projected


def formulation_rows_from_tables(
    tables: dict[str, list[dict[str, Any]]],
) -> list[dict[str, Any]]:
    """Return normalized legacy rows plus scientist-entered Batch Builder charges."""

    rows_by_key: dict[tuple[str, str, str], dict[str, Any]] = {}
    for row in tables.get("Formulations", []):
        if not isinstance(row, dict):
            continue
        key = _formulation_key(row)
        if key[0] and key[1]:
            rows_by_key[key] = row
    for row in tables.get("Batch Builder", []):
        if not isinstance(row, dict):
            continue
        projected = batch_builder_to_formulation_row(row)
        key = _formulation_key(projected)
        if key[0] and key[1]:
            existing = rows_by_key.get(key)
            if existing is None:
                rows_by_key[key] = projected
            else:
                merged = dict(existing)
                merged.update(
                    {
                        field: value
                        for field, value in projected.items()
                        if value not in (None, "")
                    }
                )
                rows_by_key[key] = merged
    return list(rows_by_key.values())


def _formulation_key(row: dict[str, Any]) -> tuple[str, str, str]:
    return (
        str(row.get("experiment_id", "")).strip(),
        str(row.get("reagent_id", "")).strip(),
        str(row.get("target_role", "")).strip(),
    )
