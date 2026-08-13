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
    "molecular_weight_g_mol",
    "planned_moles_mmol",
    "equivalents",
    "actual_moles_mmol",
    "experiment_target_mass_g",
    "experiment_default_tolerance_percent",
    "mass_variance_percent",
    "within_tolerance",
    "formula_status",
    "functional_group",
    "functional_equivalent_weight_g_eq",
    "planned_functional_equivalents",
    "actual_functional_equivalents",
    "actual_mass_from_difference_g",
    "effective_actual_mass_g",
    "planned_carrier_mass_g",
    "actual_carrier_mass_g",
    "weighing_status",
)

BATCH_CALCULATION_MODES = (
    "direct_mass",
    "batch_wt_percent",
    "pphm",
    "target_active_mass",
    "equivalents",
    "functional_equivalents",
)


def calculate_batch_mass_row(row: dict[str, Any]) -> dict[str, Any]:
    """Calculate one charge row with the same rules exposed in the workbook."""

    mode = str(row.get("calculation_mode", "")).strip()
    active_fraction = _number(row.get("stock_active_fraction"))
    if active_fraction is None:
        active_fraction = 1.0
    status = _mass_formula_status(row, mode, active_fraction)
    planned_mass: float | None = None
    if status == "READY":
        if mode == "direct_mass":
            planned_mass = _number(row.get("direct_mass_g"))
        elif mode == "batch_wt_percent":
            planned_mass = (
                _number(row.get("experiment_target_mass_g"))
                * _number(row.get("recipe_wt_percent"))
                / 100
            )
        elif mode == "pphm":
            planned_mass = (
                _number(row.get("parts_per_hundred_monomer"))
                * _number(row.get("stage_monomer_basis_g"))
                / 100
            )
        elif mode == "target_active_mass":
            planned_mass = _number(row.get("target_active_mass_g")) / active_fraction
        elif mode == "equivalents":
            planned_mass = (
                _number(row.get("target_equivalents"))
                * _number(row.get("equivalent_basis_mmol"))
                * _number(row.get("molecular_weight_g_mol"))
                / 1000
                / active_fraction
            )
        elif mode == "functional_equivalents":
            planned_mass = (
                _number(row.get("target_functional_equivalents"))
                * _number(row.get("functional_equivalent_weight_g_eq"))
                / active_fraction
            )

    active_mass = planned_mass * active_fraction if planned_mass is not None else None
    molecular_weight = _positive_number(row.get("molecular_weight_g_mol"))
    planned_moles = (
        active_mass / molecular_weight * 1000
        if active_mass is not None and molecular_weight is not None
        else None
    )
    equivalent_basis = _positive_number(row.get("equivalent_basis_mmol"))
    equivalents = (
        planned_moles / equivalent_basis
        if planned_moles is not None and equivalent_basis is not None
        else None
    )
    density = _positive_number(row.get("density_g_mL"))
    planned_volume = (
        planned_mass / density
        if planned_mass is not None and density is not None
        else None
    )
    weighing_method = str(row.get("weighing_method", "")).strip() or "direct"
    actual_mass_from_difference: float | None = None
    if weighing_method == "direct":
        effective_actual_mass = _number(row.get("actual_mass_g"))
        weighing_status = "READY" if effective_actual_mass is not None else "NOT_RECORDED"
    elif weighing_method == "by_difference":
        before = _number(row.get("source_container_before_g"))
        after = _number(row.get("source_container_after_g"))
        if before is None or after is None:
            effective_actual_mass = None
            weighing_status = "MISSING_CONTAINER_WEIGHTS"
        elif before < after or after < 0:
            effective_actual_mass = None
            weighing_status = "INVALID_CONTAINER_DIFFERENCE"
        else:
            actual_mass_from_difference = before - after
            effective_actual_mass = actual_mass_from_difference
            weighing_status = "READY"
    else:
        effective_actual_mass = None
        weighing_status = "INVALID_WEIGHING_METHOD"
    variance = (
        effective_actual_mass - planned_mass
        if effective_actual_mass is not None and planned_mass not in (None, 0)
        else None
    )
    variance_percent = (
        variance / planned_mass * 100
        if variance is not None and planned_mass not in (None, 0)
        else None
    )
    tolerance = _positive_number(row.get("mass_tolerance_percent"))
    if tolerance is None:
        tolerance = _positive_number(row.get("experiment_default_tolerance_percent"))
    within_tolerance = ""
    if variance_percent is not None and tolerance is not None:
        within_tolerance = "PASS" if abs(variance_percent) <= tolerance else "FAIL"
    actual_moles = (
        effective_actual_mass * active_fraction / molecular_weight * 1000
        if effective_actual_mass is not None
        and molecular_weight is not None
        and 0 < active_fraction <= 1
        else None
    )
    functional_equivalent_weight = _positive_number(
        row.get("functional_equivalent_weight_g_eq")
    )
    planned_functional_equivalents = (
        active_mass / functional_equivalent_weight
        if active_mass is not None and functional_equivalent_weight is not None
        else None
    )
    actual_functional_equivalents = (
        effective_actual_mass * active_fraction / functional_equivalent_weight
        if effective_actual_mass is not None
        and functional_equivalent_weight is not None
        and 0 < active_fraction <= 1
        else None
    )
    planned_carrier_mass = (
        planned_mass * (1 - active_fraction) if planned_mass is not None else None
    )
    actual_carrier_mass = (
        effective_actual_mass * (1 - active_fraction)
        if effective_actual_mass is not None
        else None
    )
    return {
        "planned_mass_g": planned_mass,
        "active_mass_g": active_mass,
        "planned_volume_mL": planned_volume,
        "planned_moles_mmol": planned_moles,
        "equivalents": equivalents,
        "mass_variance_g": variance,
        "mass_variance_percent": variance_percent,
        "within_tolerance": within_tolerance,
        "actual_moles_mmol": actual_moles,
        "functional_group": str(row.get("functional_group", "")).strip(),
        "functional_equivalent_weight_g_eq": functional_equivalent_weight,
        "planned_functional_equivalents": planned_functional_equivalents,
        "actual_functional_equivalents": actual_functional_equivalents,
        "actual_mass_from_difference_g": actual_mass_from_difference,
        "effective_actual_mass_g": effective_actual_mass,
        "planned_carrier_mass_g": planned_carrier_mass,
        "actual_carrier_mass_g": actual_carrier_mass,
        "weighing_status": weighing_status,
        "formula_status": status,
    }


def summarize_experiment_mass_plan(
    experiment: dict[str, Any], charge_rows: list[dict[str, Any]]
) -> dict[str, Any]:
    """Return the reference run-level mass balance used to verify master formulas."""

    experiment_id = str(experiment.get("experiment_id", "")).strip()
    target_mass = _positive_number(experiment.get("target_batch_mass_g"))
    tolerance = _positive_number(experiment.get("default_mass_tolerance_percent"))
    matching = [
        row
        for row in charge_rows
        if str(row.get("experiment_id", "")).strip() == experiment_id
    ]
    calculated = [
        calculate_batch_mass_row(
            {
                **row,
                "experiment_target_mass_g": target_mass,
                "experiment_default_tolerance_percent": tolerance,
            }
        )
        for row in matching
    ]
    planned_values = [
        result["planned_mass_g"]
        for result in calculated
        if result["planned_mass_g"] is not None
    ]
    actual_values = [result["effective_actual_mass_g"] for result in calculated if result["effective_actual_mass_g"] is not None]
    planned_mass = sum(planned_values)
    actual_mass = sum(actual_values)
    calculation_issues = sum(
        result["formula_status"] != "READY" for result in calculated
    )
    out_of_tolerance = sum(
        result["within_tolerance"] == "FAIL" for result in calculated
    )
    if not matching:
        status = "NO_CHARGES"
    elif calculation_issues:
        status = "CALCULATION_ISSUES"
    elif out_of_tolerance:
        status = "CHARGE_TOLERANCE_FAILURE"
    elif target_mass is None:
        status = "NO_TARGET_MASS"
    elif tolerance is None:
        status = "MISSING_TOLERANCE"
    elif abs(planned_mass - target_mass) <= target_mass * tolerance / 100:
        status = "READY"
    else:
        status = "MASS_MISMATCH"
    numerator_group = str(experiment.get("stoichiometric_numerator_group", "")).strip()
    denominator_group = str(experiment.get("stoichiometric_denominator_group", "")).strip()
    target_ratio = _positive_number(experiment.get("target_equivalent_ratio"))
    ratio_tolerance = _positive_number(experiment.get("equivalent_ratio_tolerance_percent"))
    planned_numerator = sum(
        result["planned_functional_equivalents"] or 0
        for result in calculated
        if result["functional_group"] == numerator_group
    )
    planned_denominator = sum(
        result["planned_functional_equivalents"] or 0
        for result in calculated
        if result["functional_group"] == denominator_group
    )
    actual_numerator = sum(
        result["actual_functional_equivalents"] or 0
        for result in calculated
        if result["functional_group"] == numerator_group
    )
    actual_denominator = sum(
        result["actual_functional_equivalents"] or 0
        for result in calculated
        if result["functional_group"] == denominator_group
    )
    planned_ratio = planned_numerator / planned_denominator if planned_denominator else None
    actual_ratio = actual_numerator / actual_denominator if actual_denominator else None
    if not numerator_group and not denominator_group and target_ratio is None:
        stoichiometry_status = "NOT_CONFIGURED"
    elif not numerator_group or not denominator_group or target_ratio is None:
        stoichiometry_status = "MISSING_TARGET_RATIO"
    elif ratio_tolerance is None:
        stoichiometry_status = "MISSING_RATIO_TOLERANCE"
    elif not planned_numerator:
        stoichiometry_status = "MISSING_NUMERATOR_EQ"
    elif not planned_denominator:
        stoichiometry_status = "MISSING_DENOMINATOR_EQ"
    elif abs(planned_ratio - target_ratio) <= target_ratio * ratio_tolerance / 100:
        stoichiometry_status = "READY"
    else:
        stoichiometry_status = "RATIO_MISMATCH"
    return {
        "experiment_id": experiment_id,
        "charge_count": len(matching),
        "target_batch_mass_g": target_mass,
        "planned_mass_g": planned_mass,
        "planned_vs_target_g": (
            planned_mass - target_mass if target_mass is not None else None
        ),
        "actual_mass_g": actual_mass,
        "actual_vs_target_g": (
            actual_mass - target_mass if target_mass is not None else None
        ),
        "mass_completion_fraction": (
            actual_mass / planned_mass if planned_mass > 0 else None
        ),
        "out_of_tolerance_charges": out_of_tolerance,
        "calculation_issues": calculation_issues,
        "status": status,
        "numerator_group": numerator_group,
        "denominator_group": denominator_group,
        "target_equivalent_ratio": target_ratio,
        "planned_numerator_eq": planned_numerator,
        "planned_denominator_eq": planned_denominator,
        "planned_equivalent_ratio": planned_ratio,
        "actual_numerator_eq": actual_numerator,
        "actual_denominator_eq": actual_denominator,
        "actual_equivalent_ratio": actual_ratio,
        "stoichiometry_status": stoichiometry_status,
        "charges": calculated,
    }


def _mass_formula_status(
    row: dict[str, Any], mode: str, active_fraction: float
) -> str:
    if not mode:
        return "MISSING_MODE"
    if mode not in BATCH_CALCULATION_MODES:
        return "INVALID_MODE"
    if not 0 < active_fraction <= 1:
        return "INVALID_ACTIVE_FRACTION"
    requirements = {
        "direct_mass": (("direct_mass_g", "MISSING_DIRECT_MASS"),),
        "batch_wt_percent": (
            ("experiment_target_mass_g", "MISSING_TARGET_BATCH_MASS"),
            ("recipe_wt_percent", "MISSING_RECIPE_WT_PERCENT"),
        ),
        "pphm": (
            ("parts_per_hundred_monomer", "MISSING_PPHM_BASIS"),
            ("stage_monomer_basis_g", "MISSING_PPHM_BASIS"),
        ),
        "target_active_mass": (
            ("target_active_mass_g", "MISSING_ACTIVE_TARGET"),
        ),
        "equivalents": (
            ("target_equivalents", "MISSING_EQUIVALENT_BASIS"),
            ("equivalent_basis_mmol", "MISSING_EQUIVALENT_BASIS"),
            ("molecular_weight_g_mol", "MISSING_MOLECULAR_WEIGHT"),
        ),
        "functional_equivalents": (
            ("target_functional_equivalents", "MISSING_FUNCTIONAL_EQ_TARGET"),
            ("functional_equivalent_weight_g_eq", "MISSING_FUNCTIONAL_EQUIVALENT_WEIGHT"),
        ),
    }
    for field, missing_status in requirements[mode]:
        if _positive_number(row.get(field)) is None:
            return missing_status
    return "READY"


def _number(value: Any) -> float | None:
    if value in (None, "") or isinstance(value, bool):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _positive_number(value: Any) -> float | None:
    number = _number(value)
    return number if number is not None and number > 0 else None


def excel_batch_builder_formula(column_name: str, row_number: int) -> str:
    """Return a row-local Excel formula for a calculated Batch Builder column."""

    row = row_number
    formulas = {
        "material_name": (
            f'=IF($A{row}="","",IFERROR(VLOOKUP($G{row},'
            f"'Master Reagents'!$A$2:$B$1000,2,FALSE),\"Unknown reagent\"))"
        ),
        "planned_mass_g": (
            f'=IF($A{row}="","",IF(OR($AI{row}="",AND($M{row}<>"",OR($M{row}<=0,$M{row}>1))),"",'
            f'SWITCH($AI{row},"direct_mass",IF(AND(ISNUMBER($K{row}),$K{row}>0),$K{row},""),'
            f'"batch_wt_percent",IF(AND(ISNUMBER($AJ{row}),$AJ{row}>0,ISNUMBER($AK{row}),$AK{row}>0),$AJ{row}*$AK{row}/100,""),'
            f'"pphm",IF(AND(ISNUMBER($I{row}),$I{row}>0,ISNUMBER($J{row}),$J{row}>0),$I{row}*$J{row}/100,""),'
            f'"target_active_mass",IF(AND(ISNUMBER($AL{row}),$AL{row}>0),$AL{row}/IF($M{row}="",1,$M{row}),""),'
            f'"equivalents",IF(AND(ISNUMBER($AM{row}),$AM{row}>0,ISNUMBER($AF{row}),$AF{row}>0,ISNUMBER($AD{row}),$AD{row}>0),'
            f'$AM{row}*$AF{row}*$AD{row}/1000/IF($M{row}="",1,$M{row}),""),'
            f'"functional_equivalents",IF(AND(ISNUMBER($AU{row}),$AU{row}>0,ISNUMBER($AT{row}),$AT{row}>0),'
            f'$AU{row}*$AT{row}/IF($M{row}="",1,$M{row}),""),"")))'
        ),
        "active_mass_g": (
            f'=IF($L{row}="","",$L{row}*IF($M{row}="",1,$M{row}))'
        ),
        "density_g_mL": (
            f'=IF($A{row}="","",IF($O{row}<>"",$O{row},IFERROR(VLOOKUP('
            f"$G{row},'Master Reagents'!$A$2:$G$1000,7,FALSE),\"\")))"
        ),
        "planned_volume_mL": (
            f'=IF(OR($L{row}="",$P{row}="",$P{row}<=0),"",$L{row}/$P{row})'
        ),
        "mass_variance_g": (
            f'=IF(OR($BB{row}="",$L{row}=""),"",$BB{row}-$L{row})'
        ),
        "actual_volume_mL": (
            f'=IF(OR($BB{row}="",$P{row}="",$P{row}<=0),"",$BB{row}/$P{row})'
        ),
        "feed_rate_mL_min": (
            f'=IF(OR($Q{row}="",$V{row}="",$V{row}=0),"",$Q{row}/$V{row})'
        ),
        "molecular_weight_g_mol": (
            f'=IF($A{row}="","",IFERROR(VLOOKUP($G{row},'
            f"'Master Reagents'!$A$2:$F$1000,6,FALSE),\"\"))"
        ),
        "planned_moles_mmol": (
            f'=IF(OR($N{row}="",$AD{row}="",$AD{row}<=0),"",$N{row}/$AD{row}*1000)'
        ),
        "equivalents": (
            f'=IF(OR($AE{row}="",$AF{row}="",$AF{row}<=0),"",$AE{row}/$AF{row})'
        ),
        "actual_moles_mmol": (
            f'=IF(OR($BB{row}="",$AD{row}="",$AD{row}<=0,AND($M{row}<>"",OR($M{row}<=0,$M{row}>1))),"",$BB{row}*IF($M{row}="",1,$M{row})/$AD{row}*1000)'
        ),
        "experiment_target_mass_g": (
            f'=IF($A{row}="","",IFERROR(VLOOKUP($A{row},Experiments!$A$2:$X$1000,23,FALSE),""))'
        ),
        "experiment_default_tolerance_percent": (
            f'=IF($A{row}="","",IFERROR(VLOOKUP($A{row},Experiments!$A$2:$X$1000,24,FALSE),""))'
        ),
        "mass_variance_percent": (
            f'=IF(OR($S{row}="",$L{row}="",$L{row}=0),"",$S{row}/$L{row}*100)'
        ),
        "within_tolerance": (
            f'=IF($AP{row}="","",IF(OR(IF($AN{row}<>"",$AN{row},$AO{row})="",IF($AN{row}<>"",$AN{row},$AO{row})<=0),"",'
            f'IF(ABS($AP{row})<=IF($AN{row}<>"",$AN{row},$AO{row}),"PASS","FAIL")))'
        ),
        "formula_status": (
            f'=IF($A{row}="","",IF($AI{row}="","MISSING_MODE",IF(AND($M{row}<>"",OR($M{row}<=0,$M{row}>1)),'
            f'"INVALID_ACTIVE_FRACTION",SWITCH($AI{row},'
            f'"direct_mass",IF(AND(ISNUMBER($K{row}),$K{row}>0),"READY","MISSING_DIRECT_MASS"),'
            f'"batch_wt_percent",IF(OR(NOT(ISNUMBER($AK{row})),$AK{row}<=0),"MISSING_TARGET_BATCH_MASS",'
            f'IF(AND(ISNUMBER($AJ{row}),$AJ{row}>0),"READY","MISSING_RECIPE_WT_PERCENT")),'
            f'"pphm",IF(AND(ISNUMBER($I{row}),$I{row}>0,ISNUMBER($J{row}),$J{row}>0),"READY","MISSING_PPHM_BASIS"),'
            f'"target_active_mass",IF(AND(ISNUMBER($AL{row}),$AL{row}>0),"READY","MISSING_ACTIVE_TARGET"),'
            f'"equivalents",IF(OR(NOT(ISNUMBER($AM{row})),$AM{row}<=0,NOT(ISNUMBER($AF{row})),$AF{row}<=0),'
            f'"MISSING_EQUIVALENT_BASIS",IF(AND(ISNUMBER($AD{row}),$AD{row}>0),"READY","MISSING_MOLECULAR_WEIGHT")),'
            f'"functional_equivalents",IF(OR(NOT(ISNUMBER($AU{row})),$AU{row}<=0),"MISSING_FUNCTIONAL_EQ_TARGET",'
            f'IF(AND(ISNUMBER($AT{row}),$AT{row}>0),"READY","MISSING_FUNCTIONAL_EQUIVALENT_WEIGHT")),'
            f'"INVALID_MODE"))))'
        ),
        "functional_group": f'=IF($A{row}="","",IFERROR(VLOOKUP($G{row},\'Master Reagents\'!$A$2:$V$1000,20,FALSE),""))',
        "functional_equivalent_weight_g_eq": f'=IF($A{row}="","",IFERROR(VLOOKUP($G{row},\'Master Reagents\'!$A$2:$V$1000,21,FALSE),""))',
        "planned_functional_equivalents": f'=IF(OR($N{row}="",$AT{row}="",$AT{row}<=0),"",$N{row}/$AT{row})',
        "actual_functional_equivalents": f'=IF(OR($BB{row}="",$AT{row}="",$AT{row}<=0),"",$BB{row}*IF($M{row}="",1,$M{row})/$AT{row})',
        "actual_mass_from_difference_g": f'=IF(OR($AY{row}="",$AZ{row}="",$AY{row}<$AZ{row},$AZ{row}<0),"",$AY{row}-$AZ{row})',
        "effective_actual_mass_g": f'=IF($A{row}="","",IF(OR($AX{row}="",$AX{row}="direct"),$R{row},IF($AX{row}="by_difference",$BA{row},"")))',
        "planned_carrier_mass_g": f'=IF($L{row}="","",$L{row}*(1-IF($M{row}="",1,$M{row})))',
        "actual_carrier_mass_g": f'=IF($BB{row}="","",$BB{row}*(1-IF($M{row}="",1,$M{row})))',
        "weighing_status": f'=IF($A{row}="","",IF(OR($AX{row}="",$AX{row}="direct"),IF($R{row}="","NOT_RECORDED","READY"),IF($AX{row}="by_difference",IF(OR($AY{row}="",$AZ{row}=""),"MISSING_CONTAINER_WEIGHTS",IF(OR($AY{row}<$AZ{row},$AZ{row}<0),"INVALID_CONTAINER_DIFFERENCE","READY")),"INVALID_WEIGHING_METHOD")))',
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
            f'=MAP(A2:A{end_row},I2:I{end_row},J2:J{end_row},K2:K{end_row},M2:M{end_row},'
            f'AD2:AD{end_row},AF2:AF{end_row},AI2:AI{end_row},AJ2:AJ{end_row},AK2:AK{end_row},'
            f'AL2:AL{end_row},AM2:AM{end_row},AT2:AT{end_row},AU2:AU{end_row},LAMBDA(id,pphm,basis,direct,active,mw,eqbasis,mode,wtpct,target,active_target,target_eq,functional_eq_weight,target_functional_eq,'
            f'IF(id="","",IF(OR(mode="",AND(active<>"",OR(active<=0,active>1))),"",SWITCH(mode,'
            f'"direct_mass",IF(AND(ISNUMBER(direct),direct>0),direct,""),"batch_wt_percent",IF(AND(ISNUMBER(wtpct),wtpct>0,ISNUMBER(target),target>0),wtpct*target/100,""),'
            f'"pphm",IF(AND(ISNUMBER(pphm),pphm>0,ISNUMBER(basis),basis>0),pphm*basis/100,""),'
            f'"target_active_mass",IF(AND(ISNUMBER(active_target),active_target>0),active_target/IF(active="",1,active),""),'
            f'"equivalents",IF(AND(ISNUMBER(target_eq),target_eq>0,ISNUMBER(eqbasis),eqbasis>0,ISNUMBER(mw),mw>0),target_eq*eqbasis*mw/1000/IF(active="",1,active),""),'
            f'"functional_equivalents",IF(AND(ISNUMBER(target_functional_eq),target_functional_eq>0,ISNUMBER(functional_eq_weight),functional_eq_weight>0),target_functional_eq*functional_eq_weight/IF(active="",1,active),""),"")))))'
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
            f'(P2:P{end_row}<=0),"",L2:L{end_row}/P2:P{end_row}))'
        ),
        "mass_variance_g": (
            f'=ARRAYFORMULA(IF((BB2:BB{end_row}="")+(L2:L{end_row}=""),"",BB2:BB{end_row}-L2:L{end_row}))'
        ),
        "actual_volume_mL": (
            f'=ARRAYFORMULA(IF((BB2:BB{end_row}="")+(P2:P{end_row}="")+'
            f'(P2:P{end_row}<=0),"",BB2:BB{end_row}/P2:P{end_row}))'
        ),
        "feed_rate_mL_min": (
            f'=ARRAYFORMULA(IF((Q2:Q{end_row}="")+(V2:V{end_row}="")+'
            f'(V2:V{end_row}=0),"",Q2:Q{end_row}/V2:V{end_row}))'
        ),
        "molecular_weight_g_mol": (
            f'=ARRAYFORMULA(IF(A2:A{end_row}="","",IFNA(VLOOKUP('
            f"G2:G{end_row},'Master Reagents'!A2:F1000,6,FALSE),\"\")))"
        ),
        "planned_moles_mmol": (
            f'=ARRAYFORMULA(IF((N2:N{end_row}="")+(AD2:AD{end_row}="")+'
            f'(AD2:AD{end_row}<=0),"",N2:N{end_row}/AD2:AD{end_row}*1000))'
        ),
        "equivalents": (
            f'=ARRAYFORMULA(IF((AE2:AE{end_row}="")+(AF2:AF{end_row}="")+'
            f'(AF2:AF{end_row}<=0),"",AE2:AE{end_row}/AF2:AF{end_row}))'
        ),
        "actual_moles_mmol": (
            f'=ARRAYFORMULA(IF((BB2:BB{end_row}="")+(AD2:AD{end_row}="")+(AD2:AD{end_row}<=0)+'
            f'((M2:M{end_row}<>"")*((M2:M{end_row}<=0)+(M2:M{end_row}>1))),"",BB2:BB{end_row}*IF(M2:M{end_row}="",1,M2:M{end_row})/AD2:AD{end_row}*1000))'
        ),
        "experiment_target_mass_g": (
            f'=ARRAYFORMULA(IF(A2:A{end_row}="","",IFNA(VLOOKUP(A2:A{end_row},Experiments!A2:X1000,23,FALSE),"")))'
        ),
        "experiment_default_tolerance_percent": (
            f'=ARRAYFORMULA(IF(A2:A{end_row}="","",IFNA(VLOOKUP(A2:A{end_row},Experiments!A2:X1000,24,FALSE),"")))'
        ),
        "mass_variance_percent": (
            f'=ARRAYFORMULA(IF((S2:S{end_row}="")+(L2:L{end_row}="")+(L2:L{end_row}=0),"",S2:S{end_row}/L2:L{end_row}*100))'
        ),
        "within_tolerance": (
            f'=MAP(AP2:AP{end_row},AN2:AN{end_row},AO2:AO{end_row},LAMBDA(error,rowtol,deftol,'
            f'IF(error="","",IF(OR(IF(rowtol<>"",rowtol,deftol)="",IF(rowtol<>"",rowtol,deftol)<=0),"",IF(ABS(error)<=IF(rowtol<>"",rowtol,deftol),"PASS","FAIL")))))'
        ),
        "formula_status": (
            f'=MAP(A2:A{end_row},I2:I{end_row},J2:J{end_row},K2:K{end_row},M2:M{end_row},AD2:AD{end_row},'
            f'AF2:AF{end_row},AI2:AI{end_row},AJ2:AJ{end_row},AK2:AK{end_row},AL2:AL{end_row},AM2:AM{end_row},AT2:AT{end_row},AU2:AU{end_row},'
            f'LAMBDA(id,pphm,basis,direct,active,mw,eqbasis,mode,wtpct,target,active_target,target_eq,functional_eq_weight,target_functional_eq,IF(id="","",'
            f'IF(mode="","MISSING_MODE",IF(AND(active<>"",OR(active<=0,active>1)),"INVALID_ACTIVE_FRACTION",SWITCH(mode,'
            f'"direct_mass",IF(AND(ISNUMBER(direct),direct>0),"READY","MISSING_DIRECT_MASS"),'
            f'"batch_wt_percent",IF(OR(NOT(ISNUMBER(target)),target<=0),"MISSING_TARGET_BATCH_MASS",IF(AND(ISNUMBER(wtpct),wtpct>0),"READY","MISSING_RECIPE_WT_PERCENT")),'
            f'"pphm",IF(AND(ISNUMBER(pphm),pphm>0,ISNUMBER(basis),basis>0),"READY","MISSING_PPHM_BASIS"),'
            f'"target_active_mass",IF(AND(ISNUMBER(active_target),active_target>0),"READY","MISSING_ACTIVE_TARGET"),'
            f'"equivalents",IF(OR(NOT(ISNUMBER(target_eq)),target_eq<=0,NOT(ISNUMBER(eqbasis)),eqbasis<=0),"MISSING_EQUIVALENT_BASIS",'
            f'IF(AND(ISNUMBER(mw),mw>0),"READY","MISSING_MOLECULAR_WEIGHT")),'
            f'"functional_equivalents",IF(OR(NOT(ISNUMBER(target_functional_eq)),target_functional_eq<=0),"MISSING_FUNCTIONAL_EQ_TARGET",IF(AND(ISNUMBER(functional_eq_weight),functional_eq_weight>0),"READY","MISSING_FUNCTIONAL_EQUIVALENT_WEIGHT")),"INVALID_MODE"))))))'
        ),
        "functional_group": f'=ARRAYFORMULA(IF(A2:A{end_row}="","",IFNA(VLOOKUP(G2:G{end_row},\'Master Reagents\'!A2:V1000,20,FALSE),"")))',
        "functional_equivalent_weight_g_eq": f'=ARRAYFORMULA(IF(A2:A{end_row}="","",IFNA(VLOOKUP(G2:G{end_row},\'Master Reagents\'!A2:V1000,21,FALSE),"")))',
        "planned_functional_equivalents": f'=ARRAYFORMULA(IF((N2:N{end_row}="")+(AT2:AT{end_row}="")+(AT2:AT{end_row}<=0),"",N2:N{end_row}/AT2:AT{end_row}))',
        "actual_functional_equivalents": f'=ARRAYFORMULA(IF((BB2:BB{end_row}="")+(AT2:AT{end_row}="")+(AT2:AT{end_row}<=0),"",BB2:BB{end_row}*IF(M2:M{end_row}="",1,M2:M{end_row})/AT2:AT{end_row}))',
        "actual_mass_from_difference_g": f'=ARRAYFORMULA(IF((AY2:AY{end_row}="")+(AZ2:AZ{end_row}="")+(AY2:AY{end_row}<AZ2:AZ{end_row})+(AZ2:AZ{end_row}<0),"",AY2:AY{end_row}-AZ2:AZ{end_row}))',
        "effective_actual_mass_g": f'=MAP(A2:A{end_row},R2:R{end_row},AX2:AX{end_row},BA2:BA{end_row},LAMBDA(id,direct,method,difference,IF(id="","",IF(OR(method="",method="direct"),direct,IF(method="by_difference",difference,"")))))',
        "planned_carrier_mass_g": f'=ARRAYFORMULA(IF(L2:L{end_row}="","",L2:L{end_row}*(1-IF(M2:M{end_row}="",1,M2:M{end_row}))))',
        "actual_carrier_mass_g": f'=ARRAYFORMULA(IF(BB2:BB{end_row}="","",BB2:BB{end_row}*(1-IF(M2:M{end_row}="",1,M2:M{end_row}))))',
        "weighing_status": f'=MAP(A2:A{end_row},R2:R{end_row},AX2:AX{end_row},AY2:AY{end_row},AZ2:AZ{end_row},LAMBDA(id,direct,method,before,after,IF(id="","",IF(OR(method="",method="direct"),IF(direct="","NOT_RECORDED","READY"),IF(method="by_difference",IF(OR(before="",after=""),"MISSING_CONTAINER_WEIGHTS",IF(OR(before<after,after<0),"INVALID_CONTAINER_DIFFERENCE","READY")),"INVALID_WEIGHING_METHOD")))))',
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
        "moles_mmol": row.get("planned_moles_mmol", ""),
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
        "molecular_weight_g_mol": row.get("molecular_weight_g_mol", ""),
        "equivalents": row.get("equivalents", ""),
        "actual_moles_mmol": row.get("actual_moles_mmol", ""),
        "calculation_mode": row.get("calculation_mode", ""),
        "recipe_wt_percent": row.get("recipe_wt_percent", ""),
        "target_active_mass_g": row.get("target_active_mass_g", ""),
        "target_equivalents": row.get("target_equivalents", ""),
        "mass_tolerance_percent": row.get("mass_tolerance_percent", ""),
        "mass_variance_percent": row.get("mass_variance_percent", ""),
        "within_tolerance": row.get("within_tolerance", ""),
        "formula_status": row.get("formula_status", ""),
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
