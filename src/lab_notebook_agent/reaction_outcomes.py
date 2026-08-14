from __future__ import annotations


REACTION_OUTCOME_FORMULA_COLUMNS = (
    "isolated_yield_percent",
    "actual_input_mass_g",
    "accounted_mass_g",
    "unaccounted_mass_g",
    "material_balance_closure_percent",
    "mass_balance_status",
)


def excel_reaction_outcome_formula(header: str, row_number: int) -> str:
    """Return one calculated Reaction Outcomes formula for an Excel row."""

    row = row_number
    formulas = {
        "isolated_yield_percent": f'=IF(OR($A{row}="",$D{row}="",$D{row}=0),"",$E{row}/$D{row})',
        "actual_input_mass_g": (
            f'=IF($A{row}="","",SUMIF(\'Batch Builder\'!$A$2:$A$1000,'
            f'$A{row},\'Batch Builder\'!$BB$2:$BB$1000))'
        ),
        "accounted_mass_g": f'=IF($A{row}="","",SUM($E{row},$G{row}:$J{row}))',
        "unaccounted_mass_g": f'=IF(OR($A{row}="",$L{row}=""),"",$L{row}-$M{row})',
        "material_balance_closure_percent": (
            f'=IF(OR($A{row}="",$L{row}=0),"",$M{row}/$L{row})'
        ),
        "mass_balance_status": (
            f'=IF($A{row}="","",IF($L{row}=0,"MISSING_ACTUAL_INPUT",'
            f'IF($E{row}="","MISSING_PRODUCT_RECOVERY",'
            f'IF($K{row}="","MISSING_TOLERANCE",'
            f'IF($N{row}<-$L{row}*$K{row}/100,"OVER_ACCOUNTED",'
            f'IF(ABS($N{row})<=$L{row}*$K{row}/100,"CLOSED","OPEN"))))))'
        ),
    }
    try:
        return formulas[header]
    except KeyError as exc:
        raise KeyError(f"Unknown Reaction Outcomes formula column {header!r}") from exc


def google_reaction_outcome_array_formulas(end_row: int = 1000) -> dict[str, str]:
    """Return bounded spill formulas for Google Sheets Reaction Outcomes."""

    return {
        "isolated_yield_percent": f'=ARRAYFORMULA(IF((A2:A{end_row}="")+(D2:D{end_row}="")+(D2:D{end_row}=0),"",E2:E{end_row}/D2:D{end_row}))',
        "actual_input_mass_g": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",SUMIF(\'Batch Builder\'!A2:A{end_row},id,\'Batch Builder\'!BB2:BB{end_row}))))',
        "accounted_mass_g": f'=ARRAYFORMULA(IF(A2:A{end_row}="","",E2:E{end_row}+G2:G{end_row}+H2:H{end_row}+I2:I{end_row}+J2:J{end_row}))',
        "unaccounted_mass_g": f'=ARRAYFORMULA(IF((A2:A{end_row}="")+(L2:L{end_row}=""),"",L2:L{end_row}-M2:M{end_row}))',
        "material_balance_closure_percent": f'=ARRAYFORMULA(IF((A2:A{end_row}="")+(L2:L{end_row}=0),"",M2:M{end_row}/L2:L{end_row}))',
        "mass_balance_status": f'=MAP(A2:A{end_row},E2:E{end_row},K2:K{end_row},L2:L{end_row},N2:N{end_row},LAMBDA(id,recovered,tolerance,total,delta,IF(id="","",IF(total=0,"MISSING_ACTUAL_INPUT",IF(recovered="","MISSING_PRODUCT_RECOVERY",IF(tolerance="","MISSING_TOLERANCE",IF(delta<-(total*tolerance/100),"OVER_ACCOUNTED",IF(ABS(delta)<=total*tolerance/100,"CLOSED","OPEN"))))))))',
    }
