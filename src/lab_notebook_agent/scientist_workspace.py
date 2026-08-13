from __future__ import annotations

from typing import Any


BENCH_LOG_TO_DAILY_LOG = {
    "Run ID": "experiment_id",
    "Date & time": "timestamp",
    "Stage": "process_stage",
    "Temperature (°C)": "temperature_C",
    "RPM": "rpm",
    "pH": "pH",
    "Solids (%)": "solids_percent",
    "Particle size (nm)": "particle_size_nm",
    "Conversion (%)": "conversion_percent",
    "Viscosity (cP)": "viscosity_cP",
    "Observation / action": "observation",
    "Issue tags": "issue_tags",
    "Attachment link": "attachments_url",
}

MEASUREMENTS_TO_RESULTS = {
    "Run ID": "experiment_id",
    "Sample ID": "sample_id",
    "Measurement": "measurement_type",
    "Method": "method",
    "Numeric value": "numeric_value",
    "Units": "units",
    "Condition": "condition",
    "Replicate": "replicate",
    "Uncertainty": "uncertainty",
    "Quality": "quality_flag",
    "Raw file ID": "raw_file_id",
    "Measured at": "measured_at",
    "Analyst": "analyst",
    "Interpretation / notes": "interpretation",
}

PROCESS_METRICS = (
    "Temperature (°C)",
    "RPM",
    "pH",
    "Solids (%)",
    "Particle size (nm)",
    "Conversion (%)",
    "Viscosity (cP)",
)

REACTION_MASTER_DIRECT_COLUMNS = {
    "Run ID": "A",
    "Date": "B",
    "Project": "C",
    "Process": "D",
    "Objective": "E",
    "Status": "I",
    "Operator": "H",
    "Protocol": "Q",
    "Equipment": "R",
    "Run summary": "K",
    "Reviewer": "U",
    "Reviewed at": "V",
    "Target batch mass (g)": "W",
    "Numerator group": "Y",
    "Denominator group": "Z",
    "Target equivalent ratio": "AA",
    "Template ID": "AC",
    "Template version": "AD",
}


def summarize_record_governance(
    experiment: dict[str, Any],
    templates: list[dict[str, Any]],
    signatures: list[dict[str, Any]],
    inventory_transactions: list[dict[str, Any]],
) -> dict[str, Any]:
    """Return the spreadsheet-equivalent template, material, and signoff state."""

    experiment_id = str(experiment.get("experiment_id", "")).strip()
    template_id = str(experiment.get("template_id", "")).strip()
    template_version = str(experiment.get("template_version", "")).strip()
    effective_template = any(
        str(row.get("template_id", "")).strip() == template_id
        and str(row.get("version", "")).strip() == template_version
        and str(row.get("state", "")).strip() == "effective"
        for row in templates
    )
    signed = [
        row
        for row in signatures
        if str(row.get("experiment_id", "")).strip() == experiment_id
        and str(row.get("status", "")).strip() == "signed"
    ]
    witnessed = [row for row in signed if str(row.get("witnessed_by", "")).strip()]
    if not signed:
        signature_status = "UNSIGNED"
    elif not witnessed:
        signature_status = "SIGNED_UNWITNESSED"
    else:
        signature_status = "SIGNED_WITNESSED"
    if not template_id:
        governance_status = "TEMPLATE_MISSING"
    elif not effective_template:
        governance_status = "TEMPLATE_NOT_EFFECTIVE"
    elif str(experiment.get("status", "")).strip() != "complete":
        governance_status = "IN_PROGRESS"
    elif signature_status == "UNSIGNED":
        governance_status = "SIGNOFF_REQUIRED"
    elif signature_status == "SIGNED_UNWITNESSED":
        governance_status = "WITNESS_REQUIRED"
    else:
        governance_status = "READY_TO_ARCHIVE"
    material_transactions = sum(
        str(row.get("experiment_id", "")).strip() == experiment_id
        for row in inventory_transactions
    )
    latest_signed = signed[-1] if signed else {}
    latest_witnessed = witnessed[-1] if witnessed else {}
    return {
        "experiment_id": experiment_id,
        "template_id": template_id,
        "template_version": template_version,
        "material_transactions": material_transactions,
        "signature_status": signature_status,
        "signed_by": latest_signed.get("signer", ""),
        "signed_at": latest_signed.get("signed_at", ""),
        "witnessed_by": latest_witnessed.get("witnessed_by", ""),
        "witnessed_at": latest_witnessed.get("witnessed_at", ""),
        "governance_status": governance_status,
    }


def reaction_master_excel_formula(header: str, row_number: int) -> str:
    """Return one auto-updating Reaction Master formula for an Excel row."""

    if header in REACTION_MASTER_DIRECT_COLUMNS:
        source = REACTION_MASTER_DIRECT_COLUMNS[header]
        return f'=IF(Experiments!$A{row_number}="","",Experiments!${source}{row_number})'
    run_cell = f"$A{row_number}"
    formulas = {
        "Planned mass (g)": f'=IF({run_cell}="","",SUMIF(\'Batch Builder\'!$A$2:$A$1000,{run_cell},\'Batch Builder\'!$L$2:$L$1000))',
        "Actual mass (g)": f'=IF({run_cell}="","",SUMIF(\'Batch Builder\'!$A$2:$A$1000,{run_cell},\'Batch Builder\'!$BB$2:$BB$1000))',
        "Mass variance (g)": f'=IF({run_cell}="","",SUMIF(\'Batch Builder\'!$A$2:$A$1000,{run_cell},\'Batch Builder\'!$S$2:$S$1000))',
        "Charges recorded": f'=IF({run_cell}="","",COUNTIFS(\'Batch Builder\'!$A$2:$A$1000,{run_cell},\'Batch Builder\'!$AB$2:$AB$1000,"charged"))',
        "Bench entries": f'=IF({run_cell}="","",COUNTIF(\'Bench Log\'!$A$2:$A$1000,{run_cell}))',
        "Measurements": f'=IF({run_cell}="","",COUNTIF(Measurements!$A$2:$A$1000,{run_cell}))',
        "Last activity": (
            f'=IF({run_cell}="","",IF(COUNTIF(\'Bench Log\'!$A$2:$A$1000,{run_cell})+'
            f'COUNTIF(Measurements!$A$2:$A$1000,{run_cell})=0,"",MAX('
            f'IFERROR(MAXIFS(\'Bench Log\'!$B$2:$B$1000,\'Bench Log\'!$A$2:$A$1000,{run_cell}),0),'
            f'IFERROR(MAXIFS(Measurements!$L$2:$L$1000,Measurements!$A$2:$A$1000,{run_cell}),0))))'
        ),
        "Next action": (
            f'=IF({run_cell}="","",IF($F{row_number}="complete",IF($O{row_number}=0,"Add measurements",'
            f'IF($Q{row_number}="","Add run summary","Record complete")),IF($J{row_number}=0,'
            f'"Enter batch quantities",IF($N{row_number}=0,"Begin bench log","Continue run and capture data"))))'
        ),
        "Planned vs target (g)": f'=IF(OR({run_cell}="",$U{row_number}=""),"",$J{row_number}-$U{row_number})',
        "Actual vs target (g)": f'=IF(OR({run_cell}="",$U{row_number}=""),"",$K{row_number}-$U{row_number})',
        "Mass completion (%)": f'=IF(OR({run_cell}="",$J{row_number}=0),"",$K{row_number}/$J{row_number})',
        "Out-of-tolerance charges": f'=IF({run_cell}="","",COUNTIFS(\'Batch Builder\'!$A$2:$A$1000,{run_cell},\'Batch Builder\'!$AQ$2:$AQ$1000,"FAIL"))',
        "Calculation issues": f'=IF({run_cell}="","",COUNTIFS(\'Batch Builder\'!$A$2:$A$1000,{run_cell},\'Batch Builder\'!$AR$2:$AR$1000,"<>READY"))',
        "Mass plan status": (
            f'=IF({run_cell}="","",IF(COUNTIF(\'Batch Builder\'!$A$2:$A$1000,{run_cell})=0,"NO_CHARGES",'
            f'IF($Z{row_number}>0,"CALCULATION_ISSUES",IF($Y{row_number}>0,"CHARGE_TOLERANCE_FAILURE",IF($U{row_number}="","NO_TARGET_MASS",'
            f'IF(IFERROR(VLOOKUP({run_cell},Experiments!$A$2:$AB$1000,24,FALSE),"")="","MISSING_TOLERANCE",'
            f'IF(ABS($V{row_number})<=$U{row_number}*VLOOKUP({run_cell},Experiments!$A$2:$AB$1000,24,FALSE)/100,"READY","MASS_MISMATCH")))))))'
        ),
        "Planned numerator eq": f'=IF({run_cell}="","",SUMIFS(\'Batch Builder\'!$AV$2:$AV$1000,\'Batch Builder\'!$A$2:$A$1000,{run_cell},\'Batch Builder\'!$AS$2:$AS$1000,$AB{row_number}))',
        "Planned denominator eq": f'=IF({run_cell}="","",SUMIFS(\'Batch Builder\'!$AV$2:$AV$1000,\'Batch Builder\'!$A$2:$A$1000,{run_cell},\'Batch Builder\'!$AS$2:$AS$1000,$AC{row_number}))',
        "Planned equivalent ratio": f'=IF(OR({run_cell}="",$AF{row_number}=0),"",$AE{row_number}/$AF{row_number})',
        "Actual numerator eq": f'=IF({run_cell}="","",SUMIFS(\'Batch Builder\'!$AW$2:$AW$1000,\'Batch Builder\'!$A$2:$A$1000,{run_cell},\'Batch Builder\'!$AS$2:$AS$1000,$AB{row_number}))',
        "Actual denominator eq": f'=IF({run_cell}="","",SUMIFS(\'Batch Builder\'!$AW$2:$AW$1000,\'Batch Builder\'!$A$2:$A$1000,{run_cell},\'Batch Builder\'!$AS$2:$AS$1000,$AC{row_number}))',
        "Actual equivalent ratio": f'=IF(OR({run_cell}="",$AI{row_number}=0),"",$AH{row_number}/$AI{row_number})',
        "Stoichiometry status": f'=IF({run_cell}="","",IF(AND($AB{row_number}="",$AC{row_number}="",$AD{row_number}=""),"NOT_CONFIGURED",IF(OR($AB{row_number}="",$AC{row_number}="",$AD{row_number}=""),"MISSING_TARGET_RATIO",IF(IFERROR(VLOOKUP({run_cell},Experiments!$A$2:$AB$1000,28,FALSE),"")="","MISSING_RATIO_TOLERANCE",IF($AE{row_number}=0,"MISSING_NUMERATOR_EQ",IF($AF{row_number}=0,"MISSING_DENOMINATOR_EQ",IF(ABS($AG{row_number}-$AD{row_number})<=$AD{row_number}*VLOOKUP({run_cell},Experiments!$A$2:$AB$1000,28,FALSE)/100,"READY","RATIO_MISMATCH")))))))',
        "Material transactions": f'=IF({run_cell}="","",COUNTIF(\'Inventory Transactions\'!$D$2:$D$1000,{run_cell}))',
        "Signature status": f'=IF({run_cell}="","",IF(COUNTIFS(\'Record Signatures\'!$B$2:$B$1000,{run_cell},\'Record Signatures\'!$H$2:$H$1000,"signed")=0,"UNSIGNED",IF(COUNTIFS(\'Record Signatures\'!$B$2:$B$1000,{run_cell},\'Record Signatures\'!$H$2:$H$1000,"signed",\'Record Signatures\'!$I$2:$I$1000,"<>")=0,"SIGNED_UNWITNESSED","SIGNED_WITNESSED")))',
        "Signed by": f'=IF({run_cell}="","",IFERROR(LOOKUP(2,1/((\'Record Signatures\'!$B$2:$B$1000={run_cell})*(\'Record Signatures\'!$H$2:$H$1000="signed")),\'Record Signatures\'!$D$2:$D$1000),""))',
        "Signed at": f'=IF({run_cell}="","",IFERROR(MAXIFS(\'Record Signatures\'!$E$2:$E$1000,\'Record Signatures\'!$B$2:$B$1000,{run_cell},\'Record Signatures\'!$H$2:$H$1000,"signed"),""))',
        "Witnessed by": f'=IF({run_cell}="","",IFERROR(LOOKUP(2,1/((\'Record Signatures\'!$B$2:$B$1000={run_cell})*(\'Record Signatures\'!$H$2:$H$1000="signed")*(\'Record Signatures\'!$I$2:$I$1000<>"")),\'Record Signatures\'!$I$2:$I$1000),""))',
        "Witnessed at": f'=IF({run_cell}="","",IFERROR(MAXIFS(\'Record Signatures\'!$J$2:$J$1000,\'Record Signatures\'!$B$2:$B$1000,{run_cell},\'Record Signatures\'!$H$2:$H$1000,"signed"),""))',
        "Governance status": f'=IF({run_cell}="","",IF($AL{row_number}="","TEMPLATE_MISSING",IF(COUNTIFS(\'Experiment Templates\'!$A$2:$A$1000,$AL{row_number},\'Experiment Templates\'!$D$2:$D$1000,$AM{row_number},\'Experiment Templates\'!$E$2:$E$1000,"effective")=0,"TEMPLATE_NOT_EFFECTIVE",IF($F{row_number}<>"complete","IN_PROGRESS",IF($AO{row_number}="UNSIGNED","SIGNOFF_REQUIRED",IF($AO{row_number}="SIGNED_UNWITNESSED","WITNESS_REQUIRED","READY_TO_ARCHIVE"))))))',
    }
    try:
        return formulas[header]
    except KeyError as exc:
        raise KeyError(f"Unknown Reaction Master column {header!r}") from exc


def google_reaction_master_array_formulas(end_row: int = 1000) -> dict[str, str]:
    """Return spill formulas that keep the Google Sheets Reaction Master current."""

    formulas = {
        header: f'=ARRAYFORMULA(IF(Experiments!A2:A{end_row}="","",Experiments!{column}2:{column}{end_row}))'
        for header, column in REACTION_MASTER_DIRECT_COLUMNS.items()
    }
    formulas.update(
        {
            "Planned mass (g)": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",SUMIF(\'Batch Builder\'!A2:A{end_row},id,\'Batch Builder\'!L2:L{end_row}))))',
            "Actual mass (g)": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",SUMIF(\'Batch Builder\'!A2:A{end_row},id,\'Batch Builder\'!BB2:BB{end_row}))))',
            "Mass variance (g)": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",SUMIF(\'Batch Builder\'!A2:A{end_row},id,\'Batch Builder\'!S2:S{end_row}))))',
            "Charges recorded": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",COUNTIFS(\'Batch Builder\'!A2:A{end_row},id,\'Batch Builder\'!AB2:AB{end_row},"charged"))))',
            "Bench entries": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",COUNTIF(\'Bench Log\'!A2:A{end_row},id))))',
            "Measurements": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",COUNTIF(Measurements!A2:A{end_row},id))))',
            "Last activity": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",IF(COUNTIF(\'Bench Log\'!A2:A{end_row},id)+COUNTIF(Measurements!A2:A{end_row},id)=0,"",MAX(IFERROR(MAXIFS(\'Bench Log\'!B2:B{end_row},\'Bench Log\'!A2:A{end_row},id),0),IFERROR(MAXIFS(Measurements!L2:L{end_row},Measurements!A2:A{end_row},id),0))))))',
            "Next action": f'=MAP(A2:A{end_row},F2:F{end_row},J2:J{end_row},N2:N{end_row},O2:O{end_row},Q2:Q{end_row},LAMBDA(id,status,planned,logs,measurements,summary,IF(id="","",IF(status="complete",IF(measurements=0,"Add measurements",IF(summary="","Add run summary","Record complete")),IF(planned=0,"Enter batch quantities",IF(logs=0,"Begin bench log","Continue run and capture data"))))))',
            "Planned vs target (g)": f'=ARRAYFORMULA(IF((A2:A{end_row}="")+(U2:U{end_row}=""),"",J2:J{end_row}-U2:U{end_row}))',
            "Actual vs target (g)": f'=ARRAYFORMULA(IF((A2:A{end_row}="")+(U2:U{end_row}=""),"",K2:K{end_row}-U2:U{end_row}))',
            "Mass completion (%)": f'=ARRAYFORMULA(IF((A2:A{end_row}="")+(J2:J{end_row}=0),"",K2:K{end_row}/J2:J{end_row}))',
            "Out-of-tolerance charges": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",COUNTIFS(\'Batch Builder\'!A2:A{end_row},id,\'Batch Builder\'!AQ2:AQ{end_row},"FAIL"))))',
            "Calculation issues": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",COUNTIFS(\'Batch Builder\'!A2:A{end_row},id,\'Batch Builder\'!AR2:AR{end_row},"<>READY"))))',
            "Mass plan status": f'=MAP(A2:A{end_row},U2:U{end_row},V2:V{end_row},Y2:Y{end_row},Z2:Z{end_row},LAMBDA(id,target,delta,oot,issues,IF(id="","",IF(COUNTIF(\'Batch Builder\'!A2:A{end_row},id)=0,"NO_CHARGES",IF(issues>0,"CALCULATION_ISSUES",IF(oot>0,"CHARGE_TOLERANCE_FAILURE",IF(target="","NO_TARGET_MASS",IF(IFNA(VLOOKUP(id,Experiments!A2:AB{end_row},24,FALSE),"")="","MISSING_TOLERANCE",IF(ABS(delta)<=target*VLOOKUP(id,Experiments!A2:AB{end_row},24,FALSE)/100,"READY","MASS_MISMATCH")))))))))',
            "Planned numerator eq": f'=MAP(A2:A{end_row},AB2:AB{end_row},LAMBDA(id,grp,IF(id="","",SUMIFS(\'Batch Builder\'!AV2:AV{end_row},\'Batch Builder\'!A2:A{end_row},id,\'Batch Builder\'!AS2:AS{end_row},grp))))',
            "Planned denominator eq": f'=MAP(A2:A{end_row},AC2:AC{end_row},LAMBDA(id,grp,IF(id="","",SUMIFS(\'Batch Builder\'!AV2:AV{end_row},\'Batch Builder\'!A2:A{end_row},id,\'Batch Builder\'!AS2:AS{end_row},grp))))',
            "Planned equivalent ratio": f'=ARRAYFORMULA(IF((A2:A{end_row}="")+(AF2:AF{end_row}=0),"",AE2:AE{end_row}/AF2:AF{end_row}))',
            "Actual numerator eq": f'=MAP(A2:A{end_row},AB2:AB{end_row},LAMBDA(id,grp,IF(id="","",SUMIFS(\'Batch Builder\'!AW2:AW{end_row},\'Batch Builder\'!A2:A{end_row},id,\'Batch Builder\'!AS2:AS{end_row},grp))))',
            "Actual denominator eq": f'=MAP(A2:A{end_row},AC2:AC{end_row},LAMBDA(id,grp,IF(id="","",SUMIFS(\'Batch Builder\'!AW2:AW{end_row},\'Batch Builder\'!A2:A{end_row},id,\'Batch Builder\'!AS2:AS{end_row},grp))))',
            "Actual equivalent ratio": f'=ARRAYFORMULA(IF((A2:A{end_row}="")+(AI2:AI{end_row}=0),"",AH2:AH{end_row}/AI2:AI{end_row}))',
            "Stoichiometry status": f'=MAP(A2:A{end_row},AB2:AB{end_row},AC2:AC{end_row},AD2:AD{end_row},AE2:AE{end_row},AF2:AF{end_row},AG2:AG{end_row},LAMBDA(id,numgrp,dengrp,target,numeq,deneq,ratio,IF(id="","",IF(AND(numgrp="",dengrp="",target=""),"NOT_CONFIGURED",IF(OR(numgrp="",dengrp="",target=""),"MISSING_TARGET_RATIO",IF(IFNA(VLOOKUP(id,Experiments!A2:AB{end_row},28,FALSE),"")="","MISSING_RATIO_TOLERANCE",IF(numeq=0,"MISSING_NUMERATOR_EQ",IF(deneq=0,"MISSING_DENOMINATOR_EQ",IF(ABS(ratio-target)<=target*VLOOKUP(id,Experiments!A2:AB{end_row},28,FALSE)/100,"READY","RATIO_MISMATCH")))))))))',
            "Material transactions": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",COUNTIF(\'Inventory Transactions\'!D2:D{end_row},id))))',
            "Signature status": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",IF(COUNTIFS(\'Record Signatures\'!B2:B{end_row},id,\'Record Signatures\'!H2:H{end_row},"signed")=0,"UNSIGNED",IF(COUNTIFS(\'Record Signatures\'!B2:B{end_row},id,\'Record Signatures\'!H2:H{end_row},"signed",\'Record Signatures\'!I2:I{end_row},"<>")=0,"SIGNED_UNWITNESSED","SIGNED_WITNESSED")))))',
            "Signed by": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",IFNA(LOOKUP(2,1/((\'Record Signatures\'!B2:B{end_row}=id)*(\'Record Signatures\'!H2:H{end_row}="signed")),\'Record Signatures\'!D2:D{end_row}),""))))',
            "Signed at": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",IFNA(MAXIFS(\'Record Signatures\'!E2:E{end_row},\'Record Signatures\'!B2:B{end_row},id,\'Record Signatures\'!H2:H{end_row},"signed"),""))))',
            "Witnessed by": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",IFNA(LOOKUP(2,1/((\'Record Signatures\'!B2:B{end_row}=id)*(\'Record Signatures\'!H2:H{end_row}="signed")*(\'Record Signatures\'!I2:I{end_row}<>"")),\'Record Signatures\'!I2:I{end_row}),""))))',
            "Witnessed at": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",IFNA(MAXIFS(\'Record Signatures\'!J2:J{end_row},\'Record Signatures\'!B2:B{end_row},id,\'Record Signatures\'!H2:H{end_row},"signed"),""))))',
            "Governance status": f'=MAP(A2:A{end_row},F2:F{end_row},AL2:AL{end_row},AM2:AM{end_row},AO2:AO{end_row},LAMBDA(id,runstatus,template,version,signature,IF(id="","",IF(template="","TEMPLATE_MISSING",IF(COUNTIFS(\'Experiment Templates\'!A2:A{end_row},template,\'Experiment Templates\'!D2:D{end_row},version,\'Experiment Templates\'!E2:E{end_row},"effective")=0,"TEMPLATE_NOT_EFFECTIVE",IF(runstatus<>"complete","IN_PROGRESS",IF(signature="UNSIGNED","SIGNOFF_REQUIRED",IF(signature="SIGNED_UNWITNESSED","WITNESS_REQUIRED","READY_TO_ARCHIVE"))))))))',
        }
    )
    return formulas


def plot_studio_process_formula(end_row: int = 1000) -> str:
    metric_array = "{" + ",".join(f'\"{value}\"' for value in PROCESS_METRICS) + "}"
    value_ranges = ",".join(
        f"'Bench Log'!${column}$2:${column}${end_row}"
        for column in ("D", "E", "F", "G", "H", "I", "J")
    )
    chosen = f"CHOOSE(MATCH($B$4,{metric_array},0),{value_ranges})"
    return (
        '=IFERROR(FILTER({\'Bench Log\'!$B$2:$B$'
        f'{end_row},{chosen}'
        "},'Bench Log'!$A$2:$A$"
        f'{end_row}=$B$3,{chosen}<>""),{{"No data",""}})'
    )


def plot_studio_measurement_formula(end_row: int = 1000) -> str:
    return (
        '=IFERROR(FILTER({\'Measurements\'!$B$2:$B$'
        f'{end_row},\'Measurements\'!$E$2:$E${end_row}'
        "},'Measurements'!$A$2:$A$"
        f'{end_row}=$B$3,\'Measurements\'!$C$2:$C${end_row}=$B$5,'
        f'\'Measurements\'!$E$2:$E${end_row}<>""),{{"No data",""}})'
    )


def plot_studio_measurement_choices_formula(end_row: int = 1000) -> str:
    return (
        '=IFERROR(SORT(UNIQUE(FILTER(\'Measurements\'!$C$2:$C$'
        f'{end_row},\'Measurements\'!$A$2:$A${end_row}=$B$3,'
        f'\'Measurements\'!$C$2:$C${end_row}<>""))),"")'
    )


def bench_log_to_daily_log_row(row: dict[str, Any]) -> dict[str, Any]:
    """Project a compact scientist-facing bench row into the canonical log shape."""

    return {
        canonical: row.get(label, "")
        for label, canonical in BENCH_LOG_TO_DAILY_LOG.items()
    }


def measurement_to_result_row(row: dict[str, Any]) -> dict[str, Any]:
    """Project a compact scientist-facing measurement into the canonical result shape."""

    projected = {
        canonical: row.get(label, "")
        for label, canonical in MEASUREMENTS_TO_RESULTS.items()
    }
    numeric = projected.get("numeric_value", "")
    projected["value"] = numeric
    return projected


def daily_log_to_bench_row(row: dict[str, Any]) -> dict[str, Any]:
    return {
        label: row.get(canonical, "")
        for label, canonical in BENCH_LOG_TO_DAILY_LOG.items()
    }


def result_to_measurement_row(row: dict[str, Any]) -> dict[str, Any]:
    projected = {
        label: row.get(canonical, "")
        for label, canonical in MEASUREMENTS_TO_RESULTS.items()
    }
    if projected.get("Numeric value", "") in (None, ""):
        projected["Numeric value"] = row.get("value", "")
    return projected


def daily_log_rows_from_tables(
    tables: dict[str, list[dict[str, Any]]],
) -> list[dict[str, Any]]:
    """Return legacy observations plus compact Bench Log entries without duplicates."""

    rows_by_key: dict[tuple[str, str, str], dict[str, Any]] = {}
    for row in tables.get("Daily Log", []):
        if isinstance(row, dict):
            rows_by_key[_daily_key(row)] = row
    for row in tables.get("Bench Log", []):
        if not isinstance(row, dict):
            continue
        projected = bench_log_to_daily_log_row(row)
        rows_by_key[_daily_key(projected)] = projected
    return [row for key, row in rows_by_key.items() if any(key)]


def result_rows_from_tables(
    tables: dict[str, list[dict[str, Any]]],
) -> list[dict[str, Any]]:
    """Return legacy results plus compact Measurements entries without duplicates."""

    rows_by_key: dict[tuple[str, str, str, str], dict[str, Any]] = {}
    for row in tables.get("Results", []):
        if isinstance(row, dict):
            rows_by_key[_result_key(row)] = row
    for row in tables.get("Measurements", []):
        if not isinstance(row, dict):
            continue
        projected = measurement_to_result_row(row)
        rows_by_key[_result_key(projected)] = projected
    return [row for key, row in rows_by_key.items() if any(key)]


def _daily_key(row: dict[str, Any]) -> tuple[str, str, str]:
    return (
        str(row.get("experiment_id", "")).strip(),
        str(row.get("timestamp", "")).strip(),
        str(row.get("observation", "")).strip(),
    )


def _result_key(row: dict[str, Any]) -> tuple[str, str, str, str]:
    return (
        str(row.get("experiment_id", "")).strip(),
        str(row.get("sample_id", "")).strip(),
        str(row.get("measurement_type", "")).strip(),
        str(row.get("replicate", "")).strip(),
    )
