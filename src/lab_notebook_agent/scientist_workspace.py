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
}


def reaction_master_excel_formula(header: str, row_number: int) -> str:
    """Return one auto-updating Reaction Master formula for an Excel row."""

    if header in REACTION_MASTER_DIRECT_COLUMNS:
        source = REACTION_MASTER_DIRECT_COLUMNS[header]
        return f'=IF(Experiments!$A{row_number}="","",Experiments!${source}{row_number})'
    run_cell = f"$A{row_number}"
    formulas = {
        "Planned mass (g)": f'=IF({run_cell}="","",SUMIF(\'Batch Builder\'!$A$2:$A$1000,{run_cell},\'Batch Builder\'!$L$2:$L$1000))',
        "Actual mass (g)": f'=IF({run_cell}="","",SUMIF(\'Batch Builder\'!$A$2:$A$1000,{run_cell},\'Batch Builder\'!$R$2:$R$1000))',
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
            "Actual mass (g)": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",SUMIF(\'Batch Builder\'!A2:A{end_row},id,\'Batch Builder\'!R2:R{end_row}))))',
            "Mass variance (g)": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",SUMIF(\'Batch Builder\'!A2:A{end_row},id,\'Batch Builder\'!S2:S{end_row}))))',
            "Charges recorded": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",COUNTIFS(\'Batch Builder\'!A2:A{end_row},id,\'Batch Builder\'!AB2:AB{end_row},"charged"))))',
            "Bench entries": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",COUNTIF(\'Bench Log\'!A2:A{end_row},id))))',
            "Measurements": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",COUNTIF(Measurements!A2:A{end_row},id))))',
            "Last activity": f'=MAP(A2:A{end_row},LAMBDA(id,IF(id="","",IF(COUNTIF(\'Bench Log\'!A2:A{end_row},id)+COUNTIF(Measurements!A2:A{end_row},id)=0,"",MAX(IFERROR(MAXIFS(\'Bench Log\'!B2:B{end_row},\'Bench Log\'!A2:A{end_row},id),0),IFERROR(MAXIFS(Measurements!L2:L{end_row},Measurements!A2:A{end_row},id),0))))))',
            "Next action": f'=MAP(A2:A{end_row},F2:F{end_row},J2:J{end_row},N2:N{end_row},O2:O{end_row},Q2:Q{end_row},LAMBDA(id,status,planned,logs,measurements,summary,IF(id="","",IF(status="complete",IF(measurements=0,"Add measurements",IF(summary="","Add run summary","Record complete")),IF(planned=0,"Enter batch quantities",IF(logs=0,"Begin bench log","Continue run and capture data"))))))',
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
