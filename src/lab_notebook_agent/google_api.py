from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from time import sleep
from typing import Any, Callable, Protocol
from urllib.parse import quote

from .agent import AgentRunConfig, build_agent_report
from .daily_agent import build_snapshot_daily_agent_run
from .daily_log_results import build_daily_log_results_report
from .experiment_record import build_experiment_record_report
from .formulation_normalization import build_formulation_normalization_report
from .google_sheets import (
    audit_report_against_snapshot,
    batch_update_requests_from_report,
    generated_sheet_ids_for_missing,
    google_contract_migration_requests,
    google_setup_audit_from_metadata,
    google_setup_requests_from_metadata,
    quality_conditional_format_requests,
    sheet_ids_from_metadata_payload,
    sheet_ids_from_snapshot,
    snapshot_from_tables,
    snapshot_to_tables,
    validate_snapshot,
)
from .material_scaffold import build_material_scaffold_report
from .planning import build_plan_materialization_report
from .project_notebook import (
    build_project_notebook_sync_report,
    parse_project_notebook_sheet,
)
from .plotting import (
    build_plot_report,
    google_chart_requests,
    tables_after_report,
)
from .recorded_daily_agent import build_snapshot_recorded_daily_agent_run
from .schema import SHEETS


DEFAULT_SCOPES = (
    "https://www.googleapis.com/auth/spreadsheets",
)


class SheetsApiClient(Protocol):
    def get_metadata(self, spreadsheet_id: str) -> dict[str, Any]:
        ...

    def get_values(
        self,
        spreadsheet_id: str,
        sheet_name: str,
        value_range: str,
        value_render_option: str = "FORMATTED_VALUE",
    ) -> list[list[Any]]:
        ...

    def batch_update(self, spreadsheet_id: str, requests: list[dict[str, Any]]) -> dict[str, Any]:
        ...


@dataclass(frozen=True)
class GoogleCredentialsConfig:
    service_account_file: str | None = None
    scopes: tuple[str, ...] = DEFAULT_SCOPES


class GoogleSheetsApiClient:
    def __init__(self, session: Any, base_url: str = "https://sheets.googleapis.com") -> None:
        self.session = session
        self.base_url = base_url.rstrip("/")

    @classmethod
    def from_credentials(cls, config: GoogleCredentialsConfig | None = None) -> "GoogleSheetsApiClient":
        config = config or GoogleCredentialsConfig()
        try:
            import google.auth
            from google.auth.exceptions import DefaultCredentialsError
            from google.auth.transport.requests import AuthorizedSession
            from google.oauth2 import service_account
        except ImportError as exc:
            raise RuntimeError(
                "Google API support requires optional dependencies. Install with "
                "`pip install -e .[google]` or install google-auth and requests."
            ) from exc

        try:
            if config.service_account_file:
                credentials = service_account.Credentials.from_service_account_file(
                    str(Path(config.service_account_file).expanduser()),
                    scopes=list(config.scopes),
                )
            else:
                credentials, _ = google.auth.default(scopes=list(config.scopes))
        except (DefaultCredentialsError, FileNotFoundError, ValueError) as exc:
            raise RuntimeError(f"Google credentials are not ready: {exc}") from exc
        return cls(AuthorizedSession(credentials))

    def get_metadata(self, spreadsheet_id: str) -> dict[str, Any]:
        return self.request_json(
            "GET",
            f"/v4/spreadsheets/{spreadsheet_id}",
            params={
                "fields": (
                    "spreadsheetId,properties(title,timeZone,locale),"
                    "sheets(properties(sheetId,title,gridProperties),"
                    "charts(chartId,spec(title),position),conditionalFormats)"
                )
            },
        )

    def get_values(
        self,
        spreadsheet_id: str,
        sheet_name: str,
        value_range: str,
        value_render_option: str = "FORMATTED_VALUE",
    ) -> list[list[Any]]:
        a1_range = f"{quote_sheet_name(sheet_name)}!{value_range}"
        payload = self.request_json(
            "GET",
            f"/v4/spreadsheets/{spreadsheet_id}/values/{quote(a1_range, safe='')}",
            params={"valueRenderOption": value_render_option},
        )
        values = payload.get("values", [])
        return values if isinstance(values, list) else []

    def batch_update(self, spreadsheet_id: str, requests: list[dict[str, Any]]) -> dict[str, Any]:
        return self.request_json(
            "POST",
            f"/v4/spreadsheets/{spreadsheet_id}:batchUpdate",
            body={"requests": requests},
        )

    def request_json(
        self,
        method: str,
        path: str,
        params: dict[str, Any] | None = None,
        body: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        response = self.session.request(
            method,
            f"{self.base_url}{path}",
            params=params,
            json=body,
        )
        response.raise_for_status()
        payload = response.json()
        if not isinstance(payload, dict):
            raise ValueError("Google Sheets API response was not a JSON object.")
        return payload


def capture_snapshot_from_google_sheets(
    spreadsheet_id: str,
    client: SheetsApiClient,
    value_range: str = "A1:Z1000",
    value_render_option: str = "FORMATTED_VALUE",
) -> dict[str, Any]:
    metadata = client.get_metadata(spreadsheet_id)
    sheet_ids = sheet_ids_from_metadata(metadata)
    existing_sheet_names = {
        str(properties.get("title", ""))
        for sheet in metadata.get("sheets", []) or []
        if isinstance(sheet, dict)
        and isinstance((properties := sheet.get("properties", {})), dict)
        and properties.get("title") is not None
    }
    sheets: dict[str, Any] = {}
    for spec in SHEETS:
        if spec.name in existing_sheet_names:
            values = client.get_values(
                spreadsheet_id,
                spec.name,
                value_range,
                value_render_option=value_render_option,
            )
        else:
            values = []
        sheets[spec.name] = {
            "sheet_id": sheet_ids.get(spec.name),
            "values": values,
        }
    return {
        "schema": "lab-notebook-agent-google-sheets-snapshot.v1",
        "spreadsheet_id": spreadsheet_id,
        "sheets": sheets,
    }


def run_live_google_project_notebook_sync(
    source_spreadsheet_id: str,
    target_spreadsheet_id: str,
    source_sheet_names: tuple[str, ...],
    client: SheetsApiClient,
    *,
    source_range: str = "A1:AO1200",
    parser_profile: str = "auto",
    apply: bool = False,
) -> dict[str, Any]:
    if source_spreadsheet_id == target_spreadsheet_id:
        raise ValueError(
            "Source and target spreadsheet IDs must differ; the source is always read-only."
        )
    if not source_sheet_names:
        raise ValueError(
            "At least one source sheet is required. Select tabs explicitly so a sync "
            "cannot accidentally scan an entire research workbook."
        )
    if len(set(source_sheet_names)) != len(source_sheet_names):
        raise ValueError("Source sheet names must be unique.")

    source_metadata = client.get_metadata(source_spreadsheet_id)
    source_sheet_ids = sheet_ids_from_metadata(source_metadata)
    missing_source_sheets = [
        sheet_name
        for sheet_name in source_sheet_names
        if sheet_name not in source_sheet_ids
    ]
    if missing_source_sheets:
        raise ValueError(
            "Source tabs were not found: " + ", ".join(missing_source_sheets)
        )
    source_title = str(
        (source_metadata.get("properties") or {}).get("title", "")
    )
    source_url = (
        f"https://docs.google.com/spreadsheets/d/{source_spreadsheet_id}"
    )
    parsed_sources = []
    source_summaries = []
    for sheet_name in source_sheet_names:
        values = client.get_values(
            source_spreadsheet_id,
            sheet_name,
            source_range,
            value_render_option="FORMATTED_VALUE",
        )
        parsed = parse_project_notebook_sheet(
            values,
            source_spreadsheet_id=source_spreadsheet_id,
            source_title=source_title,
            source_sheet_name=sheet_name,
            source_sheet_id=source_sheet_ids[sheet_name],
            source_url=source_url,
            source_range=source_range,
            parser_profile=parser_profile,
        )
        parsed_sources.append(parsed)
        source_summaries.append(
            {
                "source_key": parsed["source_key"],
                "source_sheet_name": sheet_name,
                "source_sheet_id": source_sheet_ids[sheet_name],
                "source_range": source_range,
                "parser_profile": parsed["parser_profile"],
                "source_fingerprint": parsed["source_fingerprint"],
                "experiment_id": parsed["experiment"].get("experiment_id", ""),
                "experiment_ids": [
                    experiment.get("experiment_id", "")
                    for experiment in parsed.get("experiments", [])
                    if isinstance(experiment, dict)
                ],
                "record_count": len(parsed["records"]),
                "warning_count": len(parsed["warnings"]),
            }
        )

    target_metadata = client.get_metadata(target_spreadsheet_id)
    target_snapshot = capture_snapshot_from_google_sheets(
        target_spreadsheet_id,
        client,
        value_range="A1:AZ5000",
    )
    target_tables = snapshot_to_tables(target_snapshot)
    report = build_project_notebook_sync_report(parsed_sources, target_tables)
    projected_tables = tables_after_report(target_tables, report)
    plot_report = build_plot_report(projected_tables)
    combined_report = dict(report)
    for key in (
        "replace_plot_data",
        "replace_plot_definitions",
        "replace_plot_dashboard",
        "existing_plot_row_counts",
    ):
        if key in plot_report:
            combined_report[key] = plot_report[key]

    existing_sheet_ids = sheet_ids_from_metadata(target_metadata)
    generated_sheet_ids = generated_sheet_ids_for_missing(existing_sheet_ids)
    effective_sheet_ids = {**existing_sheet_ids, **generated_sheet_ids}
    projected_snapshot = snapshot_from_tables(target_tables, effective_sheet_ids)
    apply_audit = audit_report_against_snapshot(
        combined_report,
        projected_snapshot,
        require_sheet_ids=True,
    )
    contract_audit = validate_snapshot(target_snapshot, require_sheet_ids=False)
    setup_required = (
        not contract_audit["valid"]
        or any(spec.name not in existing_sheet_ids for spec in SHEETS)
    )
    setup_requests = (
        google_setup_requests_from_metadata(target_metadata)
        if setup_required
        else []
    )
    data_requests = (
        batch_update_requests_from_report(combined_report, effective_sheet_ids)
        if apply_audit["valid"]
        else []
    )
    plot_data_requests = (
        batch_update_requests_from_report(plot_report, effective_sheet_ids)
        if apply_audit["valid"]
        else []
    )
    chart_requests = (
        google_chart_requests(
            plot_report,
            effective_sheet_ids,
            target_metadata,
        )
        if apply_audit["valid"]
        else []
    )
    requests = (
        setup_requests + data_requests + chart_requests
        if apply_audit["valid"]
        else []
    )
    response = (
        client.batch_update(target_spreadsheet_id, requests)
        if apply and requests
        else {}
    )
    return {
        "schema": "lab-notebook-agent-live-project-notebook-sync.v1",
        "source_spreadsheet_id": source_spreadsheet_id,
        "target_spreadsheet_id": target_spreadsheet_id,
        "source_is_read_only": True,
        "applied": bool(apply and requests),
        "source_sheets": source_summaries,
        "target_snapshot": target_snapshot,
        "sync_report": report,
        "plot_report": plot_report,
        "contract_audit": contract_audit,
        "setup_required": setup_required,
        "setup_request_count": len(setup_requests),
        "data_request_count": len(data_requests),
        "plot_data_request_count": len(plot_data_requests),
        "chart_request_count": len(chart_requests),
        "apply_audit": apply_audit,
        "batch_update_requests": requests,
        "batch_update_response": response,
    }


def run_live_google_plot_refresh(
    spreadsheet_id: str,
    client: SheetsApiClient,
    *,
    value_range: str = "A1:AZ5000",
    apply: bool = False,
) -> dict[str, Any]:
    metadata = client.get_metadata(spreadsheet_id)
    snapshot = capture_snapshot_from_google_sheets(
        spreadsheet_id,
        client,
        value_range=value_range,
    )
    tables = snapshot_to_tables(snapshot)
    plot_report = build_plot_report(tables)
    existing_sheet_ids = sheet_ids_from_metadata(metadata)
    generated_sheet_ids = generated_sheet_ids_for_missing(existing_sheet_ids)
    effective_sheet_ids = {**existing_sheet_ids, **generated_sheet_ids}
    projected_snapshot = snapshot_from_tables(tables, effective_sheet_ids)
    apply_audit = audit_report_against_snapshot(
        plot_report,
        projected_snapshot,
        require_sheet_ids=True,
    )
    contract_audit = validate_snapshot(snapshot, require_sheet_ids=False)
    setup_required = (
        not contract_audit["valid"]
        or any(spec.name not in existing_sheet_ids for spec in SHEETS)
    )
    setup_requests = (
        google_setup_requests_from_metadata(metadata) if setup_required else []
    )
    data_requests = (
        batch_update_requests_from_report(plot_report, effective_sheet_ids)
        if apply_audit["valid"]
        else []
    )
    chart_requests = (
        google_chart_requests(plot_report, effective_sheet_ids, metadata)
        if apply_audit["valid"]
        else []
    )
    requests = (
        setup_requests + data_requests + chart_requests
        if apply_audit["valid"]
        else []
    )
    response = (
        client.batch_update(spreadsheet_id, requests)
        if apply and requests
        else {}
    )
    return {
        "schema": "lab-notebook-agent-live-plot-refresh.v1",
        "spreadsheet_id": spreadsheet_id,
        "applied": bool(apply and requests),
        "snapshot": snapshot,
        "plot_report": plot_report,
        "contract_audit": contract_audit,
        "setup_required": setup_required,
        "setup_request_count": len(setup_requests),
        "data_request_count": len(data_requests),
        "plot_data_request_count": len(data_requests),
        "chart_request_count": len(chart_requests),
        "apply_audit": apply_audit,
        "batch_update_requests": requests,
        "batch_update_response": response,
    }


def run_live_google_setup(
    spreadsheet_id: str,
    client: SheetsApiClient,
    apply: bool = False,
    include_validations: bool = True,
    validation_end_row: int = 1000,
    normalize_existing_types: bool = True,
) -> dict[str, Any]:
    metadata = client.get_metadata(spreadsheet_id)
    snapshot = capture_snapshot_from_google_sheets(
        spreadsheet_id,
        client,
        value_range=f"A1:AZ{max(2, validation_end_row)}",
    )
    tables = snapshot_to_tables(snapshot)
    setup_audit = google_setup_audit_from_metadata(
        metadata,
        include_validations=include_validations,
        validation_end_row=validation_end_row,
    )
    requests = google_setup_requests_from_metadata(
        metadata,
        include_validations=include_validations,
        validation_end_row=validation_end_row,
    )
    existing_sheet_ids = sheet_ids_from_metadata(metadata)
    generated_sheet_ids = generated_sheet_ids_for_missing(existing_sheet_ids)
    effective_sheet_ids = {**existing_sheet_ids, **generated_sheet_ids}
    migration_requests = google_contract_migration_requests(
        tables,
        effective_sheet_ids,
        normalize_existing_types=normalize_existing_types,
    )
    sheets_with_conditional_formats = {
        str((sheet.get("properties") or {}).get("title", ""))
        for sheet in metadata.get("sheets", []) or []
        if isinstance(sheet, dict) and sheet.get("conditionalFormats")
    }
    conditional_format_requests = quality_conditional_format_requests(
        effective_sheet_ids,
        validation_end_row,
    )
    if "Results" in sheets_with_conditional_formats:
        results_id = effective_sheet_ids.get("Results")
        conditional_format_requests = [
            request
            for request in conditional_format_requests
            if (
                request.get("addConditionalFormatRule", {})
                .get("rule", {})
                .get("ranges", [{}])[0]
                .get("sheetId")
                != results_id
            )
        ]
    if "Deviations" in sheets_with_conditional_formats:
        deviations_id = effective_sheet_ids.get("Deviations")
        conditional_format_requests = [
            request
            for request in conditional_format_requests
            if (
                request.get("addConditionalFormatRule", {})
                .get("rule", {})
                .get("ranges", [{}])[0]
                .get("sheetId")
                != deviations_id
            )
        ]
    requests.extend(migration_requests)
    requests.extend(conditional_format_requests)
    setup_audit["summary"]["request_count"] = len(requests)
    setup_audit["summary"]["migration_request_count"] = len(migration_requests)
    setup_audit["summary"]["conditional_format_request_count"] = len(
        conditional_format_requests
    )
    response = client.batch_update(spreadsheet_id, requests) if apply and requests else {}
    return {
        "schema": "lab-notebook-agent-live-google-setup.v1",
        "spreadsheet_id": spreadsheet_id,
        "applied": bool(apply and requests),
        "metadata": metadata,
        "snapshot": snapshot,
        "setup_audit": setup_audit,
        "migration_request_count": len(migration_requests),
        "conditional_format_request_count": len(conditional_format_requests),
        "batch_update_requests": requests,
        "batch_update_response": response,
    }


def run_live_google_agent(
    spreadsheet_id: str,
    client: SheetsApiClient,
    config: AgentRunConfig | None = None,
    value_range: str = "A1:Z1000",
    apply: bool = False,
) -> dict[str, Any]:
    snapshot = capture_snapshot_from_google_sheets(spreadsheet_id, client, value_range=value_range)
    snapshot_audit = validate_snapshot(snapshot, require_sheet_ids=False)
    if snapshot_audit["valid"]:
        report = build_agent_report(snapshot_to_tables(snapshot), config=config)
    else:
        report = {
            "schema": "lab-notebook-agent-run.v1",
            "summary": {},
            "runs": [],
        }
    apply_audit = audit_report_against_snapshot(report, snapshot, require_sheet_ids=True)
    requests = batch_update_requests_from_report(report, sheet_ids_from_snapshot(snapshot)) if apply_audit["valid"] else []
    response = client.batch_update(spreadsheet_id, requests) if apply and requests else {}
    return {
        "schema": "lab-notebook-agent-live-google-run.v1",
        "spreadsheet_id": spreadsheet_id,
        "applied": bool(apply and requests),
        "snapshot": snapshot,
        "snapshot_audit": snapshot_audit,
        "agent_report": report,
        "apply_audit": apply_audit,
        "batch_update_requests": requests,
        "batch_update_response": response,
    }


def run_live_google_daily_agent(
    spreadsheet_id: str,
    client: SheetsApiClient,
    config: AgentRunConfig,
    value_range: str = "A1:Z1000",
    apply: bool = False,
) -> dict[str, Any]:
    snapshot = capture_snapshot_from_google_sheets(spreadsheet_id, client, value_range=value_range)
    daily_run = build_snapshot_daily_agent_run(snapshot, config)
    requests = daily_run.get("batch_update_requests", [])
    response = client.batch_update(spreadsheet_id, requests) if apply and requests else {}
    return {
        "schema": "lab-notebook-agent-live-google-daily-run.v1",
        "spreadsheet_id": spreadsheet_id,
        "applied": bool(apply and requests),
        "snapshot": snapshot,
        "snapshot_audit": daily_run.get("snapshot_audit", {}),
        "daily_agent_run": daily_run,
        "daily_summary": daily_run.get("daily_summary", {}),
        "agent_report": daily_run.get("agent_report", {}),
        "apply_audit": daily_run.get("apply_audit", {}),
        "batch_update_requests": requests,
        "batch_update_response": response,
    }


def run_live_google_daily_agent_watch(
    spreadsheet_id: str,
    client: SheetsApiClient,
    config: AgentRunConfig,
    value_range: str = "A1:Z1000",
    apply: bool = False,
    iterations: int = 1,
    interval_seconds: float = 60,
    sleep_fn: Callable[[float], None] = sleep,
) -> dict[str, Any]:
    if iterations < 0:
        raise ValueError("iterations must be >= 0; use 0 for continuous watch mode.")
    if interval_seconds < 0:
        raise ValueError("interval_seconds must be >= 0.")

    runs: list[dict[str, Any]] = []
    last_applied_signature = ""
    iteration = 0
    while iterations == 0 or iteration < iterations:
        iteration += 1
        run = run_live_google_daily_agent(
            spreadsheet_id,
            client,
            config=config,
            value_range=value_range,
            apply=False,
        )
        requests = run.get("batch_update_requests", [])
        signature = batch_request_signature(requests)
        should_apply = bool(apply and requests)
        if should_apply and signature == last_applied_signature:
            run["applied"] = False
            run["apply_skip_reason"] = "duplicate_batch"
            run["batch_update_response"] = {}
        elif should_apply:
            run["batch_update_response"] = client.batch_update(spreadsheet_id, requests)
            run["applied"] = True
            last_applied_signature = signature
        else:
            run["applied"] = False
            run["batch_update_response"] = {}
        runs.append(
            {
                "iteration": iteration,
                "applied": run.get("applied", False),
                "apply_skip_reason": run.get("apply_skip_reason", ""),
                "apply_audit_valid": run.get("apply_audit", {}).get("valid", False),
                "request_count": len(requests),
                "daily_summary": run.get("daily_summary", {}).get("summary", {}),
                "daily_agent_run_summary": run.get("daily_agent_run", {}).get("summary", {}),
                "run": run,
            }
        )
        if iterations != 0 and iteration >= iterations:
            break
        if interval_seconds:
            sleep_fn(interval_seconds)

    return {
        "schema": "lab-notebook-agent-live-google-daily-watch.v1",
        "spreadsheet_id": spreadsheet_id,
        "apply_requested": bool(apply),
        "iterations_requested": iterations,
        "interval_seconds": interval_seconds,
        "summary": summarize_watch_runs(runs),
        "runs": runs,
    }


def batch_request_signature(requests: list[dict[str, Any]]) -> str:
    return json.dumps(requests, sort_keys=True, separators=(",", ":"))


def summarize_watch_runs(runs: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "iteration_count": len(runs),
        "applied_iterations": sum(1 for run in runs if run.get("applied")),
        "duplicate_batches_skipped": sum(
            1 for run in runs if run.get("apply_skip_reason") == "duplicate_batch"
        ),
        "total_request_count": sum(int(run.get("request_count", 0) or 0) for run in runs),
        "last_apply_audit_valid": runs[-1].get("apply_audit_valid", False) if runs else False,
        "last_request_count": int(runs[-1].get("request_count", 0) or 0) if runs else 0,
    }


def run_live_google_experiment_record(
    spreadsheet_id: str,
    client: SheetsApiClient,
    record: dict[str, Any],
    value_range: str = "A1:Z1000",
    apply: bool = False,
) -> dict[str, Any]:
    snapshot = capture_snapshot_from_google_sheets(spreadsheet_id, client, value_range=value_range)
    snapshot_audit = validate_snapshot(snapshot, require_sheet_ids=False)
    report = build_experiment_record_report(record, tables=snapshot_to_tables(snapshot))
    apply_audit = audit_report_against_snapshot(report, snapshot, require_sheet_ids=True)
    requests = batch_update_requests_from_report(report, sheet_ids_from_snapshot(snapshot)) if apply_audit["valid"] else []
    response = client.batch_update(spreadsheet_id, requests) if apply and requests else {}
    return {
        "schema": "lab-notebook-agent-live-google-experiment-record.v1",
        "spreadsheet_id": spreadsheet_id,
        "applied": bool(apply and requests),
        "snapshot": snapshot,
        "snapshot_audit": snapshot_audit,
        "record_report": report,
        "apply_audit": apply_audit,
        "batch_update_requests": requests,
        "batch_update_response": response,
    }


def run_live_google_recorded_daily_agent(
    spreadsheet_id: str,
    client: SheetsApiClient,
    record: dict[str, Any],
    config: AgentRunConfig,
    value_range: str = "A1:Z1000",
    apply: bool = False,
) -> dict[str, Any]:
    snapshot = capture_snapshot_from_google_sheets(spreadsheet_id, client, value_range=value_range)
    recorded_daily_run = build_snapshot_recorded_daily_agent_run(snapshot, record, config)
    requests = recorded_daily_run.get("batch_update_requests", [])
    response = client.batch_update(spreadsheet_id, requests) if apply and requests else {}
    return {
        "schema": "lab-notebook-agent-live-google-recorded-daily-run.v1",
        "spreadsheet_id": spreadsheet_id,
        "applied": bool(apply and requests),
        "snapshot": snapshot,
        "snapshot_audit": recorded_daily_run.get("snapshot_audit", {}),
        "recorded_daily_run": recorded_daily_run,
        "record_report": recorded_daily_run.get("record_report", {}),
        "daily_agent_run": recorded_daily_run.get("daily_agent_run", {}),
        "apply_report": recorded_daily_run.get("apply_report", {}),
        "apply_audit": recorded_daily_run.get("apply_audit", {}),
        "batch_update_requests": requests,
        "batch_update_response": response,
    }


def run_live_google_plan_materialization(
    spreadsheet_id: str,
    client: SheetsApiClient,
    planned_date: str | None = None,
    suggestion_ids: tuple[str, ...] = (),
    value_range: str = "A1:Z1000",
    apply: bool = False,
) -> dict[str, Any]:
    snapshot = capture_snapshot_from_google_sheets(spreadsheet_id, client, value_range=value_range)
    snapshot_audit = validate_snapshot(snapshot, require_sheet_ids=False)
    if snapshot_audit["valid"]:
        report = build_plan_materialization_report(
            snapshot_to_tables(snapshot),
            planned_date=planned_date,
            suggestion_ids=suggestion_ids,
        )
    else:
        report = {
            "schema": "lab-notebook-agent-plan-materialization.v1",
            "summary": {},
            "runs": [],
        }
    apply_audit = audit_report_against_snapshot(report, snapshot, require_sheet_ids=True)
    requests = batch_update_requests_from_report(report, sheet_ids_from_snapshot(snapshot)) if apply_audit["valid"] else []
    response = client.batch_update(spreadsheet_id, requests) if apply and requests else {}
    return {
        "schema": "lab-notebook-agent-live-google-plan-materialization.v1",
        "spreadsheet_id": spreadsheet_id,
        "applied": bool(apply and requests),
        "snapshot": snapshot,
        "snapshot_audit": snapshot_audit,
        "materialization_report": report,
        "apply_audit": apply_audit,
        "batch_update_requests": requests,
        "batch_update_response": response,
    }


def run_live_google_formulation_normalization(
    spreadsheet_id: str,
    client: SheetsApiClient,
    experiment_ids: tuple[str, ...] = (),
    value_range: str = "A1:Z1000",
    apply: bool = False,
) -> dict[str, Any]:
    snapshot = capture_snapshot_from_google_sheets(spreadsheet_id, client, value_range=value_range)
    snapshot_audit = validate_snapshot(snapshot, require_sheet_ids=False)
    if snapshot_audit["valid"]:
        report = build_formulation_normalization_report(
            snapshot_to_tables(snapshot),
            experiment_ids=experiment_ids,
        )
    else:
        report = {
            "schema": "lab-notebook-agent-formulation-normalization.v1",
            "selection": {
                "requested_experiment_ids": list(experiment_ids),
            },
            "summary": {},
            "runs": [],
        }
    apply_audit = audit_report_against_snapshot(report, snapshot, require_sheet_ids=True)
    requests = batch_update_requests_from_report(report, sheet_ids_from_snapshot(snapshot)) if apply_audit["valid"] else []
    response = client.batch_update(spreadsheet_id, requests) if apply and requests else {}
    return {
        "schema": "lab-notebook-agent-live-google-formulation-normalization.v1",
        "spreadsheet_id": spreadsheet_id,
        "applied": bool(apply and requests),
        "snapshot": snapshot,
        "snapshot_audit": snapshot_audit,
        "formulation_normalization_report": report,
        "apply_audit": apply_audit,
        "batch_update_requests": requests,
        "batch_update_response": response,
    }


def run_live_google_daily_log_results_normalization(
    spreadsheet_id: str,
    client: SheetsApiClient,
    experiment_ids: tuple[str, ...] = (),
    review_date: str | None = None,
    value_range: str = "A1:Z1000",
    apply: bool = False,
) -> dict[str, Any]:
    snapshot = capture_snapshot_from_google_sheets(spreadsheet_id, client, value_range=value_range)
    snapshot_audit = validate_snapshot(snapshot, require_sheet_ids=False)
    if snapshot_audit["valid"]:
        report = build_daily_log_results_report(
            snapshot_to_tables(snapshot),
            experiment_ids=experiment_ids,
            review_date=review_date,
        )
    else:
        report = {
            "schema": "lab-notebook-agent-daily-log-results.v1",
            "selection": {
                "requested_experiment_ids": list(experiment_ids),
                "review_date": review_date or "",
            },
            "summary": {},
            "runs": [],
        }
    apply_audit = audit_report_against_snapshot(report, snapshot, require_sheet_ids=True)
    requests = batch_update_requests_from_report(report, sheet_ids_from_snapshot(snapshot)) if apply_audit["valid"] else []
    response = client.batch_update(spreadsheet_id, requests) if apply and requests else {}
    return {
        "schema": "lab-notebook-agent-live-google-daily-log-results-normalization.v1",
        "spreadsheet_id": spreadsheet_id,
        "applied": bool(apply and requests),
        "snapshot": snapshot,
        "snapshot_audit": snapshot_audit,
        "daily_log_results_report": report,
        "apply_audit": apply_audit,
        "batch_update_requests": requests,
        "batch_update_response": response,
    }


def run_live_google_material_scaffold(
    spreadsheet_id: str,
    client: SheetsApiClient,
    experiment_id: str,
    process_type: str | None = None,
    query: str = "",
    include_optional: bool = False,
    value_range: str = "A1:Z1000",
    apply: bool = False,
) -> dict[str, Any]:
    snapshot = capture_snapshot_from_google_sheets(spreadsheet_id, client, value_range=value_range)
    snapshot_audit = validate_snapshot(snapshot, require_sheet_ids=False)
    if snapshot_audit["valid"]:
        report = build_material_scaffold_report(
            snapshot_to_tables(snapshot),
            experiment_id=experiment_id,
            process_type=process_type,
            query=query,
            include_optional=include_optional,
        )
    else:
        report = {
            "schema": "lab-notebook-agent-material-scaffold.v1",
            "experiment_id": experiment_id,
            "process_type": process_type or "",
            "include_optional": include_optional,
            "query": query,
            "role_scaffold": [],
            "append_master_reagents": [],
            "append_formulations": [],
            "summary": {},
        }
    apply_audit = audit_report_against_snapshot(report, snapshot, require_sheet_ids=True)
    requests = batch_update_requests_from_report(report, sheet_ids_from_snapshot(snapshot)) if apply_audit["valid"] else []
    response = client.batch_update(spreadsheet_id, requests) if apply and requests else {}
    return {
        "schema": "lab-notebook-agent-live-google-material-scaffold.v1",
        "spreadsheet_id": spreadsheet_id,
        "applied": bool(apply and requests),
        "snapshot": snapshot,
        "snapshot_audit": snapshot_audit,
        "material_scaffold_report": report,
        "apply_audit": apply_audit,
        "batch_update_requests": requests,
        "batch_update_response": response,
    }


def google_api_doctor(
    spreadsheet_id: str | None = None,
    service_account_file: str | None = None,
) -> dict[str, Any]:
    result: dict[str, Any] = {
        "schema": "lab-notebook-agent-google-api-doctor.v1",
        "checks": [],
        "ready": False,
    }
    try:
        client = GoogleSheetsApiClient.from_credentials(
            GoogleCredentialsConfig(service_account_file=service_account_file)
        )
    except RuntimeError as exc:
        result["checks"].append(
            {
                "name": "credentials",
                "status": "failed",
                "message": str(exc),
            }
        )
        return result

    result["checks"].append(
        {
            "name": "credentials",
            "status": "passed",
            "message": "Google API credentials loaded.",
        }
    )
    if not spreadsheet_id:
        result["ready"] = True
        return result

    try:
        metadata = client.get_metadata(spreadsheet_id)
    except Exception as exc:  # pragma: no cover - provider-specific transport details.
        result["checks"].append(
            {
                "name": "spreadsheet_metadata",
                "status": "failed",
                "message": str(exc),
            }
        )
        return result

    sheet_ids = sheet_ids_from_metadata(metadata)
    missing_sheets = [spec.name for spec in SHEETS if spec.name not in sheet_ids]
    if missing_sheets:
        result["checks"].append(
            {
                "name": "spreadsheet_contract",
                "status": "failed",
                "missing_sheets": missing_sheets,
            }
        )
        return result

    result["checks"].append(
        {
            "name": "spreadsheet_contract",
            "status": "passed",
            "sheet_count": len(sheet_ids),
        }
    )
    result["ready"] = True
    return result


def sheet_ids_from_metadata(metadata: dict[str, Any]) -> dict[str, int]:
    return sheet_ids_from_metadata_payload(metadata)


def quote_sheet_name(sheet_name: str) -> str:
    return "'" + sheet_name.replace("'", "''") + "'"
