# Active Plotting Workflow

## Why the notebook needs managed plotting

The source audit found that chart coverage is sparse and inconsistent. The
219-tab `Emulsion Polymerization` Google workbook has no embedded charts. The
49-tab `BioMAG Surfactant Charge Sheet` has only three embedded charts, all on
older batch tabs (`Surfactant Batch #11`, `#12`, and `#13`); newer 300 L and
10 L tabs do not carry those charts forward.

The notebook therefore does not inherit source charts. It records canonical
numeric plot points and generates charts from those records on every refresh.
This avoids a chart silently pointing at the wrong source columns when one
source-tab layout changes.

## Managed tabs

- `Plot Data` is the numeric chart source. Each contiguous block begins with a
  managed header row and continues with point rows. Point rows retain a stable
  `plot_point_id`, experiment, stage/section, source record IDs, source ranges,
  quality flag, units, and recorded time.
- `Plot Definitions` stores the stable `plot_id`, chart type, axes and series,
  exact one-based `Plot Data` row bounds, source-point count, and `ready` or
  `insufficient_data` status.
- `Plot Dashboard` stores the chart inventory. Excel charts are embedded by
  `openpyxl`; Google charts are added, updated, repositioned, or deleted by
  stable `[LNA:plot-id]` title markers.

These three tabs are derived and fully managed. Enter or correct observations in
`Daily Log`, `Results`, or the authoritative source spreadsheet, then refresh.
Do not hand-edit plot blocks.

## What gets plotted

Imported `Project Notebook Records` can generate:

- reaction and oil temperature versus elapsed process time;
- RPM and torque/power versus elapsed process time;
- cumulative initiator addition versus elapsed process time;
- feed rate, cumulative feed percentage, radical flux, and estimated particle
  size versus feed time;
- planned versus actual charge mass by material;
- formulation composition in pphm by material.

Manual `Daily Log` rows can generate time series for temperature, RPM, pH,
particle size, viscosity, solids/conversion/residual-monomer percentages, PDI,
and Tg. Numeric `Results` rows generate cross-experiment trends grouped by
measurement type and units.

A line definition with only one numeric point is retained as
`insufficient_data` and is not embedded yet. This makes the missing next
measurement visible without pretending that one point is a trend. A column
comparison can be ready with one material.

## What to record during a run

For useful process plots, record:

- clock time or elapsed minutes on every process-log row;
- reaction temperature and oil/controller temperature separately;
- RPM and torque/power, with the source units noted;
- incremental and cumulative initiator addition;
- feed start/end time, feed rate, and cumulative feed percentage;
- radical flux and estimated/measured particle size when calculated;
- planned and actual component mass, not only the variance;
- outcome timestamp, sample ID, measurement type, numeric value, units,
  replicate, and quality flag.

Blank values remain chart gaps. Planned, observed, and estimated values remain
distinguishable through the point quality flag and separate series.

## Local workbook refresh

Dry-run:

```bash
PYTHONPATH=src python3 -m lab_notebook_agent.cli plot-notebook \
  --workbook artifacts/lab_notebook_template.xlsx \
  --report-output artifacts/plot-refresh-dry-run.json
```

Apply in place or to a copy:

```bash
PYTHONPATH=src python3 -m lab_notebook_agent.cli plot-notebook \
  --workbook artifacts/lab_notebook_template.xlsx \
  --apply \
  --workbook-output artifacts/lab_notebook_plotted.xlsx \
  --report-output artifacts/plot-refresh-applied.json
```

The local project sync also refreshes plots automatically after projecting new
source rows:

```bash
PYTHONPATH=src python3 -m lab_notebook_agent.cli project-notebook-sync \
  --source-workbook "source.xlsx" \
  --target-workbook artifacts/lab_notebook_template.xlsx \
  --source-sheet "EXACT_SOURCE_TAB" \
  --apply \
  --workbook-output artifacts/lab_notebook_synced.xlsx
```

## Live Google Sheets refresh

Dry-run and inspect `plot_report`, `apply_audit`, and
`batch_update_requests`:

```bash
PYTHONPATH=src python3 -m lab_notebook_agent.cli \
  google-plot-notebook-live \
  --spreadsheet-id TARGET_PROJECT_NOTEBOOK_ID \
  --run-output artifacts/google-plot-refresh-dry-run.json \
  --report-output artifacts/google-plot-report.json \
  --audit-output artifacts/google-plot-audit.json \
  --batch-output artifacts/google-plot-batch.json
```

Apply only after the audit is valid:

```bash
PYTHONPATH=src python3 -m lab_notebook_agent.cli \
  google-plot-notebook-live \
  --spreadsheet-id TARGET_PROJECT_NOTEBOOK_ID \
  --apply \
  --run-output artifacts/google-plot-refresh-applied.json
```

`google-project-notebook-sync-live` includes the same plot-data replacement and
chart reconciliation automatically. Source spreadsheet IDs remain read-only;
all plot writes target the project notebook.

Repeat runs are stable. Data tabs are replaced only when their semantic content
changes; timestamps alone do not cause replacement. Managed Google charts are
reconciled by plot ID, including stale-chart deletion. Charts without the
`[LNA:...]` marker are treated as user-owned and left alone.
