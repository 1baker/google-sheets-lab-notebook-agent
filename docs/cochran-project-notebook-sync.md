# Cochran Project Notebook Synchronization

## Source inventory

The Cochran Research Group Drive review identified these relevant workbooks:

- `Emulsion Polymerization`, native Google Sheet
  (`1zhdi6LkRNhQRitZjzd1mkN4kOrtrMKiQ2Wx4H_Hws0c`). The workbook has 219
  experiment tabs spanning EP-71 through EP-324. `EP-200-VG` captures the
  earlier seed/core/shell recipe layout. The latest variable-feed tabs,
  including `EP-324 Variable Feed Rates SFS Trial 2`, add detailed oxidant,
  reductant, monomer-emulsion, radical-flux, and estimated-particle-size
  schedules.
- `Biomag emulsion Recipe.xlsx`
  (`15UyTrQp16bK0D1WI0LbllKwZciheJ8d8`). Its repeated batch tables distinguish
  weight fraction, goal mass, and actual mass, then record mixing instructions,
  passes, separation, foam, sprayability, and aging or freeze-thaw observations.
- `BioMAG Surfactant Charge Sheet`
  (`1b34YnDYw8TTplfM43ZDPwrg7qsDE3qkAHUG8x7_TpCQ`) contains 49 tabs: a master
  lineage index, 300 L and 10 L batch tabs, lab-scale DES-OX trials,
  terpolymer-charge tabs, and constants. Its master index links experiment
  description and date to input lots (AG, AEHOSO/SESO, CTA, AMBN, solvents) and
  output product lots. Individual charge tabs add scale/composition inputs,
  target versus actual charge mass, mass variance, numbered instructions, and
  time-series reaction temperature, RPM, oil temperature, torque, initiator
  addition, and event notes.

Charge-sheet experiment IDs are normalized (for example, master-index product
lot `BMS-18-081123` and tab `300L Surfactant Batch #18` both map to
`BMS-018`) so lineage and run records join in the project notebook. The complete
product lot remains in `product_lot`.

The source files remain authoritative and read-only. The sync writes only to a
separate project notebook spreadsheet.

## What the workflow keeps

The normalized contract was derived from both the early and latest recipe
layouts. It preserves:

- source workbook, source tab, source range, parser version, source fingerprint,
  sync time, and a direct source link;
- experiment ID, date when present, stage, subsection, and active/retired state;
- material, normalized role, pphm, planned mass, actual mass, volume, density,
  concentration, and units;
- target/calculated values separately from actual/measured values;
- input and product lot identifiers when a selected source tab supplies them;
- seed, core, shell, and functional-shell organization;
- pre-reactor, monomer pre-emulsion, aqueous pre-emulsion, initiator, redox
  shot, and redox-feed sections;
- feed interval, start/end time, feed rate, cumulative percentage, radical flux,
  estimated particle size, temperature, RPM, solids, ratios, particle count, and
  other labeled calculations;
- charge target, actual charge, mass variance, clock time, oil temperature,
  torque, incremental initiator addition, and cumulative initiator-addition
  history for BioMAG charge sheets;
- free-text instructions and observations such as temperature peaks,
  equilibration, separation, coagulum, foam, sprayability, mixing, passes, and
  stability;
- a lossless JSON copy of additional cells on every imported row so uncommon
  calculations are not discarded by the normalized columns.

`Project Notebook Records` stores those source-shaped records.
`Source Sync` stores one state row per source tab. Imported records are also
projected into the experiment entry used by notebook search and recommendation:
components become formulation context, process/feed records become observation
context, and result records become result context.

The same sync also rebuilds `Plot Data`, `Plot Definitions`, and
`Plot Dashboard`. It creates process, feed, charge-comparison, composition, and
outcome plots from canonical rows rather than depending on sparse source
charts. See [plotting-workflow.md](plotting-workflow.md).

## Live workflow

Install the optional Google dependencies and confirm credentials:

```bash
pip install -e .[google]
PYTHONPATH=src python3 -m lab_notebook_agent.cli google-doctor \
  --spreadsheet-id TARGET_PROJECT_NOTEBOOK_ID
```

Run a dry audit first. Source tabs are required and repeatable by design; there
is no implicit “scan all 219 tabs” mode.

```bash
PYTHONPATH=src python3 -m lab_notebook_agent.cli \
  google-project-notebook-sync-live \
  --source-spreadsheet-id 1zhdi6LkRNhQRitZjzd1mkN4kOrtrMKiQ2Wx4H_Hws0c \
  --target-spreadsheet-id TARGET_PROJECT_NOTEBOOK_ID \
  --source-sheet "EP-200-VG" \
  --source-sheet "EP-324 Variable Feed Rates SFS Trial 2" \
  --source-range "A1:AO1200" \
  --run-output artifacts/cochran-project-sync-dry-run.json \
  --report-output artifacts/cochran-project-sync-report.json \
  --audit-output artifacts/cochran-project-sync-audit.json \
  --batch-output artifacts/cochran-project-sync-batch.json
```

Inspect the report, especially `source_sheets`, `warnings`,
`unchanged_sources`, `changed_sources`, and `apply_audit`. Apply the same
selection only after the audit is valid:

```bash
PYTHONPATH=src python3 -m lab_notebook_agent.cli \
  google-project-notebook-sync-live \
  --source-spreadsheet-id 1zhdi6LkRNhQRitZjzd1mkN4kOrtrMKiQ2Wx4H_Hws0c \
  --target-spreadsheet-id TARGET_PROJECT_NOTEBOOK_ID \
  --source-sheet "EP-200-VG" \
  --source-sheet "EP-324 Variable Feed Rates SFS Trial 2" \
  --source-range "A1:AO1200" \
  --apply \
  --run-output artifacts/cochran-project-sync-applied.json
```

Use `--parser-profile cochran-emulsion-ledger` for a source tab that uses
`component / wt. % / goal / actual` columns. The default `auto` profile selects
that parser when those headers are present, selects
`cochran-charge-sheet` for mass-charge plus reactor-log tabs, and otherwise
selects `cochran-charge-index` for the master material/product-lot index or uses
the staged emulsion-recipe parser.

For an exported or downloaded Excel source such as
`Biomag emulsion Recipe.xlsx`, use the local path. Dry-run and inspect the
report first:

```bash
PYTHONPATH=src python3 -m lab_notebook_agent.cli project-notebook-sync \
  --source-workbook "Biomag emulsion Recipe.xlsx" \
  --target-workbook artifacts/lab_notebook_template.xlsx \
  --source-sheet "EXACT_SOURCE_TAB_NAME" \
  --parser-profile auto \
  --report-output artifacts/biomag-project-sync-dry-run.json
```

Apply to a target copy after review:

```bash
PYTHONPATH=src python3 -m lab_notebook_agent.cli project-notebook-sync \
  --source-workbook "Biomag emulsion Recipe.xlsx" \
  --target-workbook artifacts/lab_notebook_template.xlsx \
  --source-sheet "EXACT_SOURCE_TAB_NAME" \
  --apply \
  --workbook-output artifacts/lab_notebook_biomag_synced.xlsx \
  --report-output artifacts/biomag-project-sync-applied.json
```

The source workbook is opened read-only with formulas resolved to their cached
values. Only the target workbook or requested target copy is saved.

## Repeat-run behavior

The source-tab fingerprint is calculated from its bounded normalized cell
values.

- Same fingerprint, parser profile, and parser version: no target data requests.
- Changed source: stable record IDs are compared field by field.
- New source rows: appended once.
- Changed source rows: updated in place.
- Removed source rows: retained for audit but marked `active=false`.
- Existing human experiment fields are preserved. The sync updates provenance
  fields and only fills selected descriptive fields when they are blank.

This makes the command suitable for a scheduled task after the exact source-tab
selection has been reviewed. A scheduler should archive each JSON run report and
alert on `warning_count`, an invalid `apply_audit`, or unexpected record-count
changes.

## Current safety boundary

The parser intentionally does not infer unverified reagent inventory IDs,
hazards, or safety approval. Imported experiments enter as `needs_review`.
Source warnings are recorded rather than silently forcing an uncertain mapping.
The normal experiment preflight and human review gates still apply before a
recipe is treated as runnable.
