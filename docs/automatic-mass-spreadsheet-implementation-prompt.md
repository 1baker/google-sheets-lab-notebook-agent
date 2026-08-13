# Implementation Prompt: Automatic Reaction Mass Spreadsheet

You are a senior polymer scientist and laboratory informatics engineer. Upgrade
the existing `lab-notebook-agent` workbook into an automatic, bench-usable mass
planning spreadsheet for graduate students and experienced polymer scientists.

Implement the following behavior in both the generated XLSX workbook and the
Google Sheets migration/setup path. Preserve existing laboratory records and
append new contract columns rather than shifting established columns.

## Scientist workflow

1. A scientist defines a reaction in `Experiments`, including a target total
   batch mass and a default charge-mass tolerance.
2. In `Batch Builder`, every material row has one explicit `calculation_mode`:
   `direct_mass`, `batch_wt_percent`, `pphm`, `target_active_mass`,
   `equivalents`, or `functional_equivalents`.
3. The spreadsheet calculates the as-supplied mass to weigh. It must never make
   the scientist infer which input won when several basis fields are populated.
4. The sheet separately reports as-supplied mass, active mass, volume, mmol,
   equivalents, actual mass, mass error, percent error, and tolerance status.
5. `Reaction Master` automatically rolls up target batch mass, planned batch
   mass, planned-versus-target difference, actual mass, completion percentage,
   out-of-tolerance charge count, and a concise mass-plan status for every run.

## Calculation rules

- `direct_mass`: planned mass = `direct_mass_g`.
- `batch_wt_percent`: planned mass = target batch mass × recipe wt% / 100.
- `pphm`: planned mass = stage monomer basis × pphm / 100.
- `target_active_mass`: planned as-supplied mass = target active mass ÷ stock
  active fraction.
- `equivalents`: planned active mmol = target equivalents × equivalent-basis
  mmol; planned as-supplied mass = active mmol × molecular weight / 1000 ÷
  stock active fraction.
- `functional_equivalents`: planned as-supplied mass = target functional
  equivalents × verified grams per functional equivalent ÷ stock active
  fraction.
- For `by_difference` weighing, effective actual mass is source-container mass
  before charging minus its mass after charging. Missing, negative, or reversed
  container weights must produce an explicit weighing status.
- Planned and actual inactive carrier contribution is the corresponding
  as-supplied mass multiplied by one minus the active fraction.
- A blank stock active fraction means 1.0. Zero, negative, or greater-than-one
  active fractions are invalid and must not yield a plausible mass.
- Density-backed volume requires a positive density. Molecular calculations
  require a positive molecular weight.
- Actual mass error = actual mass − planned mass. Percent error is relative to
  planned mass. `within_tolerance` is blank until an actual mass exists, then is
  `PASS` or `FAIL` using the row tolerance or the experiment default.
- Formula status must identify the first actionable problem using concise
  states such as `READY`, `MISSING_MODE`, `MISSING_TARGET_BATCH_MASS`,
  `MISSING_DIRECT_MASS`, `MISSING_PPHM_BASIS`, `MISSING_ACTIVE_TARGET`,
  `MISSING_EQUIVALENT_BASIS`, `MISSING_MOLECULAR_WEIGHT`, or
  `INVALID_ACTIVE_FRACTION`.

## Run-level safeguards

- `Reaction Master` must show `READY` only when the run has charge rows, every
  charge row is calculation-ready, and planned mass agrees with target batch
  mass within the experiment's default tolerance.
- A run with no target batch mass may still use direct, pphm, active-mass, or
  equivalents modes, but its run-level status must say `NO_TARGET_MASS` rather
  than implying that total mass balance passed.
- Do not count blank formula-filled rows as charges.
- Preserve input/calculated color cues, validation dropdowns, frozen identifier
  columns, filters, and compatibility with normalized `Formulations` records.
- When a run defines numerator and denominator functional groups, a target
  equivalent ratio, and a tolerance, roll up planned and actual equivalents by
  group and report `READY` or `RATIO_MISMATCH`. Keep unconfigured runs explicit.

## Verification cases

Prove the implementation with calculation-level tests and workbook formula
tests for at least these cases:

1. A 500 g batch with 40 wt% monomer calculates 200 g.
2. 2.0 pphm on a 300 g monomer basis calculates 6 g.
3. 5 g active target from a 25% stock calculates 20 g as supplied.
4. 0.5 equivalents on a 100 mmol basis with MW 200 g/mol and 50% active stock
   calculates 20 g as supplied, 10 g active, and 50 mmol.
5. Actual 19.7 g versus planned 20.0 g gives −0.3 g and −1.5%; it passes a 2%
   tolerance and fails a 1% tolerance.
6. Missing/invalid inputs produce an explicit status and no misleading mass.
7. Reaction Master totals and statuses reference the correct Batch Builder and
   Experiments columns in both Excel formulas and Google Sheets array formulas.
8. Existing workbook and migration tests remain green.
9. 0.05 functional equivalents at 200 g/eq and 50% active calculates 20 g.
10. Source-container weights of 100.0 g and 80.3 g yield a 19.7 g effective
    actual charge.
11. NCO and active-H charge rows roll up to a configured 1.05 ratio.

Iterate until all requirements are implemented, the full test suite passes,
the generated workbook survives OOXML/openpyxl integrity checks, and a rendered
guide makes the scientist workflow understandable without reading source code.
