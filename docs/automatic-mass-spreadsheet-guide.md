# Automatic Reaction Mass Spreadsheet Guide

Contract version: 0.10.0

## What the workbook does

The workbook converts a scientifically meaningful recipe basis into the
as-supplied mass that must be weighed. It keeps active material, solution
strength, volume, moles, equivalents, actual charge error, and batch balance
visible as separate quantities.

## Set up a reaction

1. In **Experiments**, enter a unique `experiment_id`, the
   `target_batch_mass_g`, and a `default_mass_tolerance_percent`.
2. Add one row per charge in **Batch Builder** and select exactly one
   `calculation_mode`.
3. Enter only the inputs required by that mode. The blue cells are scientist
   inputs; gray cells are calculated.
4. Confirm that `formula_status` reads `READY` before weighing.
5. Record `actual_mass_g`. The sheet calculates mass error, percent error, and
   `PASS` or `FAIL` against the row tolerance or experiment default.
6. Use **Reaction Master** to judge total planned mass, actual completion,
   out-of-tolerance charges, and overall mass-plan status.
7. Link the run to an effective **Experiment Templates** version, record
   consumed material in **Inventory Transactions**, and complete signoff in
   **Record Signatures** after the record is finalized.

## Choose the calculation mode

| Mode | Enter | The workbook calculates |
|---|---|---|
| `direct_mass` | `direct_mass_g` | The same as-supplied mass |
| `batch_wt_percent` | `recipe_wt_percent`; experiment target mass | Target mass × wt% / 100 |
| `pphm` | `parts_per_hundred_monomer`; `stage_monomer_basis_g` | Basis × pphm / 100 |
| `target_active_mass` | `target_active_mass_g`; stock active fraction | Active target ÷ active fraction |
| `equivalents` | `target_equivalents`; `equivalent_basis_mmol`; molecular weight; stock active fraction | Active mmol and as-supplied mass |
| `functional_equivalents` | `target_functional_equivalents`; functional equivalent weight; stock active fraction | Functional equivalents and as-supplied mass |

For container-by-difference charging, choose `by_difference` and enter the
source-container weights before and after the charge. `effective_actual_mass_g`
then drives variance, actual moles, functional equivalents, and carrier totals.
The direct actual-mass cell is retained for ordinary scale entry.

Configure numerator and denominator functional groups plus a target ratio in
`Experiments` when a run needs an NCO:active-H or similar stoichiometric audit.
`Reaction Master` reports the planned and actual group totals and flags a ratio
mismatch without assuming a particular chemistry.

A blank `stock_active_fraction` means neat material (1.0). For a 25% solution,
enter `0.25`, not `25`. A zero, negative, or greater-than-one fraction is
rejected so it cannot produce a plausible-looking mass.

## Bench examples

- A 500 g batch at 40 wt% gives 200 g as supplied.
- 2 pphm on 300 g monomer gives 6 g.
- A 5 g active target from a 25% stock gives 20 g as supplied.
- 0.5 equivalents on a 100 mmol basis, MW 200 g/mol, and 50% active stock gives
  50 mmol active, 10 g active, and 20 g as supplied.
- An actual charge of 19.7 g against 20.0 g is -0.3 g or -1.5%; it passes a 2%
  tolerance and fails a 1% tolerance.

## Read the safeguards

`formula_status` names the first actionable missing or invalid input. A blank
planned mass is intentional whenever status is not `READY`; it prevents a bad
input from becoming a believable weigh-out.

At run level, `Mass plan status` is `READY` only when charge rows exist, every
charge is calculation-ready, and planned total agrees with target batch mass
within the default tolerance. `NO_TARGET_MASS` means row calculations may be
usable, but total batch balance has not been proven. `CALCULATION_ISSUES`,
`CHARGE_TOLERANCE_FAILURE`, and `MASS_MISMATCH` require resolution before the
reaction is treated as ready.

`Governance status` separately evaluates record control. It reports missing or
non-effective templates, incomplete work, required signoff, required witness,
or `READY_TO_ARCHIVE`. The workbook records these events but cannot by itself
guarantee account identity, immutability, or regulatory compliance.

## Recommended review habit

Before the run, verify `formula_status`, planned total, and mass-plan status.
During charging, enter actual masses rather than overwriting calculated planned
masses. After the run, resolve each failed tolerance deliberately and document
accepted deviations in the charge notes. The spreadsheet supports scientific
judgment; it does not replace a reviewed recipe, current SDS controls, or an
approved operating procedure.
