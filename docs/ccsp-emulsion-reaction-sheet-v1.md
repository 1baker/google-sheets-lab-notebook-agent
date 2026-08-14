# CCSP emulsion-polymerization reaction sheet v1

## Decision

The first reaction sheet is a deterministic planning calculator based on the
current `CCSP-52` tab of the shared **CCSP Emulsion Polymerization** workbook.
It is intentionally not another full electronic lab notebook.

The interface is reduced to four focused views so the working recipe and feed
rates are not buried under chemistry metadata or audit prose:

1. `Reaction Plan` — the bench-facing charge plan and primary editable inputs.
2. `Feed Schedule` — editable interval endpoints and distributions with
   calculated emulsion, oxidant, and reductant rates in mL/min.
3. `Checks` — stage mass, volume, polymer solids, and design split.
4. `Assumptions` — a concise issue register; detailed evidence remains in the
   companion audit JSON and this document.

Yellow cells are validated inputs or controlled properties. Green cells are
protected calculations. The visible release state remains `NOT RELEASED` until
the run setup is complete, every feed schedule reconciles, and all `VERIFY`
assumptions are resolved.

The source workbook remains read-only. The v1 planner is generated locally by:

```bash
PYTHONPATH=src python3 -m lab_notebook_agent.cli ccsp-reaction-sheet \
  --output artifacts/ccsp_emulsion_reaction_v1.xlsx \
  --audit-output artifacts/ccsp_emulsion_reaction_audit_v1.json
```

## Values retained in v1

- Stage: seed, core, shell, or functional shell.
- Addition group: pre-reactor, monomer pre-emulsion, aqueous pre-emulsion,
  redox shot/feed, or chase.
- Controlled material identity and role.
- PHR, stage factor, allocation fraction, or explicit carry mass.
- Planned mass, density, and volume using one row-local equation.
- Stage total mass and volume.
- Polymer-forming mass and polymer-solids fraction.
- Target versus actual core/shell/functional-shell mass split.
- Core, shell, and functional-shell time intervals and calculated feed rates
  for emulsion, oxidant, and reductant streams.
- Run ID, operator, date, recipe revision, reactor/pump map, temperature,
  agitation, purge condition, and pump-calibration status.
- Live checks for increasing feed time, nondecreasing cumulative emulsion,
  100% stream allocation, scheduled-versus-recipe volume, negative rates,
  setup completeness, open assumptions, and overall release state.

Zero-quantity placeholders are omitted from the main plan. They can be added
when actually required without changing the calculation rules.

## Items deferred from v1

- Instantaneous radical-flux modeling.
- Syringe-diameter lookup tables.
- Advanced mixed-surfactant CMC calculations.
- Per-charge active fraction, active mass, molecular weight, molar ratios,
  source cell, and normalization-note columns. These remain in the calculation
  engine and audit rather than the simplified bench interface.
- Observations, measurements, outcomes, and signature workflow already owned by
  the broader lab-notebook contract.

These are not required to answer the first planning question: **what is being
charged, how much is charged, what does that amount mean chemically, and do the
stage totals and key ratios reconcile?**

## Guardrails

- Yellow recipe and feed cells use stop-style Excel validation.
- Green formulas and calculated carry-forward cells are protected from routine
  editing; sheets may still be deliberately unprotected for controlled repair.
- Invalid or zero feed intervals return a visible calculation error rather than
  a plausible zero rate.
- The `Checks` view distinguishes the narrow stage-split check from feed checks
  and the overall release gate.
- The initial workbook is intentionally `NOT RELEASED`: pump calibration and
  run metadata are blank, and MMA density plus Dowfax actives remain `VERIFY`.

## Confirmed source issues corrected or surfaced

1. `CCSP-52!G6` references `E10/F10`, so zero sodium-acetate mass displays
   `0.1277 mL`. V1 always uses the same row's mass divided by density.
2. `CCSP-52!K35:K36` are `#DIV/0!`. They omit the non-zero main APS feed at
   `E54` and use inconsistent molecular-weight divisors. V1 uses active APS
   mass and `228.18 g/mol` for both ratios.
3. Shell methyl methacrylate uses `0.89 g/mL` at `F74`, while the same material
   uses `0.94 g/mL` at `F101`. V1 uses `0.94 g/mL`, pending lot/SDS confirmation.
4. Rounded source stage factors produce an actual new-polymer split of about
   `70.38% / 9.85% / 19.77%`, not exactly `70% / 10% / 20%`. V1 shows target
   and actual values separately.
5. Source solids formulas omit at least the shell BDDMA charge and do not state
   whether “solids” means polymer-forming material or all nonvolatile material.
   The workbook shows polymer solids; the companion audit reports both polymer
   and theoretical nonvolatile definitions.
6. Dowfax 2A1 is treated as 100% active in the source. V1 preserves that source
   assumption but marks product actives as a required verification.

## Independently verified baseline

| Calculation | Result |
|---|---:|
| Seed final charge mass | 196.560 g |
| Core latex mass | 144.618 g |
| Core-shell latex mass | 162.3588 g |
| Functional-shell final mass | 201.9396 g |
| Core monomer:initiator | 684.5075 |
| Core CTA:initiator | 0.130081 |
| Core crosslinker:CTA | 14.0992 |
| Actual core/shell/functional split | 70.3809% / 9.8533% / 19.7658% |

The Python calculation engine and native Microsoft Excel recalculation produced
the same values for every entry in this table.
