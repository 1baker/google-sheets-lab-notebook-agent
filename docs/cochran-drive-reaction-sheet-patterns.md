# Cochran Research Group reaction-sheet patterns

Reviewed read-only on 2026-08-12. No Google Drive files were modified.

## Representative sources

- [Baker's Reaction Sheet](https://docs.google.com/spreadsheets/d/1xHrbEugcRTG24J7FAC7j5xtGLCum2he4-9dc6e29Upk)
- [2026 02 Sperry Rails Reactions](https://docs.google.com/spreadsheets/d/14YL7MYbIYmEVEoNhwG9hyxS8WO03Zf5WEvJ1y4VrmK8)
- [PU-BAN-Reactions-DJ](https://docs.google.com/spreadsheets/d/1Rn9dBwUnJhNN7afXkyakXe1Qrlx1gt78P8IUvSIlUCk)
- [Samsara Experimental Sheet](https://docs.google.com/spreadsheets/d/1iL4Cl88025KxSyvzOKK4qgAMbzVzlNzH4d4G8H6-9Cw)
- [Emulsion Polymerization](https://docs.google.com/spreadsheets/d/1zhdi6LkRNhQRitZjzd1mkN4kOrtrMKiQ2Wx4H_Hws0c)
- [Pyrrolidone Monomer Synthesis](https://docs.google.com/spreadsheets/d/1Ucm4AbXudFU45tSZmuB6RdIqpttM0FjAwZmKAbXX8bE)
- [Nylon 4 Synthesis](https://docs.google.com/spreadsheets/d/1wBpli3ib1S1of-3bcOfVmEcgp1eeUr5PdNKhOMBMz7E)
- [BioMAG Surfactant Charge Sheet](https://docs.google.com/spreadsheets/d/1b34YnDYw8TTplfM43ZDPwrg7qsDE3qkAHUG8x7_TpCQ)
- [Castor Oil Thiolation and Aminated Epoxidized Soybean Oil Synthesis](https://docs.google.com/spreadsheets/d/1e7ZrmSLN9M9b8aN5SlZaT8N_ZjUXZPxYkANK3r2IcqI)
- [Bot tested formulation sheet](https://docs.google.com/spreadsheets/d/1xcXE1rX_ms1ffCON2K2YTR2gEMKpUHy_6Y38OH_8Iug)
- [CTA Synthesis — Vivek Oxcart tab](https://docs.google.com/spreadsheets/d/1cUiBLIUCIEB0vUS3qrY0ueTPBKcuznhIB9TIxFxGaf0)

## Vivek emulsion-polymerization workbook scan

A complete bounded scan on 2026-08-13 inspected all 219 tabs, 156,574
populated cells, and 68,282 formulas in `Emulsion Polymerization`. Populated
content extends through row 178 and column AH; the nominal 900+ row grids are
mostly blank template area.

- Eighty-eight tabs, from `EP-234 Variable Feed Rates CSP` through
  `EP-324 Variable Feed Rates SFS Trial 2`, use explicit interval feed tables.
- Those tables calculate separate emulsion or monomer, oxidant, and reductant
  rates. Later runs also calculate radical flux and estimated particle size.
- Older tabs already place oil/feed rate, temperature, and RPM beside their
  stage recipes. `EP-200-VG` is a representative transition layout; later tabs
  use separate core, shell, and functional-shell schedules.
- The historical purple, red, blue, and green fills encode stage or stream
  groupings, not a stable input/output contract. The generated planner uses a
  simpler rule: pale yellow is editable input and pale green is calculated
  output. Muted stage colors appear only in the Stage column.

## Vivek Oxcart review

The `Vivek Oxcart` tab was reviewed beside `OXCART-1`, `AKM-OXCART-1`, and
`Thamer-OXCART` in the same workbook. This makes it possible to distinguish the
useful local adaptation from inherited template behavior.

Strengths worth preserving:

- A single target product mass drives reagent moles, masses, and volumes.
- Molecular weight and density remain visible beside each charge, making the
  calculation inspectable rather than opaque.
- Reagent equivalents and the intended two-hour separation between additions
  are kept close to the recipe.
- Vivek adapted the inherited OXCART calculator to the shorter-chain chemistry
  instead of duplicating arithmetic by hand.

Risks to correct in a reusable notebook:

- The tab is a plan/calculator, not a completed experiment record: it has no
  stable run ID, date, operator, lot IDs, actual charges, product recovery,
  yield, observations, attachments, or conclusion.
- Column labels such as `1st` through `5th` and abbreviations such as `3c2b` do
  not identify a controlled step, material, unit, or timestamp semantics.
- The right-side addition block contains partial totals and isolated clock-time
  values without a complete sequence or provenance.
- Copied labels and formulas can drift from the chemistry. For example, the
  Vivek tab retains an `Ether` row/heading while the material identity and
  neighboring template choices differ, and the same fixed historical times
  appear on multiple researcher tabs.
- There is no explicit theoretical-versus-recovered product comparison or
  material-balance closure, so a mathematically scaled recipe can look complete
  even when the physical run is undocumented.

The implementation preserves the compact scaling behavior in `Batch Builder`
and addresses these gaps with `Notebook Sections` and `Reaction Outcomes`.
Source chemistry values and procedures remain read-only and were not copied
into the generated template.

## Recurring design conventions

- Begin with target scale or target product mass, then calculate each charge.
- Separate editable inputs, calculated values, and recorded outputs visually.
- Use functional equivalents and explicit group ratios where molecular weight
  alone does not express the reaction balance.
- Account for masterbatch or stock carrier so its contribution is not charged
  twice.
- Organize charges and observations by process stage.
- Support both direct scale values and source-container weighing by difference.
- Capture time, reactor temperature, bath or jacket temperature, RPM, pressure,
  torque or power, cumulative feed, observations, and recovered mass as needed.
- Keep reusable reagent constants separate from experiment records.
- Provide one compact experiment history or master rollup rather than requiring
  scientists to inspect every reaction tab.

## Applied design choices

The workbook contract implements these as chemistry-neutral fields and formula
paths. It does not copy another scientist's experimental values, procedures, or
chemistry-specific assumptions. Functional-equivalent groups are configured in
the reagent database and per-run numerator and denominator fields, while the
existing normalized experiment, charge, bench-log, and measurement tables stay
intact.
