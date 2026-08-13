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
