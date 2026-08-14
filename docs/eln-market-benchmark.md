# Electronic lab notebook market benchmark

Reviewed from official product documentation on 2026-08-12. This comparison
identifies transferable workflow patterns; it does not claim feature parity or
regulatory equivalence with the referenced products.

| Product | Officially documented pattern | Workbook response |
|---|---|---|
| [Benchling Notebook templates](https://help.benchling.com/hc/en-us/articles/9684282597645) | Governed template collections with draft, effective, superseded, and withdrawn states; structured tables, formulas, locked columns, and review criteria | `Experiment Templates` mirrors the lifecycle and links an exact template/version to each run. Calculated Batch Builder columns remain visually and formula controlled. |
| [Benchling review processes](https://help.benchling.com/hc/en-us/articles/39952320412685-Create-and-use-review-processes) | Required fields, checked-in attachments, submitted structured rows, record locking during review, and assigned reviewers | `Reaction Master` exposes a governance gate. The spreadsheet records readiness and signoff but explicitly does not claim irreversible locking. |
| [SciNote ELN functionality](https://www.scinote.net/product/) | Projects/experiments/tasks, protocol templates, assigned inventory, stock tracking, reminders, activity history, reports, and export | Existing run plans, protocols, audit history, and reports are augmented with a linked material transaction ledger. |
| [SciNote 2026 releases](https://www.scinote.net/release-notes/2026-release-notes/) | Protocol templates can carry results and assigned inventory; equipment usage, calibration, and maintenance can be scheduled | `Experiment Templates`, `Inventory Transactions`, and `Equipment Bookings` provide corresponding spreadsheet-level structures. |
| [LabArchives signing and witnessing](https://help.labarchives.com/hc/en-us/articles/11778411173652-Signing-and-Witnessing-a-Notebook-Page) | Finalized pages can be signed, independently witnessed, and locked | `Record Signatures` captures signer, timestamp, meaning, fingerprint, witness, supersession, and revocation context. It is an audit-friendly workflow, not a cryptographic lock. |
| [LabArchives widgets](https://help.labarchives.com/hc/en-us/articles/11732040270484-Introduction-to-Widgets) | Reusable forms, checklists, tables, dropdowns, and scientific calculators | Existing controlled tables, Run Capture Plan checklists, validations, and automatic mass/stoichiometry calculators cover the useful spreadsheet analogue. |
| [RSpace features](https://www.researchspace.com/features) | Templates/forms/snippets, internal links, file galleries, chemistry tools, signing/witnessing, inventory, sharing, export, and APIs | Existing protocols, source links, attachments, chemistry calculations, signoff ledger, inventory ledger, and machine-readable contract form the portable subset. |
| [Chemotion ELN](https://chemotion.net/docs/eln/ui) | Reactions connect starting materials, reagents, and product samples; sample splits preserve physical lineage; collections organize work | Batch Builder roles, reagent IDs, Samples parent lineage, and experiment-centered Reaction Master preserve these relationships without imposing organic-synthesis-only assumptions. |
| [Signals Notebook experiment templates](https://support.revvitysignals.com/hc/en-us/articles/37319392132756-Signals-Notebook-Creating-and-Managing-Experiment-Templates) | Templates combine required fields with text, chemistry, materials tables, tasks, images, and file attachments; required content can gate close/sign transitions | `Notebook Sections` provides typed, ordered, required narrative blocks with completion state, authorship, timestamps, and attachments. It is a portable structured-record analogue, not an application-level lock. |
| [Labguru protocols](https://help.labguru.com/en/articles/5469376-creating-and-using-protocols) | Protocols become reusable experiment templates containing procedure steps, reagents, samples, calculations, and results-oriented sections | Governed templates define required capture sections; instantiated work is recorded in `Run Capture Plan`, `Batch Builder`, `Notebook Sections`, and `Reaction Outcomes`. |
| [Labguru inventory](https://www.labguru.com/inventory) | Experiments link specific stocks and samples, decrement usage, retain physical location, and support barcode-backed container handling | Reagent/lot charge rows, append-only inventory transactions, and sample lineage retain the essential provenance. Barcode printing/scanning remains outside the workbook. |
| [Uncountable experiments](https://www.support.uncountable.com/knowledge-base/experiments/) | Each experiment separates a structured recipe (ingredients and process parameters) from structured measurements and outputs | `Batch Builder` remains the recipe surface; `Measurements` and the new `Reaction Outcomes` sheet hold results, yield, recovery, and material closure without mixing them into planning cells. |

## v0.11 design decisions

1. A run selects an exact governed template ID and version. Only an `effective`
   matching version passes the governance gate.
2. Material movements are recorded as append-only transactions rather than
   silently inferred from planned quantities. This keeps planned, charged, and
   inventory-accounting meanings distinct.
3. Equipment bookings distinguish usage, calibration, maintenance, and other
   events and retain the linked experiment and responsible person.
4. Signing and witnessing are explicit events with fingerprints and
   supersession context. The workbook does not describe editable spreadsheet
   cells as an immutable or compliant electronic signature system.
5. Existing normalized science records remain intact. All schema changes are
   new sheets or appended columns so live workbook migrations preserve prior
   row meanings.
6. Required notebook narrative is normalized into typed, ordered sections rather
   than buried in a single free-text cell. Each section retains completion,
   authorship, timestamps, and attachments.
7. Reaction outcome accounting is separate from the recipe calculator. Yield is
   based on theoretical versus recovered product, while material closure reports
   product, retained samples, waste, handling loss, and explicitly expected
   non-product loss against the actual charged mass.

## Deliberately not emulated

- Identity-provider authentication, cryptographic signatures, irreversible
  record locks, granular permissions, and regulatory certification require an
  application and administrative controls beyond an XLSX or Google Sheet.
- Real-time notifications and calendar reminders require external services.
- Chemical structure drawing and substructure search require dedicated
  cheminformatics components; the current contract retains identifiers and
  links instead of pretending text cells provide those capabilities.
