# Evidence-based lab-record design

Reviewed 2026-08-13. This note records the evidence used to turn the reaction
sheet critique into the workbook contract and preflight release gates.

## Product research performed with agent-browser

The browser review covered current eLabFTW 5.6 and RSpace documentation rather
than relying on marketing summaries.

- eLabFTW treats an experiment as a linked record containing narrative,
  structured fields, steps, resources, and attachments. It preserves revisions
  and changelog events, uses soft deletion, can lock records, and can create an
  immutable JSON-plus-hash timestamp archive. This supports stable run IDs,
  append-only audit events, linked files, revision reasons, and fingerprinted
  signatures.
- eLabFTW templates and typed custom fields provide repeatable capture without
  removing free-form scientific text. Its export uses the ELN Consortium's
  RO-Crate-based interchange format. This supports governed templates,
  controlled vocabulary, narrative sections, and portable exports.
- RSpace Forms provide ordered typed fields, validation ranges, and required
  fields. A record with missing required fields cannot be signed. Published
  form edits create a new version and preserve prior versions for records that
  used them. This supports immutable template/version binding and a hard
  signoff gate.
- RSpace notebooks preserve a chronological sequence and retain document
  metadata/history when content is shared or moved. RSpace exports can bundle
  notebook records with externally stored raw files and structured metadata.
  This supports contemporaneous timestamps, file provenance, and complete
  record export rather than links that rot independently.

Primary product documentation:

- <https://doc.elabftw.net/docs/usage/traceability-and-auditability>
- <https://doc.elabftw.net/docs/usage/user-guide/experiments>
- <https://doc.elabftw.net/docs/usage/user-guide/templates>
- <https://doc.elabftw.net/docs/usage/user-guide/custom-fields>
- <https://documentation.researchspace.com/l/en/article/rozak8tlwr-forms>
- <https://documentation.researchspace.com/l/en/article/ptqn8pr7xw-notebooks>
- <https://documentation.researchspace.com/l/en/article/wxfk9gf0a0-templates>
- <https://documentation.researchspace.com/l/en/article/25mt56kamf-export-options>

## Literature review performed with LitScout

Reproducible discovery:

```bash
litscout search multi \
  "electronic laboratory notebook scientific record keeping reproducibility data integrity audit trail best practices" \
  --sources openalex,crossref,semantic_scholar --depth medium --limit 30 \
  --save --session-name labnotebook/good-record-2026-08-13

litscout search multi \
  "laboratory notebook best practices complete contemporaneous records raw data reproducibility" \
  --sources openalex,semantic_scholar --depth intense --limit 25 \
  --save --session-name labnotebook/record-practice-2026-08-13
```

The first session was enriched through OpenAlex, Crossref, and Semantic Scholar.
The second omitted Crossref after the locally configured contact parameter was
rejected with HTTP 400. The saved exports are
`artifacts/litscout-good-lab-record.json` and
`artifacts/litscout-record-practice.json`.

High-value works included:

- Bird, Willoughby, and Frey, *Laboratory notebooks in the digital era*,
  Chemical Society Reviews (2013), DOI `10.1039/C3CS60122F`.
- Milsted et al., *LabTrove: A Lightweight, Web Based, Laboratory “Blog” as a
  Route towards a Marked Up Record of Work in a Bioscience Research
  Laboratory*, PLOS ONE (2013), DOI `10.1371/journal.pone.0067460`.
- Kanza et al., *Electronic Lab Notebooks and Experimental Design Assistants*
  (2019), DOI `10.1007/164_2019_287`.
- Walsh and Cho, *Implementation and use of cloud-based electronic lab notebook
  in a bioprocess engineering teaching laboratory* (2017), DOI
  `10.1186/s13036-017-0083-2`.
- Bespalov et al., *Introduction to the EQIPD quality system* (2021), DOI
  `10.7554/eLife.63294`.

The strongest directly inspected full text was LabTrove. Its findings were:

- a record must contain enough detail for a knowledgeable researcher to repeat
  the work;
- edits must remain possible, but every version and edit reason must remain
  accessible;
- repeated work benefits from templates, while copy/paste is error-prone;
- overly rigid formalism can degrade records or create parallel shadow notes;
- samples, procedures, materials, and data files need stable identities and
  explicit links;
- raw instrument data should be attached alongside a human-viewable rendering;
- metadata collection should be automated where possible; and
- the system should preserve the input-process-output provenance chain.

## Implemented design requirements

The `0.12.0` contract implements the combined evidence as follows:

| Requirement | Workbook implementation | Enforced by |
|---|---|---|
| Stable identity and accountability | Experiment ID, operator, timestamps | `accountable_run_identity` |
| Published, repeatable structure | Effective template ID and exact version | `governed_template` |
| Authoritative method provenance | Active protocol ID, exact version, controlled URL | `controlled_protocol` |
| Equipment fitness | Registered equipment and calibration state | `equipment_readiness` |
| Plan versus actuality | Batch Builder separates calculated plan from attributable actual charges | `batch_plan_and_actuals` |
| Scientific reasoning without shadow notes | Eight ordered narrative sections | `notebook_sections` |
| Contemporaneous observations | Timestamped Bench Log | `daily_log_observations` |
| Raw-data provenance | Measurement-to-Raw Data Files identifiers | `raw_data_traceability` |
| Input-process-output closure | Reaction Outcomes yield, purity, appearance, and material balance | `reaction_outcome` |
| Exceptions remain visible | Deviations with assessment, disposition, owner, and reviewer | `deviations` |
| Independent review | Summary, completion, reviewer, and review timestamp | `review_attribution` |
| Signed state cannot drift silently | Signature must match the current record fingerprint and witness | `signature_integrity` |

`experiment-preflight` now has `planning`, `review`, and `archive` gates. The
Reaction Master also exposes section, raw-data, and outcome readiness and will
not report `READY_TO_ARCHIVE` until those record components are complete.

This is a traceable research workflow, not a claim of GLP, GMP, 21 CFR Part 11,
or other regulated-system certification.
