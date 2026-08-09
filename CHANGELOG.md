# Changelog

This file records user-visible release changes. `JOURNAL.md` contains the complete investigation
and implementation history.

## 0.3.0 — 2026-08-08

### Added

- Headless Python engine for Fixings, Exposure, Daily P&L, and Cumulative P&L.
- Strict typed input contracts, fail-closed validation, deterministic manifests, and atomic
  publication.
- Macro-free Excel input and output adapter.
- Read-only legacy `.xlsm` importer with normalized CSV/JSON output and audit files.
- SETUP-derived BOOK and Underlying configuration.
- Event ledger, explicit closure rows, operating flows, trade-entry adjustments, and simulation
  controls.
- Legacy-format `Daily Report D2`, generated from verified Python P&L output.
- Synthetic golden, regression, property, validation, determinism, performance, importer, and
  Excel-adapter tests.
- Installation, quick-start, detailed-user, methodology, decision, acceptance, and repository
  documentation.

### Corrected

- Same-day fixing eligibility now uses `Fixing Date >= Trade Date`.
- Trades dated on non-Market Days apply once on the next configured Market Date.
- Current-month TTF and PVB-family curves import from the dated prompt series rather than an empty
  same-month forward column.
- Excel output tables no longer create the table/AutoFilter corruption reported by Microsoft
  Excel.
- Windows installations now include the `tzdata` runtime dependency required to resolve the
  configured `Europe/Madrid` timezone.

### Known acceptance items

- Product-specific DAY_AHEAD weekend and holiday pricing requires final business fixtures.
- Brent and Henry Hub roll cases require named business regression examples.
- The 13-BOOK Initial P&L bridge and Logistics sign require final business acceptance.
- A historical July 2026 run used an explicit temporary zero-price policy; that policy is not a
  general default.
