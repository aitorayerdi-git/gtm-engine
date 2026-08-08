# Gas Trading Model Acceptance-Test Specification

Version: 0.3-approved-build-baseline
Policy version: 0.3.0
Status: Core golden pack executable; product/workbook-specific golden cases remain pending
Date: 2026-08-07
Governing specification: `docs/GTM_ENGINE_SPECIFICATION.md`

## 1. Purpose

This document defines the evidence required to accept the Gas Trading Model (GTM) engine as functionally correct, deterministic, auditable, and safe to connect to Excel.

The automated suite can prove that an implementation conforms to the written rules and approved economic examples. It cannot independently certify the business meaning of a source sign or price rule; selected golden cases therefore remain a production-release gate.

## 2. Authority and acceptance principles

1. Approved golden cases and confirmed requirements are the acceptance oracle.
2. The Python implementation is the executable reference for the later VBA implementation.
3. Legacy workbook formulas and reconciliation output are diagnostic evidence only.
4. Expected results SHALL be calculated independently and stored in test fixtures; tests SHALL NOT generate their expected results by calling the implementation under test.
5. Tests SHALL be deterministic and SHALL not depend on live Reuters, SharePoint, Dropbox, helper workbooks, Excel calculation state, worksheet filters, or the active workbook.
6. A production build SHALL pass all blocking tests relevant to its declared feature set.
7. A failed run SHALL leave the last published output intact.

## 3. Test artifact standard

Each fixture SHALL be a versioned directory containing, as applicable:

```text
case_manifest.json
input/
  initial_exposure.csv
  initial_pnl.csv
  trades.csv
  simulation_trades.csv
  setup.csv
  market_calendar.csv
  curve_prices.csv
  fixing_prices.csv
  operating_flows.csv
expected/
  validation.csv
  fixings.csv
  exposure.csv
  pnl.csv
  build_manifest.json
notes.md
```

`case_manifest.json` SHALL record:

- case ID and purpose;
- governing policy version;
- fixed timezone;
- all model-control dates and simulation status;
- applicable decisions from the engine-specification decision register;
- expected result: `PASS` or `FAIL`;
- expected blocking-error and warning counts;
- input and expected-output hashes;
- reviewer and approval date for golden cases.

Rows SHALL carry stable source IDs. Serialized output SHALL have canonical field order and deterministic row order.

## 4. Comparison rules

Numeric comparison follows the approved D-007 convention:

- identifiers, categories, dates, booleans, error codes, row counts, and key sets compare exactly;
- volume values compare with absolute tolerance `0.000001`;
- price values compare with absolute tolerance `0.00000001`;
- P&L and monetary amounts compare with absolute tolerance `EUR 0.01`;
- no intermediate calculation is rounded;
- `NaN`, infinity, null, blank, and zero are distinct values;
- output order compares exactly after canonical sorting.

A test SHALL fail when a required expected row is absent, an unexpected row exists, a category differs, or a numeric difference exceeds tolerance.

## 5. Test levels

### 5.1 Unit tests

Unit tests cover parsing, canonicalization, mappings, date functions, allocation formulas, event application, valuation formulas, and validation rules with no file or Excel dependency.

### 5.2 Component tests

Component tests run Preflight, Fixings, Exposure, P&L, and manifest generation separately against normalized fixture tables.

### 5.3 Pipeline tests

Pipeline tests execute a complete immutable input bundle through staged output and atomic publication.

### 5.4 Golden economic tests

Golden tests use small portfolios that a business owner can calculate independently. User approval turns them into authoritative acceptance evidence.

### 5.5 Implementation-parity tests

The same fixtures SHALL run through Python and the clean VBA implementation. Parity is required before VBA can replace Python for a production calculation.

### 5.6 Excel-adapter tests

Adapter tests verify import, display, validation reporting, user controls, and safe publication. They SHALL not redefine calculation results.

## 6. Schema and Preflight tests

| ID | Case | Required result |
|---|---|---|
| VAL-001 | Required source table/file absent | FAIL with stable missing-source code before calculation |
| VAL-002 | Required header absent or duplicated | FAIL with source and field identified |
| VAL-003 | `daily_qty` blank | FAIL; blank SHALL NOT become zero |
| VAL-004 | Invalid numeric text, `NaN`, or infinity | FAIL with source row ID |
| VAL-005 | Side outside BUY/SELL aliases | FAIL with source row ID |
| VAL-006 | End Date before Start Date | FAIL with source row ID |
| VAL-007 | Unknown or inactive book | FAIL with offending canonical key |
| VAL-008 | Underlying has no fixing method | FAIL before Fixings build |
| VAL-009 | Duplicate required price key | FAIL unless an explicit source-precedence policy resolves it |
| VAL-010 | Required curve or fixing price absent | FAIL; list every safely discoverable missing required price and publish no new output |
| VAL-011 | Simulation row has no scenario | FAIL when simulation input is enabled |
| VAL-012 | Actual trade has a scenario value | Warning or FAIL as defined by schema policy; never silently reclassify |
| VAL-013 | Duplicate `source_row_id` | FAIL in the affected source |
| VAL-014 | Manifest row count or hash differs from loaded data | FAIL before calculation |
| VAL-015 | Invalid control-date ordering | FAIL with all offending dates reported |
| VAL-016 | Initial-position date differs from cut-off | FAIL |
| VAL-017 | Unsupported currency or unit | FAIL until an explicit conversion exists |
| VAL-018 | Negative `daily_qty` with separate Side | FAIL; Side is the only sign field |
| VAL-019 | Multiple independent errors | All safely discoverable errors reported in one Preflight run |
| VAL-020 | Previous validation contains failures | New report is append-only and associated with the new Build ID; stale display rows cannot alter the result |
| VAL-021 | Initial P&L contains documentation text or inactive BOOK | FAIL; accept exactly one row for each active SETUP book |
| VAL-022 | Non-Month-Ahead trade contains a delivery slice whose normal fixing precedes Trade Date | FAIL the complete source row; do not catch up or omit the slice |
| VAL-023 | Month Ahead trade has no fixing date on or after Trade Date | FAIL; do not fix on Trade Date or use Trade Price |

## 7. Calendar and effective-date tests

| ID | Case | Required result |
|---|---|---|
| CAL-001 | Calendar contains market and non-market days | Market Date axis contains only configured market days in the inclusive model range |
| CAL-002 | Request previous market day | Return the nearest earlier configured market day, not simply date minus one |
| CAL-003 | Economic event occurs on weekend/holiday | Apply once to the first output Market Date on or after its economic date |
| CAL-004 | Deferred event lies after Historical End Date | Do not apply it; report it as unapplied |
| CAL-005 | Event date equals an output Market Date | Apply on that date |
| CAL-006 | Duplicate or non-monotonic calendar rows | FAIL Preflight |
| CAL-007 | Trade occurs on a non-market date | Retain Trade Date; apply once on first output Market Date on or after it |

## 8. Fixings tests

| ID | Case | Required result |
|---|---|---|
| FIX-001 | Initial exposure has fixing date equal to Initial Market Date | Exclude that fixing; initial eligibility is strictly later than cut-off |
| FIX-002 | Actual trade has same-day eligible fixing | Include it; incremental eligibility is on or after Trade Date |
| FIX-003 | WITHINDAY product | Fix on every calendar Delivery Day, including non-market dates |
| FIX-004 | DAY_AHEAD across ordinary weekday | Fix on Delivery Day minus one calendar day |
| FIX-005 | DAY_AHEAD across weekend/holiday | Still use Delivery Day minus one calendar day; do not roll to previous Market Day |
| FIX-006 | HEREN product | Use previous configured market day |
| FIX-007 | MONTH_AHEAD with complete preceding month | Allocate scheduled volume across eligible fixing days and conserve total delivery volume |
| FIX-008 | MONTH_AHEAD when part of schedule predates trade | Redistribute the whole affected quantity equally over remaining Market Days on/after Trade Date |
| FIX-009 | MONTH_AHEAD with no remaining eligible fixing day | FAIL Preflight; no catch-up fixing or economic output |
| FIX-010 | BRENT/HH current or opening month | Same-month Exposure/opening MtM uses next Delivery Month's curve; later months use matching curve |
| FIX-011 | Delivery spans month boundary | Split into delivery-month rows without volume loss or duplication |
| FIX-012 | SELL exposure | Fixing volume has the opposite sign to exposure volume |
| FIX-013 | Known fixing price | `fixing_amount = fixing_volume * fixing_price` |
| FIX-014 | Required fixing price missing | FAIL with exact missing-price details; never write a plausible zero amount |
| FIX-015 | Mixed initial, actual, and simulation sources | Preserve source and scenario in output grain |
| FIX-016 | Simulation OFF/ON | OFF excludes every simulation row; ON includes only the requested scenario population |
| FIX-017 | Allocation conservation | Sum of fixing volume per economic lot equals the volume scheduled to fix within the modeled horizon |
| FIX-018 | Duplicate economic fixing key | FAIL or aggregate only under an explicit documented rule |
| FIX-019 | Brent current-month case around a roll boundary | Resolve to the explicitly expected next-month curve column and price |
| FIX-020 | HH current-month case around a roll boundary | Resolve to the explicitly expected next-month curve column and price |
| FIX-021 | Active product contains `DA` in its name but SETUP maps HEREN | Use HEREN; never infer fixing method from product name |
| FIX-022 | Each active generic DAY_AHEAD PVB product crosses a weekend | Use the product's explicit SETUP method and approved weekend price fixture |

## 9. Exposure tests

| ID | Case | Required result |
|---|---|---|
| EXP-001 | Initial position only | First modeled snapshot carries the opening signed volume by key |
| EXP-002 | BUY trade starts delivery | Add signed delivered volume on the effective date |
| EXP-003 | SELL trade starts delivery | Subtract signed delivered volume on the effective date |
| EXP-004 | Fixing event closes volume | Apply opposite-signed fixing volume exactly once |
| EXP-005 | Weekend/holiday event | Defer to first later Market Date and apply once |
| EXP-006 | Position reaches zero | Emit one explicit closure row |
| EXP-007 | Position remains zero | Omit later repeated zero rows |
| EXP-008 | Closed position reopens | Emit a new row from reopening date onward |
| EXP-009 | Simulation OFF/ON | Exposure population changes only by included simulation events |
| EXP-010 | Curve price available | `exposure_mtm = exposure_volume * curve_price` |
| EXP-011 | Curve price missing | FAIL with exact missing-price details; no silent zero valuation |
| EXP-012 | New economic key | Previous exposure is zero, with no dependency on output row adjacency |
| EXP-013 | Multiple events same day/key | Aggregate deterministically and apply each source event once |
| EXP-014 | Roll-forward identity | Closing equals prior closing plus all effective trade/delivery events plus fixing events, using their declared signs |
| EXP-015 | Output uniqueness | At most one row per declared Exposure grain |

## 10. P&L tests

| ID | Case | Required result |
|---|---|---|
| PNL-001 | Existing exposure reprices | Gross exposure delta equals current MtM minus previous Market Date MtM for the same key |
| PNL-002 | New key appears | Previous MtM is zero; lookup is by economic key, not previous physical row |
| PNL-003 | Trade-entry economics | Adjustment matches the confirmed execution-price versus market-price formula |
| PNL-004 | Same-day trade and normal fixing | Apply the trade and valid same-day fixing once; adjustment uses only volume remaining afterward |
| PNL-005 | Fixing amount | Raw Fixings output remains `volume * price`; P&L recognizes its inverse economic sign once on the correct date |
| PNL-006 | Operating flows | Use `-COSTS!B:N`, `COSTS!O:P + Foto FO!R:S`, and `Foto FO!T:V` as daily BOOK-level flows |
| PNL-007 | Component summation | Total P&L equals the sum of declared components for every row and aggregate |
| PNL-008 | Gross versus adjusted exposure delta | Both remain separately auditable; trade-entry adjustment is not hidden in gross movement |
| PNL-009 | Initial P&L bridge | 13 active-book balances remain separate from daily P&L and seed cumulative BOOK P&L |
| PNL-010 | Simulation OFF/ON | OFF contains no simulation economics; ON includes only selected scenarios |
| PNL-011 | No legacy operating fallback | Use only normalized operating inputs present in the bundle; manifest counts expose their population and no legacy value is substituted |
| PNL-012 | Idempotent rerun | Same inputs and policy produce byte-equivalent normalized P&L output |
| PNL-013 | Cross-layer key | Every Fixings/Exposure reference used by P&L resolves to the same Build ID |
| PNL-014 | Aggregate reconciliation | Sum by date, book, underlying, component, source, and scenario matches independently prepared totals |
| PNL-015 | Logistics sign | Approved Ops fixture proves a stored positive cost reduces P&L under the `-COSTS` transform |

## 11. Manifest, lineage, and publication tests

| ID | Case | Required result |
|---|---|---|
| MAN-001 | Successful build | Manifest contains policy/code versions, timestamps, input/output hashes, counts, status, and parent lineage |
| MAN-002 | Cross-layer load | Fixings, Exposure, P&L, and validation report share one Build ID and matching fingerprints |
| MAN-003 | Input changes by one value | Input fingerprint and Build ID change |
| MAN-004 | Identical rerun | Economic output and normalized file hashes remain identical; run timestamp may differ only in run metadata |
| MAN-005 | Blocking validation failure | Status `FAILED`; no calculated output is published |
| MAN-006 | Calculation invariant failure | Status `FAILED`; staged diagnostics retained; last published build unchanged |
| MAN-007 | Interrupted publication | Atomic mechanism prevents a partially replaced output set |
| MAN-008 | Unapplied deferred events | Manifest reports count and IDs; production status follows policy |

## 12. Property and metamorphic tests

The suite SHALL generate many small randomized portfolios and check these properties:

1. splitting a trade into identical sub-trades does not change aggregate economics;
2. permuting input row order does not change normalized output;
3. changing a source row ID alone does not change aggregate economics;
4. a BUY and equal SELL with otherwise identical terms net to zero aggregate exposure while both remain traceable;
5. total allocated fixing volume is conserved under the method's confirmed eligibility rules;
6. every economic event is applied zero or one times, never more than once;
7. Exposure roll-forward identities hold for every key and Market Date;
8. Total P&L equals the sum of its components;
9. Simulation OFF output is unaffected by changes confined to simulation input;
10. identical inputs and policies produce identical normalized outputs;
11. increasing an applicable curve price changes MtM by `volume * price_change`, absent other changes;
12. no required missing value becomes numeric zero through parsing or serialization.

Randomized tests supplement but do not replace golden cases.

## 13. Golden economic case catalogue

Each case SHALL use the smallest possible dataset and include a manual explanation in `notes.md`.

| ID | Case | Approval requirement |
|---|---|---|
| GOLD-001 | Initial exposure only, no trades | User approves daily Exposure and P&L path |
| GOLD-002 | Single BUY, one delivery day | User approves signs, dates, fixing, and P&L |
| GOLD-003 | Single SELL, one delivery day | User approves signs, dates, fixing, and P&L |
| GOLD-004 | Same-day known trade/fixing | Valid `Fixing Date = Trade Date` outcome independently calculated and approved |
| GOLD-005 | Weekend/holiday DAY_AHEAD versus HEREN | Previous-calendar versus previous-market behavior independently calculated and approved |
| GOLD-006 | MONTH_AHEAD trade entered after schedule begins and after schedule ends | Remaining-day redistribution and blocking late-trade result independently calculated and approved |
| GOLD-007 | BRENT/HH opening/current-month case | Next-month curve behavior independently calculated and approved |
| GOLD-008 | Simulation OFF versus ON | User approves scope and scenario totals |
| GOLD-009 | Position closes to zero and later reopens | One closure row/no repeated zeros behavior independently approved |
| GOLD-010 | Missing required fixing or curve price | Confirmed D-001 behavior: fail, list missing prices, and preserve the last published build |
| GOLD-011 | Trade-entry adjustment | User approves formula and expected amount |
| GOLD-012 | Operating flows | Ops approves source extraction and Logistics sign using a known day/book |
| GOLD-013 | Initial P&L bridge | 13-book opening balance and cumulative roll-forward independently approved |
| GOLD-014 | Known current-workbook position that disappears in legacy output | User approves expected new-engine continuity |
| GOLD-015 | Known current-workbook P&L key absent from Model Control count | User approves expected complete key population |
| GOLD-016 | Brent roll-boundary regression | Market Date, Delivery Month, expected curve column, and expected price independently approved |
| GOLD-017 | HH roll-boundary regression | Market Date, Delivery Month, expected curve column, and expected price independently approved |

The suite SHALL preserve the workbook-derived cases as regression fixtures after their source rows are anonymized or reduced without changing the economics.

## 14. Python-to-VBA parity tests

For every approved fixture, the VBA implementation SHALL match Python on:

- validation PASS/FAIL state, stable codes, and affected source IDs;
- output field names and types;
- exact dates, identifiers, categories, keys, row counts, and row order;
- numeric values within the Section 4 tolerance;
- Build ID inputs, lineage, and source/output counts;
- omitted versus explicit-zero rows;
- scenario and Trade Source separation.

Parity failures SHALL be diagnosed against the specification and fixture. Python does not override a proven specification error; instead, both implementation and expected result are corrected under change control.

## 15. Excel-adapter tests

| ID | Case | Required result |
|---|---|---|
| XLS-001 | Import a passed output bundle | Excel displays the exact Build ID, status, counts, and normalized rows |
| XLS-002 | Existing tables are filtered | Adapter clears/handles filters safely before replacement |
| XLS-003 | Output table contains old extra rows | Successful import removes stale rows without damaging headers/formats |
| XLS-004 | Bundle fails coherence/hash check | Import stops; existing published tables remain unchanged |
| XLS-005 | Engine run fails | No partial table replacement and no misleading PASS status |
| XLS-006 | Excel events/calculation/screen updating disabled during import | All application settings are restored on success and failure |
| XLS-007 | Wrong workbook is active | Adapter uses explicit workbook references and cannot write to it |
| XLS-008 | Links unavailable | Previously prepared bundle can still be imported without link refresh |
| XLS-009 | Headless calculation request | Core completes without opening Excel |
| XLS-010 | User cancels or process is interrupted | Last published workbook state remains recoverable and coherent |
| XLS-011 | Validation display | Only the selected/current Build ID determines status; historical rows remain queryable |
| XLS-012 | Source workbook protection | Original `.xlsm` hash remains unchanged during test execution |

## 16. Performance and reproducibility gates

On the current Mac and a production-sized snapshot:

- a full production rebuild completes in less than 60 seconds;
- a normal daily/incremental update targets less than 10 seconds;
- the focused unit suite completes in less than 2 seconds;
- peak Python resident memory remains below 2 GiB;
- every run records elapsed time and peak resident memory as an audit baseline;
- later candidate releases identify and explain material runtime or memory regressions;
- two clean runs produce identical normalized economic outputs and hashes.

The performance fixture SHALL record machine, Python, dependency, Excel/VBA, and operating-system versions. A performance failure does not permit calculation shortcuts that weaken correctness or auditability.

## 17. Decision-to-test traceability

| Resolution | Tests enforcing it |
|---|---|
| D-001 Missing-price policy — confirmed | VAL-010, FIX-014, EXP-011, GOLD-010 |
| D-002 Brent/HH next-month valuation — resolved | FIX-010, FIX-019, FIX-020, GOLD-007, GOLD-016, GOLD-017 |
| D-003 SETUP-driven DAY_AHEAD/HEREN weekend behavior — product confirmation pending | FIX-004, FIX-005, FIX-021, FIX-022, GOLD-005 |
| D-004 Same-day valid; genuinely late trade rejected | VAL-022, VAL-023, FIX-002, FIX-008, FIX-009, PNL-004, GOLD-004, GOLD-006 |
| D-005 Direct BOOK-level operating flows — resolved; Ops sign approval pending | PNL-006, PNL-011, PNL-015, GOLD-012 |
| D-006 One explicit closure row — resolved | EXP-006 to EXP-008, GOLD-009 |
| D-007 Precision/tolerance — resolved | Every numeric assertion and VBA parity |
| D-008 SETUP/canonical PVB mappings — resolved | VAL-005, VAL-007, VAL-008 and output category parity |
| D-009 First Market Date on/after event — resolved | CAL-003, CAL-005, CAL-007 |
| D-010 Approved performance targets | Section 16 gates |
| D-011 13-book Initial P&L bridge — resolved | VAL-021, PNL-009, GOLD-013 |
| D-012 Side supplies sign — resolved | VAL-018, BUY/SELL volume cases |

## 18. Release gates

### Gate A: Contract frozen

- input/output schemas versioned;
- the methodology resolutions D-001 through D-012 are versioned;
- golden-case owners and approval method recorded.

### Gate B: Python component acceptance

- all applicable schema, Preflight, calendar, Fixings, Exposure, and P&L tests pass;
- no skipped blocking test;
- coverage includes every conditional business-rule branch;
- type checking and static checks pass.

### Gate C: Python end-to-end acceptance

- all approved golden cases pass;
- property tests pass;
- manifest/coherence/publication tests pass;
- deterministic rerun and performance gates pass.

### Gate D: Excel adapter acceptance

- all XLS tests pass on a disposable workbook copy;
- source workbook hash is unchanged;
- a business reviewer validates displayed results against the output bundle.

### Gate E: VBA acceptance, if required

- all applicable Python fixtures run through VBA;
- parity requirements pass;
- no core calculation function reads worksheets or global Excel state;
- runtime failure restores Excel state and preserves the last published build.

Production Total P&L SHALL NOT pass Gate C until GOLD-012/PNL-015 certify the operating-flow extraction and Logistics sign. A narrower Fixings/Exposure release MAY pass earlier if its manifest declares operating components `NOT_IMPLEMENTED` and it does not present an authoritative Total P&L.

## 19. Evidence bundle

Every candidate release SHALL retain:

- code commit and dependency lock hash;
- policy and schema versions;
- test command and environment versions;
- machine-readable test report;
- failing-case diagnostics, if any;
- coverage and property-test seeds;
- performance measurements;
- input/expected/actual hashes for golden cases;
- Python/VBA parity report when applicable;
- Excel source and runtime-copy hashes;
- named reviewer approvals.

## 20. Change control

Each business-rule change SHALL:

1. update the engine specification and decision status;
2. add or change a focused test that would fail under the previous behavior;
3. update affected golden expected results only with documented business approval;
4. increment the policy version when economics change;
5. preserve historical fixtures and manifests needed to reproduce prior published builds.
