# Gas Trading Model Engine Specification

Version: 0.3-approved-build-baseline
Policy version: 0.3.0
Status: Approved Python reference baseline; core golden pack received, product/workbook-specific acceptance remains pending
Date: 2026-08-07
Target implementations: Python reference engine, followed by a clean VBA/Excel implementation

## 1. Purpose

This specification defines the calculation, validation, audit, and publication behavior of the Gas Trading Model (GTM) engine. It replaces implicit workbook behavior with explicit rules that can drive implementation and automated tests.

The engine reconstructs the daily economic state of a gas trading portfolio from an end-of-day initial position, subsequent trades, fixing schedules, market prices, and operating flows. It produces normalized Fixings, Exposure, and P&L outputs by book, underlying, delivery month, trade source, and scenario.

## 2. Requirement language and status

The words `SHALL`, `SHALL NOT`, `SHOULD`, and `MAY` are normative.

Requirements carry one of four statuses:

- **CONFIRMED**: supported by direct user instruction or consistent local methodology evidence.
- **GOLDEN TEST REQUIRED**: the rule is specified, but a small independent economic result must be approved before production release.
- **ENGINEERING TARGET**: a measurable non-economic release objective.
- **REJECTED LEGACY**: observed legacy behavior that the new engine must not reproduce.

## 3. Sources and authority

When sources disagree, use this order:

1. direct user decisions recorded against this specification;
2. approved golden economic test cases;
3. this specification's confirmed requirements;
4. the editable fixing-method definitions and active mappings on `SETUP`;
5. rules implemented consistently in the complete exposure/fixing builder and `GTM_Fast_Helper.xlsm`;
6. `GTM_v2_Technical_Handover_Report_2026-08-07.docx` and normalized v2 architectural intent;
7. current v2 code, except where it is explicitly incomplete;
8. legacy formulas and reconciliation output, for corroboration or diagnosis only.

`questions_answers.txt` and its backup are excluded from the methodology evidence at the user's instruction.

Legacy reconciliation SHALL NOT determine acceptance because extensive legacy formulas stop at TRADES row 202 while the current trade population extends far beyond that row.

## 4. Product scope

### 4.1 Required deliverables

The project SHALL produce:

1. a headless Python reference engine;
2. versioned normalized input and output schemas;
3. deterministic Preflight, Fixings, Exposure, and P&L calculations;
4. an immutable Build Manifest and detailed validation report;
5. an automated acceptance-test suite;
6. a thin Excel import/export and user-interface layer;
7. a clean VBA implementation that matches the Python reference if self-contained Excel deployment remains required;
8. migration documentation and an audit trail from the existing workbook to the new engine.

### 4.2 Calculation-core boundary

The calculation core SHALL:

- accept a complete, immutable input bundle;
- validate the bundle before calculation;
- perform no UI operations;
- perform no SharePoint, Reuters, or external-workbook refresh;
- produce outputs in a staging run directory;
- publish outputs only after all blocking checks pass.

External data retrieval belongs to a separate acquisition layer. The engine consumes snapshots and records their hashes and as-of timestamps.

### 4.3 Initial exclusions

The first engine release SHALL NOT:

- repair or extend the formula-heavy legacy calculation sheets;
- treat legacy output as authoritative truth;
- save or rewrite the source `.xlsm` file;
- depend on modal dialogs, ActiveWorkbook, selected cells, worksheet filters, or Excel calculation state;
- silently substitute zero for missing required prices or invalid numeric inputs.

## 5. Architecture

### 5.1 Logical components

The Python reference implementation SHOULD use these logical components:

```text
contracts       Input/output schemas and enumerations
canonicalize    Text, date, key, and mapping normalization
validation      Schema, policy, completeness, and coherence checks
calendar        Market-day and event-effective-date functions
fixings         Fixing schedules, volumes, prices, and amounts
exposure        Event ledger and daily closing exposure snapshots
pnl             Valuation, trade-entry economics, and P&L components
manifest        Build identity, hashes, counts, versions, and lineage
pipeline        Ordered execution, stop gates, and atomic publication
io              Read-only extraction and staged output serialization
```

The later VBA implementation SHALL preserve the same separation. Core modules SHALL accept arrays/records and SHALL NOT read or write worksheets directly.

### 5.2 Pipeline order

The production pipeline SHALL run in this order:

1. load and hash the input bundle;
2. canonicalize identifiers and dates;
3. run Preflight;
4. stop if any blocking error exists;
5. create the Market Date axis;
6. build Fixings;
7. validate Fixings invariants;
8. build Exposure;
9. validate Exposure invariants;
10. build P&L, including trade-entry economics;
11. validate P&L and cross-layer invariants;
12. write the Build Manifest and staged outputs;
13. publish atomically only when the build status is `VERIFIED`, producing `PUBLISHED`.

### 5.3 Build states

Each run SHALL occupy one state:

```text
CREATED -> VALIDATED -> CALCULATED -> VERIFIED -> PUBLISHED
    |          |             |            |
    +----------+-------------+------------+-> FAILED
```

A failed build SHALL NOT overwrite the last published build.

## 6. Common types and conventions

### 6.1 Dates

- Business dates SHALL contain no time component.
- `Initial Market Date`, `Trade Date`, `Start Date`, `End Date`, `Delivery Day`, `Fixing Date`, `Market Date`, and `Previous Market Date` SHALL use ISO `YYYY-MM-DD` in serialized files.
- `Delivery Month` SHALL be stored as the first calendar day of the month.
- Build timestamps SHALL be timezone-aware. The current model timezone is `Europe/Madrid`.

### 6.2 Identifiers

- `BOOK`, `Underlying`, `Trade Source`, and `Scenario` SHALL be trimmed.
- Canonical comparison SHALL be case-insensitive.
- Output SHALL use the configured canonical spelling.
- Mappings SHALL be explicit data, not hard-coded branches hidden in calculation functions.
- Source row identity SHALL be preserved through `source_file`, `source_table`, and `source_row_id` fields in audit records.

### 6.3 Numeric values

- Volume units SHALL follow the configured underlying: MWh, bbl, or MMBtu in the current model.
- Price units SHALL follow the configured contract; currency/unit conversion SHALL be explicit where required.
- P&L and fixing amounts are EUR in the current reporting contract.
- The calculation core SHALL NOT perform an implicit currency or unit conversion. The extraction
  layer SHALL normalize prices to each underlying's configured reporting currency and unit;
  required prices with incompatible metadata SHALL block the build.
- The engine SHALL preserve full calculation precision and SHALL NOT round intermediate values for display.
- Display formatting SHALL NOT alter stored numeric values.
- Values with absolute magnitude at or below `1e-7` MAY be treated as economically zero for event/row emission.
- Volume comparisons SHALL use an absolute tolerance of `0.000001` in the configured volume unit.
- Price comparisons SHALL use an absolute tolerance of `0.00000001` in the configured price unit.
- P&L and monetary-amount comparisons SHALL use an absolute tolerance of `EUR 0.01`.
- Identifiers, dates, categories, booleans, key sets, and row counts SHALL compare exactly.

### 6.4 Enumerations

`Trade Source` SHALL be one of:

- `INITIAL`
- `ACTUAL`
- `SIMULATION`

`Side` SHALL be one of:

- `BUY`
- `SELL`

`Simulation Status` SHALL be one of:

- `ON`
- `OFF`

`Fixing Method` SHALL use canonical configured values:

- `WITHINDAY`
- `DAY_AHEAD`
- `HEREN`
- `MONTH_AHEAD`
- `BRENT_HH`

Aliases from the workbook SHALL be resolved during canonicalization.

## 7. Input bundle

### 7.1 Bundle configuration and files

Every run SHALL receive `bundle.json` with:

- `model_id`
- `schema_version`
- `policy_version`
- `initial_market_date`
- `historical_start_date`
- `historical_end_date`
- `simulation_status`
- `timezone`
- engine version

Required date relationship:

```text
historical_start_date <= historical_end_date
initial_market_date < historical_end_date
```

The bundle directory SHALL also contain `books.csv`, `underlyings.csv`,
`market_calendar.csv`, `initial_exposure.csv`, and `initial_pnl.csv`. `trades.csv`,
`curve_prices.csv`, `fixing_prices.csv`, and `operating_flows.csv` MAY be physically absent when
no row from that source is required. The loader SHALL hash every present input file. Economically
required price keys remain mandatory under D-001 regardless of whether an entire optional price
file is absent.

### 7.2 Initial exposure

Required fields:

| Field | Type | Rule |
|---|---|---|
| initial_market_date | date | Must equal the bundle cut-off |
| book | string | Required, active, canonical |
| underlying | string | Required, active, canonical |
| delivery_month | date | First day of month |
| exposure_volume | decimal/float | Required, finite, signed |
| source_row_id | string | Required and unique within source |

The engine SHALL aggregate duplicate economic input rows only after reporting their presence. The aggregate key is `book + underlying + delivery_month`.

### 7.3 Initial P&L

The initial P&L source is a separate BOOK-level closing balance at Initial Market Date. The authoritative source population is `INITIAL POSITION!A7:B19`: one row for each of the 13 active books.

Required fields:

| Field | Type | Rule |
|---|---|---|
| initial_market_date | date | Must equal the bundle cut-off |
| book | string | Required, active, canonical, unique |
| amount | decimal/float | Required, finite |
| comment | string/null | Optional source explanation |
| source_row_id | string | Required |

The engine SHALL reject documentation labels or inactive books. The saved `tblInitialPnL` has six erroneous zero rows created when the refresh macro ingested methodology descriptions; these rows are not business data.

Initial P&L SHALL remain separate from opening Exposure MtM and from daily post-cut-off P&L. It is the opening balance used only when cumulative P&L is requested.

### 7.4 Trades and simulation trades

Required normalized fields:

| Field | Type | Rule |
|---|---|---|
| source_row_id | string | Required and unique within source |
| trade_date | date | Required |
| book | string | Required and active |
| underlying | string | Required and active |
| side | enum | BUY or SELL |
| start_date | date | Required |
| end_date | date | Required; must be on or after start_date |
| daily_qty | decimal/float | Required, finite |
| execution_price | decimal/float | Required for incremental economic trades |
| trade_source | enum | ACTUAL or SIMULATION |
| scenario | string | Required for SIMULATION; blank for ACTUAL |

Derived signed quantity:

```text
signed_daily_qty = daily_qty       for BUY
signed_daily_qty = -daily_qty      for SELL
```

The engine SHALL reject negative `daily_qty`; sign belongs to `side`. A deliberate numeric zero is allowed with a warning and generates no economic event.

### 7.5 Setup and mappings

Required tables:

- active books;
- active underlyings;
- fixing method by underlying;
- canonical aliases;
- underlying price-curve mapping;
- fixing-price lookup mapping;
- optional book or pseudo-book mapping.

Every active underlying used by an eligible position or trade SHALL resolve to exactly one fixing method and required price mappings.

The reviewed active mapping export is `docs/GTM_ACTIVE_SETUP_MAPPING_v0.3.csv`. It separates the
source name, canonical exposure name, curve name, fixing-price name, and price-date basis so that
PVB aggregation does not erase product-specific fixing behavior.

### 7.6 Market calendar

Required fields:

| Field | Type | Rule |
|---|---|---|
| date | date | Unique |
| is_market_day | boolean | Required |
| calendar_id | string | Required if more than one calendar is supported |

The calendar SHALL cover all dates needed for delivery, fixing, previous-market-day lookup, deferred events, and historical output.

### 7.7 Curve prices

Required normalized fields:

| Field | Type | Rule |
|---|---|---|
| market_date | date | Required |
| underlying | string | Canonical valuation underlying |
| delivery_month | date | First day of month |
| curve_price | decimal/float | Finite |
| currency | string | Required |
| unit | string | Required |
| source_id | string | Required |
| source_as_of | timestamp/null | Source timestamp when available |

Economic key: `market_date + underlying + delivery_month`.

### 7.8 Fixing prices

Required normalized fields:

| Field | Type | Rule |
|---|---|---|
| price_lookup_date | date | Date used to retrieve the price |
| underlying | string | Canonical pricing underlying |
| fixing_price | decimal/float | Finite |
| currency | string | Required |
| unit | string | Required |
| source_id | string | Required |
| source_as_of | timestamp/null | Source timestamp when available |

Economic key: `price_lookup_date + underlying`. Fixing output and event lineage preserve Fixing
Date, Delivery Day, and Delivery Month. Their relationship to `price_lookup_date` remains explicit.

### 7.9 Operating flows

The target engine SHALL consume direct normalized operating flows rather than copying legacy P&L output.

Required components and workbook extraction rules:

- `LOGISTICAL_COSTS`: `-COSTS!B:N`, matched by daily Market Date and BOOK header;
- `FEES_AND_OPTIMIZATIONS`: `COSTS!O:P + Foto FO!R:S`, matched by daily Market Date and BOOK header and added with stored signs;
- `REPLICATION`: `Foto FO!T:V`, matched by daily Market Date and BOOK header and added with stored signs.

Operating inputs are daily flows and SHALL NOT be differenced again. They SHALL remain at BOOK level with Underlying `TOTAL / BOOK LEVEL`, blank Delivery Month, and no invented allocation to Underlying, Trade Source, Scenario, or delivery period.

The Logistics sign rule requires one approved Ops golden case before production Total P&L release because `COSTS!R4` itself requests sign confirmation.

## 8. Preflight requirements

### 8.1 Fail-closed behavior

- Preflight SHALL return a structured result object; it SHALL NOT rely on a pre-existing worksheet status cell.
- Any unexpected exception SHALL create a blocking validation item with the original exception type, message, and stage.
- A failed or interrupted preflight SHALL produce status `FAIL`.
- The pipeline SHALL stop before calculation when status is `FAIL`.
- Preflight SHALL NOT clear evidence from a previous run. Each run writes a separate immutable validation artifact.

### 8.2 Blocking validations

Preflight SHALL block on:

- missing required input table or column;
- duplicate required identifier;
- invalid or incomplete date relationship;
- populated trade row with blank or non-numeric Daily Qty;
- invalid BUY/SELL value;
- End Date before Start Date;
- missing execution price for an eligible incremental trade;
- unknown or inactive book;
- unknown or inactive underlying;
- missing or ambiguous fixing method;
- missing calendar coverage;
- duplicate curve/fixing-price key;
- incoherent Simulation state;
- missing required price for any included economic row (`D-001`, CONFIRMED);
- an incremental non-Month-Ahead trade whose complete quantity cannot be assigned to normal fixing opportunities on or after Trade Date;
- an incremental Month Ahead trade for which no fixing date on or after Trade Date remains;
- upstream schema or policy version incompatibility.

### 8.3 Warnings and information

- A deliberate numeric zero Daily Qty is allowed with a warning and creates no economic event.
- Unused future prices or calendar dates MAY be absent if they are outside every required economic lookup.
- Warnings SHALL include an affected row/key count and representative source locations.
- A warning SHALL never cause zero-value substitution.

## 9. Calendar and event-effective-date rules

### 9.1 Market Date axis

- The engine SHALL derive the Market Date axis from the configured calendar and historical range.
- Exposure snapshots SHALL use market days strictly after Initial Market Date through Historical End Date, inclusive.
- `Previous Market Date` SHALL be the preceding configured market day.
- The axis SHALL NOT be copied from a legacy output sheet.

### 9.2 Deferred events

- An event with economic date `d` SHALL apply to the first output Market Date `m` for which `m >= d`.
- Each event SHALL apply exactly once.
- The event ledger SHALL preserve both `economic_date` and `applied_market_date`.
- Events beyond Historical End Date SHALL remain unapplied and SHALL be reported, not discarded.

Trade events on non-market dates follow the same due-event rule: retain Trade Date and apply once on the first output Market Date on or after it.

## 10. Initial-position and trade eligibility

### 10.1 Initial position

- Initial Market Date represents an end-of-day state.
- Initial exposure already includes all activity through and including Initial Market Date.
- Initial-position fixing events SHALL be eligible only when Fixing Date is strictly after Initial Market Date.

### 10.2 Incremental trades

- Only trades with Trade Date strictly after Initial Market Date SHALL enter the incremental calculation.
- ACTUAL trades SHALL always be included.
- SIMULATION trades SHALL be included only when Simulation Status is `ON`.
- Scenario SHALL be preserved for simulation output.
- Incremental trade fixing eligibility SHALL include Fixing Date equal to Trade Date.

### 10.3 Delivery periods

- Start Date and End Date are inclusive.
- A trade spanning more than one delivery month SHALL be split by Delivery Month.
- Daily quantity SHALL apply to each eligible delivery day in the inclusive range unless the contract type specifies another profile.
- Support for non-flat delivery profiles requires a future schema extension and SHALL NOT be inferred from the current workbook.

## 11. Fixing-method rules

### 11.1 WITHINDAY

Confirmed rule:

```text
fixing_date = delivery_day
price_lookup_date = delivery_day
```

### 11.2 DAY_AHEAD

Confirmed rule:

```text
fixing_date = delivery_day - 1 calendar day
price_lookup_date = fixing_date
```

Every calendar delivery day is eligible. Weekend and holiday dates are not rolled back to the previous Market Day; that behavior belongs to HEREN.

Product names SHALL NOT determine methodology. Active `SETUP` mapping is authoritative. In the current mapping, `D+1 Auction`, `Mibgas Index ES`, and `MIBGAS D+1 Daily Reference` use DAY_AHEAD; `TTF DA` uses HEREN despite the `DA` text in its name. Product-level weekend price fixtures remain a golden-test requirement.

### 11.3 HEREN

Confirmed rule:

```text
fixing_date = previous_market_day(delivery_day)
```

The price lookup may use Delivery Day; the recognition date remains Fixing Date.

Every calendar delivery day is eligible. The previous Market Day fixes the next Market Day and all intervening calendar delivery days.

### 11.4 BRENT_HH

Confirmed source behavior:

- only eligible market delivery days contribute;
- Fixing Date is the previous market day of Delivery Day;
- curve valuation and fixing-price selection use different rules;
- if Market Date and Delivery Month are the same month, Exposure and opening Initial MtM use the next Delivery Month's Brent/HH curve;
- later Delivery Months use the matching Delivery Month curve.

### 11.5 MONTH_AHEAD

Confirmed rule:

- the fixing window is the set of market days in the calendar month before Delivery Month;
- total delivery volume for the month is distributed equally across eligible fixing days;
- initial positions use fixing days strictly after Initial Market Date;
- incremental trades use remaining fixing days on or after Trade Date and redistribute the whole affected monthly delivery quantity equally over those remaining dates;
- if no eligible fixing date remains, Preflight SHALL reject the trade as a blocking validation error.

### 11.6 Late-trade rejection

For an incremental trade, a normal fixing on Trade Date is valid:

```text
normal_fixing_date >= trade_date
```

The engine SHALL NOT invent a catch-up fixing on Trade Date, use Trade Price as Fixing Price, carry a past fixing forward, or silently omit an overdue delivery slice.

For WITHINDAY, DAY_AHEAD, HEREN, and BRENT_HH, Preflight SHALL reject the complete source trade row when any included delivery slice has `normal_fixing_date < trade_date`; this conservative release-1 rule preserves the row's full economic quantity. For MONTH_AHEAD, the complete affected monthly quantity MAY be redistributed over valid remaining fixing dates on or after Trade Date. If none remains, Preflight SHALL reject the trade.

Initial-position rows do not use late-trade validation because the initial position is already the closing state at its cut-off.

### 11.7 Fixing volume and amount

- Fixing Volume SHALL use the opposite sign from the open exposure it closes.
- For a valid fixing-price lookup:

```text
fixing_amount = fixing_volume * fixing_price
```

- The engine SHALL preserve full precision.
- A missing required fixing price SHALL block the build under Decision D-001 and SHALL never silently create a zero amount.

## 12. Fixings output

Required fields:

| Field | Type |
|---|---|
| fixing_date | date |
| applied_market_date | date |
| price_lookup_date | date |
| book | string |
| source_underlying | string |
| underlying | string |
| pricing_underlying | string |
| delivery_month | date |
| delivery_day | date/null |
| fixing_volume | numeric |
| fixing_price | numeric |
| fixing_amount | numeric |
| trade_source | enum |
| scenario | string/null |
| simulation_status | enum |
| build_id | string |

Economic key:

```text
fixing_date + applied_market_date + price_lookup_date + book + source_underlying
+ underlying + pricing_underlying + delivery_month + delivery_day + trade_source + scenario
```

The engine SHALL aggregate source events to one row per economic key and SHALL retain source-to-output lineage in a separate audit table.

## 13. Exposure calculation

### 13.1 Event ledger

The engine SHALL construct an event ledger containing:

- initial-position opening events;
- incremental trade events;
- fixing/closure events;
- economic date;
- applied Market Date;
- signed volume change;
- book, underlying, Delivery Month, Trade Source, and Scenario;
- source lineage.

### 13.2 Roll-forward

For each economic exposure key and Market Date:

```text
closing_exposure[m] = closing_exposure[previous_m]
                    + trade_events_applied[m]
                    + fixing_events_applied[m]
```

Initial-position exposure forms the opening state at Initial Market Date.

### 13.3 Explicit zero closures

When a previously non-zero position becomes zero, the engine SHALL emit an explicit zero snapshot on the closure Market Date.

The engine SHALL emit that one closure row, omit repeated later zero rows, and resume emitting the key if it reopens.

### 13.4 Exposure valuation

For every material open output row:

```text
exposure_mtm = exposure_volume * curve_price
```

An explicit zero-closure row has `exposure_volume = 0`, `exposure_mtm = 0`, and
`curve_price = null`. A price is not required solely to display that zero closure.

The normal price key is `Market Date + Underlying + Delivery Month`. Underlying-specific mappings SHALL be explicit. For Brent Dated and HH, a same-month Delivery Month resolves to the next Delivery Month's curve contract; later months resolve to the matching contract.

## 14. Exposure output

Required fields:

| Field | Type |
|---|---|
| market_date | date |
| previous_market_date | date |
| book | string |
| underlying | string |
| delivery_month | date |
| curve_delivery_month | date |
| exposure_volume | numeric |
| curve_price | numeric/null for an explicit zero closure |
| exposure_mtm | numeric |
| trade_source | enum/null for BOOK-level operating flows |
| scenario | string/null |
| simulation_status | enum |
| is_explicit_closure | boolean |
| build_id | string |

Economic key:

```text
market_date + book + underlying + delivery_month + trade_source + scenario
```

`Previous Market Date` is metadata and SHALL NOT be part of the economic key.

## 15. P&L calculation

### 15.1 Output grain

P&L SHALL preserve:

- Market Date;
- Previous Market Date;
- BOOK;
- Underlying;
- Delivery Month;
- Trade Source;
- Scenario.

### 15.2 Exposure movement

Base exposure movement:

```text
gross_delta_exposure_mtm = current_exposure_mtm - previous_exposure_mtm
```

If the prior economic key is absent, prior Exposure MtM is zero unless an initial-position bridge supplies an opening amount.

### 15.3 Trade-entry adjustment

Trade-entry economics SHALL be integrated into the P&L engine, not applied by an unguarded post-processing helper.

Observed formula:

```text
trade_entry_adjustment = -open_volume_at_trade_entry * execution_price
```

Confirmed adjusted exposure component:

```text
delta_exposure_mtm = gross_delta_exposure_mtm + trade_entry_adjustment
```

The engine SHALL calculate the adjustment once per Build ID. Rerunning the same build SHALL produce the same output rather than adding to saved output.

For trade-entry adjustment, `open_volume_at_trade_entry` is the volume remaining after any valid normal fixing on Trade Date. This keeps same-day fixing volume from receiving both a trade-entry adjustment and a closing fixing treatment. A non-market Trade Date retains its economic date and is reported on the first Market Date on or after it.

### 15.4 Fixing P&L

- Fixings output SHALL preserve the signed raw settlement defined in Section 11:
  `raw_fixing_amount = fixing_volume * fixing_price`.
- Because Fixing Volume has the opposite sign from the exposure it closes, the economic P&L
  contribution SHALL be `pnl_fixing_amount = -raw_fixing_amount`.
- The economic Fixing Amount SHALL be recognized on the event's applied Market Date while retaining
  the original Fixing Date.
- P&L SHALL aggregate economic fixing contributions by the P&L economic key.
- A fixing event SHALL appear exactly once.

### 15.5 Operating components

The final engine SHALL calculate or import direct normalized amounts for:

- Logistical Costs;
- Fees and Optimizations;
- Replication.

It SHALL NOT copy these values from legacy `DAILY PNL DATA` in the production target architecture.

The source-to-component transforms are:

```text
logistical_costs      = - COSTS[B:N] for Market Date and BOOK
fees_and_optimizations = COSTS[O:P] + Foto FO[R:S]
replication            = Foto FO[T:V]
```

These values are daily flows, retained at BOOK level with Underlying `TOTAL / BOOK LEVEL` and blank Delivery Month. The engine SHALL NOT fabricate a finer allocation.

The engine SHALL import only operating flows supplied by the source bundle. It SHALL NOT manufacture simulated operating costs or allocate actual BOOK-level flows to simulation scenarios unless a future explicit simulation rule provides that data and allocation.

### 15.6 Total P&L

For every P&L row:

```text
total_pnl = delta_exposure_mtm
          + fixing_amount
          + logistical_costs
          + fees_and_optimizations
          + replication
```

No component may be hidden in Total P&L.

### 15.7 Initial and cumulative P&L

Initial P&L is a separate signed BOOK-level closing balance at Initial Market Date. It SHALL NOT be written as a post-cut-off daily flow and SHALL NOT be added to opening Exposure MtM.

When cumulative reporting is requested:

```text
cumulative_pnl[book, market_date]
    = initial_pnl[book]
    + sum(total_daily_pnl[book] for dates after cut-off through market_date)
```

Daily and cumulative outputs SHALL remain distinguishable.

## 16. P&L output

Required fields:

| Field | Type |
|---|---|
| market_date | date |
| previous_market_date | date |
| book | string |
| underlying | string |
| delivery_month | date |
| exposure_mtm | numeric |
| gross_delta_exposure_mtm | numeric |
| trade_entry_adjustment | numeric |
| delta_exposure_mtm | numeric |
| fixing_amount | numeric |
| logistical_costs | numeric |
| fees_and_optimizations | numeric |
| replication | numeric |
| total_pnl | numeric |
| trade_source | enum |
| scenario | string/null |
| simulation_status | enum |
| build_id | string |

Economic key:

```text
market_date + book + underlying + delivery_month + trade_source + scenario
```

The Excel adapter MAY combine `gross_delta_exposure_mtm` and `trade_entry_adjustment` into the existing `Delta Exposure MtM` presentation column, but the engine output SHALL retain both values for audit.

When cumulative reporting is enabled, the engine SHALL also write a BOOK-level table with `market_date`, `previous_market_date`, `book`, `initial_pnl`, `daily_pnl`, `cumulative_pnl`, `simulation_status`, and `build_id`. It SHALL not allocate Initial P&L to Underlying or Delivery Month without a new explicit policy.

## 17. Simulation requirements

- Simulation Status is a build-level input.
- `OFF` SHALL exclude every SIMULATION trade and output row.
- `ON` SHALL include eligible simulation trades.
- Scenario SHALL be required and preserved for every simulation row.
- ACTUAL and INITIAL rows SHALL have an empty Scenario.
- All Fixings, Exposure, and P&L outputs in one build SHALL carry the same Simulation Status.

## 18. Build Manifest and lineage

### 18.1 Build identity

Each execution SHALL receive a unique `run_id`. The `build_id` and
`calculation_fingerprint` SHALL be deterministic calculation identities derived from normalized,
canonically ordered input content including:

- schema version;
- policy version;
- engine version;
- Initial Market Date;
- historical range;
- Simulation Status;
- canonical mappings.

Raw input-file hashes SHALL be retained separately in the manifest. Two runs with the same
fingerprint SHALL have the same Build ID and identical economic outputs even if source rows were
presented in a different order.

### 18.2 Manifest content

The Build Manifest SHALL record:

- build ID and fingerprint;
- start/end timestamps and timezone;
- engine and schema versions;
- raw and normalized input hashes and input row counts;
- policy decisions and versions;
- validation counts by severity/code;
- output row counts and hashes;
- component totals;
- performance timings, runtime platform, Python version, architecture, and peak memory;
- status and failure stage;
- parent build or source workbook identity when applicable.

### 18.3 Coherence gate

The Excel adapter SHALL import a build only when:

- status is `PUBLISHED`;
- all output files reference the same Build ID and fingerprint;
- Initial Market Date, historical range, and Simulation Status match the manifest;
- output hashes match;
- no blocking validation exists.

## 19. Validation output

Every validation item SHALL contain:

| Field | Meaning |
|---|---|
| build_id | Run identity |
| stage | Preflight, Fixings, Exposure, P&L, Publish |
| severity | ERROR, WARNING, INFO |
| code | Stable machine-readable code |
| message | Concise human explanation |
| table | Input/output table when applicable |
| source_row_id | Source row when applicable |
| economic_key | Output key when applicable |
| actual | Observed value |
| expected | Required value/rule |
| remediation | Concrete action |

Validation SHALL be append-only by Build ID. A new run SHALL NOT erase an older run's evidence.

## 20. Error handling and publication

- Core functions SHALL raise or return typed errors; they SHALL NOT display dialogs.
- The pipeline SHALL convert unexpected exceptions into a blocking `UNEXPECTED_EXCEPTION` validation item.
- A failure SHALL retain staged diagnostics and SHALL discard or quarantine incomplete economic outputs.
- Published outputs SHALL be replaced atomically.
- The Excel adapter SHALL clear filters before replacing table bodies.
- The adapter SHALL write arrays in blocks and SHALL restore Excel application state in a guaranteed cleanup path.
- The adapter SHALL write a status summary only after successful import.

## 21. Determinism and performance

- Output keys and values SHALL be independent of input row order.
- Output ordering SHALL be explicit and stable.
- Repeated runs with identical fingerprints SHALL produce identical economic output hashes.
- The engine SHALL use vectorized, indexed, or dictionary-based operations rather than worksheet-style nested scans.
- A full production-sized Python rebuild SHALL complete in less than 60 seconds on the agreed current-Mac benchmark.
- A normal daily or incremental update SHOULD complete in less than 10 seconds.
- The focused unit-test suite SHALL complete in less than 2 seconds on the benchmark machine.
- Peak Python resident memory SHALL remain below 2 GiB for the production-sized fixture.
- Every performance run SHALL record elapsed time, peak memory, machine, operating system, Python version, and input row counts.
- Subsequent builds SHALL be checked for material regression against both the target and the recorded baseline.

## 22. Excel and VBA target

### 22.1 Thin Excel adapter

The Excel layer SHALL:

- export or identify immutable source snapshots;
- invoke or exchange files with the engine;
- validate the returned manifest;
- import outputs in bulk;
- update user-facing status and audit sheets;
- create `Daily Report D2` from verified engine output, with D2 equal to the final Market Date and
  D1 equal to its configured predecessor;
- avoid business calculations.

`Daily Report D2` SHALL aggregate the D2 P&L rows into the established legacy presentation:

1. Total P&L and its components by BOOK.
2. Delta Exposure MtM and economic Fixing Amount by BOOK and Delivery Month.
3. Delta Exposure MtM by Delivery Month and the selected Underlying columns, aggregated across all
   books with zeroes shown explicitly.

The adapter SHALL use the engine's `delta_exposure_mtm`, which already includes the trade-entry
adjustment. It SHALL use the economic `fixing_amount` from P&L, not the raw settlement from Fixings.
The report SHALL reconcile to the D2 engine output within the approved P&L tolerance and SHALL not
recalculate business methodology with worksheet formulas or VBA.

### 22.2 Clean VBA port

If production must run without Python, the VBA port SHALL:

- start from clean exported/imported source modules;
- implement the same contracts and rules;
- use arrays and dictionaries rather than cell-by-cell calculations;
- contain no modal dialog in core execution;
- pass Python/VBA parity tests before release;
- retain the Python engine as the reference oracle.

## 23. Acceptance boundary

The engine is acceptable only when:

- the methodology baseline is implemented and every required golden economic case is approved;
- all blocking tests in `GTM_ACCEPTANCE_TEST_SPECIFICATION.md` pass;
- approved golden economic cases pass;
- no required input price is silently substituted;
- output schemas and Build Manifest validate;
- Python output is deterministic;
- any VBA port meets parity tolerances;
- the Excel adapter passes controlled integration tests;
- a full historical build receives independent economic review.

## 24. Methodology resolution register

Plain-English resolutions and supporting explanations are maintained in `docs/GTM_DECISION_GUIDE.md`; detailed local evidence is in `docs/GTM_LOCAL_METHODOLOGY_TRACE.md`.

### D-001: Missing required prices — CONFIRMED 2026-08-07

Decision: A missing price required by any included fixing, exposure, or P&L row SHALL block the build. The failed build SHALL publish no new economic output and SHALL preserve the last accepted output.

The engine SHALL produce a structured list of every missing required price it can identify safely. The Excel adapter SHALL display a clear message asking the user to supply those prices. Each item SHALL identify the price type, relevant date, underlying/product, delivery or contract month, and affected source row or economic key when available.

The engine SHALL NOT silently replace a required missing price with zero, a previous price, an interpolated price, or another contract's price. Any substitution requires an explicit business policy, visible input provenance, a policy-version change, and dedicated tests.

For the temporary production run covering 30 June through 10 July 2026, the user approved an explicit input-normalization policy: fixing prices are source-backed only for TTF DA and PVB Heren and are zero for every other fixing product; Exposure uses the imported source curves except that Brent Dated and Henry Hub are zero. These are deliberate, auditable input values for this run—not a general missing-price fallback. D-001 continues to block any required key left absent after the approved normalization.

A price absent from the source but unused by every included economic row SHALL NOT block the build.

### D-002: Brent/HH valuation rule — RESOLVED

Same-month Brent/HH exposure and opening Initial MtM use the next Delivery Month's curve. Later months use their matching curve.

### D-003: DAY_AHEAD weekend and holiday rule — METHOD RESOLVED / PRODUCT GOLDEN TEST REQUIRED

DAY_AHEAD uses the previous calendar day. HEREN is the method that uses the previous Market Day and covers intervening calendar delivery days. SETUP assigns the method product by product; product names never imply methodology. Weekend-price fixtures are required for the active products assigned to DAY_AHEAD.

### D-004: Trade after scheduled fixing time/window — CONFIRMED 2026-08-07

Same-day normal fixings are valid. Reject a trade row whose complete quantity cannot be scheduled without inventing a past/catch-up fixing. Month Ahead may redistribute its complete monthly quantity over remaining eligible fixing dates on or after Trade Date; reject it if none remains. Never substitute Trade Price as Fixing Price.

### D-005: Operating-flow sources and allocation — RESOLVED / GOLDEN TEST REQUIRED

Use `-COSTS!B:N`, `COSTS!O:P + Foto FO!R:S`, and `Foto FO!T:V` as daily BOOK-level flows. One Ops golden case must certify the Logistics sign before production Total P&L approval.

### D-006: Zero-row retention — RESOLVED

Emit one closure row, omit later repeated zeros, and emit again on reopening.

### D-007: Precision and tolerances — CONFIRMED 2026-08-07

No intermediate rounding. Reconciliation tolerances are Volume `0.000001`, Price `0.00000001`, and P&L `EUR 0.01`. The separate `1e-7` economic-zero threshold controls event/row emission only.

### D-008: Canonical mappings — RESOLVED

Active SETUP mappings are authoritative. PVB detail products preserve method but aggregate to `Index PVB`; unknown identifiers fail Preflight.

### D-009: Non-market Trade Date — RESOLVED

Retain Trade Date and apply the event once on the first output Market Date on or after it.

### D-010: Performance target — ENGINEERING TARGET

Full production rebuild `< 60 s`; daily/incremental update ideally `< 10 s`; focused tests `< 2 s`; peak Python memory `< 2 GiB` on the agreed benchmark.

### D-011: Initial P&L schema — RESOLVED

Use the 13 active-book closing balances as a separate opening bridge for cumulative P&L. They sum to `EUR 37,445,758.99728647` in the inspected source. Six extra saved rows are accidentally ingested documentation and must be rejected. An end-to-end golden case SHALL prove the bridge is included exactly once.

### D-012: Daily Qty sign convention — RESOLVED

Daily Qty is non-negative; Side supplies sign. Blank/non-numeric/negative blocks. A deliberate zero warns and produces no economic event.

## 25. Change control

- Every accepted decision SHALL update `policy_version`.
- Every requirement change SHALL identify affected tests.
- Code SHALL NOT implement unresolved decisions through undocumented defaults.
- The project journal SHALL record specification revisions, user decisions, and implementation evidence.
