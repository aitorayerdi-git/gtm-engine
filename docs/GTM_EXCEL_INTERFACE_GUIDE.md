# GTM v0.3 Detailed User Manual

## 1. Purpose

The GTM system reconstructs daily gas-trading exposure, fixings, and P&L from an opening state,
subsequent trades, market prices, and operating flows.

The system has three parts:

```text
Excel input workbook
        ↓
Python calculation engine
        ↓
Excel result workbook + CSV/JSON audit evidence
```

Excel is an interface, not the calculation engine. Python does not open Microsoft Excel, run
macros, or calculate workbook formulas. A build reads a saved `.xlsx` file and creates a separate
result. It never intentionally edits the source workbook.

For the normal workflow, start with:

```text
outputs/gtm_excel_v0_3/GTM_Excel_Interface_v0.3.xlsx
```

For the short daily procedure, use [GTM_QUICK_START.md](GTM_QUICK_START.md).
For a new computer or checkout, begin with the [installation guide](../INSTALL.md).

## 2. The three tasks

Every run follows the same sequence:

1. Pass information to the engine through the input tables.
2. Save and close the workbook, then run the engine.
3. Open the separate result workbook and analyse it.

## 3. Pass information to the engine

### 3.1 General data-entry rules

- Enter data only in the blue input tables.
- Use values, not formulas. The adapter rejects formulas in authoritative input cells.
- Do not rename worksheets, Excel tables, or table columns.
- Add rows inside the table. Press `Tab` in the table's last cell, use Excel's Insert Table Row,
  or resize the table after pasting. Data outside the table boundary are ignored.
- Use real Excel dates. Do not enter dates as ambiguous text such as `7/8/26`.
- Use the first day of the month for every Delivery Month, for example `2026-09-01`.
- Use stable, unique source IDs. Keep the same ID when correcting the same source record.
- Leave an optional field blank rather than entering `N/A`.
- Save a working copy before making a large import.

### 3.2 `CONTROL`

`CONTROL` defines the run.

| Field | What to enter |
|---|---|
| Model ID | Keep `GTM`. |
| Schema Version | Keep the supplied value unless the software is upgraded. |
| Policy Version | Keep the supplied value unless the business rules are upgraded. |
| Engine Version | Keep the supplied value unless the engine is upgraded. |
| Initial Market Date | The close-of-day date represented by Initial Exposure and Initial P&L. |
| Historical Start Date | The first date requested for reconstruction; normally the Initial Market Date. |
| Historical End Date | The final requested reporting date. |
| Simulation Status | `OFF` excludes simulation trades; `ON` includes them. |
| Timezone | Keep `Europe/Madrid` unless the operating convention changes. |
| Logistics Sign | Normally `-1`: a positive stored Logistics cost reduces P&L. |
| Materiality | The threshold below which exposure is treated as zero. |

The Initial Market Date is an end-of-day boundary. Initial Exposure and Initial P&L already
include activity through that date. Only later trades enter as incremental activity.

### 3.3 `BOOKS`

This table replaces the active BOOK section of the legacy `SETUP` sheet.

| Column | Meaning |
|---|---|
| Book | Exact canonical BOOK name. |
| Active | `TRUE` if the BOOK participates in the build. |

The delivered template contains the 13 reviewed BOOK names. Unknown names fail Preflight. Do not
merge similar names or invent aliases.

### 3.4 `UNDERLYINGS`

This table replaces the Underlying, methodology, and mapping sections of legacy `SETUP`.

| Column | Meaning |
|---|---|
| Source Underlying | Name used by the trade or position source. |
| Canonical Underlying | Name used in exposure and P&L output. |
| Fixing Method | `WITHINDAY`, `DAY_AHEAD`, `HEREN`, `MONTH_AHEAD`, or `BRENT_HH`. |
| Unit | `MWh`, `bbl`, `MMBtu`, or another explicitly approved unit. |
| Currency | Currency expected for prices. |
| Curve Underlying | Name used to find curve prices. |
| Fixing Price Underlying | Name used to find fixing prices. |
| Fixing Price Basis | `FIXING_DATE` or `DELIVERY_DAY`. |
| Active | Whether the mapping is active. |
| Current Month Uses Next Curve | `TRUE` for the approved Brent/HH current-month roll rule. |

The delivered template contains 18 reviewed mappings. Product names never determine the fixing
method; this table does.

Prices must match the configured currency and unit. The engine does not silently convert USD to
EUR or bbl/MMBtu to MWh. In particular, do not import legacy Brent or HH prices under a different
currency convention until that convention and its FX treatment have been approved.

### 3.5 `MARKET CALENDAR`

Supply a continuous daily calendar. Include every calendar day needed by the reporting range,
trade dates, fixing dates, and delivery periods. Do not supply only weekdays.

| Column | Meaning |
|---|---|
| Date | Calendar date. |
| Is Market Day | `TRUE` for a configured Market Day; otherwise `FALSE`. |
| Calendar ID | Normally `GTM`. |

The engine uses this table to find previous and next Market Dates. A weekend trade remains dated
on the weekend but is reported once on the first Market Date on or after it.

### 3.6 `INITIAL EXPOSURE`

Enter positions open at the close of the Initial Market Date.

| Column | Meaning |
|---|---|
| Initial Market Date | Must match `CONTROL`. |
| Book | Active BOOK name. |
| Underlying | Approved source or canonical Underlying name. |
| Delivery Month | First calendar day of the delivery month. |
| Exposure Volume | Signed opening volume: positive long, negative short. |
| Source Row ID | Stable unique identifier. |

Zero opening positions may be omitted. Do not enter trades already represented in this opening
state as later incremental trades.

### 3.7 `INITIAL PNL`

Enter the cumulative opening P&L bridge at the Initial Market Date.

| Column | Meaning |
|---|---|
| Initial Market Date | Must match `CONTROL`. |
| Book | Active BOOK name. |
| Amount | Signed opening cumulative P&L. |
| Source Row ID | Stable unique identifier. |
| Comment | Optional explanation or source note. |

Supply exactly one row for every active BOOK. Initial P&L is not Daily P&L and is not opening
Exposure MtM. It seeds Cumulative P&L only.

### 3.8 `TRADES`

Enter one row per trade.

| Column | Meaning |
|---|---|
| Source Row ID | Stable unique trade identifier. |
| Trade Date | Original economic trade date, including weekends or holidays. |
| Book | Active BOOK. |
| Underlying | Approved mapped Underlying. |
| Side | `BUY` or `SELL`. |
| Start Date | First delivery day. |
| End Date | Last delivery day, inclusive. |
| Daily Qty | Unsigned quantity for each delivery day. |
| Execution Price | Trade price in the configured currency and unit. |
| Trade Source | `ACTUAL` or `SIMULATION`. |
| Scenario | Optional simulation scenario identifier. |

Quantity rules:

- `BUY 100` creates positive exposure.
- `SELL 100` creates negative exposure.
- Entering `SELL -100` is invalid because SELL already supplies the sign.
- Blank, text, and negative quantities block the build.
- Explicit zero is accepted with a warning and creates no economic event.

A fixing on the Trade Date is valid. A trade entered after every valid fixing opportunity blocks
the build; the engine does not invent a catch-up fixing.

### 3.9 `CURVE PRICES`

Curve prices value remaining exposure.

| Column | Meaning |
|---|---|
| Market Date | Reporting Market Date. |
| Underlying | Configured Curve Underlying. |
| Delivery Month | First day of the curve contract's delivery month. |
| Curve Price | Closing curve price. |
| Currency | Must match `UNDERLYINGS`. |
| Unit | Must match `UNDERLYINGS`. |
| Source ID | Stable unique market-data identifier. |
| Source As Of | Optional source timestamp. |

Supply every key required by included open positions. Missing required prices block the build;
unused absent prices do not. Current-month Brent and HH exposure uses the next-month curve
contract, so include that contract when applicable.

### 3.10 `FIXING PRICES`

Fixing prices settle fixing events.

| Column | Meaning |
|---|---|
| Price Lookup Date | Date required by the Underlying's configured price basis. |
| Underlying | Configured Fixing Price Underlying. |
| Fixing Price | Observed fixing price. |
| Currency | Must match `UNDERLYINGS`. |
| Unit | Must match `UNDERLYINGS`. |
| Source ID | Stable unique market-data identifier. |
| Source As Of | Optional source timestamp. |

Do not fill a missing required price with zero, carry forward an old price, or interpolate unless
an approved product rule explicitly requires it.

### 3.11 `OPERATING FLOWS`

Enter daily BOOK-level operating flows.

| Column | Meaning |
|---|---|
| Market Date | Date of the daily flow. |
| Book | Active BOOK. |
| Logistics Source Amount | Amount as stored by the source; `CONTROL` applies the Logistics sign. |
| Fees And Optimizations | Signed Fees plus Optimizations amount. |
| Replication | Signed Replication amount. |
| Source Row ID | Stable unique identifier. |

These are daily amounts, not day-over-day balance changes. Do not allocate them to Underlying,
Delivery Month, or Scenario unless an authoritative source supplies that detail.

### 3.12 Importing a legacy workbook

The read-only importer converts cached values from a legacy `.xlsm` into both supported v0.3 input
formats. It does not open Microsoft Excel, run VBA, recalculate formulas, or save the source file.

Run:

```sh
.venv/bin/gtm-engine legacy-import \
  --workbook "Gas_Trading_Model 070826.xlsm" \
  --output outputs/legacy_import_070826_v0.3 \
  --historical-end 2026-07-10
```

The output directory must not already exist. The command creates:

```text
outputs/legacy_import_070826_v0.3/
├── GTM_Imported_Input.xlsx
├── legacy_import_audit.json
├── legacy_import_issues.csv
└── normalized_bundle/
    ├── bundle.json
    └── the nine normalized CSV input tables
```

`GTM_Imported_Input.xlsx` contains a `LEGACY IMPORT` sheet with the source SHA-256, import status,
and review items. The JSON audit also records exact source ranges, extracted row counts, skipped
placeholder counts, and confirmation that the source hash remained unchanged.

`--historical-end` defines the complete as-of cutoff. The importer excludes source trades whose
Trade Date is later than that date and lists their row numbers in the audit. Prices after the
cutoff do not block the build unless an included fixing uses a later lookup date, as Brent can
when its fixing price is keyed by Delivery Day.

The importer applies these mappings:

| Legacy source | New destination | Transformation |
|---|---|---|
| `SETUP` | `BOOKS`, `UNDERLYINGS` | Extract active rows and apply the approved canonical, fixing-basis, and Brent/HH roll mappings. |
| `INITIAL POSITION!C2` | `CONTROL` | Initial Market Date. |
| `INITIAL POSITION DATA!A:E` | `INITIAL EXPOSURE` | Rename Initial Exposure to Exposure Volume; add Source Row IDs; normally omit zero rows. |
| `INITIAL POSITION!A5:C17` | `INITIAL PNL` | Use exactly the 13 active BOOK rows; ignore methodology text and the total. |
| `TRADES!A:I` | `TRADES` | Map source fields; set Trade Source to `ACTUAL`; ignore calculated columns J onward. |
| `SIMULATION TRADES!A:I` | `TRADES` | Map source fields; set Trade Source to `SIMULATION`. |
| `CALENDAR!A:E` | `MARKET CALENDAR` | Copy Date and convert the saved Market Day result to TRUE/FALSE values. |
| `TTF`, `Brent Dated`, `HH` | `CURVE PRICES` | Unpivot cached history into one row per Market Date, Underlying, and Delivery Month. |
| `PVB-TTF`, `PEG-TTF` | `CURVE PRICES` | Add the cached spread to TTF to reconstruct PVB-family and PEG curves. |
| `FIXING PRICES` | `FIXING PRICES` | Unpivot into one row per Price Lookup Date and Underlying. |
| `COSTS!B:N` | `OPERATING FLOWS` | Logistics by Market Date and BOOK. |
| `COSTS!O:P` plus `Foto FO!R:S` | `OPERATING FLOWS` | Sum into Fees and Optimizations by Market Date and BOOK. |
| `Foto FO!T:V` | `OPERATING FLOWS` | Replication by Market Date and BOOK. |

Exact zero opening positions are omitted. Zero-filled curve and fixing-price placeholders are
treated as missing, not as real market prices. Explicit zero trade quantities and prices remain in
the input and appear as audit warnings. Legacy `FIXINGS DATA`, `EXPOSURE DATA`, `PNL DATA`, and
related calculated sheets remain diagnostic material; the importer never treats them as source
inputs.

Import status and build status answer different questions. `CREATED` means that conversion
succeeded. `CREATED_WITH_REVIEW_ITEMS` means that conversion succeeded but the audit found a
known business or data gap. Run `excel-build` or `build` next; only `PUBLISHED` means that the
engine accepted the economic input.

## 4. Run the engine

### 4.1 Pre-run checklist

Before every run:

1. Confirm the three dates in `CONTROL`.
2. Confirm Simulation Status.
3. Confirm that all imported rows are inside their named tables.
4. Confirm that source IDs are populated and unique.
5. Confirm that Daily Qty is unsigned and BUY/SELL supplies the sign.
6. Confirm that Delivery Month values are month-start dates.
7. Confirm that currencies and units match `UNDERLYINGS`.
8. Save and close the workbook.

### 4.2 Run with the macOS launcher

If you edited the delivered interface, double-click:

```text
scripts/GTM_Build.command
```

The launcher's default input is:

```text
outputs/gtm_excel_v0_3/GTM_Excel_Interface_v0.3.xlsx
```

If you created another working copy, drag that `.xlsx` file onto the launcher. Otherwise, the
launcher will build the delivered default workbook rather than your copy.

### 4.3 Run from Terminal

From the project directory:

```sh
.venv/bin/gtm-engine excel-build \
  --workbook path/to/your_input.xlsx \
  --output outputs/gtm_excel_runs
```

Microsoft Excel does not run during this command.

### 4.4 What the build creates

On success:

```text
outputs/gtm_excel_runs/
├── GTM_LATEST.xlsx
├── LATEST
└── runs/<run_id>/
    ├── GTM_Result.xlsx
    ├── build_manifest.json
    ├── validation.csv
    ├── fixings.csv
    ├── exposure.csv
    ├── pnl.csv
    ├── cumulative_pnl.csv
    └── event_ledger.csv
```

`GTM_LATEST.xlsx` changes only after a successful verified build.

On failure:

```text
outputs/gtm_excel_runs/failed/<run_id>/GTM_Failed.xlsx
```

The previous `GTM_LATEST.xlsx` remains unchanged.

## 5. Analyse the result

### 5.1 Start with status and validation

Open the result workbook and read these sheets first:

1. `START HERE`: Build Status must be `PUBLISHED`.
2. `VALIDATION`: there must be no `ERROR` rows. Review all `WARNING` rows.
3. `BUILD MANIFEST`: confirm the Build ID, versions, input hash, input row counts, output row
   counts, totals, and Simulation Status.

Do not accept a workbook marked `FAILED`, even if some old output rows appear plausible.

### 5.2 Understand each output

| Sheet | What it answers |
|---|---|
| `Daily Report D2` | What changed on the latest Market Date compared with the previous Market Date? |
| `EVENT LEDGER` | Which initial positions, trades, and fixings changed the state, and when were they applied? |
| `FIXINGS` | What volume fixed, on which date, at which price? |
| `EXPOSURE` | What volume remained open on each Market Date, and what was its MtM? |
| `DAILY PNL` | Which components produced the day's P&L? |
| `CUMULATIVE PNL` | How did opening Initial P&L plus later Daily P&L accumulate by BOOK? |

`Daily Report D2` follows the legacy report layout without its calculation sheets. D2 is the
latest Market Date in the successful build, and D1 is the preceding configured Market Date. The
sheet contains three sections:

1. Daily P&L by BOOK.
2. Delta Exposure MtM and Fixing Amount by BOOK and Delivery Month.
3. Delta Exposure MtM by Delivery Month and selected Underlying, aggregated across all books.

The Python engine writes the report from the verified `DAILY PNL` output. The sheet contains no
formulas or macros. It displays one decimal place to match the legacy format, but the underlying
cells retain normal Excel numeric precision. Use `DAILY PNL` for the full audit detail.

### 5.3 Read the signs correctly

- Positive Exposure Volume means long; negative means short.
- A BUY trade adds exposure; a SELL trade subtracts exposure.
- Fixing Volume is the signed change that closes exposure. Closing a long position therefore has
  negative Fixing Volume.
- `FIXINGS.Fixing Amount` is the raw signed settlement: Fixing Volume multiplied by Fixing Price.
- `DAILY PNL.Fixing Amount` is the economic P&L contribution and has the opposite sign from the raw
  settlement. Do not expect those two columns to have the same sign.
- Logistics Source Amount is multiplied by the configured Logistics Sign, normally `-1`.

### 5.4 Reconcile Daily P&L

For each Market Date and economic key:

```text
Delta Exposure MtM
    = Gross Delta Exposure MtM
    + Trade Entry Adjustment

Total P&L
    = Delta Exposure MtM
    + Fixing Amount
    + Logistical Costs
    + Fees and Optimizations
    + Replication
```

For each BOOK:

```text
Cumulative P&L
    = Initial P&L
    + sum of Daily P&L after the Initial Market Date
```

The engine retains full precision during calculation. Normal comparison tolerances are:

| Measure | Tolerance |
|---|---:|
| Volume | 0.000001 |
| Price | 0.00000001 |
| P&L | EUR 0.01 |

### 5.5 Trace one trade or position

To investigate a result:

1. Find its Source Row ID in `EVENT LEDGER`.
2. Confirm the Economic Date and Applied Market Date.
3. Follow the BOOK, Underlying, and Delivery Month into `FIXINGS` and `EXPOSURE`.
4. Find the same key in `DAILY PNL`.
5. Confirm the BOOK total in `CUMULATIVE PNL`.
6. Record the Build ID from `BUILD MANIFEST` with your analysis.

This sequence separates bad source data from fixing-schedule, valuation, or aggregation issues.

## 6. Diagnose a failed build

Open the newest `GTM_Failed.xlsx` and read `VALIDATION`.

| Error | Meaning | Correction |
|---|---|---|
| Missing required price | An included position or fixing has no exact price key. | Add the Market Date, Underlying, Delivery Month, or lookup date shown. |
| Unknown BOOK or Underlying | The source name has no active mapping. | Correct the name or add an approved configuration row. |
| Formula not allowed | An authoritative input cell contains a formula. | Paste the saved value into the cell. |
| Invalid Daily Qty | Quantity is blank, text, or negative. | Enter a numeric magnitude greater than or equal to zero. |
| Late trade | No valid fixing opportunity remains. | Correct the dates or reject the trade; do not create a catch-up fixing. |
| Calendar gap | A required calendar date is absent. | Add every missing date and mark it TRUE or FALSE. |
| Duplicate key | Two source rows identify the same unique economic or price key. | Remove the duplicate or correct its identity. |
| Currency or unit mismatch | Price metadata differ from `UNDERLYINGS`. | Correct the input or the approved configuration; do not silently convert. |

Warnings do not block publication, but they require review. An explicit zero trade, for example,
produces a warning because it may indicate an input mistake.

## 7. Run without Excel

The core engine can also read a normalized directory:

```text
bundle.json
books.csv
underlyings.csv
market_calendar.csv
initial_exposure.csv
initial_pnl.csv
trades.csv
curve_prices.csv
fixing_prices.csv
operating_flows.csv
```

Run:

```sh
.venv/bin/gtm-engine build \
  --input path/to/normalized_bundle \
  --output path/to/results
```

`legacy-import` creates this directory as `normalized_bundle/`. The Excel and normalized outputs
contain the same economic rows; small numeric differences caused by Excel's 15-digit storage must
remain within the approved volume, price, and P&L tolerances.

## 8. Current production boundary

The engine, Excel interface, read-only legacy importer, failure controls, and synthetic regression
suite are implemented. Production acceptance still requires:

- a complete approved production input snapshot;
- correction or supply of the exact fixing-price keys reported by imported-snapshot Preflight;
- an approved FX rule for Brent/HH before their USD values enter combined EUR P&L;
- product-specific DAY_AHEAD weekend fixtures;
- approved Brent and HH roll/price examples, including the currency convention;
- the one-time 13-BOOK Initial P&L reconciliation;
- a production-sized timing and reconciliation run.

These are data-migration and business-acceptance tasks. They are not silently supplied by the
empty production template.
