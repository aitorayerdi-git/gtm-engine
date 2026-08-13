# GTM v0.3 Quick Start

This is the short guide. Use the [detailed manual](GTM_EXCEL_INTERFACE_GUIDE.md) when preparing a
new dataset, importing legacy data, or investigating a result.

If the software is not installed, begin with the [installation guide](../INSTALL.md).

## What the system does

```text
Excel input workbook → Python engine → separate Excel result workbook
```

Excel holds the inputs and displays the results. Python performs the calculations. The engine
does not run Excel formulas or macros and does not change your input workbook.

## 1. Pass information to the engine

If the information is already in a legacy GTM workbook, import it first:

```sh
.venv/bin/gtm-engine legacy-import \
  --workbook "Gas_Trading_Model 070826.xlsm" \
  --output outputs/legacy_import_070826_v0.3 \
  --historical-end 2026-07-10
```

Open `legacy_import_audit.json` and review every error and warning. The macro-free working input is
`outputs/legacy_import_070826_v0.3/GTM_Imported_Input.xlsx`. The source `.xlsm` remains unchanged.
The example uses 10 July 2026 as both the reporting cutoff and the last included Trade Date.

For manual entry, open:

```text
outputs/gtm_excel_v0_3/GTM_Excel_Interface_v0.3.xlsx
```

Enter values in the blue input tables:

1. `MANUAL CHANGES`: enter all dates in `tblManualDates`. Historical End Date is the single Last Market Date shared by Foto FO and the engine.
2. `MARKET CALENDAR`: supply every calendar date needed by the model and mark each date as a
   Market Day or non-Market Day.
3. `INITIAL EXPOSURE`: enter the positions open at the close of the Initial Market Date.
4. `INITIAL PNL`: enter one opening P&L balance for every active BOOK.
5. `TRADES`: enter trades after the Initial Market Date.
6. `CURVE PRICES` and `FIXING PRICES`: enter all prices required by those positions and trades.
7. `OPERATING FLOWS`: enter daily Logistics, Fees/Optimizations, and Replication amounts when
   applicable.

`BOOKS` and `UNDERLYINGS` already contain the reviewed legacy SETUP mapping. Change them only when
the approved configuration changes.

Important rules:

- Put data inside the named Excel tables. Rows pasted below a table but outside its boundary are
  ignored.
- Use values, not formulas.
- Give each trade, initial position, price, and operating-flow row a stable, unique source ID.
- Enter `Daily Qty` as a positive or zero magnitude. `BUY` or `SELL` supplies the sign.
- Use the first day of the month for every `Delivery Month`.
- Currency and unit must match the `UNDERLYINGS` configuration.

## 2. Run the engine

Save and close the input workbook.

If you edited the delivered workbook, double-click:

```text
scripts/GTM_Build.command
```

If you made a separate working copy, drag that `.xlsx` file onto `GTM_Build.command`.

You can also use Terminal:

```sh
.venv/bin/gtm-engine excel-build \
  --workbook path/to/your_input.xlsx \
  --output outputs/gtm_excel_runs
```

On success, open:

```text
outputs/gtm_excel_runs/GTM_LATEST.xlsx
```

On failure, open the newest `GTM_Failed.xlsx` under:

```text
outputs/gtm_excel_runs/failed/
```

A failed run never replaces the last successful `GTM_LATEST.xlsx`.

## 3. Analyse the result

Read the result in this order:

1. `START HERE`: confirm that Build Status is `PUBLISHED`.
2. `VALIDATION`: confirm that there are no errors; review every warning.
3. `Daily Report D2`: review the latest Market Date against the previous Market Date. D2 is the
   final Market Date in the build; D1 is its configured predecessor.
4. `DAILY PNL`: inspect the detailed P&L components by Market Date, BOOK, Underlying, and Delivery
   Month.
5. `CUMULATIVE PNL`: check the opening Initial P&L plus subsequent Daily P&L by BOOK.
6. `EXPOSURE`: inspect remaining open volume and MtM.
7. `FIXINGS`: inspect priced fixing events.
8. `EVENT LEDGER`: trace a source row through its trade and fixing events.
9. `BUILD MANIFEST`: confirm versions, input hash, row counts, totals, and Build ID.

Never use a workbook whose status is `FAILED` as an accepted result.

## Common failures

| Message | What to do |
|---|---|
| Missing required price | Add the exact Market Date, Underlying, and Delivery Month or fixing key shown in `VALIDATION`. |
| Unknown BOOK or Underlying | Correct the name or add an approved mapping. |
| Formula not allowed | Replace the formula with its saved value. |
| Invalid Daily Qty | Enter a number greater than or equal to zero; keep the sign in BUY/SELL. |
| Late trade | Correct the dates or reject the trade; the engine will not invent a catch-up fixing. |
| Calendar gap | Add the missing calendar dates and mark each one TRUE or FALSE. |
