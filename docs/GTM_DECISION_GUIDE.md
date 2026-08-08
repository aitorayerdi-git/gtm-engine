# GTM Methodology Resolution Register

Version: 0.3
Date: 2026-08-07
Status: Business guidance and local evidence reconciled; approved for Python implementation

This file replaces the earlier questionnaire. A deeper audit found that the workbooks and helper implementations already answer nearly all of the questions. `questions_answers.txt` and its backup are excluded from this evidence.

The detailed evidence and conflicts are recorded in `GTM_LOCAL_METHODOLOGY_TRACE.md`.

## D-001 — Required price missing — CONFIRMED BY USER

If any included fixing, exposure, or P&L calculation needs a price that is absent, the complete build fails. It publishes nothing new, preserves the last accepted output, and returns a consolidated list of the missing price keys. Excel shows a clear message asking the user to supply them.

No substitute is allowed: not zero, a previous price, interpolation, or another contract. A missing source price that no included economic row uses does not block the build.

## D-002 — Brent and HH current-month curve — RESOLVED LOCALLY

If Market Date and Delivery Month are in the same calendar month, value Brent Dated and HH exposure with the next Delivery Month's curve price. For later Delivery Months, use the matching Delivery Month curve.

The rule applies both to normal Exposure valuation and to Initial Exposure when P&L constructs opening MtM. Current v2 Exposure implements this rule; current v2 P&L does not, so P&L must be corrected.

Plain-English example: on 15 July, July Brent exposure uses the August Brent curve column. September exposure uses the September column.

## D-003 — Day Ahead across weekends and holidays — METHOD RESOLVED / PRODUCT TESTS REQUIRED

Ordinary DAY_AHEAD uses the previous calendar date:

```text
fixing_date = delivery_day - 1 calendar day
```

It does not roll back to the previous Market Day. HEREN is the separate method that uses the previous Market Day and groups intervening weekend/holiday delivery days under it. BRENT_HH also uses the previous Market Day but only for market delivery days.

Plain-English example: a Monday DAY_AHEAD delivery fixes on Sunday. A Monday HEREN delivery fixes on the preceding Market Day, normally Friday, together with the intervening delivery days.

SETUP assigns methodology product by product. A product name containing `DA` does not imply DAY_AHEAD: `TTF DA` is currently mapped to HEREN. Weekend-price examples must be approved for `D+1 Auction`, `Mibgas Index ES`, and `MIBGAS D+1 Daily Reference`, the three current PVB detail products assigned to ordinary DAY_AHEAD.

## D-004 — Trade after its normal fixing — BUSINESS POLICY OVERRIDES EXISTING VBA

Same-day fixing is valid: `Fixing Date >= Trade Date`.

- Reject a non-Month-Ahead trade row if any included delivery slice has already lost its normal fixing opportunity. Do not silently omit that slice or invent a catch-up fixing.
- For MONTH_AHEAD, divide the affected total monthly delivery quantity equally over remaining prior-month Market Days on or after Trade Date.
- If no Month Ahead fixing day remains, reject the trade as a blocking error.
- Never use Trade Price as a replacement Fixing Price.

The completed legacy builder and fast helper implement catch-up-on-Trade-Date behavior. That code is retained as historical evidence but is explicitly rejected for release 1 by the later business decision.

## D-005 — Operating-flow sources, sign, and allocation — RESOLVED LOCALLY

Use daily Market Date and BOOK matching:

| Output component | Source | Sign treatment | Detail level |
|---|---|---|---|
| Logistical Costs | `COSTS!B:N` | Negate the stored cost, matching the legacy PNL formulas | BOOK level only |
| Fees and Optimizations | `COSTS!O:P` + `Foto FO!R:S` | Add values as stored | BOOK level only |
| Replication | `Foto FO!T:V` | Add values as stored | BOOK level only |

For these rows, use Underlying `TOTAL / BOOK LEVEL`, leave Delivery Month blank, and do not invent an allocation to product, month, source, or scenario.

These inputs are daily flows and are not differenced again. `COSTS!R4` contains an Ops sign-confirmation note, so the above workbook rule must receive one golden Ops sign test before production Total P&L approval.

## D-006 — Zero rows after closure — RESOLVED LOCALLY

Write one explicit zero row when a previously non-zero position closes. Omit later repeated zeros. If the position reopens, start emitting the key again.

This records the closure without filling every later date with zero rows.

## D-007 — Precision and test tolerance — CONFIRMED

Do not round intermediate calculations. Store full double-precision values; formatting is display-only.

Use `1e-7` only as the current economic materiality threshold for event/row emission. Use separate reconciliation tolerances:

```text
Volume = 0.000001
Price  = 0.00000001
P&L    = EUR 0.01
```

Money may display to two decimal places and normalized volume/price to six decimals without changing stored values.

## D-008 — Canonical mappings — RESOLVED LOCALLY

Active `SETUP` books, underlyings, and methods are authoritative. Matching is trimmed and case-insensitive; output uses the configured canonical spelling. Unknown values block Preflight.

The active PVB detail underlyings in `SETUP!M3:P11` retain their configured fixing method but aggregate to `Index PVB` for canonical exposure reporting. Main underlyings such as `Phys PVB`, `TVB`, and `AVB` remain distinct.

## D-009 — Trade Date on a weekend or holiday — RESOLVED LOCALLY

Retain the original Trade Date in the event ledger and audit trail. Apply the event once on the first output Market Date on or after Trade Date. Never move it backward to a date before the trade existed.

The same due-event rule applies to any fixing event whose economic date is not on the output Market Date axis.

## D-010 — Performance target — APPROVED ENGINEERING TARGET

Full production rebuild must take less than 60 seconds and use less than 2 GiB peak Python memory on the agreed current-Mac benchmark. The focused test suite must take less than 2 seconds. A normal daily/incremental update should ideally take less than 10 seconds. Every benchmark also records the measured result as a regression baseline.

## D-011 — Meaning of Initial P&L — RESOLVED LOCALLY

`INITIAL POSITION!A7:B19` contains 13 signed closing P&L balances—one for each active BOOK—at Initial Market Date. This is a book-level opening balance for cumulative reporting. It is separate from Initial Exposure and is not opening Exposure MtM.

```text
cumulative_pnl[book, market_date]
    = initial_pnl[book]
    + sum(daily_pnl[book] after Initial Market Date through market_date)
```

Do not report Initial P&L as a post-cut-off daily flow and do not add it to Exposure MtM.

The saved normalized table contains 19 rows only because the refresh macro accidentally ingested six documentation lines as zero-value books. The correct population is the 13 active SETUP books. This is a defect to fix, not a 19-component business schema.

The 13 inspected source balances sum to `EUR 37,445,758.99728647`. The structural 19-row check is complete; an end-to-end golden case must still prove that the opening bridge is included exactly once.

## D-012 — BUY/SELL and Daily Qty — RESOLVED LOCALLY

Daily Qty is a non-negative magnitude. BUY supplies a positive exposure sign; SELL supplies a negative sign. Negative Daily Qty is invalid because Side would apply a second sign.

Blank or non-numeric Daily Qty blocks the build. A deliberately entered numeric zero is accepted with a warning and generates no economic event. The live trade data support the convention: there are no negative Daily Qty values and 12 deliberate zeros.

## Remaining direct input

No broad methodology questionnaire remains. Direct review is limited to concrete golden results: the three DAY_AHEAD weekend-price fixtures, one Brent roll case, one HH roll case, the 13-book Initial P&L bridge, and one Ops example certifying the Logistics sign convention.

## Temporary price policy for the 30 June–10 July 2026 run

The user approved a narrow normalization policy for this historical run. Fixing prices remain source-backed for TTF DA and PVB Heren; every other fixing product is explicitly assigned zero. Exposure uses the imported source curves, except Brent Dated and Henry Hub are explicitly assigned zero.

This does not weaken D-001. The run input must contain every required price key after normalization, and the provenance must show which values were supplied by the source and which were deliberately set to zero by policy. An absent required key still blocks publication.
