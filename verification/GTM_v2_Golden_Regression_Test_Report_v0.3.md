# GTM v2 Golden Regression Test Report

Date: 2026-08-07
Engine: 0.3.0
Policy: 0.3.0
Test source: `GTM_v2_Golden_Regression_Test_Pack_for_CODEX.docx`
Source SHA-256: `9cf0e80179694859cab1c9fed7a6edc9a01a170e2bdd500764a7ed5925b8731a`

## Result

**PASS — 12 of 12 golden-pack tests passed.**

The first run passed 10 tests and failed TEST 4 and TEST 12. Both failures came from one defect:
P&L added the raw signed fixing settlement instead of its inverse economic cash-flow sign. After
the central correction, all 12 cases passed.

The full project suite also passed: 41 tests, 88.16% statement/branch coverage, Ruff clean, strict
mypy clean, and no broken Python requirements. The full suite completed in 2.02 seconds under
coverage instrumentation.

Comparison tolerances:

- Volume: `0.000001` MWh
- Price: `0.00000001`
- P&L: `EUR 0.01`

All reported numerical differences below are exact zero before tolerance.

## TEST 1 — BUY / SELL sign convention

- **PASS / FAIL:** PASS
- **Input:** BUY and SELL 100 MWh one-day Sep-26 slices; negative, blank, and explicit-zero quantity variants.
- **Expected Fixings:** None within the modeled July horizon.
- **Actual Fixings:** None within the modeled July horizon.
- **Expected Exposure:** BUY `+100`; SELL `-100`; invalid quantities create no exposure.
- **Actual Exposure:** BUY `+100`; SELL `-100`; negative quantity blocked; blank rejected by schema; explicit zero warned and created no event.
- **Expected P&L components:** Not specified for this sign/validation test.
- **Actual P&L components:** Excluded from the PASS decision.
- **Difference:** Zero for every asserted exposure and validation result.
- **Explanation if FAIL:** Not applicable.

## TEST 2 — Trade-entry Exposure and P&L

- **PASS / FAIL:** PASS
- **Input:** One-day Sep-26 slice; BUY/SELL 100 MWh at EUR 45/MWh; 01/07 curve EUR 47/MWh.
- **Expected Fixings:** None within the modeled horizon.
- **Actual Fixings:** None within the modeled horizon.
- **Expected Exposure:** BUY `+100`, MtM `+4,700`; SELL `-100`, MtM `-4,700`.
- **Actual Exposure:** BUY `+100`, MtM `+4,700`; SELL `-100`, MtM `-4,700`.
- **Expected P&L components:** BUY gross `+4,700`, entry `-4,500`, total `+200`; SELL gross `-4,700`, entry `+4,500`, total `-200`.
- **Actual P&L components:** Exact match.
- **Difference:** EUR `0.00`.
- **Explanation if FAIL:** Not applicable.

## TEST 3 — Pure market movement after trade entry

- **PASS / FAIL:** PASS
- **Input:** Continue TEST 2 BUY; curve moves from EUR 47 to EUR 49 on 02/07; no new trade.
- **Expected Fixings:** None.
- **Actual Fixings:** None.
- **Expected Exposure:** `+100`; MtM moves from EUR 4,700 to EUR 4,900.
- **Actual Exposure:** Exact match.
- **Expected P&L components:** Gross delta `+200`; entry `0`; fixing `0`; total `+200`.
- **Actual P&L components:** Exact match.
- **Difference:** EUR `0.00`.
- **Explanation if FAIL:** Not applicable.

## TEST 4 — Full fixing and exposure closure

- **PASS / FAIL:** PASS
- **Input:** Open `+100` MWh at prior curve EUR 47/MWh; full fixing at EUR 48/MWh on 02/07.
- **Expected Fixings:** Closing volume `-100`; raw amount `-4,800`; economic P&L contribution `+4,800`.
- **Actual Fixings:** Closing volume `-100`; raw amount `-4,800`; economic P&L contribution `+4,800`.
- **Expected Exposure:** Explicit closure row with volume `0` and MtM `0`.
- **Actual Exposure:** One explicit closure row with volume `0` and MtM `0`.
- **Expected P&L components:** Gross delta `-4,700`; economic fixing `+4,800`; total `+100`.
- **Actual P&L components:** Exact match.
- **Difference:** EUR `0.00`.
- **Explanation if FAIL:** Not applicable after correction.

## TEST 5 — Same-day fixing eligibility

- **PASS / FAIL:** PASS
- **Input:** Trade Date and Fixing Date 14/07; delivery 15/07; DAY_AHEAD; 100 MWh.
- **Expected Fixings:** The 14/07 fixing is eligible.
- **Actual Fixings:** One fixing dated 14/07.
- **Expected Exposure:** Trade and fixing use the same eligibility boundary and net once.
- **Actual Exposure:** Event ledger contains one `+100` trade and one `-100` fixing on 14/07; no residual exposure row.
- **Expected P&L components:** Not specified by the pack for this eligibility test.
- **Actual P&L components:** Excluded from the PASS decision.
- **Difference:** Zero for the asserted dates, counts, and volumes.
- **Explanation if FAIL:** Not applicable.

## TEST 6 — Trade after all fixing opportunities

- **PASS / FAIL:** PASS
- **Input:** Jun-26 MONTH_AHEAD delivery; May-26 fixing window; Trade Date 10/06.
- **Expected Fixings:** No catch-up fixing; blocking failure.
- **Actual Fixings:** None; Preflight returned `LATE_MONTH_AHEAD_TRADE`.
- **Expected Exposure:** None.
- **Actual Exposure:** None.
- **Expected P&L components:** None.
- **Actual P&L components:** None.
- **Difference:** Exact match.
- **Explanation if FAIL:** Not applicable.

## TEST 7 — Deferred weekend / holiday event

- **PASS / FAIL:** PASS
- **Input:** Market Dates Friday 10/07 and Monday 13/07; Saturday 11/07 BUY 100 MWh.
- **Expected Fixings:** None within the modeled horizon.
- **Actual Fixings:** None within the modeled horizon.
- **Expected Exposure:** No Saturday/Sunday row; `+100` on Monday, applied once.
- **Actual Exposure:** No 11/07 or 12/07 row; `+100` on 13/07 and still `+100`, not `+200`, on 14/07.
- **Expected P&L components:** Not specified for this event-timing test.
- **Actual P&L components:** Excluded from the PASS decision.
- **Difference:** Zero for dates, event count, and volume.
- **Explanation if FAIL:** Not applicable.

## TEST 8 — Explicit zero row on closure

- **PASS / FAIL:** PASS
- **Input:** `+100` on 01/07; close on 02/07; reopen `+50` on configured Market Date 05/07.
- **Expected Fixings:** One closing event for `-100`.
- **Actual Fixings:** One closing event for `-100`.
- **Expected Exposure:** Rows on 01/07 `+100`, 02/07 `0`, no 03/07 or 04/07 row, and 05/07 `+50`.
- **Actual Exposure:** Exact match; exactly one row has `is_explicit_closure=TRUE`.
- **Expected P&L components:** Not specified for this row-emission test.
- **Actual P&L components:** Excluded from the PASS decision.
- **Difference:** Zero for dates, row count, and volumes.
- **Explanation if FAIL:** Not applicable.

## TEST 9 — Missing required price

- **PASS / FAIL:** PASS
- **Input:** 01/07 TEST GAS Sep-26 exposure `+100`; required curve price removed.
- **Expected Fixings:** No published output after failure.
- **Actual Fixings:** No published output after failure.
- **Expected Exposure:** Build failure; no zero, carry-forward, or interpolation.
- **Actual Exposure:** Build failed before Exposure publication with `MISSING_CURVE_PRICE`; message identified `2026-07-01`, `TEST GAS`, and `2026-09-01`.
- **Expected P&L components:** None after failure.
- **Actual P&L components:** None after failure.
- **Difference:** Exact match. A separate empty-portfolio case verified that an unused absent price does not block.
- **Explanation if FAIL:** Not applicable.

## TEST 10 — Operating P&L components

- **PASS / FAIL:** PASS
- **Input:** Gross Exposure delta `+500`; economic fixing `+200`; Logistics source `+1,000`; Fees/Optimizations `+150`; Replication `+50`.
- **Expected Fixings:** Economic fixing contribution `+200`.
- **Actual Fixings:** Raw settlement `-200`; economic P&L contribution `+200`.
- **Expected Exposure:** Position revaluation contributes `+500`.
- **Actual Exposure:** Gross Exposure MtM delta `+500`.
- **Expected P&L components:** `+500 + 200 - 1,000 + 150 + 50 = -100`.
- **Actual P&L components:** Exact match across position, fixing, and BOOK-level operating rows; aggregate `-100`.
- **Difference:** EUR `0.00`.
- **Explanation if FAIL:** Not applicable.

## TEST 11 — Initial Market Date boundary

- **PASS / FAIL:** PASS
- **Input:** Initial Market Date 30/06; Trade A dated 30/06; Trade B dated 01/07.
- **Expected Fixings:** Only Trade B enters the incremental schedule.
- **Actual Fixings:** Only Trade B generated schedule/event lineage.
- **Expected Exposure:** `+100` from Trade B; Trade A omitted from incremental activity.
- **Actual Exposure:** `+100`; ledger contains Trade B and excludes Trade A.
- **Expected P&L components:** Not specified for this boundary test.
- **Actual P&L components:** Excluded from the PASS decision.
- **Difference:** Zero for source IDs, event count, and exposure.
- **Explanation if FAIL:** Not applicable.

## TEST 12 — End-to-end conservation

- **PASS / FAIL:** PASS
- **Input:** BUY 100 at EUR 45 on 01/07; curves EUR 47 and EUR 49; full fixing EUR 50 on 03/07.
- **Expected Fixings:** Closing volume `-100`; raw amount `-5,000`; economic contribution `+5,000`.
- **Actual Fixings:** Exact match.
- **Expected Exposure:** 01/07 `+100`, MtM `4,700`; 02/07 `+100`, MtM `4,900`; 03/07 explicit zero.
- **Actual Exposure:** Exact match.
- **Expected P&L components:** 01/07 `+200`; 02/07 `+200`; 03/07 `+100`; cumulative `+500`.
- **Actual P&L components:** Exact match.
- **Difference:** EUR `0.00`; actual cumulative `500` equals independent economics `100 × (50 − 45) = 500`.
- **Explanation if FAIL:** Not applicable.

### TEST 12 fixture note

The pack labels the position Sep-26 but fixes it on 03/07. None of the configured production
methods can produce that date relationship. The executable pipeline twin therefore uses one
WITHINDAY delivery on 03/07. It preserves the pack's Market Dates, quantity, prices, P&L components,
and independent EUR 500 conservation result without adding a test-only production methodology.

## Corrections made

1. P&L now converts raw signed fixing settlement to its inverse economic contribution exactly once.
2. A cross-layer invariant, `PNL_FIXING_SIGN`, verifies that conversion for every P&L key.
3. CSV load failures now identify the filename, CSV row, and `source_row_id`; blank Daily Qty remains distinct from explicit zero.
4. The engine and acceptance specifications now distinguish raw Fixings output from economic Fixing P&L.

## Remaining scope

This report accepts the 12 synthetic core cases. It does not claim completion of the separate
Brent/HH FX and source-column cases, product-specific DAY_AHEAD weekend cases, 13-book Initial P&L
reconciliation, workbook extraction adapter, or real-workbook diagnostic regressions.
