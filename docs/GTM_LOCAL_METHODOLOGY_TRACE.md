# GTM Local Methodology Trace

Version: 0.3
Date: 2026-08-07
Status: Local evidence preserved and reconciled with later business decisions

## 1. Purpose and evidence boundary

This report records the business rules recovered from the three local macro-enabled workbooks and the technical handover. It distinguishes intended methodology from incomplete or buggy implementations.

Evidence reviewed:

- `Gas_Trading_Model 070826.xlsm`: all 47 worksheets, workbook names, tables, comments, formulas, and 68 extracted VBA components;
- `GTM_Fast_Helper.xlsm`: complete fast exposure/fixing implementation and its parity checks;
- `GTM_Trade_Entry_Helper_V2.xlsm`: trade-entry P&L adjustment implementation and diagnostics;
- `docs/GTM_v2_Technical_Handover_Report_2026-08-07.docx`;
- workbook-local source values and cached outputs used only to understand structure or identify defects.

`questions_answers.txt` and `questions_answers.txt~` are explicitly excluded from the evidence set at the user's instruction. They have not been read as methodology sources, cited, modified, or deleted.

## 2. Interpretation rule for conflicting generations

The workbook evolved before the legacy design was complete and v2 was added before the older architecture was fully retired. “Newest code wins” is therefore unsafe.

For recovered methodology, the working authority is:

1. direct user instruction;
2. the editable methodology definitions and active mappings on `SETUP`;
3. rules implemented consistently in the completed legacy exposure builder and the independent fast helper;
4. the v2 architectural intent and defect descriptions in the handover;
5. current v2 code, except where the handover or complete helpers identify it as unfinished;
6. legacy formulas only as corroborating or diagnostic evidence.

This rule does not make the legacy outputs an acceptance oracle. Many legacy formulas stop at `TRADES` row 202 while the live trade population extends to row 478.

## 3. Recovered fixing methodology

The authoritative method definitions are in `SETUP!A20:B24`. The method cells in columns I and O are explicitly commented as editable inputs that drive Fixings.

| Method | Eligible delivery days | Normal Fixing Date | Allocation |
|---|---|---|---|
| WITHINDAY | Every calendar day | Delivery Day | Daily delivery quantity fixes on the same day |
| DAY_AHEAD | Every calendar day | Delivery Day minus one calendar day | Each calendar delivery day fixes on the immediately preceding calendar date; this is not a previous-market-day rule |
| HEREN | Every calendar day | Previous configured Market Day | One Market Day can fix the next Market Day and every intervening weekend/holiday delivery day |
| MONTH_AHEAD | Every calendar day of Delivery Month | Market days in the previous calendar month | Total monthly delivery exposure is divided equally across the eligible fixing days, then distributed across the delivery days |
| BRENT_HH | Configured Market Days only | Previous configured Market Day | Monthly volume is divided by market delivery days; each market delivery day fixes on the previous Market Day |

Active main-underlying assignments are:

| Underlying | Method |
|---|---|
| Brent Dated | BRENT_HH |
| HH | BRENT_HH |
| TTF DA | HEREN |
| TTF MA | MONTH_AHEAD |
| Index PVB | HEREN |
| Phys PVB | HEREN |
| TVB | HEREN |
| AVB | HEREN |
| PEG | HEREN |

The active PVB detail assignments in `SETUP!M3:P11` retain their own fixing method but aggregate to canonical output underlying `Index PVB` in the complete fast exposure implementation:

| PVB source underlying | Method | Canonical exposure underlying |
|---|---|---|
| GWDES Auction | WITHINDAY | Index PVB |
| D+1 Auction | DAY_AHEAD | Index PVB |
| Mibgas Index ES | DAY_AHEAD | Index PVB |
| MIBGAS D+1 Daily Reference | DAY_AHEAD | Index PVB |
| MIBGAS LPI | HEREN | Index PVB |
| PVB Heren DA | HEREN | Index PVB |
| Mibgas API DA | HEREN | Index PVB |
| Mibgas MA | MONTH_AHEAD | Index PVB |
| PVB Heren DA (Delivery) | HEREN | Index PVB |

## 4. Initial position and incremental-trade boundary

- `INITIAL POSITION!C2` is an end-of-day cut-off.
- Initial exposure and Initial P&L already include all activity through that date.
- Only trades with `Trade Date > Initial Market Date` are incremental.
- Initial-position fixing eligibility is strictly `Fixing Date > Initial Market Date`.
- Incremental-trade fixing eligibility includes `Fixing Date = Trade Date`.
- Start Date and End Date are inclusive.
- A trade spanning months is split by Delivery Month.

## 5. Late and same-day trades: implementation evidence versus approved policy

The complete legacy builder and the fast helper implement the same catch-up behavior:

- for WITHINDAY, DAY_AHEAD, HEREN, and BRENT_HH, use `effective_fixing_date = max(normal_fixing_date, trade_date)`;
- if the normal fixing occurred before Trade Date, close that delivery quantity on Trade Date;
- for MONTH_AHEAD, redistribute the whole affected monthly delivery quantity equally over remaining prior-month Market Days whose date is on or after Trade Date;
- if no Month Ahead fixing day remains, close the whole affected quantity on Trade Date.

This is implementation evidence, not the release-1 rule. The later business decision overrides it:

- a normal fixing on Trade Date is valid;
- a non-Month-Ahead trade row is rejected if any included delivery slice has already lost its normal fixing opportunity;
- a Month Ahead trade uses remaining dates on or after Trade Date when they exist and is rejected when none remains;
- no catch-up fixing or Trade-Price substitution is permitted.

The existing VBA catch-up branches must therefore be covered by tests that fail under v0.3 policy, so they cannot be copied accidentally into Python or a later clean VBA port.

## 6. Market Date axis and deferred events

- Output Market Dates are configured market days strictly after Initial Market Date and within the requested historical range.
- An event retains its original economic date.
- An event whose date is not an output Market Date is applied once to the first output Market Date on or after its economic date.
- This includes trades dated on weekends or holidays.
- No event may be applied early to the previous Market Date.
- Events after the historical end remain unapplied and are reported.

The current v2 Exposure function compares `eventDate = targetDate`; this is a known bug. It must use due-event logic plus an applied-event set.

## 7. Exposure and explicit closures

For each `BOOK + Underlying + Delivery Month + Trade Source + Scenario` key:

```text
closing exposure = previous closing exposure
                 + applied signed trade volume
                 + applied opposite-signed fixing volume
```

When a previously material non-zero key becomes zero, write one explicit zero row on that closure Market Date. Do not repeat zero rows on later dates. If the key reopens, emit it again. The current v2 `CaptureSnapshot` drops every zero and is therefore incomplete.

## 8. Curve valuation

For most underlyings, Exposure MtM is `exposure volume × curve price` at Market Date and Delivery Month.

Special current-month behavior:

- Brent Dated: when Market Date and Delivery Month are in the same calendar month, use the next Delivery Month's Brent curve price;
- HH: apply the identical next-month rule;
- for later Delivery Months, use the matching Delivery Month curve;
- P&L opening Initial MtM must use the same rule as Exposure.

The current Exposure engine implements this rule. The current P&L engine does not apply it when valuing Initial Exposure, which is a defect.

### 8.1 Currency boundary found in the workbook

The workbook itself labels `FIXING PRICES` as accepting `€/MWh`, `$/bbl`, and `$/MMBtu`.
The raw `Brent Dated` and `HH` history sheets load Reuters dollar-denominated contract values,
and `LoadCurveHistory` copies those values directly into the Exposure/P&L dictionaries.
Although an `EURF` forward-curve sheet exists and is required by structural Preflight, no
reference to it was found in the v2 Exposure or P&L price-selection path. `GetCurvePrice`
returns the raw Brent/HH number.

Therefore the workbook does not establish a safe, auditable USD-to-EUR transformation for
Brent/HH Total P&L. The Python calculation core must not reproduce the implicit cross-currency
addition. A future workbook extractor must either receive an explicitly approved FX conversion
contract or emit the original currency metadata and let Preflight block an incompatible reporting
build. This is an adapter/input-normalization issue, not a reason to complicate the core engine.

## 9. Price gaps

Per the confirmed user decision, every price required by an included economic row is mandatory. The engine collects all safely discoverable missing price keys, fails the complete build, publishes no new economic output, preserves the last accepted build, and asks the Excel user to supply the missing prices. It does not substitute zero, a previous price, interpolation, or another contract. An unused missing source cell does not block the build.

## 10. Quantity and sign conventions

- Daily Qty is a non-negative magnitude.
- BUY produces positive exposure; SELL produces negative exposure.
- Fixing Volume carries the opposite sign from the exposure it closes.
- Negative Daily Qty is invalid because applying Side would reverse the sign twice.
- Blank or non-numeric Daily Qty is blocking.
- A deliberate numeric zero is accepted with a warning and produces no economic event.

The live `TRADES` population corroborates the convention: 476 populated quantities are numeric, none are negative, and 12 are explicit zeros.

## 11. P&L methodology

Daily P&L retains Market Date, Previous Market Date, BOOK, Underlying, Delivery Month, Trade Source, and Scenario for position-derived components.

```text
gross exposure movement = current Exposure MtM - previous Exposure MtM
trade-entry adjustment  = - open volume after same-day fixing × execution price
adjusted exposure P&L   = gross exposure movement + trade-entry adjustment
total daily P&L         = adjusted exposure P&L
                        + economic fixing amount
                        + logistical costs
                        + fees and optimizations
                        + replication
```

Fixings output preserves the raw signed settlement
`raw_fixing_amount = fixing_volume * fixing_price`. Since Fixing Volume closes exposure with the
opposite sign, the P&L component is `economic_fixing_amount = -raw_fixing_amount`. The approved
golden closure case requires a raw `-EUR 4,800` settlement to contribute `+EUR 4,800` to P&L.

Trade-entry adjustment must be integrated into the engine or guarded by Build ID and adjustment version so it is idempotent. Under v0.3, “same-day fixing” means a valid normal fixing, not a catch-up fixing.

## 12. Operating-flow source and allocation map

The local workbook contains an exact source map:

| Component | Daily source | Recognition and sign | Output allocation |
|---|---|---|---|
| Logistical Costs | `COSTS!B:N`, matched by Market Date and BOOK header | Convert the stored cost amount to P&L sign as the legacy `PNL` formulas do: `-COSTS` | BOOK level; Underlying=`TOTAL / BOOK LEVEL`; Delivery Month blank; no unsupported allocation |
| Fees and Optimizations | `COSTS!O:P` plus `Foto FO!R:S`, matched by Market Date and BOOK | Add values as stored; the Foto FO refresh creates fee values with their P&L sign | BOOK level; Underlying=`TOTAL / BOOK LEVEL`; Delivery Month blank |
| Replication | `Foto FO!T:V`, matched by Market Date and BOOK | Add values as stored | BOOK level; Underlying=`TOTAL / BOOK LEVEL`; Delivery Month blank |

These are daily flows and must not be differenced again. The older snapshot macro subtracts cumulative values because it reads cumulative snapshots; `modDailyBasesIndependent_v11` explicitly identifies the direct daily-flow treatment.

There is one acceptance caution rather than a missing source rule: `COSTS!R4` says `CONFIRMAR SIGNO CON OPS`. The engine can implement the workbook's explicit sign transform, but a golden Ops example must certify that sign before production Total P&L approval. The current independent-daily VBA adds raw Logistics values and appears inconsistent with the legacy P&L sign formula.

## 13. Initial P&L bridge and newly found extraction defect

`INITIAL POSITION!A7:B19` contains one signed closing P&L balance for each of the 13 active books at Initial Market Date. It is separate from Initial Exposure and therefore is not opening Exposure MtM.

Required treatment:

- keep daily post-cut-off P&L separate;
- expose the Initial P&L as a book-level opening balance;
- calculate cumulative P&L as `initial_pnl + sum(daily_pnl after cut-off)`;
- do not inject Initial P&L as a daily flow and do not add it to opening Exposure MtM.

The saved `tblInitialPnL` incorrectly contains 19 rows. Only the first 13 are books. Six documentation lines (`Description` plus five fixing-method descriptions) were ingested as zero-P&L books because `ReadPnL` scans continuously until the later `TOTAL` row and does not validate against active SETUP books. The refresh must read exactly the active-book table or stop before the documentation block. Preflight must reject any Initial P&L book not active in SETUP.

The current v2 P&L engine validates `tblInitialPnL` but never consumes it. Cumulative reporting therefore omits the opening bridge.

## 14. Precision and comparison

- Use Excel/Python double precision through the calculation; do not round intermediate economics.
- Use `1e-7` as the current materiality threshold for emitting non-zero engine events/rows.
- For cross-implementation comparisons, use the approved measure-specific tolerances:

```text
Volume = 0.000001
Price  = 0.00000001
P&L    = EUR 0.01
```

- Formatting is presentation only: normalized volume/price can display six decimals and monetary output two decimals without altering stored values.

## 15. Performance

No SLA was encoded in the local artifacts, but later business guidance approved engineering targets: full rebuild under 60 seconds, daily/incremental update ideally under 10 seconds, focused tests under 2 seconds, and peak Python memory below 2 GiB on the agreed current-Mac benchmark.

## 16. Defects exposed by the recovered methodology

1. The old advertised launcher calls obsolete preflight code and can fail without actionable detail.
2. Preflight source/comments/logging contradict each other about blank Daily Qty; the handover and calculation engines require blank quantities to block.
3. Existing completed helper/legacy catch-up behavior conflicts with the approved v0.3 late-trade rejection policy and must not be ported.
4. Current v2 Exposure loses non-market-date events because it applies only exact-date events.
5. Current v2 Exposure omits explicit zero closures.
6. Current v2 P&L values Brent/HH opening exposure using Delivery Month rather than the Exposure engine's current-month next-contract rule.
7. Current v2 P&L ignores Initial P&L balances.
8. `ReadPnL` pollutes `tblInitialPnL` with six documentation rows.
9. Current direct-daily operating-flow code appears to omit the Logistics cost-to-P&L sign conversion.
10. Helper trade-entry adjustment is not idempotent and must be integrated or guarded.
11. Several current outputs belong to different build sessions; a shared Build ID and coherence gate are required.
12. Brent/HH raw prices are dollar-denominated, but the v2 Exposure/P&L path does not use the present `EURF` sheet before producing reporting totals.

## 17. Conclusion

The broad business-method questions D-002 through D-012 did not require a new questionnaire. Local artifacts provide the core calculation rules, mappings, initial-balance meaning, and operating-flow sources; later business guidance deliberately replaces the observed late-trade catch-up branches. The remaining work is implementation, targeted tests, and economic acceptance of small worked examples—not rediscovery of the model.
