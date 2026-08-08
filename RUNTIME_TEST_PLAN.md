# GTM Excel Runtime Test Plan

## Objective

Establish what Microsoft Excel actually executes, confirm how the saved workbook behaves at runtime, and collect reproducible evidence before changing VBA or running a full historical build.

## Source protection

- Never open the primary source for a write-enabled test.
- Primary source: `Gas_Trading_Model 070826.xlsm`.
- Runtime copy: `runtime_tests/2026-08-07_run01/Gas_Trading_Model_070826_runtime_run01.xlsm`.
- Record SHA-256 hashes before and after every material stage.
- Keep external-link updates and connection refreshes disabled unless a later test explicitly requires them.
- Do not execute the helper workbooks against the source file.
- Stop on any unexpected external refresh, workbook-selection ambiguity, crash, hang, or mutation outside the runtime copy.

## Evidence captured for every test

- test identifier, start/end time, Excel version, workbook path, and macro name;
- workbook hash before and after the test;
- calculation mode and workbook read-only/saved state;
- visible Excel/VBA error or message text;
- relevant `MODEL CONTROL`, `MODEL VALIDATION`, `MODEL LOG`, and output metadata before and after;
- whether any output table row count, refresh timestamp, Simulation state, Build ID, or sheet inventory changed;
- outcome: PASS, FAIL, BLOCKED, or INCONCLUSIVE, with a reason.

## Phase 0 — Baseline and isolation

### RT-00: File baseline

- Verify the source SHA-256 remains `f6f173ce398109615cc2c8986c52e4feec3249d6b5ba8b15f3c6a75cc5656b31`.
- Create the numbered runtime directory and copy.
- Verify the copy initially has the same hash as the source.
- Record source/copy sizes and modification times.

Pass criterion: hashes are identical and the source remains unchanged.

### RT-01: Excel automation check

- Query the installed Microsoft Excel version.
- Confirm AppleScript can address Excel.
- Do not open a workbook or run a macro in this test.

Pass criterion: Excel responds with a version and no workbook is mutated.

## Phase 1 — Safe open and live-state capture

### RT-10: Open runtime copy without link updates

- Open only the runtime copy in Excel.
- Suppress external-link updates at open.
- Confirm the active workbook path and filename point to the runtime directory.
- Confirm sheet count, calculation mode, read-only state, and saved state.
- Record any Protected View, macro-security, link, repair, compatibility, or corruption prompt.
- Do not refresh connections and do not save yet.

Pass criterion: the runtime copy opens normally, is the sole selected GTM target, and no external data is refreshed.

Stop conditions:

- Excel repairs or removes workbook content;
- Excel opens the source rather than the copy;
- a SharePoint/external-link refresh starts;
- workbook is read-only when the test requires write access;
- unexpected workbook(s) satisfy the helpers' target-sheet signature.

### RT-11: Initial live-state snapshot

Capture before running VBA:

- `MODEL CONTROL` Build ID, preflight status, engine/reconciliation statuses;
- `PROCESS` D1, D2, historical start/end, and full-build state;
- Simulation status;
- `FIXINGS DATA`, `EXPOSURE DATA`, and `PNL DATA` metadata/row counts;
- last populated rows of `MODEL VALIDATION` and `MODEL LOG`;
- workbook/sheet inventory visible to Excel.

Pass criterion: live values are captured without recalculation or refresh.

## Phase 2 — Compilation and launcher smoke test

### RT-20: Execute `Run_GTM_v2`

Reason for selecting this procedure first: it is the advertised public launcher, but static source indicates that it should run preflight only and should not build Fixings, Exposure, or P&L.

- Record the pre-execution state and hash.
- Invoke `Run_GTM_v2` in the runtime copy.
- Capture every Excel/VBA dialog verbatim when possible.
- Dismiss informational/error dialogs without selecting any refresh or repair action.
- Observe whether the procedure returns, hangs, raises a compile/runtime error, or starts an unexpected engine.
- Save only the runtime copy after the post-state is captured.

Expected current behavior:

- preflight runs;
- validation/control/log areas may change;
- Fixings, Exposure, and P&L build timestamps and row populations do not change;
- a controlled-rollout/information message appears.

Pass criterion: actual behavior matches the expected limited launcher behavior and Excel returns control cleanly.

Fail criteria:

- compile error or unhandled VBA error;
- preflight errors are swallowed while status remains PASS;
- malformed success `MsgBox` raises an error;
- external refresh begins;
- any output engine runs unexpectedly;
- Excel hangs or crashes;
- application calculation/events/screen state is not restored.

### RT-21: Post-launcher comparison

- Compare live state with RT-11.
- Compare saved package hash and sheet XML metadata with the baseline.
- Identify exactly which sheets/parts changed.
- Confirm the source workbook hash remains unchanged.

Pass criterion: all mutations are confined to expected validation/control/log state in the runtime copy.

## Phase 3 — Targeted validation tests

These tests are gated on RT-20/RT-21. Do not continue automatically after a failure.

### RT-30: Fail-closed preflight test

- In a new copy derived from the clean baseline, introduce one deterministic invalid input.
- Run preflight once.
- Verify it returns FAIL, records the precise row/reason, and cannot leave a stale PASS.
- Restore by discarding the test copy, not by editing the invalid value back in place.

### RT-31: Mandatory Daily Qty test

- Use a populated disposable test trade row with blank Daily Qty.
- Expected: blocking validation error in preflight, Fixings, Exposure, and trade-entry paths.
- Verify no output build begins.

### RT-32: Same-day fixing regression

- Test case: CGTO / Phys PVB; Trade Date and Fixing Date 2026-07-14; Delivery Day 2026-07-15; signed volume approximately -490.37 MWh, confirmed from source.
- Expected after the structural correction: fixing row present, volume and delivery-day price correct, exposure event applied once, P&L row present, and explicit zero closure retained when applicable.
- The current code is expected to expose a `>` versus `>=` inconsistency; this test first documents the failure before repair.

### RT-33: Deferred event and zero-closure tests

- Create/identify an event dated outside the output Market Date axis.
- Confirm current behavior and then, after repair, confirm next-valid-date application exactly once.
- Confirm a closing position produces an explicit zero exposure snapshot.

### RT-34: Trade-entry adjustment idempotence

- Use a fresh P&L build copy.
- Run `Apply_GTMv2_Trade_Entry_Adjustment_V2_External` once and capture totals/audit.
- A second run is currently expected to duplicate the adjustment; do not perform it on a retained build unless a reversible test copy is used.
- After repair, the second invocation must refuse or produce zero additional change using Build ID plus adjustment version/applied state.

## Phase 4 — Coherent build gate

Do not run a full historical build until Phases 2 and 3 pass after repairs.

Required preconditions:

- clean-compiled VBA source with source/P-code ambiguity removed;
- fail-closed preflight;
- mandatory quantity rule consistent everywhere;
- same-day and deferred-event rules consistent between Fixings and Exposure;
- explicit zero closures;
- output tables unfiltered before clear/resize;
- common Build ID/coherence checks implemented;
- idempotent or integrated trade-entry adjustment;
- blocking policy for missing prices defined.

Then run in one Excel session:

1. preflight;
2. initial-position refresh only if required;
3. Fixings;
4. Exposure;
5. P&L;
6. trade-entry adjustment if still external;
7. technical/internal/economic validators.

Legacy reconciliation is diagnostic only and cannot determine acceptance.

## Current authorization boundary

The user authorized controlled runtime execution on 2026-08-07. This authorizes opening and testing an isolated copy in local Microsoft Excel. It does not authorize refreshing external SharePoint data, altering the source workbook, or running a full historical rebuild before the stated gates pass.

## Run 01 status

- RT-00: PASS.
- RT-01: PASS.
- RT-10: PASS with limitation; Excel ignored the requested read-only flag, but the isolated copy was closed without saving and remained hash-identical.
- RT-11: PASS; live-state capture identified stale validation rows and a PnL table/control-count mismatch.
- RT-20: FAIL. `Run_GTM_v2` invoked an older `PREFLIGHT` path, changed the displayed model version from alpha.3 to alpha.1, and returned `Unexpected error 0` with Preflight status FAIL and no validation-row explanation.
- RT-21: PASS for containment. Fixings/Exposure/PnL table row counts did not change; Excel application state was restored; the copy was closed without saving; source and copy hashes remained identical.
- Stop gate applied: RT-30 onward will not be run against the current VBA architecture until the execution/debug strategy is revised.
