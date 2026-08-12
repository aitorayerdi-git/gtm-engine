# Project Journal

## Standing architectural context supplied by the user

- The mixed legacy/v2 architecture is the result of premature evolution: the legacy implementation was not complete when v2 development began.
- Accordingly, overlapping modules, incomplete handoffs, duplicate calculations, and unfinished reconciliation paths must be interpreted as migration debt between two unfinished generations—not as a clean replacement of a stable legacy baseline.
- The system boundary includes three macro-enabled workbooks:
  - `Gas_Trading_Model 070826.xlsm`: main trading model;
  - `GTM_Fast_Helper.xlsm`: additional helper workbook;
  - `GTM_Trade_Entry_Helper_V2.xlsm`: additional trade-entry helper workbook.
- The project also includes `docs/GTM_v2_Technical_Handover_Report_2026-08-07.docx`, which is a required source for intended architecture, operating procedure, known limitations, and development history.
- Findings in this journal will distinguish among: intended behavior documented in the handover; visible VBA source behavior; saved workbook state; and behavior confirmed by controlled execution in Microsoft Excel.

## 2026-08-07 — Initial workbook and VBA inspection

### Scope and working rule

- Project objective supplied by the user: build a reliable system of Excel/VBA macros for gas trading.
- Current task: read the existing macro-enabled workbook first and report its functionality, defects, structure, risks, and other material findings.
- Standing project rule: maintain this journal with detailed, cumulative updates as investigation and implementation proceed.
- Source workbook: `Gas_Trading_Model 070826.xlsm`.
- Inspection is read-only. The source workbook has not been modified.

### Source identity and package inventory

- File type: Microsoft Excel Open XML macro-enabled workbook (`.xlsm`).
- SHA-256: `f6f173ce398109615cc2c8986c52e4feec3249d6b5ba8b15f3c6a75cc5656b31`.
- The Open XML package contains 47 worksheet XML parts, an embedded `xl/vbaProject.bin`, two external-link parts, a calculation chain, seven Excel table definitions, comments, VML drawings, and volatile-dependency metadata.
- The package expands to approximately 411 MB. Two worksheet XML parts are exceptionally large: `sheet10.xml` is approximately 242 MB and `sheet11.xml` is approximately 69 MB. Other large parts include `sheet32.xml` at approximately 42 MB and `calcChain.xml` at approximately 9 MB. This is a material performance, maintainability, and corruption-risk signal and will be investigated further.
- The embedded VBA project is approximately 4.0 MB.

### VBA inventory

- 68 VBA components were detected in total. Most are empty worksheet or workbook class modules.
- 19 standard modules contain substantive code, totalling approximately 15,900 source lines and 12,328 non-comment/non-attribute code lines.
- 23 public entry points were identified.
- Main GTM v2 modules:
  - `modGTMv2EngineCore`: common engine helpers plus the public `Run_GTM_v2` launcher. Despite its name and comments, the inspected source launcher performs preflight only; it does not run the fixings, exposure, or P&L engines.
  - `modGTMv2ExposureEngineV2`: construction of normalized exposure data.
  - `modGTMv2FixingsEngineV4`: construction of normalized fixing data.
  - `modGTMv2PnLEngine`: P&L construction from exposure, fixings, price curves, and book-level flows.
  - `modGTMv2PnLReconciliationV3`: reconciliation of new P&L output against legacy output.
  - `modGTMv2PreflightFixV2`: validation and repair of required workbook structures.
  - `modNewArchitectureSetup`: creation and status management for the newer model architecture.
- Supporting/newer modules:
  - `modBuildFixingsD2Data4`: builds a D2-format fixings dataset.
  - `modBuildDeltaPnLByMarketDateD21`: builds delta P&L by market date.
  - `modDailyBasesIndependent_v11`: rebuild/update/report/clear operations for independent daily bases.
  - `Módulo9`, `Módulo10`, `Módulo11`, `modRefreshInitialPositionData`, and `modRefreshDailyExposureData1`: initial-position setup and refresh flows plus alternate exposure/fixings rebuilds.
- Legacy/utility modules:
  - `Módulo2`: legacy exposure calculation.
  - `Módulo3`: refreshes the FO snapshot.
  - `Módulo4`: refreshes PVB/PEG/TTF data.
  - `Módulo8`: audits external links and related workbook objects.

### Initial safety and integrity findings

- Static VBA analysis detected file/IO, environment-variable, command-execution, and `CreateObject` keywords. These are not by themselves evidence of malicious behavior: some are expected for file selection, SharePoint/CSV refreshes, and late-bound dictionaries. Each concrete use is being traced.
- Two SharePoint URLs under `https://cepsacorp.sharepoint.com/sites/Shared-CGC2/` were present in the VBA project.
- Static analysis reported that VBA source and compiled P-code differ (commonly called “VBA stomping”). This detector is experimental and old edited workbooks can produce false positives, but the mismatch means source-only review is not sufficient to prove what Excel will execute. P-code comparison and a clean recompile/export are recommended before trusting the workbook in production.
- `ThisWorkbook` contains no event code, and the inspected worksheet class modules are empty. No auto-open macro was detected in source. The main operations appear to require explicit macro execution or assignment to workbook controls/shapes.

### Inspection status

- Completed: package inventory, source hash, VBA component extraction, module/procedure inventory, first-pass security scan.
- Completed: workbook sheet/name/link mapping, GTM v2 call graph, first source-versus-P-code comparison, cached-output inspection, and first concrete compile/runtime/financial-logic defect review.
- Not yet performed: opening or recalculating the workbook in Excel, executing any macro, refreshing an external connection, or changing any workbook cell or VBA module.

## 2026-08-07 — Functional map and defect report

### What the workbook is designed to do

The workbook combines two overlapping generations of gas-trading calculations.

1. Legacy macros refresh source sheets and produce exposure, fixings, bases, and daily P&L outputs directly in workbook tables.
2. The newer GTM v2 architecture attempts to normalize those calculations into three engines:
   - a fixings engine that constructs `FIXINGS DATA`;
   - an exposure engine that constructs `EXPOSURE DATA`;
   - a P&L engine that constructs `PNL DATA` and compares it with the legacy `DAILY PNL DATA` result.
3. Preflight macros validate model structures, date inputs, prices, and trade rows and then write PASS/FAIL information to `MODEL CONTROL`, `MODEL VALIDATION`, and `MODEL LOG`.
4. Separate reconciliation routines compare new outputs with legacy outputs.
5. Refresh utilities retrieve or copy data for the FO snapshot, hub curves such as PVB/PEG/TTF, initial positions, and daily exposure.

No source-level `Workbook_Open` or worksheet event was found. The workflows appear to start only when a user explicitly runs a public procedure or invokes an assigned workbook control.

### Actual v2 execution topology

- `Run_GTM_v2` is not an end-to-end orchestrator. It invokes the older preflight entry point and then displays a controlled-rollout message. It does not invoke the fixings, exposure, or P&L builders.
- The substantive builders must currently be invoked individually:
  - `Build_GTMv2_Fixings_Data_v3`;
  - `Build_GTMv2_Exposure_Data_v2`;
  - the GTM v2 P&L builder in `modGTMv2PnLEngine`;
  - the separate V3 P&L reconciliation procedure when a diagnostic comparison is required.
- All three v2 engines invoke `Run_GTMv2_Preflight_Alpha3` indirectly through `Application.Run`.
- Each engine surrounds that preflight call with `On Error Resume Next` and then relies on the pre-existing PASS/FAIL cell in `MODEL CONTROL`. If preflight raises an error while that cell still contains PASS from an earlier run, the engine can continue after a failed validation attempt. This is a critical stale-state defect.
- Dependencies are checked primarily through the simulation-state text in cell `B2` of the upstream output. The code does not consistently require matching build IDs, timestamps, source-input ranges, reconciliation status, or freshness.

### Current saved workbook state

These observations are from stored cell/formula caches and metadata. Excel has not been opened or recalculated, so live values may change on a real calculation or refresh.

- Workbook calculation mode is saved as Manual, with `forceFullCalc` enabled. Users can therefore see internally inconsistent cached results until a deliberate full calculation is completed.
- `PROCESS!C10`, the FO-update control cell, contains cached `#VALUE!`.
- The process sheet reports `FULL REBUILD REQUIRED`, because the current simulation status is OFF while the last full-build status is `NOT BUILT`.
- `MODEL CONTROL` currently shows a preflight PASS and PASS for the saved fixings and exposure reconciliations, but FAIL for the P&L reconciliation.
- Stored output-refresh dates do not describe one coherent build:
  - `FIXINGS DATA`: 2026-08-01;
  - `EXPOSURE DATA`: 2026-07-31;
  - `PNL DATA`: 2026-08-03;
  - latest saved preflight: 2026-08-07.
- The most recent saved V3 P&L diagnostic is FAIL, with 1,082 mismatched keys, an absolute difference of approximately 91,993,228.30, and a signed economic difference of approximately 6,329,027.57.
- The saved P&L engine log reports 3,429 rows and 68 book-level fallback rows. The log contains repeated P&L failures beginning on 2026-07-25.

### Cached formula-error inventory

The following counts describe cached error cells in the saved workbook and are signals for diagnosis, not results of a fresh Excel recalculation:

- `Foto FO`: 4,551 cached errors, comprising 4,503 `#VALUE!` and 48 `#N/A`; approximately 3,576 cells contain external-workbook formulas.
- `TTF`: 322 cached errors, mainly `#DIV/0!`.
- `PVB-TTF`: 846 cached errors and approximately 8,457 external-workbook formula cells.
- `PEG-TTF`: 846 cached errors and approximately 8,460 external-workbook formula cells.
- `Historical_DA`: 1,760 cached errors.
- `PROCESS`: one cached `#VALUE!` at `C10`.

### Data-volume and performance findings

- `FIXINGS D2 by BOOK Calc` is the largest worksheet part: approximately 230.6 MiB of XML, 286,299 populated cells, and 271,471 formula cells.
- `FIXINGS D2 DATA` is approximately 66.1 MiB of XML, with roughly 1,540,194 populated cells across 256,699 rows.
- `DAILY EXPOSURE DATA` is approximately 40.0 MiB of XML, with roughly 943,496 populated cells across 117,937 rows.
- `calcChain.xml` is approximately 9 MiB. Combined with Manual calculation, external links, and hundreds of thousands of formulas, this makes full recalculation and file saving expensive and increases the likelihood of stale displayed state.

### External dependencies

- Two saved external-workbook links point to CEPSA SharePoint resources:
  - `Foto FO.xlsx`;
  - `LOCATIONSP.xlsx`.
- The FO and market-data refresh modules can change the workbook from external state rather than only from local inputs.
- A controlled test must therefore use a copy and explicitly decide whether links/connections are disabled, mocked, or allowed to refresh.

### Critical preflight defects

1. **Preflight exceptions are swallowed by the callers.** As noted above, a stale PASS can authorize downstream execution after `Application.Run` failed.
2. **Blank Daily Qty has contradictory implementation/documentation remnants.** The handover establishes the current business rule: a populated trade row must have numeric Daily Qty, and blank/non-numeric values must block the build. Current blocking code is therefore directionally correct. However, comments and saved validation text still say blank quantity is accepted as zero, and `MODEL VALIDATION` contains both rules in different rows. Those stale statements and any helper paths that silently coerce blank quantity to zero must be removed.
3. **The trade-row conditional is closed too early.** In `modGTMv2PreflightFixV2`, an `End If` closes the `If Not TradeRowBlank` block before the price/date/book and other field validation. Those checks can consequently execute for blank rows and can reuse stale `reasonText` state.
4. **The blank-quantity counter is never incremented.** `blankQtyRows` is declared and displayed but the inspected source never increases it, so the report always says zero.
5. **The success message call is malformed in source.** The line continuation concatenates `vbInformation` into the prompt and passes the title string where the buttons argument belongs. A clean source recompile is expected to reveal a type/argument problem here.
6. **Unexpected-error handling can preserve stale PASS.** The handler restores application state but does not reliably change `MODEL CONTROL!B6` to FAIL or write a validation/log record.
7. **Output-table checks validate column counts, not semantic headers.** A table with the expected number of columns but incorrect/reordered meanings may pass.
8. **Curve errors are warnings only.** Missing or erroneous price-curve inputs do not necessarily stop a build, even when they materially affect valuation.
9. **Preflight clears shared diagnostics.** It clears the prior `MODEL VALIDATION` body before writing its own rows, which can erase engine and reconciliation evidence while other status cells still retain FAIL.

### Fixings-engine findings

- Source comments identify a different public entry-point name and alpha version from the actual procedure and version constant. This creates operational ambiguity for buttons, documentation, and `Application.Run` calls.
- The current saved run reports 150 missing fixing-price keys while fixings reconciliation is PASS.
- Missing prices are represented with zero-valued amounts; pricing completeness is not a blocking condition.
- The reconciliation emphasizes selected D2 volumes/aggregates and does not prove complete historical pricing correctness.
- Module-level diagnostic counters are not consistently reset at the start of each run, so repeat runs in one Excel session can accumulate stale counts.

### Exposure-engine findings

- The source header/comment, actual public entry point, and alpha version constant disagree.
- Blank Daily Qty blocks this builder, which agrees with the current business rule in the handover. The defect is inconsistent enforcement and stale workbook guidance, not the blocking behavior itself.
- Missing curve-price keys are valued at zero and treated as a warning rather than a blocking financial-data error.
- Upstream fixings are checked by simulation state but not by build freshness or successful reconciliation identity.
- Its module-level missing-price counter is not consistently reset between runs in the same VBA session.

### P&L-engine and reconciliation findings

1. **The saved output is materially different from legacy.** The current V3 diagnostic has 1,082 mismatch keys and differences measured in millions. The legacy benchmark is independently confirmed to truncate many formulas at TRADES row 202, so this is a diagnostic difference—not proof that v2 is economically wrong or a valid acceptance failure.
2. **The new P&L calculation is not independent of the legacy result.** Logistics, fees, and replication components are copied from legacy `DAILY PNL DATA` as fallback values. Their exact reconciliation therefore does not independently validate those calculations.
3. **Fixing dates appear to require an exact output market-date match.** Non-market calendar fixing dates can be omitted instead of being rolled into the next market date. The separate V3 reconciler describes legacy non-market rows as rolled forward, creating a likely semantic mismatch.
4. **Two reconciliation definitions disagree.** The engine's internal key includes the previous market date, while the separate V3 reconciliation explicitly excludes it as metadata.
5. **The documented pseudo-book mapping is absent in the inspected legacy-fixing path.** The code keys the supplied book value directly.
6. **Missing opening-curve prices are valued at zero and reported as warnings.** This can generate artificial daily deltas rather than stopping on incomplete market data.
7. **Freshness is not enforced.** The P&L builder can use saved exposure and fixings outputs from different dates, as the present workbook demonstrates.
8. **Module-level missing/fallback counters are not fully reset for repeat runs.** Diagnostics may depend on whether Excel was restarted.

### Legacy/helper macro findings

- `Build_FIXINGS_D2_DATA` hard-codes the book area `B3:B15` and underlying columns `B:S`.
- It regards nonblank book labels as active without consistently applying the adjacent Active flag.
- It deletes and recreates the entire output worksheet, so formulas, formatting, names, controls, or user additions on that sheet can be lost.
- Missing prices silently create zero amounts.
- It forces calculation, events, and screen updating to particular final values instead of restoring the caller's prior application state. `Build_Delta_PnL_by_Market_Date_D2` has a similar state-restoration risk.

### VBA source-integrity finding

- Static analysis flags a source/P-code mismatch. The extracted compiled-code representation differs substantially from the visible source.
- This can result from VBA stomping, stale compiled code after editing, or limitations in the experimental decoder. It is not proof of malicious intent, but it means the visible source cannot yet be assumed to be exactly what Excel will execute.
- The correct remediation path is to export all modules, inspect them, import them into a clean trusted workbook/project, compile explicitly in the VBA editor, and then compare behavior in an isolated test copy.

### Risk-ranked conclusion

- **Critical:** swallowed preflight failures plus stale PASS state; source/P-code uncertainty; no true end-to-end orchestrator; no completed coherent build with independent economic acceptance.
- **High:** missing prices converted to zero; output freshness/build identity not enforced; P&L dependence on legacy fallback data; same-day/event-timing inconsistencies; likely market-date roll mismatch; destructive sheet rebuilds. The large legacy reconciliation difference belongs here as a diagnostic backlog, not as an authoritative acceptance result.
- **Medium:** contradictory blank-quantity rules; shared diagnostic clearing; counters not reset; hard-coded ranges; version/entry-point drift; poor restoration of Excel application state.
- **Operational:** very large sheets, Manual calculation, numerous external formulas, SharePoint dependencies, and mixed-age outputs make the workbook slow and allow apparently plausible but stale results.

### Initial repair direction

Before adding trading features, the safest sequence is:

1. preserve the original and establish a disposable test copy;
2. export and clean-recompile the VBA project to remove source/P-code ambiguity;
3. make preflight fail closed and eliminate reliance on a stale status cell;
4. enforce the established mandatory numeric Daily Qty rule everywhere, and define authoritative blocking/roll rules for missing prices and non-market fixing dates;
5. add immutable build IDs/timestamps and require matching, reconciled upstream builds;
6. build one explicit orchestrator with deterministic phase order and error propagation;
7. make application-state restoration exception-safe;
8. establish test fixtures and reconciliation tolerances before replacing legacy fallback calculations.

## 2026-08-07 — Local Excel execution capability

- Microsoft Excel is installed on the local Mac and AppleScript automation is available.
- The workbook can therefore be exercised in Microsoft's real calculation/VBA engine rather than through a compatibility implementation.
- No execution has yet been performed. Running a macro is a separate controlled-test phase because it can mutate workbook data, delete/recreate sheets, recalculate large outputs, and contact external SharePoint sources.
- Proposed execution protocol: duplicate the file, record its starting hash, launch the copy with link/refresh behavior controlled, run one named entry point at a time, capture Excel/VBA errors and relevant status sheets, save only the test artifact, and compare ending files and output values with the baseline.

## 2026-08-07 — Expanded system review: helper workbooks and handover

### Artifact identities

- `GTM_Fast_Helper.xlsm`
  - size: approximately 257 KiB compressed;
  - SHA-256: `1d80c21339a45ab148c7a53f661e60348a9d46350c7b6590ace0c10fa708667d`;
  - one blank worksheet (`Hoja1`) and a VBA project; no external-link package parts.
- `GTM_Trade_Entry_Helper_V2.xlsm`
  - size: approximately 132 KiB compressed;
  - SHA-256: `3bdeff531141dad4b4d1208ce13848399845ede5bc97d8e123572e1fcaa9526c`;
  - one blank worksheet (`Hoja1`) and a VBA project; no external-link package parts.
- `docs/GTM_v2_Technical_Handover_Report_2026-08-07.docx`
  - size: approximately 47 KiB;
  - SHA-256: `123ba79cb7f24071eb01faa6279624b197c740ebf25232de3b9ccf86e138838c`;
  - full document text reviewed; no tracked changes, comments, footnotes, endnotes, or embedded media were detected.

Both helpers are intentionally blank macro containers. They do not contain model data and do not open a model file themselves. Their source locates an already open workbook by required worksheet names, excludes the helper's own `ThisWorkbook`, and writes directly into the selected target workbook. No `Workbook_Open`/`Auto_Open` code was found, and the inspected source does not invoke a shell or open arbitrary files. Late-bound `CreateObject` usage is for `Scripting.Dictionary`.

### `GTM_Fast_Helper.xlsm` — functional inventory

The helper contains four substantive standard modules and four public operations:

1. `Rebuild_Exposure_And_Daily_Exposure_Fast_External`
   - a performance rewrite of the unfinished legacy exposure rebuild;
   - reads `INITIAL POSITION DATA`, ACTUAL `TRADES`, `SETUP`, `CALENDAR`, and `PROCESS` from the open target model;
   - writes the legacy-format `EXPOSURE` blocks and the full `DAILY EXPOSURE DATA` sheet;
   - replaces repeated full-dictionary scans and cell-by-cell reads with arrays, indexed market days, memoized previous-market-day lookup, and dictionary buckets.
2. `Validate_Fast_Vs_Existing_Daily_Exposure_External`
   - snapshots the currently saved legacy-format exposure outputs into memory;
   - runs the fast rebuild;
   - compares row/key coverage and values at strict tolerance.
3. `Apply_GTMv2_Trade_Entry_Adjustment_External`
   - the first external trade-entry adjustment implementation;
   - superseded by the separate V2 helper because it omitted valid trade rows/audit cases.
4. `Diagnose_CGTO_TTF_Trade_Entry_External` and `Diagnose_CGTO_TTF_Trade_Entry_V2_External`
   - development diagnostics hard-coded to CGTO / TTF-related cases on 2026-07-09;
   - write diagnostic sheets into the primary workbook and are not general validators.

### `GTM_Fast_Helper.xlsm` — material cautions

- The fast exposure rebuild is a legacy reconstruction/regression tool, not the GTM v2 normalized Exposure engine. It reads ACTUAL `TRADES` only and does not include `SIMULATION TRADES`.
- It writes directly to the live `EXPOSURE` and `DAILY EXPOSURE DATA` outputs.
- The procedure named `Validate_Fast_Vs_Existing_Daily_Exposure_External` is destructive: it saves the old output only in memory, runs the fast builder over the target sheets, and does not restore the original output after reporting the comparison. If the rebuild fails after partial writes, there is no transaction/rollback.
- `ReadTradesFast` uses `Num()` for Daily Qty, which silently converts blank and non-numeric values to zero, and it converts an invalid BUY/SELL value to zero signed quantity. This conflicts with the now-established blocking validation rule.
- The fast fixing generator generally preserves same-day trade fixings: Month Ahead uses `>= Trade Date`, and other methods clamp a fixing date earlier than the trade date up to the trade date. That makes it semantically different from the strict `> Trade Date` logic still present in the v2 engines/helper paths.
- Target-workbook resolution happens before Excel application-state values are captured. If resolution itself fails, the error path attempts to restore uninitialized `oldCalc`, `oldEvents`, and `oldScreen` values and can obscure the original error.
- The validator proves equivalence to the saved legacy output under its chosen rules; it cannot prove independent economic correctness, particularly because legacy is now known to omit later trade rows.

### `GTM_Trade_Entry_Helper_V2.xlsm` — functional inventory

The workbook contains one substantive module and one public entry point:

- `Apply_GTMv2_Trade_Entry_Adjustment_V2_External`
  - requires an already built `tblPnLData` in an open GTM workbook;
  - reads ACTUAL trades row by row and includes simulations only when the global Simulation status is ON;
  - requires numeric Daily Qty and execution Price for rows it processes;
  - calculates the trade-date open volume by delivery month and fixing method;
  - computes execution adjustment as `-Open Volume × Execution Price` using signed open volume;
  - adds the adjustment to `Delta Exposure MtM` and `Total PnL` for an existing economic key, or appends a dedicated 16-column P&L row when the key is absent;
  - clears/recreates the `TRADE ENTRY ADJUSTMENTS V2` audit sheet contents and writes row-level audit details.

### `GTM_Trade_Entry_Helper_V2.xlsm` — material defects and limitations

1. **Not idempotent.** The source explicitly warns not to run it twice. There is no PNL Build ID, adjustment-version marker, already-applied flag, or rollback. A second run adds the same economics again.
2. **Same-day eligibility remains strict.** `CalculateOpenTradeVolume` uses `Fixing Date > Trade Date` for daily methods and Month Ahead. This conflicts with the handover's target rule that incremental trades are eligible when `Fixing Date >= Trade Date` and with the stated requirement that Fixings and Exposure share one rule. The trade-entry economics must be tested against the same-day regression case before this condition is changed blindly.
3. **No coherence gate.** It does not verify upstream Build ID, PNL build timestamp/range, Initial Market Date identity beyond reading the cell, upstream reconciliation, or adjustment version.
4. **Target identity is structural only.** It matches required worksheet names and rejects multiple candidates, but does not verify an expected filename/model ID/version.
5. **Application-state capture is late.** `ResolveTargetWorkbook` runs before the old calculation/events/screen values are stored, so a target-resolution error can be mishandled during restoration.
6. **Partial updates are possible.** PNL rows are modified in place before the audit is written. A later error has no transaction or automatic restoration.
7. **Simulation status is weakly validated.** Only exact `ON` activates simulation reads; unexpected values are effectively treated as OFF and can still be stamped into appended rows.

### Source/P-code status of the helpers

- The same experimental static detector flags source/P-code mismatch in both helper files.
- The disassembler produces malformed procedure signatures for portions of the compiled streams, including nonsensical identifiers in the V2 helper. This may be stale compiled state or a modern Mac Excel decoder limitation, but it reinforces the need to export/import and compile cleanly before relying on source-level repairs.

### Handover report — intended architecture and authoritative context

The report describes an unfinished migration rather than a completed legacy-to-v2 replacement. Its target processing sequence is:

1. preflight;
2. refresh initial position only when the cut-off/source changes;
3. build normalized `FIXINGS DATA`;
4. build normalized `EXPOSURE DATA`;
5. build normalized `PNL DATA`;
6. apply the temporary external trade-entry adjustment;
7. run independent technical and economic validation.

The report establishes these current business/design decisions:

- Initial Market Date is an end-of-day position; only trades with Trade Date after it are incremental.
- Daily Qty is mandatory on a populated trade row; blank/non-numeric values are blocking.
- ACTUAL and SIMULATION source identity and Scenario must be retained.
- Fixings and Exposure must share signed-volume, fixing-eligibility, and event-timing rules.
- Trade-entry execution-price economics are necessary but temporarily external.
- Operating P&L components remain transitional/legacy-sourced.
- Legacy reconciliation is informational/diagnostic, not an acceptance gate.
- Production acceptance requires technical consistency, internal model consistency, and independent economic recomputation.

### Handover report — version and evidence reconciliation

- The report names `Gas_Trading_Model 030826.xlsm`, while the supplied current primary workbook is `Gas_Trading_Model 070826.xlsm`. It must therefore be read as handover/design history from an earlier named build.
- Its main date parameters do match the supplied `070826` workbook exactly:
  - Initial Market Date: 2026-06-30;
  - historical range: 2026-06-30 through 2026-07-28;
  - comparison dates: D1 2026-07-27 and D2 2026-07-28.
- The current workbook also contains the helper-created audit/diagnostic sheets named in the report, indicating that those external tools were used against this model lineage.
- The report's most consequential finding is independently confirmed in the current file: extensive legacy formulas reference only `TRADES!$...$3:$...$202`, while the current TRADES table/filter area extends to approximately row 506. The truncation is present at very large scale in `FIXINGS D2 by BOOK Calc` and is also present throughout `PNL` and control formulas.
- Therefore, extending those formula ranges is not a suitable remediation: it would retain the incomplete legacy design while further increasing the already extreme calculation load.
- The saved V3 reconciliation FAIL must be preserved as evidence but cannot certify or reject v2. Known independent problems—missing P&L keys, disappearing positions, same-day eligibility, stale mixed-age builds, zero closures, and curve completeness—must be tested directly.

### Handover remediation backlog incorporated into the project

Priority items from the report, now part of the working defect inventory:

1. change incremental trade fixing eligibility from `>` to `>=` consistently in Fixings and Exposure, retaining strict `>` for the end-of-day initial position;
2. defer non-output-date events to the next valid Market Date and apply each event exactly once;
3. retain explicit zero rows when an exposure position closes;
4. propagate one common Build ID and enforce coherent Simulation status, Initial Market Date, and historical range across Fixings, Exposure, P&L, and trade-entry adjustment;
5. make trade-entry adjustment idempotent or integrate it into the P&L engine;
6. unfilter every ListObject before deleting or resizing output rows;
7. confirm the Brent/HH opening-valuation curve rule;
8. retest the known disappearing PROP.TRADING / TTF DA / Mar-27 position and missing CGTO Jul-27 P&L keys;
9. replace binary legacy PASS/FAIL semantics with diagnostic classifications;
10. run the deterministic CGTO / Phys PVB same-day case before any full rebuild.

### Revised test order

The next execution phase should not begin with a full historical run. The evidence-backed order is:

1. clean-copy and clean-compile gate for all three VBA projects;
2. unit/controlled test of preflight failure propagation and mandatory Daily Qty;
3. same-day CGTO / Phys PVB regression test;
4. deferred-event and explicit-zero-closure tests;
5. trade-entry idempotence test;
6. only then one coherent single-session historical build with common Build ID;
7. independent source-to-output economic checks, with legacy used only to help explain differences.

### Document-rendering note

- The report text and OOXML structure were reviewed completely.
- Visual DOCX rendering was attempted with the packaged renderer but could not start because its `pdf2image` runtime dependency is unavailable in this environment. This does not affect the extracted technical content; no edited DOCX is being delivered.

## 2026-08-07 — Runtime execution authorized

- The user authorized proceeding with controlled Microsoft Excel runtime testing.
- The detailed executable runbook is maintained in `RUNTIME_TEST_PLAN.md`.
- The authorized boundary is an isolated local copy with external-link updates and SharePoint refresh disabled.
- The original workbook, external data sources, and full historical rebuild remain outside the first smoke-test stage.
- First live macro target: `Run_GTM_v2`, because it is the advertised launcher and static inspection predicts that it performs preflight only.

### Runtime run 01 — baseline evidence

- Runtime directory: `runtime_tests/2026-08-07_run01`.
- Runtime workbook: `Gas_Trading_Model_070826_runtime_run01.xlsm`.
- Source and copy both have size 24,315,176 bytes.
- Source and initial copy hashes are identical: `f6f173ce398109615cc2c8986c52e4feec3249d6b5ba8b15f3c6a75cc5656b31`.
- Microsoft Excel for Mac responded successfully to automation and reports version `16.111.3`.
- Excel created a default blank `Book1` when first launched; it was closed without saving before the GTM copy was opened.
- RT-00 result: PASS.
- RT-01 result: PASS.

### Runtime run 01 — RT-10/RT-11 safe-open evidence

- The isolated workbook was opened in Excel with automation security forced to disable macros, application events disabled during open, alerts suppressed, and `update links = do not update links`.
- Excel opened the correct runtime path and reported:
  - workbook name `Gas_Trading_Model_070826_runtime_run01.xlsm`;
  - 47 worksheets;
  - VBA project present;
  - calculation mode Manual;
  - saved state True.
- Although the AppleScript open command requested read-only mode, Excel reported `read only = false`. Runtime isolation must therefore depend on the numbered copy and explicit no-save/controlled-save behavior, not on the AppleScript flag.
- No repair, compatibility, link-update, corruption, or macro-security dialog was emitted during this controlled open.
- Live table counts before any macro execution:
  - `tblFixingsData`: 10,278 rows;
  - `tblExposureData`: 3,399 rows;
  - `tblPnLData`: 3,578 rows.
- `MODEL CONTROL` reports PnL rows = 3,429, while the live `tblPnLData` contains 3,578 rows: a 149-row metadata discrepancy. This is consistent with helper-appended rows not updating the engine's control count, but the exact provenance must be verified.
- Live metadata confirms mixed build ages:
  - Fixings refresh: 2026-08-01 18:06:28;
  - Exposure refresh: 2026-07-31 10:28:31;
  - PnL refresh: 2026-08-03 18:01:21.
- Live output status remains Fixings PASS, Exposure PASS, and PnL legacy reconciliation FAIL.
- `MODEL VALIDATION` contains stale rows beneath the latest PASS result:
  - rows 48-54 belong to build `GTM2-20260807-164925`;
  - row 55 and row 56 still belong to the earlier failed build `GTM2-20260807-163437`;
  - rows 57-59 contain its unkeyed SUMMARY values.
- The current validation row 52 correctly states that Daily Qty is mandatory, but recent `MODEL LOG` entries still state `blank Daily Qty accepted as zero`. This is a live, user-visible contradiction between validation detail and log narrative.
- The workbook was closed without saving. Source and runtime-copy hashes remained identical to baseline.
- RT-10 result: PASS with limitation (read-only request ignored).
- RT-11 result: PASS; two new state-integrity defects recorded (stale validation rows and stale PnL row-count metadata).

### Runtime run 01 — RT-20/RT-21 launcher smoke-test evidence

- The copy was reopened with `msoAutomationSecurityLow`, workbook events disabled during open, and external-link updates disabled. Events were restored before the macro invocation.
- `Run_GTM_v2` was invoked successfully through Excel's native `run VB macro` automation command.
- The macro completed, but the test result was FAIL:
  - `MODEL CONTROL` displayed model version `2.0.0-alpha.1`, replacing the saved alpha.3 value;
  - new Build ID: `GTM2-20260807-182259`;
  - Preflight status: FAIL;
  - Blocking errors: 1;
  - Warnings: 0;
  - execution time: 0 seconds;
  - `MODEL LOG` appended `PREFLIGHT | FAIL | Unexpected error 0:` for the runtime-copy filename.
- No corresponding detailed failure row was written to `MODEL VALIDATION`. Instead, the old preflight path cleared some recent rows and left older stale alpha.3 rows beneath the blank area. This confirms that the error handler can fail without creating actionable validation evidence.
- The live behavior confirms that the advertised launcher is wired to an obsolete preflight implementation rather than the current Alpha.3 preflight.
- Fixings, Exposure, and PnL table row counts remained unchanged at 10,278 / 3,399 / 3,578, so the launcher did not run those engines.
- Excel application state was restored after the failure: calculation Manual, events True, screen updating True.
- macOS denied `System Events` accessibility control, demonstrating that modal VBA dialogs cannot be reliably inspected/clicked through the present GUI-automation channel. Direct Excel AppleScript object-model reads remain usable.
- The workbook was closed without saving. Source and runtime-copy hashes again remained identical to the original baseline.
- RT-20 result: FAIL.
- RT-21 containment result: PASS.
- Stop gate applied: no targeted economic or full-build macro was run.

## 2026-08-07 — Recommended headless architecture pivot

The first runtime test demonstrates that GUI-driven VBA should not be the primary development/debug loop. The recommended target is:

1. **Pure calculation core outside Excel.** Implement preflight, Fixings, Exposure, P&L, reconciliation, event timing, and Build Manifest logic as deterministic code with no workbook/UI calls.
2. **Normalized input/output contracts.** Treat Initial Position, Trades, Simulation Trades, Setup/Methods, Calendar, Curves, Fixing Prices, Costs, and operating flows as versioned tables with explicit schemas.
3. **Thin Excel adapter.** Retain only small VBA procedures that export/read table data, call the external engine, import validated outputs, and update user-facing status. No business logic or modal `MsgBox` should live in the adapter.
4. **Headless automated tests.** Use small CSV/Parquet/JSON fixtures and a normal test runner for same-day eligibility, deferred events, zero closures, signs, missing prices, simulation switching, and trade-entry idempotence.
5. **Transactional publishing.** Compute and validate outputs outside the workbook; write to staging tables/files; replace visible Excel outputs only after all checks pass. Failed runs must leave the last accepted build untouched.
6. **Immutable Build Manifest.** Assign a Build ID with input hashes, code version, date range, Simulation state, policy version, and component counts. All downstream outputs must reference the same manifest.
7. **Excel used only for final integration smoke tests.** The GUI is then needed only to prove that the thin adapter, formatting, controls, and external refresh integration work—not to debug financial logic.

### Why this is simpler

- Business logic becomes source-controlled, diffable, and testable without Excel.
- Modal dialogs, stale application state, P-code ambiguity, and workbook corruption no longer obstruct core debugging.
- Large legacy formula sheets can remain frozen as historical evidence and later be archived.
- Helper workbooks become unnecessary once their functions are incorporated into the external engine or thin adapter.
- Independent economic validation can call the same normalized input contract while implementing separate check calculations, rather than comparing one opaque workbook output with another.

### Migration sequence

1. export and preserve all VBA source as evidence;
2. define schemas and business-policy decisions before porting calculations;
3. build read-only workbook extractors that never save `.xlsm` files;
4. implement and test Preflight first;
5. port Fixings, then Exposure, then P&L/trade-entry adjustment;
6. compare against targeted independently verified cases, not legacy PASS/FAIL;
7. build the thin Excel adapter and retire superseded helper/legacy execution paths;
8. keep one small Excel integration suite for the final user workflow.

## 2026-08-07 — Executable engine and acceptance specification drafted

The project scope has now been converted from workbook observations and migration recommendations into two versioned specification documents:

- `docs/GTM_ENGINE_SPECIFICATION.md` defines the target product, data contracts, calculation pipeline, business-rule status, outputs, validation, audit lineage, failure behavior, Excel boundary, and Python-to-VBA migration target.
- `docs/GTM_ACCEPTANCE_TEST_SPECIFICATION.md` defines deterministic fixtures, comparison tolerances, unit/component/pipeline tests, independently approved golden economic cases, property tests, Python-to-VBA parity, Excel-adapter tests, performance gates, evidence retention, and release gates.

Both documents are version `0.1-draft`. They deliberately distinguish confirmed requirements from provisional interpretations and unresolved business decisions. This prevents the new implementation from quietly converting incomplete legacy behavior into a new de facto specification.

### Product boundary now specified

The required target is:

1. a headless Python reference engine for Preflight, Fixings, Exposure, P&L, validation, and manifest generation;
2. immutable, versioned input snapshots and normalized output contracts;
3. calculation and validation stages with fail-closed stop gates;
4. staged output followed by atomic publication only after all blocking checks pass;
5. a thin Excel adapter for input/output and user-facing controls;
6. a later clean VBA implementation, if required, that is tested against the same fixtures and must match Python within approved tolerances;
7. preservation of the Python engine and tests as the long-term executable oracle even after VBA parity is achieved.

External Reuters, SharePoint, Dropbox, and helper-workbook retrieval is separated from the pure calculation engine. The core consumes a complete input bundle with source hashes and as-of timestamps. It does not refresh links, select cells, use `ActiveWorkbook`, show modal dialogs, or depend on Excel calculation/application state.

### Core rules recorded from available evidence

- The initial position is an end-of-day cut-off. Initial fixing eligibility is strictly after Initial Market Date.
- Incremental trades can be eligible on Trade Date, subject to the unresolved late/scheduled-fixing policy.
- Delivery periods are inclusive and split across delivery months.
- Simulation OFF excludes simulation trades; Simulation ON includes the configured scenario population while retaining source/scenario lineage.
- Fixing methods are normalized as WITHINDAY, DAY_AHEAD, HEREN, MONTH_AHEAD, and BRENT_HH.
- Fixing volume carries the opposite sign to open exposure, and fixing amount is fixing volume times fixing price.
- Economic events are recorded in a ledger and applied once to the first output Market Date on or after their effective date.
- Exposure is rolled forward by economic key rather than worksheet row position and records an explicit closing-to-zero transition.
- Exposure MtM is exposure volume times the applicable curve price.
- P&L retains gross exposure movement and trade-entry adjustment as separate auditable values, then reports adjusted movement and explicit operating components.
- No legacy operating-flow fallback is allowed in a production result.
- All output layers share an immutable Build ID and source/output fingerprints.
- Failed builds do not overwrite the last published output.
- Validation history is keyed by Build ID and retained; stale worksheet rows cannot determine current status.

### Acceptance strategy now specified

The acceptance suite contains:

- schema and Preflight tests `VAL-001` through `VAL-020`;
- calendar/effective-date tests `CAL-001` through `CAL-007`;
- Fixings tests `FIX-001` through `FIX-018`;
- Exposure tests `EXP-001` through `EXP-015`;
- P&L tests `PNL-001` through `PNL-014`;
- manifest/publication tests `MAN-001` through `MAN-008`;
- twelve property/metamorphic test families;
- fifteen small golden economic cases requiring independent expected results and business approval;
- Python-to-VBA parity requirements;
- Excel-adapter integration tests `XLS-001` through `XLS-012`;
- contract, Python component, Python end-to-end, Excel adapter, and optional VBA release gates.

Legacy output is explicitly excluded as an acceptance oracle. It may be retained as diagnostic comparison evidence. Business correctness will be grounded in approved, independently calculated golden cases, including cases derived from known workbook anomalies.

### Unresolved decision register

Twelve decisions are explicitly registered in the engine specification:

1. `D-001` — required-price gaps: fail, defer, or another explicit policy. Proposed default: block the build for every price required by an included economic row.
2. `D-002` — exact BRENT/HH current/opening-month curve and valuation behavior. No default is safe without a worked example.
3. `D-003` — DAY_AHEAD weekend/holiday fixing day. Proposed default: previous configured market day.
4. `D-004` — trade entered after its scheduled fixing time or after no scheduled fixing days remain. Proposed default: reject as unsupported until an explicit rule exists.
5. `D-005` — authoritative sources, keys, signs, and allocation rules for operating cost, revenue, and shipping. Proposed release treatment: mark these components NOT_IMPLEMENTED and do not publish authoritative Total P&L until resolved.
6. `D-006` — zero-row retention after a position closes. Proposed default: emit one explicit closure row; omit later zeros until reopening.
7. `D-007` — precision and comparison tolerances. Proposed default: no intermediate rounding; volume `1e-6 MWh`, price `1e-8`, P&L `EUR 0.01` at row and aggregate levels.
8. `D-008` — authoritative canonical mappings. Proposed default: active SETUP mappings only, with unknown values blocking Preflight.
9. `D-009` — trades dated on a non-market day. Proposed default: preserve Trade Date and apply their economic event on the first configured Market Date on or after it.
10. `D-010` — performance target. Proposed default: production-sized Python build under 60 seconds and 2 GiB on the current Mac; focused unit suite under 2 seconds.
11. `D-011` — normalized initial P&L schema and opening-balance behavior. Proposed default: keep it as a separate auditable opening bridge until its components are confirmed.
12. `D-012` — Daily Qty sign convention. Proposed default: Daily Qty is a non-negative magnitude and Side supplies the sign; negative Daily Qty blocks Preflight.

Each decision is traced to the tests that it blocks or leaves provisional. In particular, unresolved operating-flow policy blocks authoritative production Total P&L but does not need to block a deliberately narrower Fixings/Exposure engine release whose manifest declares P&L incomplete.

### Immediate next gate

Implementation should begin only after the user answers the first blocking decision batch or explicitly accepts the proposed defaults. D-002 and D-005 require direct business input; neither can be recovered reliably from the currently observed legacy workbook behavior. Once decisions are recorded, the next concrete artifact should be the versioned Python package skeleton plus normalized schema models and failing tests for the approved first golden cases.

## 2026-08-07 — All specification questions expanded in plain English

At the user's request, every unresolved business question D-001 through D-012 has been unpacked in `docs/GTM_DECISION_GUIDE.md`.

For each decision, the guide now records:

- what the question means in the model;
- why the answer changes calculation, reporting, validation, or performance behavior;
- a concrete gas-trading example;
- the practical policy choices;
- the recommended safe initial choice;
- the exact information required from the user.

The expansion exposed two useful clarifications that were implicit in the shorter decision register:

1. D-007 contains two separate concepts: business rounding and test-comparison tolerance. The specification must distinguish whether amounts are rounded per row before aggregation or only when displayed.
2. D-012 must also settle the handling of zero-quantity trades, not only negative quantities.

The engine specification now points to the plain-English guide. No unresolved decision has been silently resolved, and implementation remains gated on user confirmation or an explicitly narrowed release scope.

## 2026-08-07 — D-001 missing-price policy confirmed

The user accepted the recommended missing-price policy and added a user-interface requirement: the system must show a message asking for the missing prices.

The confirmed rule is:

- any missing price required by an included Fixing, Exposure, or P&L row blocks the build;
- the failed build publishes no new economic output and preserves the last accepted build;
- the engine returns a structured list of every safely discoverable missing required price;
- the Excel adapter displays a clear request for those prices;
- each missing-price item identifies its price type, relevant date, underlying/product, delivery or contract month, and affected source row or economic key when available;
- the engine never silently substitutes zero, a previous price, an interpolated price, or another contract's price;
- a price absent from the source but unused by all included economic rows does not block the build.

The user's question about the earlier follow-up was clarified. That follow-up asked whether any product should be allowed to use an automatic substitute, such as Monday's price when Tuesday's price is missing. No such exception is currently approved. D-001 is therefore recorded as CONFIRMED; a future exception would require a named product rule, a policy-version change, and dedicated tests.

The acceptance requirements VAL-010, FIX-014, EXP-011, and GOLD-010 have been updated from provisional behavior to exact expected failure, diagnostic, and publication behavior.

## 2026-08-07 — Deep local methodology recovery and specification correction

### Evidence boundary and search depth

The user challenged the earlier decision questionnaire because the local project artifacts should already contain the methodology. A deeper audit was therefore performed before requesting any further business input.

The search covered:

- all 47 worksheets in `Gas_Trading_Model 070826.xlsm`;
- workbook-defined names, 7 structured tables, cell comments, selected values, and normalized formula patterns;
- all 68 VBA components extracted from the main workbook;
- all modules in `GTM_Fast_Helper.xlsm` and `GTM_Trade_Entry_Helper_V2.xlsm`;
- the complete text of `docs/GTM_v2_Technical_Handover_Report_2026-08-07.docx`;
- the main workbook's saved normalized outputs as structural/defect evidence, not as an acceptance oracle.

The analysis artifacts are stored under `analysis/deep_methodology_2026-08-07/`. Source workbooks were not modified or executed during this phase.

At the user's explicit instruction, `questions_answers.txt` and `questions_answers.txt~` are excluded from the evidence set. They were not used to resolve, support, or word any methodology rule. They remain untouched on disk.

### Authority rule adopted for the evolved architecture

The user's architectural context is controlling: v2 was added before the legacy model was complete, so neither “newest module wins” nor “legacy output wins” is safe.

The methodology trace now prioritizes:

1. direct user instruction;
2. the editable definitions and active mappings on `SETUP`;
3. rules implemented consistently in the completed legacy builder and the independent fast helper;
4. v2 architectural intent and the handover's explicit defect descriptions;
5. current v2 code only where it is not identified as incomplete;
6. legacy formulas for corroboration and diagnosis, never as an acceptance oracle.

### Core methodology recovered from the workbook

`SETUP!A20:B24` contains full descriptions for all five fixing methods, and comments on the method cells state that these editable values drive FIXINGS dynamically:

- WITHINDAY: every calendar delivery day fixes on itself.
- DAY_AHEAD: every calendar delivery day fixes on the immediately preceding calendar day.
- HEREN: every calendar delivery day fixes on the previous Market Day; one fixing day covers the next Market Day plus all intervening weekend/holiday delivery days.
- MONTH_AHEAD: all calendar delivery days in a month are allocated across Market Days in the previous calendar month.
- BRENT_HH: only market delivery days are eligible and each fixes on the previous Market Day.

This disproves the earlier proposed D-003 default. Ordinary DAY_AHEAD does not use the previous Market Day; that is the HEREN rule.

Active main-underlying method assignments were recovered directly from `SETUP!G3:J11`. Active PVB detail assignments were recovered from `SETUP!M3:P11`. The complete fast helper proves that PVB detail products retain their configured method but aggregate to canonical exposure underlying `Index PVB`.

### Late-trade and same-day rule recovered

Two independent complete implementations agree:

- non-Month-Ahead methods use `effective fixing date = max(normal fixing date, Trade Date)`;
- a same-day normal fixing is eligible;
- overdue delivery quantity closes on Trade Date;
- Month Ahead volume is redistributed equally over remaining prior-month Market Days on or after Trade Date;
- if no Month Ahead fixing date remains, the whole affected quantity closes on Trade Date.

The earlier D-004 recommendation to reject the trade was therefore incorrect and has been removed from the specification.

The trade-entry helper's strict `fixing date > Trade Date` open-volume test is consistent with this rule: the adjustment applies only to volume still open after any same-day fixing. Exposure itself must add the trade and the fixing once on that date.

### Event timing and zero closures

The handover and fast helper give a complete effective-date rule:

- output dates are configured Market Days;
- an event on a non-output date is applied once to the first Market Date on or after its economic date;
- original economic date remains in the audit ledger;
- one explicit zero row is emitted when a material position closes;
- repeated later zeros are omitted until the key reopens.

Current v2 Exposure applies only `eventDate = targetDate` and drops every zero snapshot. Both behaviors are implementation defects, not policy questions.

### Brent and HH curve rule

The v2 Exposure engine and the handover agree that a Brent Dated or HH exposure whose Delivery Month equals the valuation month uses the next Delivery Month's curve. A later delivery month uses its matching curve.

The rule also applies when P&L values Initial Exposure to construct opening MtM. Current v2 P&L omits the same-month next-contract adjustment, so its opening valuation must be aligned with Exposure.

### Operating-flow methodology and sign trace

The workbook contains an exact direct source and allocation map:

- Logistical Costs: `COSTS!B:N`, matched by daily Market Date and BOOK, converted to P&L sign as `-COSTS`;
- Fees and Optimizations: `COSTS!O:P + Foto FO!R:S`, added with stored signs;
- Replication: `Foto FO!T:V`, added with stored signs;
- all three remain at BOOK level with Underlying `TOTAL / BOOK LEVEL` and blank Delivery Month;
- these are daily flows and must not be differenced again.

The independent daily module identifies the same columns and daily-flow treatment, but it adds raw Logistics values without the legacy PNL sign inversion. This appears to be a code defect. `COSTS!R4` also says `CONFIRMAR SIGNO CON OPS`, so one known Ops example is retained as a production golden-test requirement. This is now a narrow validation request rather than a broad methodology question.

### Initial P&L meaning and newly discovered 19-row defect

`INITIAL POSITION!A7:B19` contains 13 signed P&L balances, exactly one for each active SETUP book, and the sheet states that position and P&L represent the closing state at Initial Market Date.

The specified treatment is therefore:

- Initial P&L is a separate BOOK-level opening balance;
- it does not form opening Exposure MtM;
- it is not a post-cut-off daily flow;
- cumulative P&L equals the opening balance plus daily P&L after the cut-off.

The saved `tblInitialPnL` contains 19 rows because `modRefreshInitialPositionData.ReadPnL` reads continuously from the book table to the later TOTAL row. It mistakenly ingests `Description` and five methodology-description rows as zero-P&L books. The correct population is 13, not 19. Preflight must require exactly active SETUP books, and the refresh routine must stop before the documentation block or read a structured source range.

The current v2 P&L engine validates `tblInitialPnL` but never uses it, so cumulative reporting currently omits the opening P&L bridge.

### Daily Qty and precision

The local inputs and engines resolve the sign convention:

- Daily Qty is a non-negative magnitude;
- BUY creates positive exposure and SELL negative exposure;
- fixing volume has the opposite sign from the exposure it closes;
- blank, non-numeric, or negative Daily Qty blocks;
- deliberate numeric zero is accepted with a warning and produces no event.

The live TRADES population contains 476 populated numeric quantities, no negative values, and 12 explicit zeros.

No financial intermediate rounding was found. The v2 engines use `1e-7` as economic materiality. The fast helper's existing parity check uses `1e-6 + 1e-12 × abs(reference)` as allowed numeric difference. Stored values remain full precision; number formats are display-only.

### Performance classification

No runtime or memory SLA is encoded in the local methodology. Performance is therefore removed from the business-decision gate. The first production-sized Python run must establish a measured baseline and later builds must report regressions. An operational SLA can be approved before deployment without changing economic rules.

### Documents corrected

The following documents were revised to version 0.2 methodology baseline:

- `docs/GTM_LOCAL_METHODOLOGY_TRACE.md` — new detailed rule/evidence/conflict report;
- `docs/GTM_DECISION_GUIDE.md` — replaced the questionnaire with resolved D-001 through D-012 rules;
- `docs/GTM_ENGINE_SPECIFICATION.md` — corrected fixing dates, catch-up behavior, curve rules, event timing, zero closures, operating sources/signs, Initial P&L, quantity signs, tolerances, and performance classification;
- `docs/GTM_ACCEPTANCE_TEST_SPECIFICATION.md` — converted provisional cases into exact expected results and added validation for polluted Initial P&L rows and the Ops Logistics sign.

The earlier journal entries are preserved as historical records of what was known at the time. This entry supersedes their “unresolved decision” conclusions wherever they conflict with the deeper evidence.

### Current readiness conclusion

There is no remaining broad methodology questionnaire. The model scope is sufficient to build the headless Python reference engine and deterministic tests. The next direct business input should be approval or correction of small golden-case outputs, especially the single Ops Logistics-sign example, rather than another abstract list of questions.

## 2026-08-07 — Additional business guidance compared with recovered local methodology

### Scope and evidence handling

The user supplied an additional business decision set for D-001 through D-012 and asked for a comparison with the most recent deep local audit. This was treated as new business guidance, not as a claim about what the current VBA already does. The comparison therefore keeps three things separate:

1. rules confirmed by both the business guidance and local workbook/helper evidence;
2. deliberate business-policy changes that override existing VBA/helper behavior;
3. remaining implementation or acceptance details that the supplied guidance does not erase.

No workbook, helper, VBA source, engine specification, or acceptance specification was changed during this reporting step. The only edit was this required journal update. `questions_answers.txt` and its backup remain excluded, unused, and untouched.

### Overall result

The new guidance is strongly consistent with the recovered methodology. It confirms D-001, D-002, D-006, D-008, D-009, D-011, D-012, and the architectural rule that Legacy is diagnostic evidence rather than the acceptance oracle. D-003 and D-005 agree with the recovered evidence but need narrowly scoped product/sign acceptance tests. D-004 deliberately replaces the recovered catch-up behavior. D-007 replaces the current generic parity tolerance with measure-specific tolerances. D-010 turns the previous performance-baseline proposal into explicit operational targets.

### Decision-by-decision comparison

#### D-001 — Missing prices: confirmed

The guidance exactly matches the current recovered specification:

- only prices required by included Fixings, open Exposure, or P&L calculations block;
- unused missing source prices do not block;
- zero, carry-forward, interpolation, or another-contract substitution is forbidden without a future explicit product rule;
- the failed build publishes nothing new and reports the missing keys.

The new wording makes the minimum error key especially clear: Market Date, Underlying, and Delivery Month or the relevant fixing key. It also confirms that the Python reference engine must be stricter than any VBA path that merely counts a missing key and then completes.

#### D-002 — Brent/HH curve: confirmed; regressions now mandatory

The supplied rule matches the v2 Exposure code, the handover, and the recent trace:

- current-month Brent Dated and HH exposure uses the next-month curve contract;
- later Delivery Months use the matching curve month;
- Exposure and opening Initial MtM must use the same lookup rule;
- no discretionary or liquidity roll is permitted.

The implementation requirement is now sharper: add one explicit Brent and one explicit HH golden regression around a roll boundary, each naming Market Date, Delivery Month, expected curve column, and expected price.

#### D-003 — DAY_AHEAD weekends/holidays: qualified agreement, not a contradiction

The new warning not to generalize “Monday uses Friday” agrees with the local evidence. `SETUP` defines method behavior and is authoritative:

- `Day Ahead` means Delivery Day minus one calendar day, so a Monday delivery normally fixes on Sunday;
- `Metodología Heren` means previous configured Market Day, so Monday normally maps to Friday when the intervening days are not Market Days;
- `Brent & HH` also uses a previous Market Day but has market-day delivery eligibility.

The remaining question is narrower than the earlier questionnaire suggested. Method semantics are already explicit. What still needs product-level confirmation/testing is that the actual price feeds support the weekend behavior assigned in `SETUP`. The three currently active PVB detail products mapped to `Day Ahead` are:

- `D+1 Auction`;
- `Mibgas Index ES`;
- `MIBGAS D+1 Daily Reference`.

`TTF DA` is not configured with the generic Day Ahead method; it is configured as `Metodología Heren`. Therefore a product name containing “DA” must never be used to infer its fixing method.

#### D-004 — Late trades: deliberate policy override of existing implementations

This is the principal conflict.

The completed legacy builder `M_dulo10.bas` and the independent fast helper implement catch-up behavior:

- non-Month-Ahead fixing date becomes `max(normal fixing date, Trade Date)`;
- Month Ahead volume is reallocated across remaining Market Days on or after Trade Date;
- when no Month Ahead fixing day remains, the entire quantity closes on Trade Date.

The new business decision supersedes that behavior for release 1:

- same-day fixing is valid, so eligibility is `Fixing Date >= Trade Date`;
- a trade with no valid fixing opportunity remaining is a blocking validation error;
- no catch-up fixing is invented;
- Trade Price is not substituted as Fixing Price.

The existing catch-up code must therefore be retained only as historical evidence, not ported into the Python reference engine.

One edge must be encoded conservatively: for a non-Month-Ahead trade row spanning delivery slices before and after Trade Date, any slice whose fixing opportunity has passed cannot be silently dropped or caught up. The safe release-1 interpretation is to reject the whole source row if its full quantity cannot be scheduled through valid fixing opportunities. For Month Ahead, the locally recovered allocation across remaining eligible dates can still be used when such dates exist, because those are valid future/same-day opportunities; if none exists, reject.

Tests must separately cover:

- normal fixing before Trade Date: reject;
- normal fixing equal to Trade Date: accept;
- at least one Month Ahead fixing day equal to or after Trade Date: schedule on the remaining eligible dates and conserve total volume;
- no Month Ahead fixing day remaining: reject and publish nothing.

#### D-005 — Operating P&L: source and grain confirmed; Logistics sign test still required

The new guidance matches the recovered source layout exactly:

- `COSTS!B:N` supplies Logistics;
- `COSTS!O:P` supplies Fees;
- `Foto FO!R:S` supplies Optimizations;
- `Foto FO!T:V` supplies Replication;
- all are daily Market Date flows and must not be differenced as D2 minus D1;
- BOOK-only source data remains at BOOK level;
- no Underlying, Delivery Month, Trade Source, Scenario, or simulation allocation may be fabricated.

This confirms that these components remain in release-1 Total P&L and reverses the much earlier provisional suggestion to exclude them.

The guidance does not resolve the one concrete local sign conflict. Legacy PNL formulas negate stored Logistics values, while the independent daily v11 code adds raw Logistics, and `COSTS!R4` says `CONFIRMAR SIGNO CON OPS`. The source and grain are closed; one worked Ops golden case is still required to certify whether the engine should use `-COSTS` for Total P&L.

#### D-006 — Explicit closure zero: exact agreement

Emit exactly one zero row when a previously open key closes, omit repeated zeros while it stays closed, and resume rows if it reopens. This matches the handover and recovered helper behavior and confirms that current v2 `CaptureSnapshot`, which removes every zero, is defective.

#### D-007 — Precision: calculation principle confirmed; test tolerances replaced

Both sources agree on full precision through calculation and aggregation, no trade-by-trade rounding, and display-only rounding unless an explicit contract rule says otherwise.

The supplied guidance replaces the current generic helper comparison tolerance with the following acceptance tolerances:

- Volume: `0.000001`;
- Price: `0.00000001`;
- P&L: `EUR 0.01`.

The recovered `1e-7` threshold may remain a separately named economic-zero/output-emission threshold if required, but it must not be confused with reconciliation tolerance. The current specifications still state `1e-6 + 1e-12 × abs(reference)` as a generic comparison tolerance and must be corrected before implementation acceptance.

#### D-008 — SETUP mappings: exact agreement and active export

The new guidance confirms that active `SETUP` configuration is authoritative; unknown values fail Preflight; names remain distinct unless an alias is explicitly configured; Legacy formula labels do not create aliases.

The active mapping recovered from `SETUP` is:

- 13 BOOKs: `CGA_SHT1`, `CGA_AVB`, `CGTINDEX`, `CGA_TVB`, `CGA_FS`, `CGTO`, `PROP.TRADING`, `PIRINEOS`, `CGC_BUNKER`, `PVB FLOW`, `BIOMETHANE`, `COBERTURAS CLIENTES`, `CGA_GS`.
- Main underlyings: `Brent Dated -> Brent & HH`, `HH -> Brent & HH`, `TTF DA -> Metodología Heren`, `TTF MA -> Month Ahead`, `Index PVB -> Metodología Heren`, `Phys PVB -> Metodología Heren`, `TVB -> Metodología Heren`, `AVB -> Metodología Heren`, `PEG -> Metodología Heren`.
- PVB detail mappings, all with canonical Exposure underlying `Index PVB`: `GWDES Auction -> Withinday`; `D+1 Auction -> Day Ahead`; `Mibgas Index ES -> Day Ahead`; `MIBGAS D+1 Daily Reference -> Day Ahead`; `MIBGAS LPI -> Metodología Heren`; `PVB Heren DA -> Metodología Heren`; `Mibgas API DA -> Metodología Heren`; `Mibgas MA -> Month Ahead`; `PVB Heren DA (Delivery) -> Metodología Heren`.

This is the requested one-time active mapping export. Canonical aggregation applies only to the configured PVB detail products; it does not merge main underlyings such as `Phys PVB`, `TVB`, or `AVB` into `Index PVB`.

#### D-009 — Non-Market-Date Trade Date: exact agreement

The original Trade Date is preserved for audit. Its event is applied once on the first configured output Market Date on or after it, never backward. The economic delivery/fixing dates remain their real dates. This matches the planned `eventDate <= targetDate` plus applied-event dictionaries correction.

#### D-010 — Performance: explicit targets now supplied

The local files contained measurement hooks but no approved SLA, so the current v0.2 documents only require a baseline. The new guidance now supplies operational targets:

- full production Python build under 60 seconds;
- normal daily/incremental update ideally under 10 seconds;
- focused/unit tests under 2 seconds;
- peak Python memory under 2 GiB.

These should become measurable engineering acceptance gates on the current Mac and a frozen production-sized input snapshot. The full-build and memory limits are hard targets; “ideally” makes the 10-second daily target an optimization target until explicitly promoted to a hard release gate.

#### D-011 — Initial P&L: definition confirmed; requested 19-row check already resolved structurally

The guidance agrees with the recovered business meaning:

- Initial Market Date is the end-of-day opening state for the reconstruction;
- Initial P&L is a separate cumulative opening bridge by BOOK;
- it is not first-day Daily P&L;
- it is not Exposure MtM;
- later cumulative P&L equals Initial P&L plus daily P&L after Initial Market Date.

The requested check of the 19 saved rows has been performed against the source sheet and refresh logic:

- `INITIAL POSITION!A7:B19` contains exactly 13 real active-BOOK balances;
- those 13 balances sum to `37,445,758.99728647`;
- the other 6 saved rows are `Description` plus five fixing-method description lines, all accidentally read as zero-P&L books by `ReadPnL`;
- therefore the 19-row normalized table is a refresh-boundary defect, not a 19-record business dataset.

This removes the structural uncertainty. Final business acceptance should still reconcile the 13 opening balances once in an end-to-end cumulative example to prove that they are included exactly once and not double-counted.

#### D-012 — BUY/SELL and Daily Qty: exact agreement

Daily Qty is a non-negative magnitude. BUY applies positive sign; SELL applies negative sign. Blank, non-numeric, and negative quantity block. Explicit numeric zero is accepted with a warning and creates no economic event. The live trade population corroborates this convention: no populated negative quantities and 12 explicit zeros were found.

### Architecture and acceptance-oracle conclusion

The new broader architecture statement exactly matches the adopted recovery rule. Legacy remains useful for tracing intent and finding defects, but cannot be the acceptance truth because hard-coded ranges omit later trades and because legacy/v2/helper paths conflict. Acceptance must come from:

- explicit business decisions;
- SETUP-driven configuration;
- deterministic unit and property tests;
- independently calculated golden cases;
- coherent end-to-end builds with conservation and reconciliation checks;
- Python/VBA parity after the Python engine is accepted.

### Specification impact identified, but not applied in this reporting step

Before code implementation, the v0.2 engine and acceptance specifications need a controlled v0.3 amendment for:

1. D-004: remove catch-up-on-Trade-Date behavior and replace fully late trades with blocking validation; retain `>=` for same-day eligibility;
2. D-007: use separate Volume, Price, and P&L reconciliation tolerances;
3. D-010: encode the explicit runtime/memory targets;
4. D-003: describe SETUP-driven product mapping and add weekend fixtures for the three active generic Day Ahead PVB products;
5. D-002: add explicit Brent and HH roll-boundary fixtures;
6. D-011: record that the 19-row audit found 13 economic rows plus 6 documentation rows;
7. D-005: retain the Ops Logistics-sign golden gate.

The present report therefore establishes what changed without silently rewriting the earlier evidence. Source workbooks remain unchanged.

## 2026-08-07 — Project-local Python environment created

At the user's instruction, a project-local virtual environment was created at `.venv` with the installed system Python 3.13.5. System Python and system packages were not modified.

Installed engine/development dependencies:

- Pydantic 2.13.4 for versioned normalized contracts and validation;
- pytest 9.1.1 for component and end-to-end tests;
- Hypothesis 6.165.2 for property tests;
- pytest-cov 7.1.0 and coverage 7.15.4 for test coverage evidence;
- psutil 7.2.2 for peak-memory and performance measurements;
- mypy 2.3.0 for static type checks;
- Ruff 0.16.2 for linting and formatting checks;
- openpyxl 3.1.5 for future read-only Excel adapter/extraction work;
- current pip, setuptools, and wheel build tooling.

All subsequent Python commands for this project will use `.venv/bin/python` or executables under `.venv/bin`. No source workbook was opened, saved, or changed during environment creation.

## 2026-08-07 — v0.3 headless reference engine implemented and verified

### Outcome

The first executable v0.3 Python reference engine is now implemented. It calculates Fixings,
Exposure, daily P&L, cumulative P&L, and an event ledger from a normalized, versioned input
bundle without opening Excel. It is not yet a production replacement for the workbook because
the complete workbook-to-bundle extraction layer and a small set of business golden approvals
remain outstanding. The calculation baseline itself is runnable and tested.

The code deliberately does not mirror the workbook's Byzantine architecture. The workbook and
its VBA are treated as evidence for source locations, mappings, and business-method recovery.
They are not treated as a module design. The Python implementation has four explicit layers:

1. strict immutable input/output contracts;
2. fail-closed Preflight and schedule/price validation;
3. pure calculation modules using keyed state and an event ledger;
4. file loading and atomic publication outside the calculation core.

There are no worksheet formulas, GUI calls, global workbook state, hidden mutable calculation
sheets, or legacy-output dependencies in the core.

### Versioned documentation and configuration

The following v0.3 artifacts now define the approved build baseline:

- `docs/GTM_ENGINE_SPECIFICATION.md`;
- `docs/GTM_ACCEPTANCE_TEST_SPECIFICATION.md`;
- `docs/GTM_DECISION_GUIDE.md`;
- `docs/GTM_LOCAL_METHODOLOGY_TRACE.md`;
- `docs/GTM_ACTIVE_SETUP_MAPPING_v0.3.csv`;
- `README.md` and `pyproject.toml`.

The active mapping export contains 31 records: 13 BOOK records and 18 Underlying records. It
preserves the SETUP-defined source name, canonical name, unit, fixing method, curve name, fixing
price name, fixing-price date basis, current-month roll rule, and source range. No alias is inferred
from Legacy formulas or name similarity.

### Implemented package

The package under `src/gtm_engine` contains:

- `models.py`: schema/policy/engine version `0.3.0`, Pydantic input/output/audit contracts;
- `decimal_utils.py`: full-precision Decimal helpers and distinct materiality/tolerances;
- `canonicalize.py`: explicit SETUP registry and canonical lookup;
- `calendar.py`: configured calendar axis and fixing-date rules;
- `validation.py`: structural, semantic, mapping, late-trade, and duplicate-key Preflight;
- `fixings.py`: fixing and trade schedules, conservation, required prices, and pricing;
- `exposure.py`: event application, one-time closure zeros, curve resolution, and opening MtM;
- `pnl.py`: exposure movement, trade-entry, fixing, operating, daily, and cumulative P&L;
- `invariants.py`: conservation and output coherence checks;
- `manifest.py`: deterministic fingerprint/Build ID, hashes, counts, totals, runtime, and memory;
- `pipeline.py`: one-way fail-closed orchestration;
- `io.py`: normalized CSV/JSON loading and atomic result publication;
- `cli.py`: the headless `gtm-engine build` command.

The core accepts `bundle.json`, BOOK and Underlying configuration, a complete calendar, Initial
Exposure, Initial P&L, trades, curve prices, fixing prices, and daily book-level operating flows.
Successful builds are first verified in memory, then written under a unique run directory, and
only then become `LATEST` through an atomic replacement. Failed builds retain validation evidence
under `failed/<run_id>` and cannot replace the last accepted output.

### Business rules encoded

The engine implements the supplied D-001 through D-012 decisions:

- required missing prices block; unused absent prices do not;
- a required price must match the configured currency and unit;
- same-month Brent/HH exposure and opening MtM select the next-month curve contract;
- fixing methodology is taken from SETUP, including distinct DAY_AHEAD and HEREN behavior;
- same-day normal fixing is valid, genuinely late source rows block, and no catch-up fixing or
  Trade Price substitution is invented;
- Logistics, Fees/Optimizations, and Replication are direct daily book-level flows, never D2-D1
  differences and never arbitrarily allocated;
- exactly one explicit zero exposure row is emitted on closure;
- calculations retain full precision and use separate volume, price, and P&L tolerances;
- unknown BOOK/Underlying names fail instead of being guessed;
- weekend/holiday Trade Dates are applied once on the first subsequent configured Market Date;
- Initial P&L remains an opening cumulative bridge, separate from first-day Daily P&L and MtM;
- Daily Qty is non-negative magnitude, BUY/SELL supplies sign, and explicit zero warns without
  creating an economic event.

### Workbook price-path finding

A further read-only trace of `modGTMv2PnLEngine.bas`, `modGTMv2ExposureEngineV2.bas`, `FIXING
PRICES`, `Brent Dated`, `HH`, `TTF`, `PVB-TTF`, `PEG-TTF`, and `EURF` confirmed the following:

- monthly raw curves are read as dated matrix rows and delivery-month columns;
- current-month TTF uses the TTF day-ahead series; later months use the TTF forward matrix;
- PVB and PEG curves are constructed as TTF plus the corresponding monthly spread;
- the workbook describes Brent and HH fixing prices as `$/bbl` and `$/MMBtu`;
- raw Brent/HH matrices are passed directly through `GetCurvePrice`;
- `EURF` exists and is structurally required, but no use of it was found in the v2 Exposure/P&L
  price-selection path.

This means the workbook does not currently provide an auditable Brent/HH USD-to-EUR rule for a
single EUR reporting total. The Python engine does not conceal this defect. Prices must arrive at
the core with compatible metadata, and a future extractor must either apply an explicitly approved
FX contract or allow Preflight to block the incompatible input. The workbook's cached Reuters
history rows were also largely empty/paused in this saved file, so a reliable headless production
feed cannot assume that opening the `.xlsm` as a ZIP/XML container supplies refreshed market data.

`FIXING PRICES` also contains formula/default zeros even though its instruction says missing prices
should be blank. A workbook extractor must distinguish source-backed zero prices from formulas
that return zero because no observation exists. Treating every cached zero as a valid price would
violate D-001; treating every zero as missing would be fail-closed but could reject a genuine zero
market price. This distinction belongs in the source adapter and its provenance rules.

### Automated evidence

The final verification command ran formatter checking, linting, strict static typing, the full
test suite with branch coverage, and dependency validation.

Results:

- Ruff format check: 22 files already formatted;
- Ruff lint: all checks passed;
- mypy strict mode: no issues in 14 source files;
- pytest: 28 tests passed in 1.83 seconds;
- combined statement/branch coverage: 87.73%, above the configured 85% gate;
- dependency check: no broken requirements;
- synthetic performance case: 476 trades and 13,804 ledger events completed in 0.88 seconds while
  coverage instrumentation was active, below its 5-second test gate and the 2 GiB memory gate.

The test set covers calendar and fixing policy, deterministic golden economics, validation and
input-order determinism, property-based conservation/netting, normalized IO and CLI publication,
failed-publication isolation, and synthetic performance. During the final validation extension, a
real defect was found and fixed: unknown BOOK/Underlying validation for a single Initial Exposure
row was accidentally nested under the duplicate-opening-key warning. Regression tests now prove
that unknown opening keys, a calendar gap, and duplicate daily operating-flow keys all block.

Test coverage is only evidence that code paths were exercised; it is not itself proof that the
business methodology is correct. The stronger evidence is the set of explicit business-rule,
conservation, deterministic-output, publication-safety, and independently calculated golden tests.

### Source integrity and exclusions

No workbook was saved or modified. Final SHA-256 values remain:

- `Gas_Trading_Model 070826.xlsm` — `f6f173ce398109615cc2c8986c52e4feec3249d6b5ba8b15f3c6a75cc5656b31`;
- `GTM_Fast_Helper.xlsm` — `1d80c21339a45ab148c7a53f661e60348a9d46350c7b6590ace0c10fa708667d`;
- `GTM_Trade_Entry_Helper_V2.xlsm` — `3bdeff531141dad4b4d1208ce13848399845ede5bc97d8e123572e1fcaa9526c`.

`questions_answers.txt` and its backup remain excluded: they were not used as methodology evidence,
were not modified, and were not deleted.

### Remaining production acceptance work

The following are deliberately not claimed complete:

1. a read-only workbook/source-feed adapter that produces a complete normalized bundle with
   trustworthy price provenance;
2. an approved Brent/HH currency-conversion contract if EUR Total P&L is required;
3. explicit business-approved DAY_AHEAD weekend fixtures for the three active generic DAY_AHEAD
   PVB products;
4. one workbook-specific Brent roll example and one HH roll example with expected source column
   and price;
5. one approved Logistics-sign example and end-to-end reconciliation of the 13 real Initial P&L
   balances exactly once;
6. a full production-sized benchmark and then a deliberately thin Excel/VBA adapter using these
   same contracts and fixtures.

These gates are intentionally small and concrete. They prevent the incomplete or conflicting
parts of the old workbook from becoming accidental requirements in the clean engine.

## 2026-08-07 — FX audit clarification and Excel-runtime stop

The user correctly challenged the statement that the Brent/HH FX rule was simply “missing” and
requested the secure spreadsheet-analysis route. The distinction is now recorded more precisely:

- the workbook does contain an `EURF` sheet with EUR/USD spot and forward observations;
- the legacy control material explicitly monitors `EUR/USD`, so FX data was part of the intended
  model environment;
- `FIXING PRICES!A2` explicitly describes mixed source-price units: `€/MWh`, `$/bbl`, and
  `$/MMBtu`;
- the inspected `CURVES` Brent and HH formulas select values directly from the raw `Brent Dated`
  and `HH` matrices;
- the inspected legacy `PNL` formula multiplies Brent and HH exposure directly by those `CURVES`
  values alongside euro-denominated gas components;
- the v2 Exposure and P&L VBA loaders also copy Brent and HH matrix values directly and their
  `GetCurvePrice` functions return those values without consulting `EURF`;
- the only v2 VBA references to `EURF` found so far are structural/preflight checks that require
  the sheet and inspect it for formula errors.

Therefore it is too broad to say that the workbook has no FX information. The evidence-supported
statement is narrower: an explicit, operational Brent/HH USD-to-EUR transformation has not been
found in the legacy or v2 valuation paths inspected. The exact approved date/tenor convention is
still not established by those paths.

An attempted follow-up through Microsoft Excel confirmed representative formulas, but Excel
became unstable/crashed from the user's perspective even with macros force-disabled. No macro was
enabled or run, external-link updates were disabled, and the isolated workbook was closed without
saving by the completed read. The user correctly noted that launching Excel is unnecessary for
this audit. Excel runtime inspection is stopped and must not be resumed for methodology discovery
unless the user explicitly requests it again.

The intended secure route is a headless spreadsheet-analysis library or the already extracted
OOXML/formula/VBA evidence. Microsoft Excel GUI/runtime is reserved only for a later, explicitly
authorized integration smoke test. No source workbook was modified.

## 2026-08-07 — Golden Regression Test Pack intake and execution-readiness review

### Artifact inspected

The new independent test specification was found in the singular `verification` directory:

- `verification/GTM_v2_Golden_Regression_Test_Pack_for_CODEX.docx`;
- size: 42,619 bytes;
- SHA-256 before and after review:
  `9cf0e80179694859cab1c9fed7a6edc9a01a170e2bdd500764a7ed5925b8731a`;
- document date/version: working specification, 7 August 2026;
- rendered length: seven pages.

The complete document text and every table were extracted. All seven pages were rendered and
visually inspected. There are no comments, tracked changes, embedded workbooks, or external data
attachments. Two harmless layout wraps occur in the operating-P&L calculation on page 4 and the
required-report field list on page 6; neither changes the semantic content. Microsoft Word was
used only to create a temporary review PDF because LibreOffice and the packaged Python renderer's
`pdf2image` dependency were unavailable. The DOCX was closed without saving. Microsoft Excel was
not launched, and the test-pack hash confirms that the source document was not changed.

### What the pack establishes

The pack is a genuine independent acceptance oracle for the economic core. It contains:

- 12 deterministic synthetic tests;
- the approved volume, price, and P&L tolerances;
- BUY/SELL and Daily Qty conventions;
- same-day and late-fixing boundaries;
- deferred non-Market-Date event behavior;
- explicit-zero closure behavior;
- required-price failure behavior;
- operating-flow signs and source ranges;
- the Initial Market Date trade boundary;
- a three-day cumulative conservation example whose independent result is EUR 500;
- 10 cross-layer invariants;
- a required per-test report schema;
- four diagnostic real-workbook regressions that must not use Legacy as their expected truth.

This agrees with the v0.3 specification's core principles and replaces the earlier status
"named business golden cases remain pending" for the economic cases explicitly present here.
It does not, by itself, supply every previously requested production golden fixture.

### Case-by-case mapping to the current v0.3 engine

| Pack test | Readiness and current-engine mapping |
|---|---|
| TEST 1 — BUY/SELL | Executable. `Side` supplies sign; negative quantity blocks; explicit zero warns. A blank CSV value fails during strict bundle loading, although the reporting harness should add source-row context to that load failure. |
| TEST 2 — trade entry | Executable as a 100 MWh one-delivery-day slice whose Delivery Month key is Sep-26. Current trade-entry formula is structurally aligned with the expected +/- EUR 200. |
| TEST 3 — market movement | Executable as continuation of TEST 2; keyed previous MtM and zero entry adjustment are implemented. |
| TEST 4 — full fixing/closure | Executable and already exposes a current P&L sign defect; see the probe below. |
| TEST 5 — same-day fixing | Executable. The current schedule rule uses `Fixing Date >= Trade Date`, and an existing test already covers this boundary. |
| TEST 6 — late Month Ahead | Executable. Current Preflight emits `LATE_MONTH_AHEAD_TRADE` and produces no economic output. |
| TEST 7 — deferred weekend event | Executable with the exact July dates. Current event ledger retains the Saturday Trade Date and applies the trade once on Monday. |
| TEST 8 — one closure zero | Executable. The fixture must explicitly declare 05/07/2026 a Market Date if the requested reopening row must literally be dated Sunday 05/07; otherwise the general deferred-event rule would place it on the next configured Market Date. |
| TEST 9 — missing price | Executable. Current errors identify Market Date, Underlying, and Delivery Month and do not require unused prices. |
| TEST 10 — operating P&L | Executable. `+1000` source Logistics becomes `-1000`, Fees and Optimizations enter as the combined `+150` normalized field, Replication is `+50`, and the expected total is `-100`. |
| TEST 11 — initial boundary | Executable. Current inclusion rule is exactly `Trade Date > Initial Market Date`. |
| TEST 12 — conservation | The arithmetic oracle is clear, but the exact dated case is not directly generatable end to end by the current production methodologies: a Sep-26 delivery cannot normally receive a full fixing on 03/07/2026 under WITHINDAY, DAY_AHEAD, HEREN, MONTH_AHEAD, or BRENT_HH. It can immediately run as a component-level economic conservation test. For a true pipeline twin, the synthetic delivery/fixing dates should be aligned to a valid configured method rather than adding a test-only production method. |

The phrases `BUY 100 MWh Sep-26` and `BUY 100 Sep-26` must be normalized carefully. The current
contract defines `daily_qty` per eligible delivery day, while the pack expects total exposure of
100 MWh. The smallest faithful fixture is therefore a single delivery-day slice within Sep-26 (or
an explicitly calculated daily quantity whose full-period total is exactly 100 MWh). This is a
fixture-encoding detail, not a reason to reinterpret the engine's Daily Qty contract.

### Readiness probe and defect discovered

A read-only in-memory probe was run for TEST 4 using:

- trade: BUY 100 MWh, entered 01/07/2026 at EUR 47/MWh;
- one WITHINDAY delivery/fixing on 02/07/2026;
- 01/07 curve: EUR 47/MWh;
- 02/07 fixing price: EUR 48/MWh.

The v0.3 result was:

```text
Fixing volume                -100 MWh
Raw Fixing Amount            -EUR 4,800
Gross Delta Exposure MtM     -EUR 4,700
Current P&L fixing component -EUR 4,800
Current Total P&L            -EUR 9,500
Golden expected Total P&L    +EUR   100
```

The cause is precise. `price_fixings` correctly preserves the closing event's signed raw amount as
`fixing_volume * fixing_price = -4,800`, but `build_pnl` currently adds that raw settlement amount
as if it were the economic P&L contribution. For P&L, the contribution must have the opposite
sign: `-raw_fixing_amount = +4,800` for closing a long exposure, and symmetrically `-4,800` for
closing a short exposure. TEST 4 and TEST 12 will fail until this is corrected. The pipeline still
labels this mathematically coherent but economically wrong result `VERIFIED`, which also proves
that the post-calculation invariants need the new independent conservation assertion.

No production code was changed during this intake/readiness step. The probe was deliberately kept
outside the permanent test suite so that the approved pack can be encoded as traceable named
fixtures in the next implementation step.

### Remaining acceptance coverage not supplied by this pack

The pack is sufficient to begin economic-core testing, but it does not close these previously
identified production gates:

1. a business-sourced Brent roll example and HH roll example with Market Date, Delivery Month,
   exact source curve column, and expected price;
2. the Brent/HH USD-to-EUR conversion date/tenor contract if Total P&L must be reported in EUR;
3. product-specific weekend/holiday prices for the three active generic DAY_AHEAD PVB products;
4. the 13-book Initial P&L opening bridge reconciled exactly once into cumulative P&L;
5. the read-only workbook-to-normalized-bundle adapter needed for the four real-workbook
   diagnostic regressions.

### Readiness decision

The system is ready to be tested headlessly in Python now. The correct next step is to encode the
12 named cases as versioned fixtures, add a deterministic expected-versus-actual report with the
fields required by Section 16, run the synthetic cases in the pack's prescribed order, and fix
failures against the independent economics. Excel is not required for those tests.

This is readiness to begin a falsifiable test-and-fix cycle, not a claim that v0.3 already passes.
The TEST 4 probe proves that it currently does not. Full production acceptance must remain open
until the real-workbook adapter and the additional business fixtures listed above exist.

## 2026-08-07 — Golden Regression Test Pack implemented, defects fixed, and 12/12 passed

### Executable pack

The 12 approved synthetic cases from
`verification/GTM_v2_Golden_Regression_Test_Pack_for_CODEX.docx` are now permanent executable
tests in `tests/test_regression_pack.py`. Each test uses independently stated expected amounts;
the expected results are not generated by the implementation under test.

The first formal run produced the predicted result:

```text
10 passed
2 failed: TEST 4 and TEST 12
```

Both failures isolated the same economic defect. `FixingRow.fixing_amount` correctly stored the
raw signed settlement as `fixing_volume * fixing_price`, but `PnlRow.fixing_amount` copied that raw
sign. For a long exposure closed by fixing volume `-100` at EUR 48, the raw amount is
`-EUR 4,800`, while the economic P&L contribution is `+EUR 4,800`. The old path produced
`-4,700 - 4,800 = -9,500`; the approved result is `-4,700 + 4,800 = +100`.

### Corrections

The following focused changes were made:

1. `src/gtm_engine/pnl.py` now preserves the raw signed amount in Fixings output and applies
   `economic_fixing_amount = -raw_fixing_amount` when building P&L. The inverse works symmetrically
   for long and short exposure.
2. `src/gtm_engine/invariants.py` now checks `PNL_FIXING_SIGN` for every Market Date and economic
   key. A future regression cannot remain `VERIFIED` merely because Total P&L is arithmetically
   consistent with a wrongly signed component.
3. `src/gtm_engine/io.py` now reports the filename, CSV row number, and `source_row_id` when strict
   row parsing fails. This makes a blank Daily Qty a traceable blocking load error while preserving
   the distinct explicit-zero warning behavior.
4. `docs/GTM_ENGINE_SPECIFICATION.md`, `docs/GTM_ACCEPTANCE_TEST_SPECIFICATION.md`, and
   `docs/GTM_LOCAL_METHODOLOGY_TRACE.md` now distinguish raw fixing settlement from the economic
   P&L component.
5. `tests/test_io_cli.py` includes a regression proving that blank quantity errors retain CSV row
   and source-row identity.

No schema-breaking field rename or test-only production fixing method was introduced.

### Pack execution details

All 12 cases now pass:

- TEST 1: BUY/SELL sign, negative and blank rejection, and explicit-zero warning;
- TEST 2: BUY `+EUR 200` and SELL `-EUR 200` trade-entry economics;
- TEST 3: `+EUR 200` pure curve movement;
- TEST 4: raw `-EUR 4,800` closing settlement converted to economic `+EUR 4,800`, producing
  `+EUR 100` closure P&L;
- TEST 5: `Fixing Date = Trade Date` is eligible and applied once in both event layers;
- TEST 6: late Month Ahead trade blocks with no catch-up output;
- TEST 7: Saturday trade applies once on Monday and never on Friday or twice;
- TEST 8: one explicit zero closure row, no repeated zeros, and row emission resumes on reopening;
- TEST 9: required missing price blocks with exact key details, while an unused absent price does
  not block;
- TEST 10: `+500 + 200 - 1,000 + 150 + 50 = -EUR 100` using actual engine rows at the correct
  position/BOOK grains;
- TEST 11: only `Trade Date > Initial Market Date` enters incremental activity;
- TEST 12: daily P&L `+200`, `+200`, `+100` sums to `+EUR 500`, exactly matching
  `100 × (50 - 45)`.

TEST 12's source prose labels the position Sep-26 while giving it a full 03/07 fixing. No active
GTM method can produce that relationship. The executable pipeline twin uses a one-day WITHINDAY
delivery on 03/07, preserving the stated Market Dates, volume, prices, P&L components, and EUR 500
conservation result. This keeps the test economically exact without contaminating the production
method set. The report records this normalization explicitly.

### Permanent test report

`verification/GTM_v2_Golden_Regression_Test_Report_v0.3.md` contains every field required by
Section 16 of the pack for all 12 cases:

- Test ID and PASS/FAIL;
- Input;
- Expected and Actual Fixings;
- Expected and Actual Exposure;
- Expected and Actual P&L components;
- Difference;
- failure explanation field.

Every asserted numeric difference is exactly zero before applying tolerance.

### Final verification

The final clean command ran formatting, linting, strict type checking, the complete suite with
branch coverage, and dependency validation:

```text
Ruff format: 23 files already formatted
Ruff lint: all checks passed
mypy strict: no issues in 14 source files
pytest: 41 passed in 2.02 seconds
coverage: 88.16% (gate: 85%)
pip check: no broken requirements
```

The full suite includes the synthetic performance gate, which passed. The pack-only rerun was
`12 passed in 0.17 seconds`.

Microsoft Excel was not launched. No `.xlsm` source or helper workbook was opened, saved, or
modified. The golden-pack DOCX remains unchanged at SHA-256
`9cf0e80179694859cab1c9fed7a6edc9a01a170e2bdd500764a7ed5925b8731a`.

### Acceptance boundary after this run

The 12-case synthetic economic core is now accepted by its supplied golden pack. This result does
not close the separate Brent/HH FX/source-column fixtures, the product-specific DAY_AHEAD weekend
fixtures, the 13-book Initial P&L reconciliation, the workbook-to-bundle adapter, or the four
real-workbook diagnostics. Those remain explicit later gates rather than hidden qualifications on
the 12/12 result.

## 2026-08-08 — Complete Excel interface build initiated; authoring-runtime decision required

The user authorized construction of the complete Excel input/output system around the accepted
Python reference engine. The selected design is a clean normalized Excel interface, headless
read-only input extraction, fail-closed Python calculation, atomic publication, and verified
result import. The legacy workbook will remain a source/methodology reference rather than become
the new interface architecture. Excel GUI calculation and legacy macros are outside the design.

The required spreadsheet-authoring runtime for this environment, `@oai/artifact-tool` loaded
through `load_workspace_dependencies`, is not available in the current session. The spreadsheet
workflow explicitly prohibits silently substituting another `.xlsx` authoring library unless the
user authorizes that fallback. The project already has `openpyxl` installed in `.venv`, and it can
implement the complete headless adapter without opening Excel, but explicit authorization is
required before using it to create or modify the final interface workbook.

No adapter code or workbook artifact was created before resolving this tooling choice. No source
workbook was opened or changed. The active implementation plan records the workbook contract,
input adapter, output publication, end-to-end tests, visual verification, and operating
documentation as the remaining steps.

## 2026-08-08 — Complete macro-free Excel input/output system delivered

### Tooling decision and scope

The user explicitly authorized `openpyxl` after the required spreadsheet-authoring runtime proved
unavailable. `openpyxl>=3.1,<4` is now a core project dependency, and `types-openpyxl>=3.1,<4` is a
development dependency. Both are installed in `.venv` through the editable project install.

The implemented system does not translate the economic engine into worksheet formulas or VBA.
It establishes the intended one-way boundary:

```text
macro-free Excel input workbook
            ↓
strict Excel contract loader
            ↓
authoritative Python v0.3 engine
            ↓
normalized CSV/JSON evidence + separate Excel result workbook
```

The legacy `.xlsm` files remain read-only methodology and migration sources. They are not opened
or calculated during the new daily workflow.

### Workbook contract

`outputs/gtm_excel_v0_3/GTM_Excel_Interface_v0.3.xlsx` is the production input template. It has 18
visible worksheets and 17 named Excel tables:

- `START HERE`: status, daily workflow, and colour legend;
- `CONTROL`: versioned dates, simulation switch, timezone, Logistics sign, and materiality;
- nine input tables: `BOOKS`, `UNDERLYINGS`, `MARKET CALENDAR`, `INITIAL EXPOSURE`, `INITIAL PNL`,
  `TRADES`, `CURVE PRICES`, `FIXING PRICES`, and `OPERATING FLOWS`;
- six calculated output tables: `VALIDATION`, `FIXINGS`, `EXPOSURE`, `DAILY PNL`, `CUMULATIVE PNL`,
  and `EVENT LEDGER`;
- `BUILD MANIFEST`: versions, hashes, row counts, output totals, validation counts, timing, memory,
  platform, and failure stage.

The template seeds the reviewed active SETUP mapping from
`docs/GTM_ACTIVE_SETUP_MAPPING_v0.3.csv`: 13 BOOK records and 18 Underlying records. It does not
invent an Initial Market Date, holiday calendar, opening state, trades, or prices. Those required
production inputs remain blank and visibly marked for the operator.

Inputs use blue text, required control dates use yellow fill, and Python-owned outputs use green
tabs and headers. Every working table has a fixed name, exact header contract, filters, frozen
headers, typed number/date formats, and bounded column widths. Categorical inputs use Excel data
validation. Negative Daily Qty receives a visible conditional-format warning before Python applies
the blocking validation rule.

The workbook contains no VBA, external workbook links, or economic formulas. The adapter rejects
formulas in every authoritative input cell and identifies the table and cell. This avoids trusting
stale cached formula values in a headless run.

### Runtime implementation

`src/gtm_engine/excel.py` implements:

- strict `.xlsx` contract validation;
- exact Excel-table-to-Pydantic conversion;
- source-workbook SHA-256 capture and an unchanged-source check after the build;
- reviewed SETUP mapping import;
- clean workbook generation;
- output-table replacement with typed values and explicit financial formats;
- success and failure status updates on `START HERE`;
- complete Build Manifest export;
- atomic workbook saving and atomic `GTM_LATEST.xlsx` replacement;
- retained diagnostic workbooks for economic and workbook-load failures;
- a full-workbook formula scan used by verification.

`gtm-engine` now exposes two Excel commands:

```text
gtm-engine excel-template --output <interface.xlsx> [--mapping <setup.csv>]
gtm-engine excel-build --workbook <interface.xlsx> --output <result-root>
```

`scripts/GTM_Build.command` supplies the normal macOS workflow. With no arguments it builds the
delivered interface and writes under `outputs/gtm_excel_runs`. A user can drag another GTM `.xlsx`
workbook onto the launcher, or pass workbook and output paths as arguments. The launcher does not
open Microsoft Excel.

On success, the engine first publishes normalized evidence under `runs/<run_id>/`, writes
`GTM_Result.xlsx` there, verifies that the source hash is unchanged, and atomically replaces
`GTM_LATEST.xlsx`. On an economic failure it writes `GTM_Failed.xlsx` under `failed/<run_id>/`. On
a contract/load failure it creates a separate `failed/load-.../GTM_Failed.xlsx` when the workbook
still contains the required diagnostic table. Neither failure path changes the previous latest
workbook.

### Tests added

`tests/test_excel_adapter.py` adds five focused end-to-end tests covering:

1. exact InputBundle round-trip, sheet/table contract, reviewed mapping counts, no formulas, and
   unchanged source hash;
2. successful build publication, Excel/CSV row-count parity, populated status and manifest, and
   byte-identical latest workbook promotion;
3. missing-price failure, retained validation, empty economic outputs, and preservation of the
   prior latest workbook;
4. input-formula rejection with table/cell identity plus a retained load-failure workbook;
5. template CLI creation and missing-sheet, wrong-suffix, and missing-mapping failures.

The macOS launcher also completed a separate synthetic end-to-end run:

```text
status: PUBLISHED
build_id: GTM3-F52CCC089F701794EBFF
validation counts: ERROR 0, WARNING 0, INFO 0
latest workbook created with mode 0644
```

### Final verification

The complete project gate finished cleanly:

```text
Ruff format: 26 files already formatted
Ruff lint: all checks passed
mypy strict: no issues in 15 source files
pytest: 46 passed in 4.34 seconds
coverage: 89.55% (gate: 85%)
pip check: no broken requirements
Golden Test Pack: 12/12 remain passing within the full suite
```

The final workbook received a separate headless integrity and presentation-contract pass:

```text
worksheets: 18, all visible and ordered
named tables: 17
BOOK rows: 13
Underlying rows: 18
formula cells: 0
formula-error values: 0
VBA parts: 0
external-link parts: 0
ZIP integrity: PASS
file permissions: 0644
```

No Microsoft Excel or Quick Look GUI was launched for this delivery. The workbook was checked
through its OOXML package and `openpyxl` object model after the user asked to keep the remaining
work headless.

### Final hashes

```text
GTM_Excel_Interface_v0.3.xlsx
568e1bf9220182091eec4c0e7df40e1298bc23c2b94f68f9647e39c4a82aa505

Gas_Trading_Model 070826.xlsm
f6f173ce398109615cc2c8986c52e4feec3249d6b5ba8b15f3c6a75cc5656b31

GTM_Fast_Helper.xlsm
1d80c21339a45ab148c7a53f661e60348a9d46350c7b6590ace0c10fa708667d

GTM_Trade_Entry_Helper_V2.xlsm
3bdeff531141dad4b4d1208ce13848399845ede5bc97d8e123572e1fcaa9526c

GTM_v2_Golden_Regression_Test_Pack_for_CODEX.docx
9cf0e80179694859cab1c9fed7a6edc9a01a170e2bdd500764a7ed5925b8731a
```

The three legacy workbooks and the Golden Test Pack retain their recorded hashes. No source file
was modified.

### Operating documentation and acceptance boundary

`docs/GTM_EXCEL_INTERFACE_GUIDE.md` now gives the daily workflow, sheet meanings, success/failure
behavior, safety rules, and fresh-template command. `README.md` describes the Excel commands and
launcher and now installs the complete environment with `.[dev]`.

The Excel interface, adapter, publication path, and automated verification are complete. A first
production run still requires an operator-approved data snapshot in the normalized input tables.
That snapshot must include the real configured Market Date calendar, opening positions, the
reviewed Initial P&L bridge, post-cut-off trades, and all economically required prices. Automatic
migration from the legacy `.xlsm` layout is a separate migration task; it is not hidden inside the
new workbook. Product-specific DAY_AHEAD weekend acceptance, business-sourced Brent/HH roll
fixtures, and the one-time 13-book Initial P&L reconciliation remain the previously recorded
business acceptance gates rather than defects in the Excel transport layer.

## 2026-08-08 — Quick Start and detailed operator manual completed

The original Excel interface guide was too brief to answer three basic operator questions in a
single, reliable sequence: how to pass information to the engine, how to run it, and how to
analyse the result. It also did not explain table boundaries, source IDs, sign conventions,
required-price coverage, or the difference between a synthetic test pass and a production-data
reconciliation.

Two separate manuals now serve different needs.

### One-page daily guide

`docs/GTM_QUICK_START.md` is a 588-word operating guide. It covers:

1. the seven practical data-entry steps in the Excel interface;
2. the rules that prevent rows from being ignored or rejected;
3. launcher and Terminal execution;
4. successful and failed output locations;
5. the required order for reading status, validation, P&L, cumulative P&L, exposure, fixings,
   event ledger, and manifest;
6. concise corrections for the six most common failures.

The guide explicitly warns that data outside a named Excel table are ignored, formulas are
rejected, Delivery Month must be the first day of the month, Daily Qty is unsigned, and currencies
and units must match the Underlying configuration. It also explains that a working copy must be
dragged onto `GTM_Build.command`; otherwise the launcher builds its default delivered workbook.

### Detailed user manual

`docs/GTM_EXCEL_INTERFACE_GUIDE.md` was rewritten as the 2,872-word
`GTM v0.3 Detailed User Manual`. It covers:

- the Excel → Python → result architecture;
- general data-entry rules and safe table-row expansion;
- every field in `CONTROL`, `BOOKS`, `UNDERLYINGS`, `MARKET CALENDAR`, `INITIAL EXPOSURE`,
  `INITIAL PNL`, `TRADES`, `CURVE PRICES`, `FIXING PRICES`, and `OPERATING FLOWS`;
- opening-state and incremental-trade boundaries;
- BUY/SELL, opening-exposure, raw-fixing, economic-fixing, and Logistics sign conventions;
- exact missing-price, date, source-ID, currency, and unit expectations;
- manual mapping from the authoritative legacy source sheets and ranges into the new normalized
  input tables;
- the warning not to treat legacy Fixings, Exposure, or P&L outputs as new-engine inputs;
- pre-run checks, launcher use, the Terminal command, and the complete output directory layout;
- the order for result review and a source-row-to-P&L trace procedure;
- explicit Daily P&L and cumulative P&L reconciliation equations;
- approved volume, price, and P&L comparison tolerances;
- a failure-diagnosis table with concrete corrections;
- the Excel-free CSV/JSON route and the fact that no persisted production bundle exists yet;
- the outstanding production-data, legacy-import, business-fixture, and reconciliation gates.

The Brent/HH section now states the actual safety boundary: prices must match the configured
currency and unit, and legacy Brent/HH prices must not be imported under a different convention
until FX treatment has been approved. This avoids implying that the current engine performs a
currency conversion that it does not implement.

`README.md` now links separately to the Quick Start and detailed manual. Relative document links,
headings, paths, command names, sheet names, table field names, output locations, and formula/sign
descriptions were checked against the implemented v0.3 workbook contract and CLI. No engine code,
workbook artifact, or source `.xlsm` file was changed during this documentation task.

## 2026-08-08 — Repository installation guide added and verified

The forthcoming repository needed installation instructions that begin from a clean computer and
do not assume the project-local `.venv` already exists. The earlier README contained only one
developer reinstall command and was not sufficient for an operator, a new contributor, or a CI
maintainer.

`INSTALL.md` now provides a complete 983-word repository installation guide with:

- Python 3.13 as the tested runtime target and an explicit compatibility-test requirement for
  later Python versions;
- Git, network/package-mirror, operating-system, and optional Excel requirements;
- clone and repository-root instructions;
- isolated `.venv` creation for macOS/Linux and Windows PowerShell;
- separate operator (`pip install -e .`) and developer (`pip install -e '.[dev]'`) installations;
- `pyproject.toml` as the sole dependency authority;
- operator verification through CLI discovery, dependency validation, and macro-free template
  generation;
- the complete developer formatting, lint, strict typing, test, coverage, and dependency gate;
- reproducible generation of the Excel interface from the reviewed SETUP mapping;
- macOS launcher permissions and the Terminal alternative;
- first-use and upgrade procedures that preserve working input workbooks;
- troubleshooting for Python versions, missing commands/packages, launcher permissions, paths
  containing spaces, economic validation failures, and package-download problems;
- a concise support-information checklist that excludes production data from public issues.

`README.md` now has a Documentation section linking the installation guide, Quick Start, and
detailed manual. Both user manuals link back to `INSTALL.md`, giving a complete path from clean
checkout to installation, data entry, execution, and analysis.

The documented operator and developer verification commands were executed rather than reviewed
only as prose:

```text
gtm-engine --help: build, excel-template, and excel-build present
pip check: no broken requirements
install smoke template: CREATED at /private/tmp/GTM_install_check.xlsx
Ruff format: 26 files already formatted
Ruff lint: all checks passed
mypy strict: no issues in 15 source files
pytest: 46 passed in 1.97 seconds
documentation relative links: PASS
```

The smoke template was written outside the repository. No engine code, delivered workbook,
legacy `.xlsm` source, or test-pack artifact was modified during this installation-documentation
task.

## 2026-08-08 — Read-only legacy workbook importer delivered

### Request and safety boundary

The user requested an importer that can take the relevant source information from
`Gas_Trading_Model 070826.xlsm`. The implementation reads the workbook headlessly through
`openpyxl` with `data_only=True`, `read_only=True`, and external-link loading disabled. It never
opens Microsoft Excel, enables macros, recalculates workbook formulas, or saves the source file.

The importer computes the source SHA-256 before and after extraction and refuses publication if
the hash changes. The source remained unchanged at:

```text
f6f173ce398109615cc2c8986c52e4feec3249d6b5ba8b15f3c6a75cc5656b31
```

The first implementation probe exposed slow random-cell access in openpyxl's read-only mode. That
run was interrupted before any output was published. The importer now streams each source range
with `iter_rows`; the complete real-workbook extraction takes about 0.5 seconds.

### Extraction contract

The importer uses source inputs and configuration, never legacy calculated result sheets:

- controls: `INITIAL POSITION!C2`, `PROCESS!C15:D15`, and `SIMULATION TRADES!Q1`;
- BOOK and Underlying configuration: the active sections of `SETUP`, including canonical PVB
  aggregation, fixing methods, delivery-date price bases, and the Brent/HH next-contract rule;
- Market Calendar: cached Date and Market Day values from `CALENDAR!A4:E1100`;
- opening exposure: material rows from `INITIAL POSITION DATA!A2:E8209`;
- opening P&L: exactly the 13 active BOOK rows in `INITIAL POSITION!A5:C17`;
- actual and simulation trades: source columns A:I, plus the simulation scenario field; calculated
  columns J onward are excluded;
- curves: cached TTF, Brent Dated, and HH histories are unpivoted; PVB-family and PEG prices are
  reconstructed from the cached TTF curve plus `PVB-TTF` or `PEG-TTF` spreads;
- fixing prices: the wide `FIXING PRICES` matrix is unpivoted by lookup date and source Underlying;
- operating flows: `COSTS!B:N` supplies Logistics, `COSTS!O:P` and `Foto FO!R:S` supply Fees and
  Optimizations, and `Foto FO!T:V` supplies Replication.

Exact zero opening positions are omitted because they do not form part of the material opening
state. Zero curve and fixing-price placeholders are treated as missing, not as valid market
prices. Explicit zero trade quantities and execution prices remain present and auditable. Legacy
`FIXINGS DATA`, `EXPOSURE DATA`, `PNL DATA`, and related calculated sheets are explicitly excluded
from authoritative input.

USD source values are retained as USD for Brent and HH; the importer does not relabel them as EUR.
The audit blocks acceptance of combined EUR P&L until an approved FX conversion rule is added.

### Code and command

`src/gtm_engine/legacy_import.py` implements the source parser, audit model, atomic artifact
publication, and embedded `LEGACY IMPORT` workbook sheet. `src/gtm_engine/io.py` now exposes the
public atomic `write_bundle` function. `src/gtm_engine/cli.py` adds:

```sh
.venv/bin/gtm-engine legacy-import \
  --workbook "Gas_Trading_Model 070826.xlsm" \
  --output outputs/legacy_import_070826_v0.3
```

Optional `--historical-start` and `--historical-end` ISO-date arguments override `PROCESS!C15:D15`.
The importer never overwrites an existing destination directory.

Every import publishes four artifacts:

```text
GTM_Imported_Input.xlsx
normalized_bundle/
legacy_import_audit.json
legacy_import_issues.csv
```

The Excel input and normalized bundle contain equivalent economic records. The workbook remains
macro-free and formula-free; it contains no VBA parts or external links.

### Real-workbook extraction result

The delivered snapshot contains:

```text
BOOKS                 13
Underlyings           18
Calendar rows       1,097
Material openings     101
Initial P&L rows        13
Trades                 476
Curve prices         5,288
Fixing prices          572
Operating flows         93
```

The importer skipped 8,107 exact-zero opening grid rows and 18,077 zero fixing-price placeholders.
The audit status is `CREATED_WITH_REVIEW_ITEMS`, with two errors, two warnings, and three
information items:

- ERROR: the material Brent opening requires an approved USD-to-EUR treatment;
- ERROR: material positions reference empty fixing-price series for Index PVB, Phys PVB, TTF MA,
  and TVB;
- WARNING: 12 explicit-zero Daily Qty rows were retained;
- WARNING: 91 explicit-zero execution-price rows were retained for business review;
- INFO: 476 legacy comments remain traceable by source row but are not economic inputs;
- INFO: legacy calculated outputs were excluded;
- INFO: an engine Preflight is still required after import.

The normalized-bundle Preflight completed in 1.1 seconds. It reached the engine contract cleanly
and reported 54 exact missing fixing-price keys plus the 12 zero-quantity warnings. The missing
keys were TVB 1, Phys PVB 20, Brent Dated 13, and Index PVB 20, all in the July 2026 delivery
month. No schema, calendar, BOOK, Underlying, duplicate-key, or importer-load error occurred. The
FX rule remains an import-audit error because v0.3 does not yet have a safe combined-currency P&L
validation layer.

### Verification

`tests/test_legacy_import.py` adds a synthetic end-to-end legacy workbook, CLI success and
fail-closed cases, normalized/Excel round-trip checks, and an optional pinned inventory test for
the private real workbook. The importer module has 90% branch coverage.

The complete project gate passed:

```text
Ruff format: 28 files already formatted
Ruff lint: all checks passed
mypy strict: no issues in 16 source files
pytest: 49 passed in 5.95 seconds
coverage: 89.69% (gate: 85%)
pip check: no broken requirements
```

The real Excel/CSV reconciliation passed at the approved tolerances. Maximum storage differences
introduced by Excel were `5E-11` for volume/flow, `6.7E-15` for price, and `4E-12` for P&L amount.
The workbook also passed ZIP integrity, full formula scan, VBA-part scan, and external-link scan.

Documentation was updated in `README.md`, `INSTALL.md`, `docs/GTM_QUICK_START.md`, and
`docs/GTM_EXCEL_INTERFACE_GUIDE.md`. All local documentation links resolve.

### Delivered artifact hashes

```text
GTM_Imported_Input.xlsx
7fcd28988fd9be1197595364fe3c6c710b65c635007c8259726cd7c43f2ba5b3

legacy_import_audit.json
d7678d9f466eba4100b08530646a9bbe5b8418516122f9ea9c13c5009400bd78

legacy_import_issues.csv
7f9dec18e8077fdd110ad2808b11ab0633ac46a32bd56aa40dc30f01d44e49f3

normalized_bundle/bundle.json
d30b3b2183561339249e229d2609f4fffd215f268d435ee119753412ed779d98
```

## 2026-08-08 — Legacy import rebuilt to the 10 July 2026 cutoff

The user clarified that the working snapshot should consider trades only through 10 July 2026.
The importer now treats Historical End Date as the complete as-of cutoff: it excludes actual and
simulation trades whose Trade Date is later than the cutoff and records their source rows in the
audit. This keeps later trades out of validation as well as calculation.

The real workbook was re-imported with:

```sh
.venv/bin/gtm-engine legacy-import \
  --workbook "Gas_Trading_Model 070826.xlsm" \
  --output <verified-staging-directory> \
  --historical-end 2026-07-10
```

The verified cutoff import replaced
`outputs/legacy_import_070826_v0.3/GTM_Imported_Input.xlsx` and its companion normalized/audit
files. The previous 28 July import remains recoverable at
`outputs/legacy_import_070826_v0.3_before_2026-07-10_cutoff/`.

The rebuilt input contains 423 trades with Trade Dates no later than 10 July; 53 later source
trades were excluded. It also contains 2,313 curve observations and 38 operating-flow rows for the
shorter reporting horizon. The 13 BOOK rows, 18 Underlying mappings, 1,097 calendar rows, 101
material opening positions, 13 Initial P&L rows, and 572 non-placeholder fixing observations are
unchanged.

The user's expectation that the cutoff removes all missing fixing prices was tested rather than
assumed. Engine Preflight still reports 18 required missing keys:

```text
Index PVB      8  (1, 2, 3, 6, 7, 8, 9, and 10 July)
Phys PVB       8  (1, 2, 3, 6, 7, 8, 9, and 10 July)
TVB            1  (2 July)
Brent Dated    1  (10 July fixing; Delivery-Day lookup is 13 July)
```

The PVB-family source columns contain zero placeholders on these dates; the importer does not
turn them into prices. Brent uses the approved Delivery-Day lookup basis, so the 10 July fixing
requires the 13 July price even though the reporting cutoff is 10 July. The separate Brent
USD-to-EUR rule also remains undefined. Preflight additionally retains 12 explicit-zero Daily Qty
warnings.

The cutoff workbook passed the input-contract load, formula scan, VBA scan, external-link scan,
ZIP integrity check, and source-hash check. A retained failed Preflight workbook with all exact
keys is under `outputs/legacy_import_070826_v0.3/preflight_results/failed/`. The original legacy
workbook remains unchanged at SHA-256
`f6f173ce398109615cc2c8986c52e4feec3249d6b5ba8b15f3c6a75cc5656b31`.

## 2026-08-08 — Excel table-repair defect found and corrected

### User-visible failure

Microsoft Excel reported that `GTM_Imported_Input.xlsx` contained invalid content and offered to
repair it. Its recovery log removed the AutoFilter and Table features from `table3.xml` through at
least `table17.xml`. This was a genuine compatibility defect in the generated workbook. Earlier
ZIP, formula, VBA, external-link, and `openpyxl` round-trip checks did not exercise Excel's rule
against overlapping filter ownership and therefore did not detect it.

### Exact cause

Each normal input/output sheet was created with an Excel structured Table. A structured Table
already owns its AutoFilter through the corresponding `xl/tables/tableN.xml` part. Immediately
after adding that table, `_add_table_sheet` also assigned the same range to
`sheet.auto_filter.ref`. `openpyxl` consequently wrote two filters over one range:

1. the correct table-owned `<autoFilter>` inside `xl/tables/tableN.xml`; and
2. a second worksheet-level `<autoFilter>` inside `xl/worksheets/sheetN.xml`, accompanied by an
   `_xlnm._FilterDatabase` defined name.

Excel rejected the duplicated definitions and removed the affected tables. The recovery pattern
confirms the diagnosis: the accepted `LEGACY IMPORT` and `CONTROL` tables had no worksheet-level
filter, while all rejected tables created through `_add_table_sheet` had both filters. The BUILD
MANIFEST table also had no duplicate filter in the input workbook.

The same erroneous assignment existed in the result-table and manifest replacement paths. A
successful or failed Python build could therefore have reintroduced the defect even if the input
template had been corrected.

### Repair

`src/gtm_engine/excel.py` no longer assigns a worksheet AutoFilter to any structured Table range:

- `_add_table_sheet` leaves filtering to the Table;
- `_replace_table_records` resizes the Table without adding a worksheet filter;
- `_replace_manifest` resizes the Table without adding a worksheet filter.

The visible filter buttons remain available because the Table's own AutoFilter remains intact.
No trading data, configuration, cutoff rule, format, or calculation behavior changed.

`tests/test_excel_adapter.py` now asserts that every worksheet containing a Table has no separate
worksheet AutoFilter. The assertion covers both a newly created input workbook and a populated
result workbook, preventing the defect from returning through either generation path.

### Rebuilt delivered workbook

The legacy source was imported again with Historical End Date `2026-07-10`. The corrected workbook
replaced the exact delivered path:

```text
outputs/legacy_import_070826_v0.3/GTM_Imported_Input.xlsx
```

Its new SHA-256 is:

```text
a7381914a3fe02c1c04337e235f040267a3e6139327a176d6a1e909ec326829b
```

The repaired workbook retains the intended cutoff inventory: 19 sheets, 18 structured Tables,
423 trades, Historical End Date 10 July 2026, and maximum Trade Date 10 July 2026. It remains
formula-free and macro-free.

### Verification evidence

Raw OOXML and independent workbook-library inspection of the delivered file now report:

```text
worksheet-level AutoFilters     0
_xlnm._FilterDatabase names     0
table-owned AutoFilters        18
structured Tables              18
formula cells                   0
VBA parts                       0
external-link parts             0
ZIP errors                      0
```

The repaired input was also passed through `excel-build`. Preflight failed for the already known
18 missing price keys and retained the 12 explicit-zero-quantity warnings; this is the intended
fail-closed economic result. The newly generated `GTM_Failed.xlsx` also retained all 18 Tables,
with zero worksheet AutoFilters, zero formulas, zero VBA parts, and zero external links. This
confirms that the result-writing path no longer reintroduces the Excel defect.

The complete project gate passed:

```text
Ruff format: 28 files already formatted
Ruff lint: all checks passed
mypy strict: no issues in 16 source files
pytest: 49 passed
coverage: 89.72% (gate: 85%)
pip check: no broken requirements
```

The workbook was still open in the user's Excel session during final verification, as shown by
its `~$GTM_Imported_Input.xlsx` owner file. Excel's AppleScript interface returned parameter errors
while that session remained in its repaired/open state. The session was not force-closed and no
open workbook was saved. The user must close the old in-memory workbook without saving, then open
the corrected file from disk. A clean native-Excel reopen is the remaining user-visible smoke
check.

## 2026-08-08 — Cutoff import repaired and daily report published

### Business scope applied

The user confirmed that the historical build stops at 10 July 2026 and that there should be no
price gaps through that date. The following temporary, run-specific price policy was applied:

- fixing prices are retained from source only for TTF DA and PVB Heren;
- every other required fixing price is explicitly populated with zero;
- Exposure uses the imported price curves, except Brent Dated and Henry Hub are explicitly zero;
- Brent and Henry Hub remain present in the data and outputs rather than being removed.

These zeroes are recorded with `POLICY:ZERO:` provenance in a separate run input. They are not a
silent missing-price fallback and do not change the general D-001 fail-closed rule. The repaired
legacy import remains source-faithful; the policy-normalized run input is a separate auditable
artifact.

### Missing-curve diagnosis and importer correction

The first policy build resolved all required fixing keys but exposed 27 missing Exposure curve
keys for TTF DA, Index PVB, and Phys PVB on 30 June and the included July Market Dates. Inspection
of the source workbook and extracted VBA showed that these prices were present but stored outside
the forward matrix:

- `TTF!AS7:BA601` holds the dated TTF prompt/day-ahead series;
- current-month TTF valuation uses that prompt series;
- current-month PVB-family valuation adds the monthly PVB–TTF spread to the TTF prompt value;
- at month end, the prompt contract maps to the next Delivery Month, so the 30 June observation
  maps to July.

`src/gtm_engine/legacy_import.py` now imports this single-date series, skips duplicate same-month
forward rows, and creates current prompt rows for TTF DA/MA and the configured spread-derived
PVB-family products. Later delivery months continue to use the forward matrix. Brent and Henry Hub
import behavior was not removed; their values are overridden only in the separate policy input.

Representative recovered prices are:

```text
Market Date  Underlying  Delivery Month  Price
2026-06-30   TTF DA      2026-07-01      43.665
2026-06-30   Index PVB   2026-07-01      43.11878260869564
2026-07-01   TTF DA      2026-07-01      42.98
2026-07-01   Index PVB   2026-07-01      42.67880434782605
2026-07-10   TTF DA      2026-07-01      48.28
2026-07-10   Index PVB   2026-07-01      49.1485
```

The corrected cutoff import contains 2,376 curve-price rows, up from 2,313. Its SHA-256 is:

```text
7d63fad158cc5eeec280d11f66581200ffc9c5068f1107c458cc52e18225835e
```

### Policy-normalized run input

The separate input `outputs/legacy_import_070826_v0.3/GTM_Run_Input_Zero_Price_Policy.xlsx`
contains:

```text
fixing rows                                      590
required zero fixing rows added                   18
existing non-TTF-DA/non-PVB-Heren fixings zeroed 132
curve rows                                      2376
Brent/HH curve rows explicitly zeroed             486
```

Its SHA-256 is:

```text
e99b7d40b68e10a8b495ad65d3c6b20bf7a8bb957622d82305ad1d2ad5569d02
```

### Published engine result

The engine completed and published run `6c37c08f-b250-4a82-978e-b39e164e9dd4`, build
`GTM3-B924E63617A0D7D1ADF6`, in 0.77 seconds. Preflight reported zero errors and 12 warnings, all
for deliberate zero Daily Qty trades. The stable report is
`outputs/legacy_import_070826_v0.3/daily_report_results/GTM_LATEST.xlsx`; it byte-matches the
run-specific `GTM_Result.xlsx`. Report SHA-256:

```text
34f5bac19cfce1ed3ba548e6d1e6ccd9d5279a2dcf01592bb709f4efc9f31947
```

The published report contains 44 fixing rows, 1,311 Exposure rows, 1,513 Daily P&L rows, and 104
Cumulative P&L rows. The economic reconciliation is:

```text
Daily P&L through 2026-07-10                EUR  4,644,118.310609552378243818915381368
Opening Initial P&L                         EUR 37,445,758.99728646877
Final Cumulative P&L at 2026-07-10          EUR 42,089,877.30789602114824381891538137
```

The opening balance plus the complete daily flow equals the final cumulative balance exactly.
All Brent/HH Exposure rows have zero curve price and zero MtM. All material non-Brent/HH Exposure
rows through 10 July use non-zero imported curves. Fixing output retains non-zero source prices
only for TTF DA; the configured zero-policy products have zero fixing amount.

### Verification evidence

The implementation gate passed after the importer change:

```text
Ruff format: 28 files unchanged
Ruff lint: all checks passed
mypy strict: no issues in 16 source files
pytest: 49 passed
coverage: 89.65% (gate: 85%)
pip check: no broken requirements
```

The final workbook passed ZIP and OOXML integrity inspection. It contains 18 sheets, 17 structured
Tables, no worksheet AutoFilters, no formulas, no VBA parts, and no external links. A native
Microsoft Excel smoke test opened it cleanly, read `status=PUBLISHED` and 18 sheets, and closed it
without saving. The earlier Excel-repair warning is therefore resolved for both the corrected
import and the published report.

## 2026-08-08 — Legacy-format Daily Report D2 added to the Python output workbook

### User clarification

The user reported that `GTM_LATEST.xlsx` contained the detailed engine tables but lacked the
business-facing `Daily Report D2` tab from `Gas_Trading_Model 070826.xlsm`. The required comparison
for the current cutoff is:

```text
D1 = 2026-07-09
D2 = 2026-07-10
```

D2 is the final Market Date in a successful build. D1 is the configured Previous Market Date for
that D2 output, not the previous calendar date inferred independently by Excel.

### Legacy report inspection

The source tab was inspected as both formula and cached-value content and rendered through a
macro-free copy. It is a static presentation sheet, not a calculation sheet. Its three sections
are:

1. Delta P&L by BOOK, split into Delta Exposure MtM, Fixing Amount, Logistical Costs, Fees and
   Optimizations, and Replication.
2. Delta Exposure MtM and Fixing Amount by BOOK and Delivery Month.
3. Delta Exposure MtM by Delivery Month and selected Underlying for all books, with explicit zero
   values.

The selected Underlying columns are Brent Dated, HH, TTF DA, TTF MA, Index PVB, Phys PVB, TVB,
AVB, and PEG. The legacy sheet uses Arial 10, one-decimal numeric display, `mmm/yy` delivery months,
plain gridlines, right-aligned BOOK labels, and centered values. Its old cell values were generated
for 30 June versus 1 July and were not formulas. This confirmed that the new Python adapter should
write a verified report directly rather than reproduce the legacy pivot/formula architecture.

### Implementation

`src/gtm_engine/excel.py` now creates `Daily Report D2` in every successful result workbook and
places it immediately before `EVENT LEDGER`. Failed workbooks do not retain a stale daily report.
The adapter determines D1 and D2 from Cumulative P&L's final Market Date and its Previous Market
Date, then uses only the matching D2 `PnlRow` records.

The report mappings are:

```text
Total Delta PnL                  = sum(total_pnl)
Sum Delta Exposure MtM           = sum(delta_exposure_mtm)
Sum Delta Fixing Amount          = sum(economic fixing_amount from P&L)
Sum Delta Logistical Costs       = sum(logistical_costs)
Sum Delta Fees and Optimizations = sum(fees_and_optimizations)
Sum Delta Replication            = sum(replication)
```

`delta_exposure_mtm` already includes the engine's trade-entry adjustment. The report therefore
does not omit or double-count entry economics. `fixing_amount` is the economic P&L contribution,
not the opposite-signed raw settlement shown on `FIXINGS`.

Section 1 groups by BOOK. Section 2 groups Delta Exposure MtM and Fixing Amount by BOOK and
Delivery Month. Section 3 groups Delta Exposure MtM by Delivery Month and selected Underlying
across all books. Its month axis begins in January of D2's year and runs through December two years
later, extending further only if included engine rows require it. This reproduces the legacy
January 2026–December 2028 view for the present build.

The sheet contains no formulas, macros, external links, copied legacy calculations, or hidden
support ranges. The stored values come from the authoritative Python result; one decimal place is
displayed to match the legacy report.

### Published 9 July versus 10 July result

The policy-normalized input was rebuilt after the adapter change. The accepted run is:

```text
Run ID       36ec0ee1-84d2-47c3-a5a7-c4f4ee4ed386
Build ID     GTM3-B924E63617A0D7D1ADF6
Status       PUBLISHED
Errors       0
Warnings     12 deliberate zero Daily Qty rows
```

The stable workbook remains
`outputs/legacy_import_070826_v0.3/daily_report_results/GTM_LATEST.xlsx`. Its new SHA-256 is:

```text
ff7f7ee9008ec9c9ebbb7aa3175c4b8fdc7723575e81e180a7c63587cd1431f7
```

Section 1 contains seven BOOK rows. Its D2 total is `EUR -560,886.77484712650889153`, matching the
engine's 10 July Total P&L. The components are:

```text
Delta Exposure MtM            -78,104.74803982686303984
Fixing Amount                -495,386.47741935477408
Logistical Costs               -1,484.72999999992267169
Fees and Optimizations           8,063.360740000004
Replication                      6,025.8198720550469
Total P&L                     -560,886.77484712650889153
```

Section 2 contains 48 material BOOK/Delivery Month rows. Its Delta Exposure MtM and Fixing Amount
totals match the same two Section 1 components. Section 3 covers January 2026 through December 2028
and sums to the same `EUR -78,104.74803982686303984` Delta Exposure MtM because every material D2
exposure Underlying is included. Brent and HH remain visible and zero under the approved temporary
policy.

### Verification

The new automated report regression proves D1/D2 selection, section 1 totals, section 2 grouping,
section 3 Underlying placement, absence of formulas, and absence of worksheet AutoFilters. The
complete gate passed:

```text
Ruff format: 27 files already formatted
Ruff lint: all checks passed
mypy strict: no issues in 16 source files
pytest: 50 passed
coverage: 89.89% (gate: 85%)
pip check: no broken requirements
```

The final workbook contains 19 sheets, including `Daily Report D2`, and 17 structured Tables. It
contains no formulas. A visual render confirmed that the report matches the legacy hierarchy,
headers, column widths, numeric format, month format, alignment, and explicit-zero matrix while
keeping the sections compact. Native Microsoft Excel opened the workbook cleanly, returned
`status=PUBLISHED`, reported 19 sheets, read D1 as 9 July and D2 as 10 July, and closed without
saving.

The short and detailed manuals now identify `Daily Report D2` as the first business output to read
after Status and Validation. The engine specification records the report's D1/D2 semantics,
aggregation map, sign choice, and reconciliation requirement.

## 2026-08-08 — Source repository organized for public GitHub publication

The GTM v0.3 source tree was converted into a Git repository with `main` as its initial branch. A
GitHub repository was created at `https://github.com/vasilybelokurov/gtm-engine`; it was changed to
public visibility at the user's direction before the first source push. The local `origin` remote
uses the repository's authenticated SSH URL.

### Publication boundary

The repository is intended to make the Python engine, Excel adapters, tests, specifications, and
operating instructions reproducible without publishing confidential trading data or generated
artifacts. The checked publication set is approximately 796 KiB and contains source code,
documentation, test code, CI configuration, the golden-test specification/report, and the
technical handover document.

The following stay local and are explicitly ignored:

- every `.xls`, `.xlsx`, `.xlsm`, and `.xlsb` workbook, including the 23 MiB legacy model and its
  helper workbooks;
- `outputs/`, including imported inputs and published daily-report workbooks;
- `analysis/`, including extracted OOXML, VBA binaries, and reverse-engineering evidence;
- `runtime_tests/`, `local_data/`, `.venv/`, caches, logs, and temporary Excel owner files;
- `questions_answers.txt` and its backup, in accordance with the earlier instruction to disregard
  them.

The boundary was checked with `git check-ignore`; no private workbook, generated report, extracted
analysis tree, runtime workbook, or virtual environment appears in the proposed tracked-file list.
A scan of the proposed source and documentation set found no embedded access tokens, private keys,
or passwords.

### Repository and documentation work

The root now includes:

- a detailed `README.md` covering architecture, business scope, the Excel and normalized-bundle
  workflows, commands, outputs, acceptance boundaries, design principles, and document navigation;
- `INSTALL.md` with clean-clone installation, verification, common operations, upgrade, and
  troubleshooting instructions;
- `CONTRIBUTING.md`, `SECURITY.md`, and `CHANGELOG.md` for maintainers;
- `docs/REPOSITORY_DATA_POLICY.md`, which makes the source/data boundary and safe publication rules
  explicit;
- `.gitignore`, `.gitattributes`, `.editorconfig`, and `.python-version` for predictable local and
  cross-platform behavior;
- `.github/workflows/ci.yml`, which runs formatting, lint, strict type checking, the coverage gate,
  and dependency-consistency checks on pushes and pull requests;
- `src/gtm_engine/py.typed`, so downstream tooling recognizes the package as typed.

`pyproject.toml` now carries the README and repository/documentation metadata and includes the
typing marker in package data. All local Markdown links in the publication set resolve.

### Pre-publication release gate

The editable package was reinstalled from the finalized metadata. The complete local gate passed:

```text
Ruff format: 28 files formatted/already formatted
Ruff lint: all checks passed
mypy strict: no issues in 16 source files
pytest: 50 passed
coverage: 89.89% (gate: 85%)
pip check: no broken requirements
CLI smoke test: passed
typed-package marker: present
```

The first commit and push, clean-clone verification, and GitHub Actions result will be recorded in
a follow-up entry after publication is complete.

## 2026-08-08 — Public repository published and independently verified

The reviewed 54-file publication set was committed as:

```text
Commit   2c931ca6996df0f1ca8b0df6cf6b53677a622665
Subject  Initial GTM v0.3 reference engine
Branch   main
Remote   git@github.com:vasilybelokurov/gtm-engine.git
Web      https://github.com/vasilybelokurov/gtm-engine
Access   Public
```

The first HTTPS push was rejected because the GitHub CLI OAuth token did not carry GitHub's
separate `workflow` scope and the commit includes `.github/workflows/ci.yml`. The machine's
existing authenticated GitHub SSH identity was verified as `vasilybelokurov`; `origin` was changed
to SSH and the same commit then pushed successfully without removing or weakening CI. Local `main`
now tracks `origin/main`.

### Clean public-clone verification

The repository was cloned from its public HTTPS URL into a new temporary directory, with no use of
the original checkout, its `.venv`, or any ignored workbook. A new Python 3.13 environment was
created and the documented developer installation command succeeded:

```text
python3.13 -m venv .venv
.venv/bin/python -m pip install -e '.[dev]'
```

The clean checkout then passed:

```text
Ruff format check: 43 files already formatted
Ruff lint: all checks passed
mypy strict: no issues in 16 source files
pytest: 49 passed, 1 skipped in 3.02 seconds
pip check: no broken requirements
CLI help smoke test: passed
```

The skipped test is the intentional optional inventory check for the private
`Gas_Trading_Model 070826.xlsm`. Its absence proves that a public checkout does not depend on or
contain the legacy production workbook. All synthetic engine, adapter, regression, property, and
performance tests passed.

### Hosted CI evidence

GitHub Actions independently ran the repository CI workflow for the initial commit and completed
successfully:

```text
Run ID      31252425332
Commit      2c931ca6996df0f1ca8b0df6cf6b53677a622665
Conclusion  success
URL         https://github.com/vasilybelokurov/gtm-engine/actions/runs/31252425332
```

The source repository is therefore public, installable without the local data tree, and guarded
by the same formatting, lint, type, test-coverage, and dependency checks used locally. Operational
workbooks, imported data, calculated results, extracted workbook evidence, and local environments
remain outside Git under the documented data policy.

### CI runtime maintenance

The follow-up CI run for the publication-evidence commit passed every job but reported that the
older `actions/checkout@v4` and `actions/setup-python@v5` definitions target deprecated Node 20.
GitHub's official action documentation was checked on 8 August 2026. Both actions now document v7
as the current major; checkout v5 and setup-python v6 introduced Node 24, while v7 carries the
current action internals and usage examples. The workflow was therefore updated to
`actions/checkout@v7` and `actions/setup-python@v7`. It retains explicit `contents: read`
permissions and the fixed Python 3.13 target.

## 2026-08-08 — Windows timezone dependency corrected

The first automated gate in a newly installed Windows checkout exposed a packaging omission.
Python's `zoneinfo` could not resolve the configured `Europe/Madrid` timezone because Windows does
not normally provide an IANA timezone database and the project did not declare the PyPI `tzdata`
fallback. Builds failed at pipeline timestamp initialization with `ZoneInfoNotFoundError`; this
was an installation defect, not an economic-engine failure.

`pyproject.toml` now declares `tzdata>=2026.1,<2027` for `sys_platform == 'win32'`. The platform
marker avoids installing redundant timezone data on systems that normally provide the IANA
database. `INSTALL.md` and `CHANGELOG.md` record the Windows behavior.

After installing `tzdata 2026.3`, the complete Windows gate passed:

```text
Ruff format check: 28 files already formatted
Ruff lint: all checks passed
mypy strict: no issues in 16 source files
pytest: 49 passed, 1 skipped in 41.12 seconds
coverage: 87.93% (gate: 85%)
pip check: no broken requirements
```

The project was reinstalled from the corrected editable metadata before this final gate; pip
resolved `tzdata 2026.3` as a runtime requirement and reported no broken dependencies. The skipped
test remains the intentional private legacy-workbook inventory test. The longer test runtime
reflects the managed Windows execution environment and coverage instrumentation.

## 2026-08-08 — Standalone Daily Report D2 verification workbook generated

The policy-normalized verification input was rebuilt with status `PUBLISHED`, Build ID
`GTM3-B924E63617A0D7D1ADF6`, zero validation errors, and the 12 expected deliberate-zero-quantity
warnings. The authoritative Python result selected D1 9 July 2026 and D2 10 July 2026.

A standalone, local-only workbook was written to
`verification/GTM_Daily_Report_D2_2026-07-10.xlsx`. It contains only the generated
`Daily Report D2` sheet, with 108 rows, 11 columns, and no formulas. Section 1 reconciles to D2
Total P&L of `EUR -560,886.7748471264`. The source input
`verification/GTM_Run_Input_Zero_Price_Policy.xlsx` remained unchanged and does not contain the
report sheet. The generated `.xlsx` remains excluded from Git under the repository data policy.

## 2026-08-08 — Daily Delta P&L matrix by Market Date and BOOK generated

A second standalone, local-only verification workbook was generated from the published
`GTM3-B924E63617A0D7D1ADF6` cumulative P&L output:
`verification/GTM_Delta_PnL_by_Market_Date_and_BOOK.xlsx`.

The workbook contains one sheet, `Delta PnL by BOOK`, with Market Date in the first column and the
13 active BOOKS in the remaining columns. Its eight data rows cover the configured Market Dates
from 1 July through 10 July 2026. Each cell is the engine's daily Total P&L aggregated by Market
Date and BOOK; absent movements are displayed as explicit zeroes. The sheet contains no formulas
and includes a frozen date column/header, filters, monetary formatting, and a red/white/green
value scale.

The matrix total is `EUR 4,644,118.310609552378243818915`. It reconciles to the sum of
`pnl.csv.total_pnl`, `EUR 4,644,118.310609552378243818913`; the residual is below representable
economic significance and far inside the approved EUR 0.01 P&L tolerance. The output workbook's
SHA-256 is `5c37954cdb737004027837bd108a0182d03f87b7e39fa01553afd51b2369c6dc` and the file remains
excluded from Git under the repository data policy.

## 2026-08-09 — Initial workbook D2 diagnostic comparison for 9/10 July 2026

The local-only workbook `verification/Gas_Trading_Model 070826.xlsm` was inspected read-only at
the user's request. Its SHA-256 is
`bdb1e3870f38437b2d464ea3c3feb89931e31ad4a8deadda44f98c777ce614d7`. No macro, formula
recalculation, link update, external refresh, or save was performed. The workbook was treated as
diagnostic evidence only, not as an acceptance oracle.

The saved workbook does not contain one coherent legacy result for D1 9 July and D2 10 July:

- `Daily Report D2` is stale at D1 30 June / D2 1 July;
- `Daily Delta PnL` contains no values or formulas;
- the cached `DAILY PNL` row for 10 July is zero in every component and BOOK;
- `DAILY PNL DATA`, refreshed 20 July, contains a 10 July total of
  `EUR 4,706.789812054237`;
- `PNL DATA`, refreshed 3 August, contains a different 10 July total of
  `EUR 417,286.513907052471`.

Against the published Python result `GTM3-B924E63617A0D7D1ADF6`, whose 10 July Total P&L is
`EUR -560,886.774847126509`, the older `DAILY PNL DATA` snapshot differs by
`EUR +565,593.564659180746`. Its component comparison is:

```text
Component                   Python                 DAILY PNL DATA       Legacy - Python
Delta Exposure MtM          -78,104.748039826863   -10,867.120800000736  +67,237.627239826127
Economic Fixing Amount     -495,386.477419354774         0.000000000000 +495,386.477419354774
Logistical Costs             -1,484.729999999923    +1,484.729999999923   +2,969.459999999846
Fees and Optimizations       +8,063.360740000004    +8,063.360740000004        0
Replication                  +6,025.819872055047    +6,025.819872055047        0
```

The newer `PNL DATA` layer is structurally closer to Python. Its adjusted Exposure differs only
by `EUR +3.058248294`, entirely from Brent Dated, which the temporary run policy sets explicitly
to zero. Its Fixing Amount is `EUR +495,388.203698590195`: it uses the raw settlement sign instead
of the inverse economic P&L sign and also includes `EUR +1.726279235` of Brent that is zero under
the temporary policy. It contains no Logistics, Fees/Optimizations, or Replication rows.

The older `DAILY PNL DATA` exposure snapshot is not simply Python before trade-entry adjustment;
it differs materially from Python gross Exposure by BOOK. It includes BOOK movements absent from
the normalized build and omits others, consistent with the previously documented incomplete
formula architecture and stale/incoherent build generations. Its Fixing component is wholly
absent. Fees and Replication agree exactly with Python, while Logistics has the opposite sign;
the source `COSTS` row for 10 July stores positive costs totalling `EUR 1,484.73` and still carries
the workbook note `CONFIRMAR SIGNO CON OPS`.

Conclusion: the differences are not one unexplained calculation residual. They are the combined
effect of stale report dates, disconnected calculation generations, missing legacy fixing P&L,
raw-versus-economic fixing sign, omission of operating flows in v2, the explicit temporary
Brent/HH zero-price policy, the unresolved Logistics sign, and materially different older
Exposure populations. No legacy layer can be accepted as the expected result without first
selecting and repairing a coherent methodology path.

## 2026-08-09 — Test 1 fixing-methodology input created

A local-only validation workbook was created at `verification/Input test fixing.xlsx` to begin a
simple business review of fixing behavior across every active underlying. It uses Initial Market
Date 30 December 2025, Historical Start Date 31 December 2025, and Historical End Date 10 July
2026. Initial Exposure is empty and Initial P&L contains one required zero row for each of the 13
active BOOKS. Operating flows are empty so the eventual result isolates trade, fixing, exposure,
and trade-entry effects.

The workbook contains one ACTUAL BUY trade in BOOK `CGTO` for each of the 18 active source
underlyings. Every trade has Trade Date 31 December 2025, delivery from 1 January through
31 December 2026, and execution price 100 in the configured price contract. The 16 EUR/MWh gas
underlyings use Daily Qty 10; Brent Dated and HH use deliberate zero quantities pending a separate
USD and unit-specific test. The calendar, fixing-price table, and curve-price table were copied
unchanged from `GTM_Run_Input_Zero_Price_Policy.xlsx`.

The generated workbook was loaded back through the strict Excel adapter. It contains 18 trades
(16 material), 13 zero Initial P&L rows, 1,097 calendar rows, 590 fixing-price rows, and 2,376
curve-price rows. The three copied market-data tables compare exactly with the source bundle.

An in-memory preflight diagnostic established that the unchanged fixing-price table does not
cover all required lookup keys from January through 10 July 2026 for these full-year trades. The
current test therefore fails closed with missing-fixing-price validation rather than producing
economic output. No missing price was synthesized or silently replaced. Before generating the
test output, the business test must either provide an approved synthetic fixing-price grid for
the required dates or narrow the trade and historical horizon to the available price coverage.

## 2026-08-09 — Test 1 fixing, P&L eligibility, and FX policy implemented

Business clarification separated the fixing schedule from its P&L recognition. Every material
fixing closes Exposure according to the configured methodology. Daily products retrieve their
price by Delivery Day. Month Ahead allocates the monthly delivery volume across eligible Fixing
Dates in the preceding month and retrieves a distinct price for each tranche's Fixing Date.

The underlying contract now declares whether raw fixing settlement enters P&L. TTF DA, Brent
Dated, HH, PVB Heren DA, and PVB Heren DA (Delivery) are included for Test 1. Other products retain
informative non-zero synthetic fixing prices but contribute zero fixing P&L because their economics
will be supplied through Replication P&L. This prevents the previous zero-price convention from
hiding whether the correct price key was selected.

An FX input table was added. Rates are currency units per EUR, with exact-date or latest-prior
lookup and fail-closed validation. Brent Dated remains bbl and USD/bbl; HH remains MMBtu and
USD/MMBtu. Raw fixing and Exposure amounts remain USD. Fixing P&L converts at EURUSD spot on Fixing
Date, Exposure MtM at Market Date, and trade-entry adjustment at applied Market Date.

`verification/Input test fixing.xlsx` was regenerated with 18 BUY trades in CGTO, Daily Qty 10,
delivery 1 January through 31 December 2026, 6,606 synthetic fixing prices, 15,678 curve prices,
and 134 EURUSD observations fixed at 1.20 USD/EUR. The source calendar was retained and extended
only with 30 December 2025, required as the previous Market Date for the first 31 December output.
The strict Excel round trip and in-memory build completed `VERIFIED` with zero errors, warnings, or
information items, producing 11,011 fixing rows, 10,985 Exposure rows, and 10,986 P&L rows. The
final user-designed output workbook has not yet been generated.

The final repository gate passed on Windows: Ruff format and lint passed, mypy strict reported no
issues in 17 source files, pytest completed with 51 passed and 1 intentional private-workbook
skip, coverage was 88.02% against the 85% gate, and pip reported no broken requirements.

## 2026-08-09 — Test 1 fixing validation output generated

The local-only workbook `verification/output test fixing.xlsx` was generated from the verified
Test 1 input. `Fixing volume by fixing date` contains every calendar date from 31 December 2025
through 10 July 2026 and aggregates fixing volume by source underlying. `Fixing PnL by fixing
date` uses the same 192-day axis and reports economic fixing P&L in EUR on Fixing Date, including
explicit zeroes for products excluded from fixing P&L.

`Fixing PnL by market date` contains the 133 configured Market Dates in the same period and moves
each economic fixing contribution to its applied Market Date. All three matrices preserve the 18
source underlyings as separate columns. The fixing-date and market-date P&L matrices both total
EUR 460,920, proving that deferral to Market Date changes timing presentation but neither loses nor
duplicates fixing P&L. The workbook contains values only and remains excluded from Git.

## 2026-08-09 — Test 1 regenerated without delivery volume

The definitive `verification/Input test fixing.xlsx` and
`verification/output test fixing.xlsx` workbooks were overwritten after adding the explicit
delivery-election contract. The input retains the same 18 trade rows and contains no delivery
elections. The existing `PVB Heren DA (Delivery)` trade is retained with Daily Qty zero; the new
`TTF DA (Delivery)` product is derived only through delivery elections and therefore has no direct
trade row.

TTF DA and PVB Heren DA base fixing prices are zero. Their `(Delivery)` price series remain
available for future elected delivery volume, but neither contributes volume or P&L in this run.
The strict build completed `VERIFIED` with zero errors and one expected zero-quantity warning.
Both delivery columns total exactly zero in all three output matrices. The output contains 192
calendar dates in each Fixing Date sheet, 133 Market Dates in the Market Date sheet, and all 19
active source-underlying columns.

## 2026-08-09 — Delivery categories removed from authoritative inputs

The physical-delivery design was simplified following business review. `TTFDA Heren` and
`PVB Heren` are now the only authoritative source products for this methodology in UNDERLYINGS,
TRADES, and FIXING PRICES. Their fixing-price series contain the real Delivery Day prices (50 and
70 respectively in Test 1); zero prices are no longer used to control P&L eligibility.

`TTFDA Delivery` and `PVB Heren Delivery` now exist only as engine-derived reporting categories.
An approved DELIVERY ELECTION splits the applicable daily source volume internally, and the
delivery leg reuses its base product's fixing-price series. The non-delivery leg remains excluded
from fixing P&L for recognition through Replication P&L, while the elected delivery leg contributes
fixing P&L.

The definitive Test 1 input and output workbooks were overwritten. The input contains 17 trades,
17 configured source underlyings, no delivery trades, no delivery fixing-price rows, and no
delivery elections. The strict build completed `VERIFIED` with zero errors, warnings, or
information items. Both derived delivery columns appear in all three output matrices and total
exactly zero for the current no-delivery test.

## 2026-08-09 — Full July 2026 delivery scenario created

Two additional local-only workbooks were created without replacing the no-delivery baseline:
`verification/Input test fixing delivery Jul26.xlsx` and
`verification/output test fixing delivery Jul26.xlsx`. All authoritative inputs remain identical
to the baseline except for two DELIVERY ELECTIONS decided on 30 June 2026. Each election assigns
the full BUY volume of 10 MWh/day in BOOK CGTO from 1 through 31 July 2026 to delivery, one for
`TTFDA Heren` and one for `PVB Heren`.

The generated Excel input was reopened through the strict adapter and rebuilt successfully. The
run completed `VERIFIED` with zero errors, warnings, or information items. Through the Historical
End Date of 10 July, the schedules contain 130 MWh of fixing volume in each derived delivery
category because 13 July delivery days have fixed by that cutoff. The report shows EUR 6,500 of
TTFDA Delivery fixing P&L and EUR 9,100 of PVB Heren Delivery fixing P&L; both base-product fixing
P&L columns remain zero. The full test suite passed with 52 tests and one intentional private
legacy-workbook skip.

## 2026-08-09 — Fixing report Excel table repair

Excel Desktop reported a recovery operation on the three fixing-report tables. The report writer
had assigned an AutoFilter to each worksheet range and then added an Excel Table over the same
range, whose table definition carries its own AutoFilter. OpenPyXL accepted the overlapping
definitions, but Excel removed the tables during recovery.

The redundant worksheet-level AutoFilter was removed from the generator. Both
`output test fixing.xlsx` and `output test fixing delivery Jul26.xlsx` were regenerated and
overwritten without changing their calculated values. Package-level verification confirms that
each workbook contains exactly three table parts, one table per worksheet, and no overlapping
worksheet-level AutoFilter definitions.

## 2026-08-09 — Exposure validation report created

The latest full-July-delivery input was processed into a separate local-only workbook,
`verification/output test fixing exposure Jul26.xlsx`; no existing input or output was
overwritten. The original three fixing matrices remain present.

`Exposure data` adds 10,986 auditable product-level rows with Market Date, Previous Market Date,
BOOK, canonical Underlying, Delivery Month, Trade Source, Scenario, Exposure Volume, Curve Price,
Exposure MtM, Gross Delta Exposure MtM, Trade Entry Adjustment, final Delta Exposure MtM,
Currency, explicit-closure flag, and Build ID. It is stored as a native Excel table suitable for
filtering and pivot analysis.

`Exposure by Market Date` and `Delta Exposure MtM` provide yellow dropdown selectors for Market
Date and BOOK. Their formulas aggregate the detail table by canonical Underlying. The exposure
view shows both volume and MtM. The delta view shows Gross Delta Exposure MtM, Trade Entry
Adjustment, and their final Delta Exposure MtM result. The workbook uses hidden, named validation
ranges and requests a full automatic Excel recalculation on open.

Package verification confirmed four valid table parts, no overlapping worksheet/table filters,
valid selector formulas, two dropdowns per dynamic sheet, and exact detail-row reconciliation with
the verified engine result. The transient Windows temporary-directory failure in the full test run
was rerun at the affected test level and passed.

## 2026-08-09 — Exposure selector views changed to delivery-month matrices

The two interactive exposure views were revised after business review. Both retain the yellow
Market Date and BOOK selectors. Their result area is now a matrix whose first column contains all
36 Delivery Months from January 2026 through December 2028 and whose remaining columns contain the
nine canonical underlyings present in the verified engine result.

`Exposure by Market Date` reports Exposure Volume for each Delivery Month and Underlying.
`Delta Exposure MtM` reports the final P&L component named Delta Exposure MtM for the same matrix.
The underlying detail table still retains Exposure MtM, Gross Delta Exposure MtM, and Trade Entry
Adjustment for audit and reconciliation. Formula, dropdown, date-boundary, table-package, and
workbook-reload checks passed on the definitive
`verification/output test fixing exposure Jul26.xlsx` file.

## 2026-08-09 — Clean-clone synthetic model reconstruction

The repository was made self-contained for reconstruction of the reviewed synthetic fixing,
full-July-delivery, and exposure validation model. `scripts/create_fixing_test.py` no longer
requires a prior local Excel workbook. When `--source` is omitted it generates the complete input
from versioned code and the reviewed mapping: 13 BOOKS, 17 BUY trades, zero opening P&L, calendar,
curve prices, real TTFDA/PVB fixing prices, the remaining synthetic fixing series, EURUSD rates,
and optional July 2026 delivery elections.

The generator writes the input workbook, reloads it through the strict Excel adapter, requires a
`VERIFIED` engine result with no validation findings, and only then writes the fixing and exposure
report. Generated workbooks remain ignored reproducible artifacts; no private or production data
was added to Git.

An end-to-end portable test now invokes this clean-checkout path in a temporary directory and
checks the rebuilt input, engine result, report sheet contract, and January 2026 through December
2028 exposure-matrix boundaries. A manual reconstruction in a new local directory also completed
`VERIFIED` with 17 trades, two delivery elections, and the expected 130 MWh fixing volume in each
derived delivery category through 10 July. Repository metadata and installation documentation now
point to `aitorayerdi-git/gtm-engine`.

The final clean-clone gate passed: Ruff format and lint, mypy strict, and pip dependency checks
reported no issues; pytest completed with 53 passed and one intentional private-workbook skip;
branch coverage was 87.80% against the 85% required gate.

## 2026-08-11 - Foto FO cost update workflow and manual compensation

A macro-enabled local Input workbook was built from the authorized workbook without adding any
private workbook or production data to Git. The versioned VBA module opens the approved Foto FO
SharePoint source, validates its required sheets, headers, and Market Date coverage, calculates
BOOK-level logistics, fee, optimization, and replication flows, publishes them to OPERATING
FLOWS, and persists the cumulative source state needed for the next delta. MANUAL CHANGES now
provides an UPDATE FOTO FO button plus timestamp, user, source, Market Date, and OK/ERROR status;
UPDATE LOG retains an execution audit.

A COSTS worksheet was added with the reviewed Gas Trading Model structure: Market Date, logistics
cost by BOOK, exchange fees, total, source, timestamp, user, and comment. Automatic runs write
AUTO rows. If a Market Date was missed, an operator can enter the daily delta in a MANUAL row.
On the next automatic Market Date the macro subtracts manual rows after the prior saved state and
before the new target date from the cumulative Foto FO delta. Once the automatic state advances,
the same manual row is outside the adjustment interval and is not applied again.

The local Foto FO validation completed OK for 13 active BOOKs, with three non-zero BOOKs and 664
OPERATING FLOWS rows. A two-case regression inserted a EUR 123.45 manual CGA_SHT1 adjustment; the
next automatic flow differed from the baseline by exactly EUR 123.45. MarketView was disabled
only inside isolated Excel test instances and restored to LoadBehavior 3 after each test. The
generated workbooks and error screenshots remain ignored local artifacts under the repository
data policy.

## 2026-08-12 - Missed Foto FO date guard

The COSTS SOURCE dropdown now offers only `MANUAL`; `AUTO` is reserved for rows published by the
macro. Before opening Foto FO or changing any output, the update checks every configured Market
Date strictly between the saved state and Last Market Date. Each skipped date must have one
valid `MANUAL` COSTS row with numeric or blank BOOK amounts and a non-empty comment. A confirmed
zero-cost day is represented explicitly by a zero-valued MANUAL row.

If any skipped date lacks that acknowledgement, the update fails closed, lists the missing dates,
and leaves FOTO FO STATE unchanged. Excel regression checks confirmed the MANUAL-only dropdown,
the missed-date error, the non-advancing state, and the existing EUR 123.45 compensation result.

## 2026-08-12 - Editable UK holiday calendar

MANUAL CHANGES now contains an editable `tblUKHolidays` table seeded with the official England
and Wales bank holidays for 2025 through 2028. Operators can change a date or description, add a
row, or deactivate a holiday with the YES/NO dropdown. The table, rather than hard-coded VBA, is
the holiday authority.

The user-facing heading is `LONDON BANK HOLIDAYS (EDITABLE)`. London has no separate GOV.UK bank
holiday division; it follows the official England and Wales dates. Scotland and Northern Ireland
holidays are therefore not included.

MARKET CALENDAR calculates Is Market Day as a weekday that is not an active holiday in that
table. Excel verification classified Saturday 29 August 2026 and the UK Summer bank holiday on
Monday 31 August as non-market days, and Tuesday 1 September as a market day.

## 2026-08-12 - Foto FO opening cost baseline

COSTS now permits a single `BASELINE` row in addition to `MANUAL`; `AUTO` remains macro-only.
BASELINE represents the cumulative Foto FO balance through the Market Date immediately before
the first automatic run. It advances the effective opening boundary, is subtracted once from the
first calculated cumulative delta, is never published as an OPERATING FLOWS/P&L row, and becomes
inactive automatically after FOTO FO STATE advances beyond its date.

The update rejects multiple baselines, a baseline on or after Last Market Date, non-numeric
amounts, and a missing explanatory comment. The Excel regression proved that a EUR 123.45 opening
balance reduces the first automatic CGA_SHT1 flow by exactly EUR 123.45 without changing the
pre-existing OPERATING FLOWS population for the baseline date. Manual compensation and missed-date
guard regressions continued to pass.
## 2026-08-12 - Idempotent first AUTO rebuild and fee sign correction

Re-running the first AUTO Market Date after correcting BASELINE or MANUAL inputs now rebuilds the
row from the current Foto FO cumulative snapshots, baseline, and intervening manual deltas. It no
longer carries the previously published first-day flow through the same-date incremental path.

The Canones cumulative snapshot now retains the sign stored by Foto FO, consistently with the
COSTS contract and the manual baseline. Excel serial dates loaded through `Value2` are accepted
for Canones, optimizations, replication, and MAIN; previously those array values were skipped by
`IsDate`. The saved 11 August case now rebuilds its logistics values from corrected inputs rather
than carrying the first erroneous publication.
## 2026-08-12 - Foto FO costs restricted to Delta de costes

Business review clarified that the COSTS worksheet must source costs exclusively from
`Delta de costes`. Canones and Optimizaciones continue to populate the Fees and Optimizations
component of OPERATING FLOWS, while Index replication and MAIN continue to populate Replication;
those P&L components are no longer copied into COSTS. Logistics deltas are rounded to cents after
baseline and manual compensation.
