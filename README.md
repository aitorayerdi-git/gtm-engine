# GTM Python Engine

GTM v0.3 is a headless reference engine for the Gas Trading Model. It reconstructs gas-trading
fixings, open exposure, daily profit and loss (P&L), and cumulative P&L from a controlled opening
state, trades, calendars, prices, and operating flows.

Python is the calculation authority. Excel supplies inputs and displays outputs; the engine does
not depend on workbook formulas, VBA, or an open Excel process.

```text
Legacy .xlsm or normalized source data
                    ↓
          macro-free Excel input
                    ↓
          authoritative Python engine
                    ↓
   Excel report + CSV/JSON audit evidence
```

The repository contains code, synthetic tests, reviewed mappings, specifications, and operating
instructions. It deliberately excludes production workbooks, trades, market prices, generated
reports, and extracted legacy evidence. See [Data and repository policy](docs/REPOSITORY_DATA_POLICY.md).

## Release status

Version 0.3 implements and tests:

- strict fail-closed input validation;
- SETUP-driven BOOK, Underlying, fixing-method, curve, and fixing-price mappings;
- opening positions and post-cut-off actual or simulation trades;
- WITHINDAY, DAY_AHEAD, HEREN, MONTH_AHEAD, and BRENT_HH fixing schedules;
- same-day fixing eligibility and release-1 late-trade rejection;
- weekend or holiday event deferral to the next configured Market Date;
- exact required-price validation without silent carry-forward or interpolation;
- the approved Brent and Henry Hub current-month curve selection rule;
- explicit one-time zero rows when positions close;
- trade-entry, fixing, operating, daily, and cumulative P&L;
- read-only legacy `.xlsm` import into normalized files and a macro-free Excel interface;
- deterministic Build IDs, output hashes, validation evidence, and atomic publication;
- a legacy-format `Daily Report D2` tab generated from verified Python output.

The portable quality gate includes a complete clean-checkout reconstruction of the synthetic
fixing, delivery, and exposure workbooks. The optional private-workbook inventory test skips when
the legacy workbook is not present. The performance target is less than 60 seconds for a full
production rebuild and less than 2 GiB peak memory.

Production acceptance remains a business decision. It requires approval of the named golden
cases, especially product-specific DAY_AHEAD weekend pricing, the Logistics sign, the opening
13-BOOK Initial P&L bridge, and Brent/HH roll examples. A temporary zero-price policy used for one
historical July 2026 run is documented in the specification; it is not a general engine default.

## Start here

For a new checkout:

```sh
git clone https://github.com/aitorayerdi-git/gtm-engine.git
cd gtm-engine
python3.13 -m venv .venv
.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install -e '.[dev]'
.venv/bin/pytest -q
```

Windows PowerShell uses `.venv\Scripts\python.exe` and `.venv\Scripts\gtm-engine.exe`. The
[installation guide](INSTALL.md) gives complete macOS, Linux, and Windows instructions, including
verification and troubleshooting.

After installation, read the [one-page Quick Start](docs/GTM_QUICK_START.md). The
[detailed Excel manual](docs/GTM_EXCEL_INTERFACE_GUIDE.md) defines every input, output, sign, and
reconciliation step.

## Reproduce the synthetic validation model from a clean clone

No local workbook is required to rebuild the reviewed synthetic fixing, full-July-delivery, and
exposure test. From the repository root run:

```sh
.venv/bin/python scripts/create_fixing_test.py \
  --input outputs/reproduction/Input_test_fixing_delivery_Jul26.xlsx \
  --output outputs/reproduction/output_test_fixing_exposure_Jul26.xlsx \
  --july-delivery
```

Windows PowerShell:

```powershell
.venv\Scripts\python.exe scripts\create_fixing_test.py `
  --input outputs\reproduction\Input_test_fixing_delivery_Jul26.xlsx `
  --output outputs\reproduction\output_test_fixing_exposure_Jul26.xlsx `
  --july-delivery
```

The script creates the calendar, 17 trades, curve prices, fixing prices, FX rates, zero opening
balances, and the two July 2026 delivery elections from versioned code and the reviewed mapping.
It then writes the input, reloads it through the strict Excel adapter, requires a `VERIFIED` build,
and writes all fixing and exposure report sheets. Generated `.xlsx` files remain intentionally
outside Git because they are reproducible artifacts.

## Normal Excel workflow

### 1. Create or import an input workbook

Create an empty macro-free interface with the reviewed SETUP mapping:

```sh
.venv/bin/gtm-engine excel-template \
  --output outputs/gtm_excel_v0_3/GTM_Excel_Interface_v0.3.xlsx \
  --mapping docs/GTM_ACTIVE_SETUP_MAPPING_v0.3.csv
```

Alternatively, import cached values from a private legacy workbook without opening Excel or
running its macros:

```sh
.venv/bin/gtm-engine legacy-import \
  --workbook "/path/to/Gas_Trading_Model.xlsm" \
  --output outputs/legacy_import_v0.3 \
  --historical-end 2026-07-10
```

The cutoff also acts as the last included Trade Date. The importer records excluded rows, source
ranges, the source hash, extracted counts, and review items.

### 2. Build

Save and close the input workbook, then run:

```sh
.venv/bin/gtm-engine excel-build \
  --workbook outputs/legacy_import_v0.3/GTM_Imported_Input.xlsx \
  --output outputs/gtm_excel_runs
```

On success, open `outputs/gtm_excel_runs/GTM_LATEST.xlsx`. On failure, open the newest
`GTM_Failed.xlsx` under `outputs/gtm_excel_runs/failed/` and read `VALIDATION`. A failed build never
replaces the last published result.

macOS users can also drag an input workbook onto `scripts/GTM_Build.command`.

### 3. Read the result

Use this order:

1. `START HERE`: require `PUBLISHED` status.
2. `VALIDATION`: require zero errors and review every warning.
3. `Daily Report D2`: compare the final Market Date, D2, with its configured predecessor, D1.
4. `DAILY PNL`, `EXPOSURE`, and `FIXINGS`: inspect the detailed economic rows.
5. `CUMULATIVE PNL`: reconcile opening Initial P&L plus later daily flows.
6. `EVENT LEDGER`: trace initial, trade, and fixing events.
7. `BUILD MANIFEST`: record the Build ID, versions, hashes, counts, totals, and runtime evidence.

`Daily Report D2` reproduces the legacy presentation without reproducing its calculation sheets.
It shows P&L by BOOK, P&L by BOOK and Delivery Month, and Delta Exposure MtM by Delivery Month and
selected Underlying. The cells contain Python-generated values, not formulas.

## Run without Excel

The core engine can build a normalized directory directly:

```sh
.venv/bin/gtm-engine build \
  --input path/to/normalized_bundle \
  --output outputs/normalized_runs
```

The bundle contains:

```text
bundle.json
books.csv
underlyings.csv
market_calendar.csv
initial_exposure.csv
initial_pnl.csv
trades.csv                 optional
curve_prices.csv           optional as a file; required keys still fail closed
fixing_prices.csv          optional as a file; required keys still fail closed
operating_flows.csv        optional
```

The Excel and directory adapters feed the same typed `InputBundle` to the calculation pipeline.

## Repository layout

```text
src/gtm_engine/             engine, adapters, validation, and CLI
tests/                      portable synthetic, golden, property, and regression tests
docs/                       specifications, mappings, decisions, and user manuals
verification/               regression-test source and signed-off test report
scripts/                    macOS launcher and workbook-preview support
README.md                   project overview and operating entry point
INSTALL.md                  installation, upgrade, and troubleshooting guide
CONTRIBUTING.md             development and change-control rules
SECURITY.md                 private-data and vulnerability-reporting rules
CHANGELOG.md                release history
JOURNAL.md                  chronological project investigation and implementation record
RUNTIME_TEST_PLAN.md        controlled legacy runtime-test plan
```

Private workbooks may exist locally at the repository root for importer testing, but `.gitignore`
prevents Git from tracking them. Generated artifacts belong under `outputs/`, which is also
ignored.

## Development

Install the development extras, then run the complete gate:

```sh
.venv/bin/ruff format --check src tests scripts
.venv/bin/ruff check src tests scripts
.venv/bin/mypy src
.venv/bin/pytest --cov=gtm_engine --cov-report=term-missing
.venv/bin/python -m pip check
```

Coverage must remain at or above 85%. Economic-rule changes require an independently calculated
regression case, a specification update, and a `JOURNAL.md` entry. See [CONTRIBUTING.md](CONTRIBUTING.md).

GitHub Actions runs the same portable gate on every push and pull request.

## Design principles

- **Fail closed.** Missing required data, unknown mappings, invalid signs, and duplicate keys block
  publication.
- **Preserve evidence.** Every result carries a Build ID, fingerprint, hashes, counts, totals, and
  validation items.
- **Keep calculations headless.** Core modules accept typed data and return typed results; they do
  not read workbooks or display dialogs.
- **Use full precision.** Economic calculations use `Decimal`; rounding occurs only for display or
  an explicit business rule.
- **Keep Excel thin.** Excel holds values and presentation. Python remains authoritative.
- **Do not treat Legacy as truth.** Legacy output is diagnostic evidence. Acceptance comes from
  business rules, golden cases, invariants, and independent reconciliation.

## Documentation map

| Document | Purpose |
|---|---|
| [INSTALL.md](INSTALL.md) | Clean installation, verification, upgrades, and troubleshooting |
| [Quick Start](docs/GTM_QUICK_START.md) | Short daily workflow |
| [Excel Interface Guide](docs/GTM_EXCEL_INTERFACE_GUIDE.md) | Detailed operator manual |
| [Engine Specification](docs/GTM_ENGINE_SPECIFICATION.md) | Authoritative functional and technical contract |
| [Acceptance Test Specification](docs/GTM_ACCEPTANCE_TEST_SPECIFICATION.md) | Required economic and engineering evidence |
| [Decision Guide](docs/GTM_DECISION_GUIDE.md) | Plain-English business decisions |
| [Methodology Trace](docs/GTM_LOCAL_METHODOLOGY_TRACE.md) | Recovered legacy evidence and conflicts |
| [Golden Regression Report](verification/GTM_v2_Golden_Regression_Test_Report_v0.3.md) | Test-pack implementation and result |
| [JOURNAL.md](JOURNAL.md) | Detailed chronological audit trail |

## Data protection and licence

Never commit or attach production trades, prices, workbooks, credentials, or generated reports.
Use synthetic fixtures in tests and private channels for sensitive evidence. Read [SECURITY.md](SECURITY.md)
before sharing diagnostics.

This repository has no open-source licence. Treat its contents as private project material unless
the owner grants redistribution rights explicitly.
