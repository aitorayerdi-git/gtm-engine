# GTM v0.3 Installation Guide

## 1. Requirements

Install these before cloning the repository:

- Python 3.13, the tested runtime target. The package metadata permits later versions, but they
  require a separate compatibility run;
- Git;
- internet access, or access to an internal Python package mirror, for the first dependency
  installation;
- macOS, Linux, or Windows for the Python engine.

Microsoft Excel is optional. You need Excel, or another compatible `.xlsx` editor, only to edit
and view the Excel interface. The Python engine does not start Excel and does not run macros.

On Windows, installation also supplies the IANA timezone database used for the configured
`Europe/Madrid` build timezone. This is declared as a Windows-only package dependency; no manual
timezone-package installation is required.

Check Python before continuing:

```sh
python3.13 --version
```

The command must report Python 3.13.x. If `python3.13` is unavailable but `python3` reports 3.13.x,
use `python3` in the commands below.

## 2. Clone the repository

```sh
git clone https://github.com/vasilybelokurov/gtm-engine.git
cd gtm-engine
```

Run all later commands from the repository root—the directory containing `pyproject.toml`.

The source repository is public, so cloning does not require a GitHub account. Authentication is
needed only for operations such as pushing changes. Never put a personal access token in a
command, script, or file committed to Git.

## 3. Create the local Python environment

The project uses a private environment named `.venv`. Do not install its dependencies into the
system Python.

### macOS or Linux

```sh
python3.13 -m venv .venv
.venv/bin/python -m pip install --upgrade pip
```

### Windows PowerShell

```powershell
py -3.13 -m venv .venv
.venv\Scripts\python.exe -m pip install --upgrade pip
```

Activation is optional. The documentation uses the environment's full executable path so that
each command runs with the correct Python.

## 4. Install the project

Choose one installation type.

### Operator installation

This installs the engine and Excel adapter:

```sh
.venv/bin/python -m pip install -e .
```

Windows PowerShell:

```powershell
.venv\Scripts\python.exe -m pip install -e .
```

### Developer installation

This also installs pytest, coverage, Ruff, mypy, Hypothesis, and type stubs:

```sh
.venv/bin/python -m pip install -e '.[dev]'
```

Windows PowerShell:

```powershell
.venv\Scripts\python.exe -m pip install -e ".[dev]"
```

`pyproject.toml` is the dependency authority. Do not maintain a separate hand-written list of
packages.

## 5. Verify an operator installation

### macOS or Linux

```sh
.venv/bin/gtm-engine --help
.venv/bin/python -m pip check
.venv/bin/gtm-engine excel-template \
  --output outputs/install_check/GTM_install_check.xlsx \
  --mapping docs/GTM_ACTIVE_SETUP_MAPPING_v0.3.csv
```

### Windows PowerShell

```powershell
.venv\Scripts\gtm-engine.exe --help
.venv\Scripts\python.exe -m pip check
.venv\Scripts\gtm-engine.exe excel-template `
  --output outputs\install_check\GTM_install_check.xlsx `
  --mapping docs\GTM_ACTIVE_SETUP_MAPPING_v0.3.csv
```

The checks pass when:

- `gtm-engine --help` lists `build`, `excel-template`, `excel-build`, and `legacy-import`;
- `pip check` reports `No broken requirements found`;
- the template command reports `"status": "CREATED"`;
- `outputs/install_check/GTM_install_check.xlsx` exists and opens as a macro-free workbook.

The install-check workbook contains blank dates and business data, so it is not expected to pass
an economic build.

## 6. Verify a developer installation

Run the complete quality gate from the repository root:

```sh
.venv/bin/ruff format --check src tests scripts
.venv/bin/ruff check src tests scripts
.venv/bin/mypy src
.venv/bin/pytest --cov=gtm_engine --cov-report=term-missing -q
.venv/bin/python -m pip check
```

The installation is ready for development when:

- formatting and lint checks pass;
- mypy reports no issues;
- all tests pass;
- total test coverage meets the 85% project gate;
- `pip check` reports no broken requirements.

In a clean clone, the portable suite passes and the optional real-workbook inventory test skips.
That test runs only when `Gas_Trading_Model 070826.xlsm` is present at the repository root.

## 7. Prepare the Excel interface

The repository does not commit Excel workbooks. Create the interface locally at:

```text
outputs/gtm_excel_v0_3/GTM_Excel_Interface_v0.3.xlsx
```

You can recreate it from the reviewed mapping at any time:

```sh
.venv/bin/gtm-engine excel-template \
  --output outputs/gtm_excel_v0_3/GTM_Excel_Interface_v0.3.xlsx \
  --mapping docs/GTM_ACTIVE_SETUP_MAPPING_v0.3.csv
```

This command creates the workbook structure and seeds 13 BOOK and 18 Underlying mappings. It does
not invent a production calendar, opening state, trades, or prices.

Private `.xlsm` sources, generated `.xlsx` workbooks, `outputs/`, `analysis/`, and runtime-test
snapshots are excluded by `.gitignore`. Keep them local. Read
[the repository data policy](docs/REPOSITORY_DATA_POLICY.md) before adding source material.

### Optional: create the input files from a legacy workbook

If the legacy `.xlsm` is available, convert it without opening Excel or enabling macros:

```sh
.venv/bin/gtm-engine legacy-import \
  --workbook "/path/to/Gas_Trading_Model 070826.xlsm" \
  --output outputs/legacy_import_070826_v0.3
```

The output directory must be new. The command writes a macro-free input workbook, an equivalent
normalized bundle, and two audit files. Read `legacy_import_audit.json` before building; a
successful import does not waive missing-price, FX, date, or mapping validation.

## 8. Prepare the macOS launcher

The launcher is optional and is specific to macOS:

```sh
chmod +x scripts/GTM_Build.command
```

Double-clicking the launcher builds the generated default workbook. To build another working
copy, drag that `.xlsx` file onto the launcher.

If macOS blocks the launcher, use the Terminal command instead:

```sh
.venv/bin/gtm-engine excel-build \
  --workbook path/to/your_input.xlsx \
  --output outputs/gtm_excel_runs
```

Windows and Linux users should use the equivalent command-line invocation. No Windows launcher is
included at present.

## 9. First use

After installation:

1. Read [the Quick Start](docs/GTM_QUICK_START.md).
2. Use [the detailed manual](docs/GTM_EXCEL_INTERFACE_GUIDE.md) when preparing a new dataset or
   importing legacy values.
3. Fill a working `.xlsx` input workbook, or create one with `legacy-import`.
4. Save and close it.
5. Run `excel-build`.
6. Open `outputs/gtm_excel_runs/GTM_LATEST.xlsx` only after a `PUBLISHED` result.

## 10. Upgrade an existing checkout

Save or commit your work before upgrading. Then pull the repository and reinstall its declared
dependencies:

```sh
git pull --ff-only
.venv/bin/python -m pip install -e '.[dev]'
.venv/bin/python -m pip check
.venv/bin/pytest -q
```

Operators who did not install the developer extras should replace `-e '.[dev]'` with `-e .` and
may omit pytest.

Do not overwrite a working input workbook during an upgrade. Generate a new template and migrate
the inputs when the schema version changes.

## 11. Troubleshooting

### Python version error

Symptom:

```text
Package requires a different Python
```

Correction: create `.venv` with Python 3.13, then reinstall the project.

### `gtm-engine` not found

Use the executable inside `.venv`:

```sh
.venv/bin/gtm-engine --help
```

On Windows use `.venv\Scripts\gtm-engine.exe`.

### Missing Python package

Reinstall from the repository root:

```sh
.venv/bin/python -m pip install -e '.[dev]'
```

### `Permission denied` for `GTM_Build.command`

Run:

```sh
chmod +x scripts/GTM_Build.command
```

Or use `gtm-engine excel-build` directly.

### Paths contain spaces

Put paths in quotes:

```sh
.venv/bin/gtm-engine excel-build \
  --workbook "/path/with spaces/input.xlsx" \
  --output "/path/with spaces/results"
```

### Workbook build fails after installation

Installation and economic validation are separate. Open the retained `GTM_Failed.xlsx` and read
`VALIDATION`. Missing prices, invalid dates, unknown mappings, and incomplete calendars are input
problems rather than installation failures.

### Dependency download fails

Confirm network access and proxy or package-mirror settings. In a controlled corporate
environment, configure pip to use the approved internal package index rather than installing
packages from unapproved sources.

## 12. Installation support information

When reporting an installation problem, include:

```sh
python3.13 --version
.venv/bin/python -m pip --version
.venv/bin/python -m pip check
.venv/bin/gtm-engine --help
```

Also include the operating system, processor architecture, complete error message, and the Git
commit being installed. Do not attach production trades or prices to a public issue.
