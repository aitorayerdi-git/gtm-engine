# Repository data policy

## Purpose

The Git repository stores reproducible source code and reviewable documentation. It does not act
as a production data store, workbook archive, or calculation-output archive.

## Tracked material

The repository may track:

- Python source and packaging metadata;
- synthetic tests and independently calculated golden cases;
- text specifications, decisions, mappings, manuals, and audit history;
- small reference documents supplied for methodology or test-pack review;
- launchers and developer tooling that contain no private data.

## Local-only material

Keep these files outside Git:

- legacy `.xlsm`, `.xlsb`, `.xls`, and generated `.xlsx` workbooks;
- `outputs/` build directories and published or failed reports;
- `analysis/` workbook extractions and third-party analysis dependencies;
- `runtime_tests/` workbook copies and runtime evidence;
- production trades, positions, prices, operating flows, and model outputs;
- temporary owner files such as `~$workbook.xlsx`;
- credentials, tokens, local environment files, caches, and logs.

`.gitignore` excludes these names and locations. Do not use `git add -f` to override the policy.

## Optional private-workbook test

`tests/test_legacy_import.py` contains one inventory regression that looks for
`Gas_Trading_Model 070826.xlsm` at the repository root. It skips when the file is absent. Local
maintainers may place an authorized copy there to run the additional test; Git ignores it.

The portable CI suite creates synthetic legacy workbooks in temporary directories and does not
need private data.

## Reproducing the synthetic validation model

The fixing, July-delivery, and exposure validation workbooks do not need to be committed. A clean
checkout can recreate them from versioned code and `docs/GTM_ACTIVE_SETUP_MAPPING_v0.3.csv` by
running `scripts/create_fixing_test.py` without `--source`. The generator creates synthetic
calendar, trade, price, FX, and opening-balance inputs, reloads the generated workbook through the
strict adapter, and requires a verified build before writing the report.

This reproducible synthetic path does not replace authorized private inputs for a production run.
It proves that the engine and reviewed test scenario can be reconstructed without local evidence.

## Reproducing a local report

1. Obtain the authorized source workbook through the approved private channel.
2. Place it in a local path ignored by Git.
3. Run `legacy-import` into `outputs/`.
4. Review the import audit and supply any approved missing-data policy through a separate input
   artifact.
5. Run `excel-build` and retain the run manifest with the distributed result.

Do not commit the source, imported input, policy-normalized input, result workbook, or run folder.
Record only non-sensitive rule changes and verification summaries in documentation.

## Before every push

Run:

```sh
git status --short
git diff --cached --stat
git diff --cached
```

Check filenames as well as file contents. If sensitive material entered Git history, stop before
pushing and remove it from the history through the approved incident process; deleting it in a
later commit is insufficient.
