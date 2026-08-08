# Contributing to GTM

GTM is an auditable financial engine. A change is complete only when its business rule, code,
tests, documentation, and evidence agree.

## Development setup

Follow [INSTALL.md](INSTALL.md), then install the development dependencies:

```sh
.venv/bin/python -m pip install -e '.[dev]'
```

Create a branch from the current `main` branch. Keep each change focused; do not mix methodology
changes with unrelated cleanup.

## Quality gate

Run this gate before every commit:

```sh
.venv/bin/ruff format --check src tests scripts
.venv/bin/ruff check src tests scripts
.venv/bin/mypy src
.venv/bin/pytest --cov=gtm_engine --cov-report=term-missing
.venv/bin/python -m pip check
```

Coverage must remain at or above 85%. GitHub Actions runs the same portable checks.

## Business-rule changes

A methodology change must include:

1. The rule in plain English, with its effective policy version.
2. An update to `docs/GTM_ENGINE_SPECIFICATION.md` and any affected decision or user guide.
3. An independently calculated regression case with exact expected values.
4. Tests for validation, signs, dates, aggregation, and edge conditions that the rule affects.
5. A detailed entry in `JOURNAL.md` recording the decision, implementation, and verification.

Do not make Legacy output the acceptance oracle. Legacy workbooks provide evidence and migration
inputs; golden cases, business decisions, invariants, and independent reconciliation define
correct behavior.

## Code rules

- Keep the calculation core independent of Excel, file I/O, dialogs, and process state.
- Use typed models at module boundaries and `Decimal` for economic arithmetic.
- Fail closed when required inputs, mappings, dates, units, currencies, or prices are ambiguous.
- Keep functions small and explicit. Prefer clear dictionaries and loops to opaque abstractions.
- Preserve deterministic ordering, Build IDs, output hashes, and atomic publication.
- Do not add worksheet formulas or VBA as a second calculation engine.
- Update public documentation when a CLI command, workbook sheet, schema, or workflow changes.

## Tests

Use synthetic data for portable tests. A useful economic test states the complete independent
calculation in its assertions, including signs and dates.

The optional private-workbook inventory test may run locally when the legacy workbook is present.
It must skip in a clean clone. Never weaken portable coverage by relying on private files.

## Data protection

Read [docs/REPOSITORY_DATA_POLICY.md](docs/REPOSITORY_DATA_POLICY.md). Never stage production
workbooks, trades, prices, outputs, extracted workbook XML, credentials, or runtime snapshots.
Before committing, inspect the exact staged set:

```sh
git status --short
git diff --cached --stat
git diff --cached
```

## Commits and review

Use an imperative, specific subject such as `Add D2 report reconciliation`. Explain business
impact and tests in the commit body when the change is not obvious.

Reviewers should be able to answer four questions:

1. Which rule changed?
2. Which inputs and outputs can change?
3. Which independent test proves the new result?
4. Which audit record identifies the decision and verification?
