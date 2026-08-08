# Security and private-data handling

This project processes commercially sensitive trading and market data. The source repository is
public, but all operational data, workbooks, generated results, and local analysis artifacts must
remain private. The tracked tests use synthetic fixtures.

## Do not commit

- production or client workbooks;
- trades, positions, prices, fees, costs, or P&L outputs;
- generated `GTM_LATEST.xlsx`, failed workbooks, CSV results, or manifests;
- extracted workbook XML, VBA binaries, screenshots, or runtime snapshots;
- credentials, tokens, `.env` files, or signed URLs;
- logs or bug reports that contain source rows or economic keys from production data.

The repository's `.gitignore` blocks the normal private and generated locations. It is a guardrail,
not a substitute for reviewing `git diff --cached` before a commit.

## Report a vulnerability

Do not open a public issue containing exploit details or sensitive data. Use a private GitHub
security advisory for this repository or contact the repository owner through an agreed private
channel. Include the affected version, a minimal synthetic reproduction, impact, and proposed
mitigation. Remove all production identifiers and values.

## Operational guidance

- Run the engine with the least filesystem access needed for its input and output directories.
- Keep Python and declared dependencies patched; run `pip check` after upgrades.
- Review unexpected dependency or lock-file changes before installation.
- Do not enable macros in legacy source workbooks for import. The importer reads cached values.
- Keep input and result workbooks separate. A failed run must never replace the last published
  result.
- Verify Build ID, input hash, status, and validation evidence before distributing a report.

## Supported versions

Security fixes target the current `main` branch and the latest v0.3 release line. Older local
snapshots receive no guaranteed updates.
