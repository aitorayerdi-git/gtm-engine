"""Command-line entry point for the headless GTM engine."""

from __future__ import annotations

import argparse
import json
import sys
from datetime import date
from pathlib import Path

from .excel import (
    ExcelAdapterError,
    build_excel_workbook,
    create_excel_template,
    load_setup_mapping,
    publish_excel_load_failure,
)
from .io import BundleLoadError, load_bundle, publish_result
from .legacy_import import (
    LegacyImportError,
    import_legacy_workbook,
    refresh_legacy_curve_table,
)
from .models import BookConfig, BuildStatus, UnderlyingConfig
from .pipeline import build


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="gtm-engine")
    subcommands = parser.add_subparsers(dest="command", required=True)
    build_command = subcommands.add_parser("build", help="Build a normalized input bundle")
    build_command.add_argument("--input", required=True, type=Path)
    build_command.add_argument("--output", required=True, type=Path)
    template_command = subcommands.add_parser(
        "excel-template", help="Create a macro-free Excel interface workbook"
    )
    template_command.add_argument("--output", required=True, type=Path)
    template_command.add_argument(
        "--mapping",
        type=Path,
        help="Optional reviewed SETUP mapping CSV used to seed BOOKS and UNDERLYINGS",
    )
    excel_build_command = subcommands.add_parser(
        "excel-build", help="Build a GTM Excel interface workbook"
    )
    excel_build_command.add_argument("--workbook", required=True, type=Path)
    excel_build_command.add_argument("--output", required=True, type=Path)
    legacy_command = subcommands.add_parser(
        "legacy-import",
        help="Convert a legacy GTM .xlsm snapshot into v0.3 input artifacts",
    )
    legacy_command.add_argument("--workbook", required=True, type=Path)
    legacy_command.add_argument("--output", required=True, type=Path)
    legacy_command.add_argument(
        "--historical-start",
        type=date.fromisoformat,
        help="Optional YYYY-MM-DD override for legacy PROCESS!C15",
    )
    legacy_command.add_argument(
        "--historical-end",
        type=date.fromisoformat,
        help="Optional YYYY-MM-DD override for legacy PROCESS!D15",
    )
    refresh_command = subcommands.add_parser(
        "excel-refresh-curves",
        help="Rebuild CURVE PRICES from cached provider sheets in an Excel snapshot",
    )
    refresh_command.add_argument("--workbook", required=True, type=Path)
    refresh_command.add_argument("--historical-end", type=date.fromisoformat)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    if args.command == "excel-refresh-curves":
        try:
            count, latest = refresh_legacy_curve_table(
                args.workbook, historical_end=args.historical_end
            )
        except (LegacyImportError, OSError, KeyError, ValueError) as exc:
            print(json.dumps({"status": "FAILED", "stage": "CurveRefresh", "error": str(exc)}))
            return 2
        print(
            json.dumps(
                {
                    "status": "REFRESHED",
                    "row_count": count,
                    "latest_market_date": latest.isoformat() if latest else None,
                },
                sort_keys=True,
            )
        )
        return 0
    if args.command == "legacy-import":
        try:
            imported = import_legacy_workbook(
                args.workbook,
                args.output,
                historical_start=args.historical_start,
                historical_end=args.historical_end,
            )
        except (LegacyImportError, BundleLoadError, OSError) as exc:
            print(json.dumps({"status": "FAILED", "stage": "LegacyImport", "error": str(exc)}))
            return 2
        print(
            json.dumps(
                {
                    "status": imported.report.status,
                    "output": str(imported.output_directory),
                    "normalized_bundle": str(imported.normalized_bundle),
                    "workbook": str(imported.excel_workbook),
                    "audit": str(imported.audit_json),
                    "issue_counts": imported.report.as_dict()["issue_counts"],
                    "row_counts": imported.report.extracted_row_counts,
                },
                sort_keys=True,
            )
        )
        return 0

    if args.command == "excel-template":
        try:
            books: tuple[BookConfig, ...] = ()
            underlyings: tuple[UnderlyingConfig, ...] = ()
            if args.mapping is not None:
                books, underlyings = load_setup_mapping(args.mapping)
            template = create_excel_template(
                args.output,
                books=books,
                underlyings=underlyings,
            )
        except ExcelAdapterError as exc:
            print(json.dumps({"status": "FAILED", "stage": "ExcelTemplate", "error": str(exc)}))
            return 2
        print(json.dumps({"status": "CREATED", "workbook": str(template)}, sort_keys=True))
        return 0

    if args.command == "excel-build":
        try:
            published, destination, result_workbook = build_excel_workbook(
                args.workbook, args.output
            )
        except ExcelAdapterError as exc:
            diagnostic = publish_excel_load_failure(args.workbook, args.output, str(exc))
            print(
                json.dumps(
                    {
                        "status": "FAILED",
                        "stage": "ExcelLoad",
                        "error": str(exc),
                        "workbook": str(diagnostic) if diagnostic else None,
                    },
                    sort_keys=True,
                )
            )
            return 2
        print(
            json.dumps(
                {
                    "status": published.manifest.status.value,
                    "build_id": published.manifest.build_id,
                    "run_id": published.manifest.run_id,
                    "output": str(destination),
                    "workbook": str(result_workbook),
                    "validation_counts": published.manifest.validation_counts,
                },
                sort_keys=True,
            )
        )
        return 0 if published.manifest.status is BuildStatus.PUBLISHED else 2

    if args.command != "build":
        return 2
    try:
        bundle = load_bundle(args.input)
    except BundleLoadError as exc:
        print(json.dumps({"status": "FAILED", "stage": "Load", "error": str(exc)}))
        return 2

    result = build(bundle)
    published, destination = publish_result(result, args.output)
    print(
        json.dumps(
            {
                "status": published.manifest.status.value,
                "build_id": published.manifest.build_id,
                "run_id": published.manifest.run_id,
                "output": str(destination),
                "validation_counts": published.manifest.validation_counts,
            },
            sort_keys=True,
        )
    )
    return 0 if published.manifest.status is BuildStatus.PUBLISHED else 2


if __name__ == "__main__":
    sys.exit(main())
