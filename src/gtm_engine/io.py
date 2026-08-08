"""Normalized CSV/JSON bundle loading and atomic result publication."""

from __future__ import annotations

import csv
import json
import os
import shutil
import tempfile
from datetime import date, datetime
from decimal import Decimal
from enum import Enum
from hashlib import sha256
from pathlib import Path
from typing import Any

from pydantic import BaseModel, ValidationError

from .decimal_utils import decimal_text
from .models import (
    BookConfig,
    BuildConfig,
    BuildResult,
    BuildStatus,
    CumulativePnlRow,
    CurvePrice,
    EventLedgerRow,
    ExposureRow,
    FixingPrice,
    FixingRow,
    InitialExposure,
    InitialPnl,
    InputBundle,
    MarketCalendarDay,
    OperatingFlow,
    PnlRow,
    Trade,
    UnderlyingConfig,
    ValidationItem,
)


class BundleLoadError(ValueError):
    pass


def _file_hash(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _clean_csv_row(row: dict[str, str | None]) -> dict[str, Any]:
    cleaned: dict[str, Any] = {}
    for key, value in row.items():
        if key is None:
            continue
        normalized_key = key.strip()
        if value is None:
            cleaned[normalized_key] = None
            continue
        text = value.strip()
        cleaned[normalized_key] = None if text == "" else text
    return cleaned


def _read_models[ModelT: BaseModel](
    path: Path, model: type[ModelT], *, required: bool
) -> tuple[ModelT, ...]:
    if not path.exists():
        if required:
            raise BundleLoadError(f"Required input file is missing: {path.name}")
        return ()
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            if reader.fieldnames is None:
                return ()
            rows: list[ModelT] = []
            for line_number, row in enumerate(reader, start=2):
                cleaned = _clean_csv_row(row)
                try:
                    rows.append(model.model_validate(cleaned))
                except (ValidationError, ValueError) as exc:
                    source_row_id = cleaned.get("source_row_id")
                    source_context = f", source_row_id={source_row_id}" if source_row_id else ""
                    raise BundleLoadError(
                        f"Cannot load {path.name} at CSV row {line_number}{source_context}: {exc}"
                    ) from exc
            return tuple(rows)
    except BundleLoadError:
        raise
    except (OSError, UnicodeError, csv.Error) as exc:
        raise BundleLoadError(f"Cannot load {path.name}: {exc}") from exc


def load_bundle(directory: str | Path) -> InputBundle:
    root = Path(directory).expanduser().resolve()
    config_path = root / "bundle.json"
    if not config_path.exists():
        raise BundleLoadError("Required input file is missing: bundle.json")
    try:
        config = BuildConfig.model_validate(json.loads(config_path.read_text(encoding="utf-8")))
    except (OSError, json.JSONDecodeError, ValidationError) as exc:
        raise BundleLoadError(f"Cannot load bundle.json: {exc}") from exc

    files: tuple[tuple[str, type[BaseModel], bool], ...] = (
        ("books.csv", BookConfig, True),
        ("underlyings.csv", UnderlyingConfig, True),
        ("market_calendar.csv", MarketCalendarDay, True),
        ("initial_exposure.csv", InitialExposure, True),
        ("initial_pnl.csv", InitialPnl, True),
        ("trades.csv", Trade, False),
        ("curve_prices.csv", CurvePrice, False),
        ("fixing_prices.csv", FixingPrice, False),
        ("operating_flows.csv", OperatingFlow, False),
    )
    loaded: dict[str, tuple[BaseModel, ...]] = {}
    hashes = {"bundle.json": _file_hash(config_path)}
    for filename, model, required in files:
        path = root / filename
        loaded[filename] = _read_models(path, model, required=required)
        if path.exists():
            hashes[filename] = _file_hash(path)

    return InputBundle(
        config=config,
        books=loaded["books.csv"],  # type: ignore[arg-type]
        underlyings=loaded["underlyings.csv"],  # type: ignore[arg-type]
        calendar=loaded["market_calendar.csv"],  # type: ignore[arg-type]
        initial_exposure=loaded["initial_exposure.csv"],  # type: ignore[arg-type]
        initial_pnl=loaded["initial_pnl.csv"],  # type: ignore[arg-type]
        trades=loaded["trades.csv"],  # type: ignore[arg-type]
        curve_prices=loaded["curve_prices.csv"],  # type: ignore[arg-type]
        fixing_prices=loaded["fixing_prices.csv"],  # type: ignore[arg-type]
        operating_flows=loaded["operating_flows.csv"],  # type: ignore[arg-type]
        input_hashes=hashes,
    )


def _cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, Decimal):
        return decimal_text(value)
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, Enum):
        return str(value.value)
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    return str(value)


def _write_models(path: Path, rows: tuple[BaseModel, ...], model: type[BaseModel]) -> None:
    fields = list(model.model_fields)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, lineterminator="\n")
        writer.writeheader()
        for row in rows:
            values = row.model_dump(mode="python")
            writer.writerow({field: _cell(values.get(field)) for field in fields})


def write_bundle(bundle: InputBundle, directory: str | Path) -> Path:
    """Publish a normalized input bundle without exposing a partial directory."""

    destination = Path(directory).expanduser().resolve()
    if destination.exists():
        raise BundleLoadError(f"Bundle output already exists: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    staging = Path(tempfile.mkdtemp(prefix=f".{destination.name}-", dir=destination.parent))
    try:
        (staging / "bundle.json").write_text(
            bundle.config.model_dump_json(indent=2) + "\n", encoding="utf-8"
        )
        tables: tuple[tuple[str, tuple[BaseModel, ...], type[BaseModel]], ...] = (
            ("books.csv", bundle.books, BookConfig),
            ("underlyings.csv", bundle.underlyings, UnderlyingConfig),
            ("market_calendar.csv", bundle.calendar, MarketCalendarDay),
            ("initial_exposure.csv", bundle.initial_exposure, InitialExposure),
            ("initial_pnl.csv", bundle.initial_pnl, InitialPnl),
            ("trades.csv", bundle.trades, Trade),
            ("curve_prices.csv", bundle.curve_prices, CurvePrice),
            ("fixing_prices.csv", bundle.fixing_prices, FixingPrice),
            ("operating_flows.csv", bundle.operating_flows, OperatingFlow),
        )
        for filename, rows, model in tables:
            _write_models(staging / filename, rows, model)
        os.replace(staging, destination)
        destination.chmod(0o755)
    except Exception:
        shutil.rmtree(staging, ignore_errors=True)
        raise
    return destination


def _write_result_directory(root: Path, result: BuildResult) -> None:
    _write_models(root / "validation.csv", result.validation, ValidationItem)
    if result.manifest.status is not BuildStatus.FAILED:
        _write_models(root / "fixings.csv", result.fixings, FixingRow)
        _write_models(root / "exposure.csv", result.exposure, ExposureRow)
        _write_models(root / "pnl.csv", result.pnl, PnlRow)
        _write_models(root / "cumulative_pnl.csv", result.cumulative_pnl, CumulativePnlRow)
        _write_models(root / "event_ledger.csv", result.event_ledger, EventLedgerRow)
    (root / "build_manifest.json").write_text(
        result.manifest.model_dump_json(indent=2), encoding="utf-8"
    )


def publish_result(result: BuildResult, output_root: str | Path) -> tuple[BuildResult, Path]:
    root = Path(output_root).expanduser().resolve()
    root.mkdir(parents=True, exist_ok=True)
    if result.manifest.status is BuildStatus.FAILED:
        destination = root / "failed" / result.manifest.run_id
        destination.mkdir(parents=True, exist_ok=False)
        _write_result_directory(destination, result)
        return result, destination

    published_manifest = result.manifest.model_copy(update={"status": BuildStatus.PUBLISHED})
    published = result.model_copy(update={"manifest": published_manifest})
    runs = root / "runs"
    runs.mkdir(parents=True, exist_ok=True)
    staging = runs / f".staging-{result.manifest.run_id}"
    destination = runs / result.manifest.run_id
    staging.mkdir(parents=False, exist_ok=False)
    _write_result_directory(staging, published)
    os.replace(staging, destination)

    latest_temp = root / f".LATEST-{result.manifest.run_id}"
    latest_temp.write_text(str(destination) + "\n", encoding="utf-8")
    os.replace(latest_temp, root / "LATEST")
    return published, destination
