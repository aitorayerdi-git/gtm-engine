"""Deterministic calculation identity and manifest helpers."""

from __future__ import annotations

import json
import platform
from collections import Counter
from collections.abc import Iterable
from datetime import datetime
from hashlib import sha256
from typing import Any

import psutil

from .decimal_utils import ZERO, decimal_text
from .models import (
    BuildConfig,
    BuildManifest,
    BuildStatus,
    CumulativePnlRow,
    EventLedgerRow,
    ExposureRow,
    FixingRow,
    InputBundle,
    PnlRow,
    Severity,
    ValidationItem,
    jsonable,
)


def canonical_json(value: Any) -> str:
    return json.dumps(
        jsonable(value),
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )


def sha256_text(value: str) -> str:
    return sha256(value.encode("utf-8")).hexdigest()


def calculation_fingerprint(bundle: InputBundle) -> str:
    payload = bundle.model_dump(mode="json", exclude={"input_hashes"})
    for field in (
        "books",
        "underlyings",
        "calendar",
        "initial_exposure",
        "initial_pnl",
        "trades",
        "delivery_elections",
        "curve_prices",
        "fixing_prices",
        "fx_rates",
        "operating_flows",
    ):
        payload[field] = sorted(payload[field], key=canonical_json)
    return sha256_text(canonical_json(payload))


def build_id_from_fingerprint(fingerprint: str) -> str:
    return f"GTM3-{fingerprint[:20].upper()}"


def hash_rows(rows: Iterable[object]) -> str:
    return sha256_text(canonical_json(list(rows)))


def current_memory_bytes() -> int:
    return int(psutil.Process().memory_info().rss)


def validation_counts(items: Iterable[ValidationItem]) -> dict[str, int]:
    counts = Counter(item.severity.value for item in items)
    return {severity.value: counts.get(severity.value, 0) for severity in Severity}


def output_hashes(
    validations: tuple[ValidationItem, ...],
    fixings: tuple[FixingRow, ...],
    exposure: tuple[ExposureRow, ...],
    pnl: tuple[PnlRow, ...],
    cumulative: tuple[CumulativePnlRow, ...],
    event_ledger: tuple[EventLedgerRow, ...],
) -> dict[str, str]:
    return {
        "validation": hash_rows(validations),
        "fixings": hash_rows(fixings),
        "exposure": hash_rows(exposure),
        "pnl": hash_rows(pnl),
        "cumulative_pnl": hash_rows(cumulative),
        "event_ledger": hash_rows(event_ledger),
    }


def component_totals(
    fixings: tuple[FixingRow, ...],
    exposure: tuple[ExposureRow, ...],
    pnl: tuple[PnlRow, ...],
) -> dict[str, str]:
    closing_exposure: dict[tuple[object, ...], ExposureRow] = {}
    for row in exposure:
        key = (
            row.book,
            row.underlying,
            row.delivery_month,
            row.trade_source,
            row.scenario,
        )
        current = closing_exposure.get(key)
        if current is None or row.market_date > current.market_date:
            closing_exposure[key] = row
    return {
        "fixing_volume": decimal_text(sum((row.fixing_volume for row in fixings), ZERO)),
        "fixing_amount": decimal_text(sum((row.fixing_amount for row in fixings), ZERO)),
        "closing_exposure_volume": decimal_text(
            sum((row.exposure_volume for row in closing_exposure.values()), ZERO)
        ),
        "gross_delta_exposure_mtm": decimal_text(
            sum((row.gross_delta_exposure_mtm for row in pnl), ZERO)
        ),
        "trade_entry_adjustment": decimal_text(
            sum((row.trade_entry_adjustment for row in pnl), ZERO)
        ),
        "logistical_costs": decimal_text(sum((row.logistical_costs for row in pnl), ZERO)),
        "fees_and_optimizations": decimal_text(
            sum((row.fees_and_optimizations for row in pnl), ZERO)
        ),
        "replication": decimal_text(sum((row.replication for row in pnl), ZERO)),
        "total_pnl": decimal_text(sum((row.total_pnl for row in pnl), ZERO)),
    }


def make_manifest(
    *,
    run_id: str,
    build_id: str,
    fingerprint: str,
    status: BuildStatus,
    started_at: datetime,
    finished_at: datetime,
    config: BuildConfig,
    input_hashes: dict[str, str],
    input_row_counts: dict[str, int],
    validations: tuple[ValidationItem, ...],
    fixings: tuple[FixingRow, ...] = (),
    exposure: tuple[ExposureRow, ...] = (),
    pnl: tuple[PnlRow, ...] = (),
    cumulative: tuple[CumulativePnlRow, ...] = (),
    event_ledger: tuple[EventLedgerRow, ...] = (),
    peak_memory_bytes: int = 0,
    failure_stage: str | None = None,
) -> BuildManifest:
    elapsed = (finished_at - started_at).total_seconds()
    hashes = output_hashes(validations, fixings, exposure, pnl, cumulative, event_ledger)
    totals = component_totals(fixings, exposure, pnl)
    row_counts = {
        "validation": len(validations),
        "fixings": len(fixings),
        "exposure": len(exposure),
        "pnl": len(pnl),
        "cumulative_pnl": len(cumulative),
        "event_ledger": len(event_ledger),
        "unapplied_events": sum(row.applied_market_date is None for row in event_ledger),
    }
    recorded_hashes = dict(sorted(input_hashes.items()))
    recorded_hashes.setdefault("normalized_bundle", fingerprint)
    return BuildManifest(
        run_id=run_id,
        build_id=build_id,
        calculation_fingerprint=fingerprint,
        status=status,
        started_at=started_at,
        finished_at=finished_at,
        timezone=config.timezone,
        engine_version=config.engine_version,
        schema_version=config.schema_version,
        policy_version=config.policy_version,
        simulation_status=config.simulation_status,
        input_hashes=recorded_hashes,
        input_row_counts=dict(sorted(input_row_counts.items())),
        output_hashes=hashes,
        row_counts=row_counts,
        component_totals=totals,
        validation_counts=validation_counts(validations),
        elapsed_seconds=elapsed,
        peak_memory_bytes=peak_memory_bytes,
        runtime_platform=platform.platform(),
        python_version=platform.python_version(),
        machine_architecture=platform.machine(),
        failure_stage=failure_stage,
    )
