"""One-way fail-closed GTM build pipeline."""

from __future__ import annotations

from datetime import datetime
from uuid import uuid4
from zoneinfo import ZoneInfo

from .calendar import CalendarIndex
from .canonicalize import Registry
from .exposure import (
    build_unpriced_exposure,
    opening_mtms,
    price_exposure,
    validate_required_curve_prices,
)
from .fixings import (
    build_schedules,
    price_fixings,
    validate_required_fixing_prices,
)
from .invariants import validate_fixing_conservation, validate_output_invariants
from .manifest import (
    build_id_from_fingerprint,
    calculation_fingerprint,
    current_memory_bytes,
    make_manifest,
)
from .models import (
    BuildResult,
    BuildStatus,
    EventLedgerRow,
    EventType,
    FixingEvent,
    InputBundle,
    Severity,
    TradeEvent,
    TradeSource,
    ValidationItem,
)
from .pnl import build_pnl
from .validation import has_errors, validate_bundle


def _input_row_counts(bundle: InputBundle) -> dict[str, int]:
    return {
        "books": len(bundle.books),
        "underlyings": len(bundle.underlyings),
        "calendar": len(bundle.calendar),
        "initial_exposure": len(bundle.initial_exposure),
        "initial_pnl": len(bundle.initial_pnl),
        "trades": len(bundle.trades),
        "curve_prices": len(bundle.curve_prices),
        "fixing_prices": len(bundle.fixing_prices),
        "operating_flows": len(bundle.operating_flows),
    }


def _ordered_validations(
    validations: tuple[ValidationItem, ...],
) -> tuple[ValidationItem, ...]:
    return tuple(
        sorted(
            validations,
            key=lambda row: (
                row.stage,
                row.severity.value,
                row.code,
                row.table or "",
                row.source_row_id or "",
                row.economic_key or "",
                row.message,
            ),
        )
    )


def _event_ledger(
    bundle: InputBundle,
    registry: Registry,
    fixing_events: tuple[FixingEvent, ...],
    trade_events: tuple[TradeEvent, ...],
) -> tuple[EventLedgerRow, ...]:
    rows: list[EventLedgerRow] = []
    for position in bundle.initial_exposure:
        book = registry.book(position.book)
        underlying = registry.underlying(position.underlying)
        if book is None or underlying is None:
            continue
        rows.append(
            EventLedgerRow(
                event_id=f"INITIAL-{position.source_row_id}",
                event_type=EventType.INITIAL_OPEN,
                economic_date=bundle.config.initial_market_date,
                applied_market_date=bundle.config.initial_market_date,
                book=book.book,
                underlying=underlying.canonical_underlying,
                delivery_month=position.delivery_month,
                signed_volume_change=position.exposure_volume,
                trade_source=TradeSource.INITIAL,
                scenario=None,
                source_row_id=position.source_row_id,
            )
        )
    for trade_event in trade_events:
        rows.append(
            EventLedgerRow(
                event_id=trade_event.event_id,
                event_type=EventType.TRADE,
                economic_date=trade_event.economic_date,
                applied_market_date=trade_event.applied_market_date,
                book=trade_event.book,
                underlying=trade_event.underlying,
                delivery_month=trade_event.delivery_month,
                signed_volume_change=trade_event.signed_volume,
                trade_source=trade_event.trade_source,
                scenario=trade_event.scenario,
                source_row_id=trade_event.source_row_id,
            )
        )
    for fixing_event in fixing_events:
        rows.append(
            EventLedgerRow(
                event_id=fixing_event.event_id,
                event_type=EventType.FIXING,
                economic_date=fixing_event.fixing_date,
                applied_market_date=fixing_event.applied_market_date,
                book=fixing_event.book,
                underlying=fixing_event.underlying,
                delivery_month=fixing_event.delivery_month,
                signed_volume_change=fixing_event.fixing_volume,
                trade_source=fixing_event.trade_source,
                scenario=fixing_event.scenario,
                source_row_id=fixing_event.source_row_id,
            )
        )
    rows.sort(
        key=lambda row: (
            row.economic_date,
            row.event_type.value,
            row.book,
            row.underlying,
            row.delivery_month,
            row.event_id,
        )
    )
    return tuple(rows)


def _failed_result(
    *,
    bundle: InputBundle,
    run_id: str,
    build_id: str,
    fingerprint: str,
    started_at: datetime,
    validations: tuple[ValidationItem, ...],
    peak_memory: int,
    stage: str,
) -> BuildResult:
    validations = _ordered_validations(validations)
    finished_at = datetime.now(ZoneInfo(bundle.config.timezone))
    manifest = make_manifest(
        run_id=run_id,
        build_id=build_id,
        fingerprint=fingerprint,
        status=BuildStatus.FAILED,
        started_at=started_at,
        finished_at=finished_at,
        config=bundle.config,
        input_hashes=bundle.input_hashes,
        input_row_counts=_input_row_counts(bundle),
        validations=validations,
        peak_memory_bytes=peak_memory,
        failure_stage=stage,
    )
    return BuildResult(manifest=manifest, validation=validations)


def build(bundle: InputBundle) -> BuildResult:
    """Run the complete in-memory calculation. No files or UI are touched."""
    zone = ZoneInfo(bundle.config.timezone)
    started_at = datetime.now(zone)
    run_id = str(uuid4())
    fingerprint = calculation_fingerprint(bundle)
    build_id = build_id_from_fingerprint(fingerprint)
    peak_memory = current_memory_bytes()
    validations: tuple[ValidationItem, ...] = ()

    try:
        registry = Registry.from_bundle(bundle)
        calendar = CalendarIndex(bundle.calendar, bundle.config)
        validations = validate_bundle(bundle, build_id)
        peak_memory = max(peak_memory, current_memory_bytes())
        if has_errors(validations):
            return _failed_result(
                bundle=bundle,
                run_id=run_id,
                build_id=build_id,
                fingerprint=fingerprint,
                started_at=started_at,
                validations=validations,
                peak_memory=peak_memory,
                stage="Preflight",
            )

        fixing_events, trade_events, schedule_issues = build_schedules(
            bundle, registry, calendar, build_id
        )
        validations += schedule_issues
        validations += validate_fixing_conservation(fixing_events, trade_events, build_id)
        validations += validate_required_fixing_prices(fixing_events, bundle, build_id)
        if has_errors(validations):
            return _failed_result(
                bundle=bundle,
                run_id=run_id,
                build_id=build_id,
                fingerprint=fingerprint,
                started_at=started_at,
                validations=validations,
                peak_memory=max(peak_memory, current_memory_bytes()),
                stage="Fixings",
            )

        unpriced_exposure = build_unpriced_exposure(
            bundle, registry, calendar, fixing_events, trade_events
        )
        validations += validate_required_curve_prices(bundle, registry, unpriced_exposure, build_id)
        if has_errors(validations):
            return _failed_result(
                bundle=bundle,
                run_id=run_id,
                build_id=build_id,
                fingerprint=fingerprint,
                started_at=started_at,
                validations=validations,
                peak_memory=max(peak_memory, current_memory_bytes()),
                stage="Exposure",
            )

        fixings = price_fixings(fixing_events, bundle, build_id)
        exposure = price_exposure(unpriced_exposure, bundle, build_id)
        initial_mtm = opening_mtms(bundle, registry)
        pnl, cumulative = build_pnl(
            bundle,
            registry,
            calendar,
            exposure,
            fixings,
            fixing_events,
            trade_events,
            initial_mtm,
            build_id,
        )
        validations += validate_output_invariants(fixings, exposure, pnl, build_id)
        if has_errors(validations):
            return _failed_result(
                bundle=bundle,
                run_id=run_id,
                build_id=build_id,
                fingerprint=fingerprint,
                started_at=started_at,
                validations=validations,
                peak_memory=max(peak_memory, current_memory_bytes()),
                stage="Verify",
            )

        validations = _ordered_validations(validations)
        ledger = _event_ledger(bundle, registry, fixing_events, trade_events)
        finished_at = datetime.now(zone)
        peak_memory = max(peak_memory, current_memory_bytes())
        manifest = make_manifest(
            run_id=run_id,
            build_id=build_id,
            fingerprint=fingerprint,
            status=BuildStatus.VERIFIED,
            started_at=started_at,
            finished_at=finished_at,
            config=bundle.config,
            input_hashes=bundle.input_hashes,
            input_row_counts=_input_row_counts(bundle),
            validations=validations,
            fixings=fixings,
            exposure=exposure,
            pnl=pnl,
            cumulative=cumulative,
            event_ledger=ledger,
            peak_memory_bytes=peak_memory,
        )
        return BuildResult(
            manifest=manifest,
            validation=validations,
            fixings=fixings,
            exposure=exposure,
            pnl=pnl,
            cumulative_pnl=cumulative,
            event_ledger=ledger,
        )
    except Exception as exc:
        unexpected = ValidationItem(
            build_id=build_id,
            stage="Pipeline",
            severity=Severity.ERROR,
            code="UNEXPECTED_EXCEPTION",
            message=f"{type(exc).__name__}: {exc}",
            remediation="Retain the failed evidence and correct the engine or input contract.",
        )
        return _failed_result(
            bundle=bundle,
            run_id=run_id,
            build_id=build_id,
            fingerprint=fingerprint,
            started_at=started_at,
            validations=(*validations, unexpected),
            peak_memory=max(peak_memory, current_memory_bytes()),
            stage="Pipeline",
        )
