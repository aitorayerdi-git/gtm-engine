"""Fixing schedules, price requirements, and deterministic fixing output."""

from __future__ import annotations

from collections import defaultdict
from datetime import date, timedelta
from decimal import Decimal
from hashlib import sha256

from .calendar import (
    CalendarIndex,
    add_months,
    delivery_month_slices,
    eligible_delivery_days,
    month_end,
    normal_fixing_date,
)
from .canonicalize import Registry, normalize_text
from .decimal_utils import ZERO, is_material
from .models import (
    FixingEvent,
    FixingMethod,
    FixingPrice,
    FixingRow,
    InputBundle,
    PriceDateBasis,
    Severity,
    TradeEvent,
    TradeSource,
    ValidationItem,
)
from .validation import included_trade


def _event_id(*parts: object) -> str:
    payload = "|".join(str(part) for part in parts)
    return sha256(payload.encode("utf-8")).hexdigest()[:24]


def _scenario(source: TradeSource, value: str | None) -> str | None:
    return value if source is TradeSource.SIMULATION else None


def _price_lookup_date(
    fixing_date: date,
    delivery_day: date,
    basis: PriceDateBasis,
) -> date:
    return delivery_day if basis is PriceDateBasis.DELIVERY_DAY else fixing_date


def _price_basis(method: FixingMethod) -> PriceDateBasis:
    """Daily products price by delivery day; Month Ahead prices each fixing tranche."""
    return (
        PriceDateBasis.FIXING_DATE
        if method is FixingMethod.MONTH_AHEAD
        else PriceDateBasis.DELIVERY_DAY
    )


def _fixing_event(
    *,
    calendar: CalendarIndex,
    book: str,
    source_underlying: str,
    canonical_underlying: str,
    pricing_underlying: str,
    basis: PriceDateBasis,
    delivery_month: date,
    delivery_day: date,
    fixing_date: date,
    fixing_volume: Decimal,
    trade_source: TradeSource,
    scenario: str | None,
    source_row_id: str,
    sequence: int,
) -> FixingEvent:
    applied = calendar.first_output_market_on_or_after(fixing_date)
    return FixingEvent(
        event_id=_event_id(
            "FIXING",
            source_row_id,
            sequence,
            fixing_date,
            delivery_day,
            decimal_text_for_id(fixing_volume),
        ),
        fixing_date=fixing_date,
        applied_market_date=applied,
        price_lookup_date=_price_lookup_date(fixing_date, delivery_day, basis),
        book=book,
        source_underlying=source_underlying,
        underlying=canonical_underlying,
        pricing_underlying=pricing_underlying,
        delivery_month=delivery_month,
        delivery_day=delivery_day,
        fixing_volume=fixing_volume,
        trade_source=trade_source,
        scenario=_scenario(trade_source, scenario),
        source_row_id=source_row_id,
    )


def decimal_text_for_id(value: Decimal) -> str:
    return format(value.normalize(), "f") if value else "0"


def _equal_allocations(total: Decimal, count: int) -> tuple[Decimal, ...]:
    """Split a total and put the Decimal division remainder in the final slice."""
    if count <= 0:
        raise ValueError("allocation count must be positive")
    per_item = total / Decimal(count)
    return (per_item,) * (count - 1) + (total - per_item * Decimal(count - 1),)


def build_schedules(
    bundle: InputBundle,
    registry: Registry,
    calendar: CalendarIndex,
    build_id: str,
) -> tuple[tuple[FixingEvent, ...], tuple[TradeEvent, ...], tuple[ValidationItem, ...]]:
    fixing_events: list[FixingEvent] = []
    trade_events: list[TradeEvent] = []
    issues: list[ValidationItem] = []

    for position in sorted(
        bundle.initial_exposure,
        key=lambda row: (
            row.book.casefold(),
            row.underlying.casefold(),
            row.delivery_month,
            row.source_row_id,
        ),
    ):
        if not is_material(position.exposure_volume, bundle.config.materiality):
            continue
        book_row = registry.book(position.book)
        underlying_row = registry.underlying(position.underlying)
        if book_row is None or underlying_row is None:
            continue
        first_delivery = position.delivery_month
        last_delivery = month_end(position.delivery_month)
        sequence = 0

        if underlying_row.fixing_method is FixingMethod.MONTH_AHEAD:
            prior_month = add_months(position.delivery_month, -1)
            fixing_dates = tuple(
                value
                for value in calendar.market_dates_between(
                    prior_month,
                    position.delivery_month - timedelta(days=1),
                )
                if value > bundle.config.initial_market_date
            )
            delivery_days = eligible_delivery_days(
                first_delivery, last_delivery, underlying_row.fixing_method, calendar
            )
            if not fixing_dates or not delivery_days:
                issues.append(
                    ValidationItem(
                        build_id=build_id,
                        stage="Preflight",
                        severity=Severity.ERROR,
                        code="INITIAL_NO_FIXING_OPPORTUNITY",
                        message=(
                            "Material Initial Exposure has no fixing opportunity after the cut-off."
                        ),
                        table="initial_exposure",
                        source_row_id=position.source_row_id,
                        actual=position.delivery_month.isoformat(),
                        expected="at least one fixing date strictly after Initial Market Date",
                        remediation="Correct the initial state or extend the configured calendar.",
                    )
                )
                continue
            allocations = iter(
                _equal_allocations(
                    -position.exposure_volume,
                    len(fixing_dates) * len(delivery_days),
                )
            )
            for fixing_date in fixing_dates:
                for delivery_day in delivery_days:
                    sequence += 1
                    fixing_events.append(
                        _fixing_event(
                            calendar=calendar,
                            book=book_row.book,
                            source_underlying=underlying_row.source_underlying,
                            canonical_underlying=underlying_row.canonical_underlying,
                            pricing_underlying=underlying_row.fixing_price_underlying or "",
                            basis=_price_basis(underlying_row.fixing_method),
                            delivery_month=position.delivery_month,
                            delivery_day=delivery_day,
                            fixing_date=fixing_date,
                            fixing_volume=next(allocations),
                            trade_source=TradeSource.INITIAL,
                            scenario=None,
                            source_row_id=position.source_row_id,
                            sequence=sequence,
                        )
                    )
            continue

        eligible = tuple(
            delivery_day
            for delivery_day in eligible_delivery_days(
                first_delivery, last_delivery, underlying_row.fixing_method, calendar
            )
            if normal_fixing_date(delivery_day, underlying_row.fixing_method, calendar)
            > bundle.config.initial_market_date
        )
        if not eligible:
            issues.append(
                ValidationItem(
                    build_id=build_id,
                    stage="Preflight",
                    severity=Severity.ERROR,
                    code="INITIAL_NO_FIXING_OPPORTUNITY",
                    message=(
                        "Material Initial Exposure has no fixing opportunity after the cut-off."
                    ),
                    table="initial_exposure",
                    source_row_id=position.source_row_id,
                    actual=position.delivery_month.isoformat(),
                    expected="at least one fixing date strictly after Initial Market Date",
                    remediation="Correct the initial state or extend the configured calendar.",
                )
            )
            continue
        delivery_allocations = _equal_allocations(-position.exposure_volume, len(eligible))
        for delivery_day, fixing_volume in zip(eligible, delivery_allocations, strict=True):
            sequence += 1
            fixing_date = normal_fixing_date(delivery_day, underlying_row.fixing_method, calendar)
            fixing_events.append(
                _fixing_event(
                    calendar=calendar,
                    book=book_row.book,
                    source_underlying=underlying_row.source_underlying,
                    canonical_underlying=underlying_row.canonical_underlying,
                    pricing_underlying=underlying_row.fixing_price_underlying or "",
                    basis=_price_basis(underlying_row.fixing_method),
                    delivery_month=position.delivery_month,
                    delivery_day=delivery_day,
                    fixing_date=fixing_date,
                    fixing_volume=fixing_volume,
                    trade_source=TradeSource.INITIAL,
                    scenario=None,
                    source_row_id=position.source_row_id,
                    sequence=sequence,
                )
            )

    for trade in sorted(bundle.trades, key=lambda row: (row.trade_date, row.source_row_id)):
        if not included_trade(trade, bundle) or not is_material(
            trade.daily_qty, bundle.config.materiality
        ):
            continue
        book_row = registry.book(trade.book)
        underlying_row = registry.underlying(trade.underlying)
        if book_row is None or underlying_row is None or trade.end_date < trade.start_date:
            continue
        sequence = 0
        for delivery_month, overlap_first, overlap_last in delivery_month_slices(
            trade.start_date, trade.end_date
        ):
            delivery_days = eligible_delivery_days(
                overlap_first, overlap_last, underlying_row.fixing_method, calendar
            )
            if not delivery_days:
                continue
            signed_month_volume = trade.signed_daily_qty * Decimal(len(delivery_days))
            applied_trade_date = calendar.first_output_market_on_or_after(trade.trade_date)
            trade_events.append(
                TradeEvent(
                    event_id=_event_id(
                        "TRADE",
                        trade.source_row_id,
                        delivery_month,
                        decimal_text_for_id(signed_month_volume),
                    ),
                    economic_date=trade.trade_date,
                    applied_market_date=applied_trade_date,
                    book=book_row.book,
                    source_underlying=underlying_row.source_underlying,
                    underlying=underlying_row.canonical_underlying,
                    delivery_month=delivery_month,
                    signed_volume=signed_month_volume,
                    trade_source=trade.trade_source,
                    scenario=_scenario(trade.trade_source, trade.scenario),
                    source_row_id=trade.source_row_id,
                    execution_price=trade.execution_price,
                )
            )

            if underlying_row.fixing_method is FixingMethod.MONTH_AHEAD:
                prior_month = add_months(delivery_month, -1)
                fixing_dates = tuple(
                    value
                    for value in calendar.market_dates_between(
                        prior_month, delivery_month - timedelta(days=1)
                    )
                    if value >= trade.trade_date
                )
                if not fixing_dates:
                    continue
                allocations = iter(
                    _equal_allocations(
                        -signed_month_volume,
                        len(fixing_dates) * len(delivery_days),
                    )
                )
                for fixing_date in fixing_dates:
                    for delivery_day in delivery_days:
                        sequence += 1
                        fixing_events.append(
                            _fixing_event(
                                calendar=calendar,
                                book=book_row.book,
                                source_underlying=underlying_row.source_underlying,
                                canonical_underlying=underlying_row.canonical_underlying,
                                pricing_underlying=underlying_row.fixing_price_underlying or "",
                                basis=_price_basis(underlying_row.fixing_method),
                                delivery_month=delivery_month,
                                delivery_day=delivery_day,
                                fixing_date=fixing_date,
                                fixing_volume=next(allocations),
                                trade_source=trade.trade_source,
                                scenario=trade.scenario,
                                source_row_id=trade.source_row_id,
                                sequence=sequence,
                            )
                        )
                continue

            for delivery_day in delivery_days:
                fixing_date = normal_fixing_date(
                    delivery_day, underlying_row.fixing_method, calendar
                )
                if fixing_date < trade.trade_date:
                    # Preflight rejects the full source row before output is usable.
                    continue
                sequence += 1
                fixing_events.append(
                    _fixing_event(
                        calendar=calendar,
                        book=book_row.book,
                        source_underlying=underlying_row.source_underlying,
                        canonical_underlying=underlying_row.canonical_underlying,
                        pricing_underlying=underlying_row.fixing_price_underlying or "",
                        basis=_price_basis(underlying_row.fixing_method),
                        delivery_month=delivery_month,
                        delivery_day=delivery_day,
                        fixing_date=fixing_date,
                        fixing_volume=-trade.signed_daily_qty,
                        trade_source=trade.trade_source,
                        scenario=trade.scenario,
                        source_row_id=trade.source_row_id,
                        sequence=sequence,
                    )
                )

    fixing_events.sort(
        key=lambda row: (
            row.fixing_date,
            row.book,
            row.underlying,
            row.delivery_month,
            row.delivery_day,
            row.trade_source,
            row.scenario or "",
            row.event_id,
        )
    )
    trade_events.sort(
        key=lambda row: (
            row.economic_date,
            row.book,
            row.underlying,
            row.delivery_month,
            row.trade_source,
            row.scenario or "",
            row.event_id,
        )
    )
    return tuple(fixing_events), tuple(trade_events), tuple(issues)


def _fixing_price_index(prices: tuple[FixingPrice, ...]) -> dict[tuple[date, str], FixingPrice]:
    return {(row.price_lookup_date, normalize_text(row.underlying)): row for row in prices}


def validate_required_fixing_prices(
    events: tuple[FixingEvent, ...],
    bundle: InputBundle,
    build_id: str,
) -> tuple[ValidationItem, ...]:
    index = _fixing_price_index(bundle.fixing_prices)
    missing: dict[tuple[date, str], FixingEvent] = {}
    incompatible: dict[tuple[date, str, str], tuple[FixingEvent, str, str]] = {}
    profiles = {
        normalize_text(row.source_underlying): row for row in bundle.underlyings if row.active
    }
    for event in events:
        if event.applied_market_date is None:
            continue
        key = (event.price_lookup_date, normalize_text(event.pricing_underlying))
        price = index.get(key)
        if price is None:
            missing.setdefault(key, event)
            continue
        profile = profiles[normalize_text(event.source_underlying)]
        if normalize_text(price.currency) != normalize_text(profile.currency):
            incompatible.setdefault(
                (*key, "currency"),
                (event, price.currency, profile.currency),
            )
        if normalize_text(price.unit) != normalize_text(profile.unit):
            incompatible.setdefault(
                (*key, "unit"),
                (event, price.unit, profile.unit),
            )
    missing_items = tuple(
        ValidationItem(
            build_id=build_id,
            stage="Preflight",
            severity=Severity.ERROR,
            code="MISSING_FIXING_PRICE",
            message=(
                "Required fixing price is missing: "
                "Market Date="
                f"{event.applied_market_date.isoformat() if event.applied_market_date else ''}, "
                f"Underlying={event.pricing_underlying}, "
                f"Delivery Month={event.delivery_month.isoformat()}, "
                f"Fixing Date={event.fixing_date.isoformat()}, "
                f"Price Lookup Date={event.price_lookup_date.isoformat()}."
            ),
            table="fixing_prices",
            source_row_id=event.source_row_id,
            economic_key=(
                f"{event.price_lookup_date.isoformat()}|{event.pricing_underlying}|"
                f"{event.delivery_month.isoformat()}|{event.delivery_day.isoformat()}"
            ),
            expected="one finite fixing price for the required lookup key",
            remediation="Supply the missing price and rebuild.",
        )
        for _, event in sorted(missing.items(), key=lambda item: item[0])
    )
    contract_items: list[ValidationItem] = []
    for contract_key, (event, actual, expected) in sorted(incompatible.items()):
        dimension = contract_key[2]
        contract_items.append(
            ValidationItem(
                build_id=build_id,
                stage="Preflight",
                severity=Severity.ERROR,
                code=f"FIXING_PRICE_{dimension.upper()}_MISMATCH",
                message=(
                    f"Required fixing price has incompatible {dimension}: "
                    f"Price Lookup Date={event.price_lookup_date.isoformat()}, "
                    f"Underlying={event.pricing_underlying}."
                ),
                table="fixing_prices",
                source_row_id=event.source_row_id,
                economic_key=f"{event.price_lookup_date.isoformat()}|{event.pricing_underlying}",
                actual=actual,
                expected=expected,
                remediation="Normalize the price to the configured contract before rebuilding.",
            )
        )
    return missing_items + tuple(contract_items)


def price_fixings(
    events: tuple[FixingEvent, ...],
    bundle: InputBundle,
    build_id: str,
) -> tuple[FixingRow, ...]:
    index = _fixing_price_index(bundle.fixing_prices)
    aggregate: dict[tuple[object, ...], Decimal] = defaultdict(lambda: ZERO)
    metadata: dict[tuple[object, ...], FixingEvent] = {}
    for event in events:
        if event.applied_market_date is None:
            continue
        aggregate_key: tuple[object, ...] = (
            event.fixing_date,
            event.applied_market_date,
            event.price_lookup_date,
            event.book,
            event.source_underlying,
            event.underlying,
            event.pricing_underlying,
            event.delivery_month,
            event.delivery_day,
            event.trade_source,
            event.scenario,
        )
        aggregate[aggregate_key] += event.fixing_volume
        metadata.setdefault(aggregate_key, event)

    rows: list[FixingRow] = []
    for aggregate_key in sorted(aggregate, key=lambda value: tuple(str(part) for part in value)):
        volume = aggregate[aggregate_key]
        if not is_material(volume, bundle.config.materiality):
            continue
        event = metadata[aggregate_key]
        price = index[(event.price_lookup_date, normalize_text(event.pricing_underlying))]
        applied_market_date = event.applied_market_date
        if applied_market_date is None:
            continue
        rows.append(
            FixingRow(
                fixing_date=event.fixing_date,
                applied_market_date=applied_market_date,
                price_lookup_date=event.price_lookup_date,
                book=event.book,
                source_underlying=event.source_underlying,
                underlying=event.underlying,
                pricing_underlying=event.pricing_underlying,
                delivery_month=event.delivery_month,
                delivery_day=event.delivery_day,
                fixing_volume=volume,
                fixing_price=price.fixing_price,
                fixing_amount=volume * price.fixing_price,
                currency=price.currency,
                trade_source=event.trade_source,
                scenario=event.scenario,
                simulation_status=bundle.config.simulation_status,
                build_id=build_id,
            )
        )
    return tuple(rows)
