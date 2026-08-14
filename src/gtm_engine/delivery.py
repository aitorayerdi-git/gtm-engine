"""Validated daily split between standard and physical-delivery product legs."""

from __future__ import annotations

from collections import defaultdict
from datetime import date, timedelta
from decimal import Decimal

from .canonicalize import Registry, normalize_text
from .decimal_utils import ZERO, is_material
from .models import (
    InputBundle,
    Severity,
    Side,
    Trade,
    TradeSource,
    UnderlyingConfig,
    ValidationItem,
)
from .validation import included_trade

DELIVERY_VARIANTS = {
    "TTFDA HEREN": "TTFDA Delivery",
    "PVB HEREN": "PVB Heren Delivery",
}
DeliveryKey = tuple[str, str, Side, TradeSource, str | None, date]


def _days(first: date, last: date) -> tuple[date, ...]:
    return tuple(first + timedelta(days=offset) for offset in range((last - first).days + 1))


def _key(
    book: str,
    underlying: str,
    side: Side,
    trade_source: TradeSource,
    scenario: str | None,
) -> tuple[str, str, Side, TradeSource, str | None]:
    return (normalize_text(book), normalize_text(underlying), side, trade_source, scenario)


def split_delivery_trades(
    bundle: InputBundle, build_id: str
) -> tuple[InputBundle, tuple[ValidationItem, ...]]:
    registry = Registry.from_bundle(bundle)
    issues: list[ValidationItem] = []
    elections: dict[DeliveryKey, Decimal] = {}

    for row in bundle.delivery_elections:
        profile = registry.underlying(row.underlying)
        base_name = normalize_text(row.underlying)
        if base_name not in DELIVERY_VARIANTS:
            issues.append(
                ValidationItem(
                    build_id=build_id,
                    stage="Preflight",
                    severity=Severity.ERROR,
                    code="DELIVERY_UNDERLYING_NOT_ELIGIBLE",
                    message=f"Delivery election is not supported for {row.underlying}.",
                    table="delivery_elections",
                    source_row_id=row.source_row_id,
                )
            )
        if registry.book(row.book) is None or profile is None:
            issues.append(
                ValidationItem(
                    build_id=build_id,
                    stage="Preflight",
                    severity=Severity.ERROR,
                    code="DELIVERY_MAPPING_MISSING",
                    message="Delivery election BOOK or Underlying is unknown or inactive.",
                    table="delivery_elections",
                    source_row_id=row.source_row_id,
                )
            )
        elif normalize_text(row.unit) != normalize_text(profile.unit):
            issues.append(
                ValidationItem(
                    build_id=build_id,
                    stage="Preflight",
                    severity=Severity.ERROR,
                    code="DELIVERY_UNIT_MISMATCH",
                    message=f"Delivery election unit {row.unit} does not match {profile.unit}.",
                    table="delivery_elections",
                    source_row_id=row.source_row_id,
                )
            )
        if row.end_date < row.start_date or row.start_date.replace(day=1) != row.end_date.replace(
            day=1
        ):
            issues.append(
                ValidationItem(
                    build_id=build_id,
                    stage="Preflight",
                    severity=Severity.ERROR,
                    code="DELIVERY_DATE_RANGE_INVALID",
                    message="Delivery election range must be ordered and remain within one month.",
                    table="delivery_elections",
                    source_row_id=row.source_row_id,
                )
            )
            continue
        if row.decision_date >= row.start_date.replace(day=1):
            issues.append(
                ValidationItem(
                    build_id=build_id,
                    stage="Preflight",
                    severity=Severity.ERROR,
                    code="DELIVERY_DECISION_LATE",
                    message="Delivery election must be decided before the delivery month starts.",
                    table="delivery_elections",
                    source_row_id=row.source_row_id,
                )
            )
        if row.delivery_daily_qty < ZERO:
            issues.append(
                ValidationItem(
                    build_id=build_id,
                    stage="Preflight",
                    severity=Severity.ERROR,
                    code="DELIVERY_QTY_NEGATIVE",
                    message="Delivery Daily Qty must be a non-negative magnitude.",
                    table="delivery_elections",
                    source_row_id=row.source_row_id,
                )
            )
        if row.trade_source is TradeSource.SIMULATION and not row.scenario:
            issues.append(
                ValidationItem(
                    build_id=build_id,
                    stage="Preflight",
                    severity=Severity.ERROR,
                    code="DELIVERY_SCENARIO_MISSING",
                    message="Simulation delivery election requires a Scenario.",
                    table="delivery_elections",
                    source_row_id=row.source_row_id,
                )
            )
        if row.trade_source is TradeSource.ACTUAL and row.scenario:
            issues.append(
                ValidationItem(
                    build_id=build_id,
                    stage="Preflight",
                    severity=Severity.ERROR,
                    code="DELIVERY_ACTUAL_SCENARIO",
                    message="Actual delivery election cannot carry a Scenario.",
                    table="delivery_elections",
                    source_row_id=row.source_row_id,
                )
            )
        for delivery_day in _days(row.start_date, row.end_date):
            daily_key = (
                *_key(row.book, row.underlying, row.side, row.trade_source, row.scenario),
                delivery_day,
            )
            if daily_key in elections:
                issues.append(
                    ValidationItem(
                        build_id=build_id,
                        stage="Preflight",
                        severity=Severity.ERROR,
                        code="DELIVERY_RANGE_OVERLAP",
                        message="Delivery election ranges overlap for the same economic key.",
                        table="delivery_elections",
                        source_row_id=row.source_row_id,
                        economic_key="|".join(str(value) for value in daily_key),
                    )
                )
            else:
                elections[daily_key] = row.delivery_daily_qty

    eligible = tuple(
        row
        for row in bundle.trades
        if included_trade(row, bundle, registry)
        and normalize_text(row.underlying) in DELIVERY_VARIANTS
    )
    candidates: dict[DeliveryKey, list[Trade]] = defaultdict(list)
    for trade in eligible:
        for delivery_day in _days(trade.start_date, trade.end_date):
            candidates[
                (
                    *_key(
                        trade.book, trade.underlying, trade.side, trade.trade_source, trade.scenario
                    ),
                    delivery_day,
                )
            ].append(trade)

    allocations: dict[tuple[str, date], Decimal] = {}
    for election_key, elected in elections.items():
        rows = sorted(candidates.get(election_key, []), key=lambda row: row.source_row_id)
        available = sum((row.daily_qty for row in rows), ZERO)
        if elected > available:
            issues.append(
                ValidationItem(
                    build_id=build_id,
                    stage="Preflight",
                    severity=Severity.ERROR,
                    code="DELIVERY_QTY_EXCEEDS_POSITION",
                    message="Elected delivery quantity exceeds same-side available daily volume.",
                    table="delivery_elections",
                    economic_key="|".join(str(value) for value in election_key),
                    actual=str(elected),
                    expected=f"<= {available}",
                )
            )
            continue
        remaining = elected
        for index, trade in enumerate(rows):
            allocation = (
                remaining if index == len(rows) - 1 else elected * trade.daily_qty / available
            )
            allocations[(trade.source_row_id, election_key[-1])] = allocation
            remaining -= allocation

    if issues:
        return bundle, tuple(issues)

    expanded: list[Trade] = []
    eligible_ids = {row.source_row_id for row in eligible}
    expanded.extend(row for row in bundle.trades if row.source_row_id not in eligible_ids)
    for trade in eligible:
        delivery_name = DELIVERY_VARIANTS[normalize_text(trade.underlying)]
        for delivery_day in _days(trade.start_date, trade.end_date):
            delivery_qty = allocations.get((trade.source_row_id, delivery_day), ZERO)
            base_qty = trade.daily_qty - delivery_qty
            for suffix, underlying, quantity in (
                ("BASE", trade.underlying, base_qty),
                ("DELIVERY", delivery_name, delivery_qty),
            ):
                if not is_material(quantity, bundle.config.materiality):
                    continue
                expanded.append(
                    trade.model_copy(
                        update={
                            "source_row_id": f"{trade.source_row_id}|{delivery_day}|{suffix}",
                            "underlying": underlying,
                            "start_date": delivery_day,
                            "end_date": delivery_day,
                            "daily_qty": quantity,
                        }
                    )
                )
    expanded.sort(key=lambda row: (row.trade_date, row.source_row_id))
    derived_profiles: list[UnderlyingConfig] = []
    for base_name, delivery_name in DELIVERY_VARIANTS.items():
        profile = registry.underlying(base_name)
        if profile is None:
            continue
        derived_profiles.append(
            profile.model_copy(
                update={
                    "source_underlying": delivery_name,
                    "fixing_price_underlying": profile.fixing_price_underlying,
                    "include_fixing_in_pnl": True,
                }
            )
        )
    return bundle.model_copy(
        update={
            "trades": tuple(expanded),
            "underlyings": (*bundle.underlyings, *derived_profiles),
        }
    ), ()
