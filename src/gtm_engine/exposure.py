"""Event roll-forward and curve valuation for daily closing exposure."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import date
from decimal import Decimal

from .calendar import CalendarIndex, add_months
from .canonicalize import Registry, normalize_text
from .decimal_utils import ZERO, is_material
from .models import (
    CurvePrice,
    ExposureRow,
    FixingEvent,
    InputBundle,
    Severity,
    TradeEvent,
    TradeSource,
    ValidationItem,
)

ExposureKey = tuple[str, str, date, TradeSource, str | None]


@dataclass(frozen=True)
class UnpricedExposure:
    market_date: date
    previous_market_date: date
    key: ExposureKey
    exposure_volume: Decimal
    curve_underlying: str
    curve_delivery_month: date
    is_explicit_closure: bool


def _curve_profiles(bundle: InputBundle) -> dict[str, tuple[str, bool, str, str]]:
    profiles: dict[str, tuple[str, bool, str, str]] = {}
    for row in bundle.underlyings:
        if not row.active:
            continue
        key = normalize_text(row.canonical_underlying)
        profile = (
            row.curve_underlying or row.canonical_underlying,
            row.current_month_uses_next_curve,
            row.currency,
            row.unit,
        )
        existing = profiles.get(key)
        if existing is not None and existing != profile:
            raise ValueError(
                f"Conflicting curve profiles for canonical underlying {row.canonical_underlying}"
            )
        profiles[key] = profile
    return profiles


def curve_contract_month(
    market_date: date,
    delivery_month: date,
    current_month_uses_next_curve: bool,
) -> date:
    if current_month_uses_next_curve and (
        market_date.year == delivery_month.year and market_date.month == delivery_month.month
    ):
        return add_months(delivery_month, 1)
    return delivery_month


def opening_volumes(bundle: InputBundle, registry: Registry) -> dict[ExposureKey, Decimal]:
    result: dict[ExposureKey, Decimal] = defaultdict(lambda: ZERO)
    for row in bundle.initial_exposure:
        book = registry.book(row.book)
        underlying = registry.underlying(row.underlying)
        if book is None or underlying is None:
            continue
        key: ExposureKey = (
            book.book,
            underlying.canonical_underlying,
            row.delivery_month,
            TradeSource.INITIAL,
            None,
        )
        result[key] += row.exposure_volume
    return dict(result)


def build_unpriced_exposure(
    bundle: InputBundle,
    registry: Registry,
    calendar: CalendarIndex,
    fixing_events: tuple[FixingEvent, ...],
    trade_events: tuple[TradeEvent, ...],
) -> tuple[UnpricedExposure, ...]:
    profiles = _curve_profiles(bundle)
    state = opening_volumes(bundle, registry)
    changes: dict[date, dict[ExposureKey, Decimal]] = defaultdict(lambda: defaultdict(lambda: ZERO))

    for trade_event in trade_events:
        if trade_event.applied_market_date is None:
            continue
        key: ExposureKey = (
            trade_event.book,
            trade_event.underlying,
            trade_event.delivery_month,
            trade_event.trade_source,
            trade_event.scenario,
        )
        changes[trade_event.applied_market_date][key] += trade_event.signed_volume
    for fixing_event in fixing_events:
        if fixing_event.applied_market_date is None:
            continue
        key = (
            fixing_event.book,
            fixing_event.underlying,
            fixing_event.delivery_month,
            fixing_event.trade_source,
            fixing_event.scenario,
        )
        changes[fixing_event.applied_market_date][key] += fixing_event.fixing_volume

    output: list[UnpricedExposure] = []
    for market_date in calendar.output_market_dates:
        previous_state = dict(state)
        for key, volume_change in changes.get(market_date, {}).items():
            state[key] = state.get(key, ZERO) + volume_change

        previous_market_date = calendar.previous_output_market_day(market_date)
        all_keys = sorted(
            set(previous_state) | set(state),
            key=lambda value: (value[0], value[1], value[2], value[3].value, value[4] or ""),
        )
        for key in all_keys:
            prior_volume = previous_state.get(key, ZERO)
            current_volume = state.get(key, ZERO)
            current_material = is_material(current_volume, bundle.config.materiality)
            closure = is_material(prior_volume, bundle.config.materiality) and not current_material
            if not current_material and not closure:
                continue
            canonical_underlying = key[1]
            curve_underlying, use_next, _, _ = profiles[normalize_text(canonical_underlying)]
            output.append(
                UnpricedExposure(
                    market_date=market_date,
                    previous_market_date=previous_market_date,
                    key=key,
                    exposure_volume=current_volume if current_material else ZERO,
                    curve_underlying=curve_underlying,
                    curve_delivery_month=curve_contract_month(market_date, key[2], use_next),
                    is_explicit_closure=closure,
                )
            )
    return tuple(output)


def _curve_index(
    prices: tuple[CurvePrice, ...],
) -> dict[tuple[date, str, date], CurvePrice]:
    return {
        (row.market_date, normalize_text(row.underlying), row.delivery_month): row for row in prices
    }


def required_curve_keys(
    bundle: InputBundle,
    registry: Registry,
    unpriced: tuple[UnpricedExposure, ...],
) -> dict[tuple[date, str, date], tuple[str, str, str]]:
    requirements: dict[tuple[date, str, date], tuple[str, str, str]] = {}
    profiles = _curve_profiles(bundle)
    for key, volume in opening_volumes(bundle, registry).items():
        if not is_material(volume, bundle.config.materiality):
            continue
        curve_underlying, use_next, currency, unit = profiles[normalize_text(key[1])]
        contract_month = curve_contract_month(bundle.config.initial_market_date, key[2], use_next)
        requirement = (
            bundle.config.initial_market_date,
            normalize_text(curve_underlying),
            contract_month,
        )
        requirements[requirement] = (
            f"opening BOOK={key[0]}, Underlying={key[1]}, Delivery Month={key[2]}",
            currency,
            unit,
        )
    for row in unpriced:
        if row.is_explicit_closure:
            continue
        requirement = (
            row.market_date,
            normalize_text(row.curve_underlying),
            row.curve_delivery_month,
        )
        _, _, currency, unit = profiles[normalize_text(row.key[1])]
        requirements[requirement] = (
            f"BOOK={row.key[0]}, Underlying={row.key[1]}, Delivery Month={row.key[2]}",
            currency,
            unit,
        )
    return requirements


def validate_required_curve_prices(
    bundle: InputBundle,
    registry: Registry,
    unpriced: tuple[UnpricedExposure, ...],
    build_id: str,
) -> tuple[ValidationItem, ...]:
    index = _curve_index(bundle.curve_prices)
    requirements = required_curve_keys(bundle, registry, unpriced)
    missing = [item for item in requirements.items() if item[0] not in index]
    missing_items = tuple(
        ValidationItem(
            build_id=build_id,
            stage="Preflight",
            severity=Severity.ERROR,
            code="MISSING_CURVE_PRICE",
            message=(
                "Required curve price is missing: "
                f"Market Date={key[0].isoformat()}, Underlying={key[1]}, "
                f"Delivery Month={key[2].isoformat()} ({context})."
            ),
            table="curve_prices",
            economic_key=f"{key[0].isoformat()}|{key[1]}|{key[2].isoformat()}",
            expected="one finite curve price for the required valuation key",
            remediation="Supply the missing price and rebuild.",
        )
        for key, (context, _, _) in sorted(missing, key=lambda item: item[0])
    )
    incompatible: list[ValidationItem] = []
    for key, (_, expected_currency, expected_unit) in sorted(requirements.items()):
        price = index.get(key)
        if price is None:
            continue
        for dimension, actual, expected in (
            ("currency", price.currency, expected_currency),
            ("unit", price.unit, expected_unit),
        ):
            if normalize_text(actual) == normalize_text(expected):
                continue
            incompatible.append(
                ValidationItem(
                    build_id=build_id,
                    stage="Preflight",
                    severity=Severity.ERROR,
                    code=f"CURVE_PRICE_{dimension.upper()}_MISMATCH",
                    message=(
                        f"Required curve price has incompatible {dimension}: "
                        f"Market Date={key[0].isoformat()}, Underlying={key[1]}, "
                        f"Delivery Month={key[2].isoformat()}."
                    ),
                    table="curve_prices",
                    economic_key=f"{key[0].isoformat()}|{key[1]}|{key[2].isoformat()}",
                    actual=actual,
                    expected=expected,
                    remediation=(
                        "Normalize the price to the configured contract before rebuilding."
                    ),
                )
            )
    return missing_items + tuple(incompatible)


def price_exposure(
    unpriced: tuple[UnpricedExposure, ...],
    bundle: InputBundle,
    build_id: str,
) -> tuple[ExposureRow, ...]:
    index = _curve_index(bundle.curve_prices)
    output: list[ExposureRow] = []
    for row in unpriced:
        book, underlying, delivery_month, trade_source, scenario = row.key
        if row.is_explicit_closure:
            price = None
            mtm = ZERO
        else:
            price_row = index[
                (
                    row.market_date,
                    normalize_text(row.curve_underlying),
                    row.curve_delivery_month,
                )
            ]
            price = price_row.curve_price
            mtm = row.exposure_volume * price
        output.append(
            ExposureRow(
                market_date=row.market_date,
                previous_market_date=row.previous_market_date,
                book=book,
                underlying=underlying,
                delivery_month=delivery_month,
                curve_delivery_month=row.curve_delivery_month,
                exposure_volume=row.exposure_volume,
                curve_price=price,
                exposure_mtm=mtm,
                trade_source=trade_source,
                scenario=scenario,
                simulation_status=bundle.config.simulation_status,
                is_explicit_closure=row.is_explicit_closure,
                build_id=build_id,
            )
        )
    return tuple(output)


def opening_mtms(
    bundle: InputBundle,
    registry: Registry,
) -> dict[ExposureKey, Decimal]:
    profiles = _curve_profiles(bundle)
    index = _curve_index(bundle.curve_prices)
    result: dict[ExposureKey, Decimal] = {}
    for key, volume in opening_volumes(bundle, registry).items():
        if not is_material(volume, bundle.config.materiality):
            result[key] = ZERO
            continue
        curve_underlying, use_next, _, _ = profiles[normalize_text(key[1])]
        contract_month = curve_contract_month(bundle.config.initial_market_date, key[2], use_next)
        price = index[
            (
                bundle.config.initial_market_date,
                normalize_text(curve_underlying),
                contract_month,
            )
        ].curve_price
        result[key] = volume * price
    return result
