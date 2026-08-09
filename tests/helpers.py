from __future__ import annotations

from collections.abc import Callable
from datetime import UTC, date, datetime, timedelta
from decimal import Decimal

from gtm_engine.calendar import add_months, delivery_month_slices
from gtm_engine.models import (
    BookConfig,
    BuildConfig,
    CurvePrice,
    FixingMethod,
    FixingPrice,
    FxRate,
    InitialExposure,
    InitialPnl,
    InputBundle,
    MarketCalendarDay,
    OperatingFlow,
    PriceDateBasis,
    SimulationStatus,
    Trade,
    UnderlyingConfig,
)

D = Decimal
DEFAULT_INITIAL_PNL = D("1000")
DEFAULT_FIXING_PRICE = D("20")
DEFAULT_CURVE_PRICE = D("25")


def calendar_rows(
    first: date,
    last: date,
    holidays: set[date] | None = None,
) -> tuple[MarketCalendarDay, ...]:
    holidays = holidays or set()
    rows: list[MarketCalendarDay] = []
    cursor = first
    while cursor <= last:
        rows.append(
            MarketCalendarDay(
                date=cursor,
                is_market_day=cursor.weekday() < 5 and cursor not in holidays,
            )
        )
        cursor += timedelta(days=1)
    return tuple(rows)


def underlying(
    name: str = "GAS",
    method: FixingMethod = FixingMethod.DAY_AHEAD,
    *,
    canonical: str | None = None,
    curve_underlying: str | None = None,
    fixing_price_underlying: str | None = None,
    price_basis: PriceDateBasis = PriceDateBasis.FIXING_DATE,
    current_month_uses_next_curve: bool = False,
    unit: str = "MWh",
    currency: str = "EUR",
    include_fixing_in_pnl: bool = True,
) -> UnderlyingConfig:
    return UnderlyingConfig(
        source_underlying=name,
        canonical_underlying=canonical or name,
        fixing_method=method,
        unit=unit,
        currency=currency,
        curve_underlying=curve_underlying,
        fixing_price_underlying=fixing_price_underlying,
        fixing_price_basis=price_basis,
        current_month_uses_next_curve=current_month_uses_next_curve,
        include_fixing_in_pnl=include_fixing_in_pnl,
    )


def base_bundle(
    *,
    initial_market_date: date = date(2026, 1, 2),
    historical_end_date: date = date(2026, 1, 12),
    underlyings: tuple[UnderlyingConfig, ...] | None = None,
    initial_exposure: tuple[InitialExposure, ...] = (),
    trades: tuple[Trade, ...] = (),
    operating_flows: tuple[OperatingFlow, ...] = (),
    simulation_status: SimulationStatus = SimulationStatus.OFF,
    holidays: set[date] | None = None,
    initial_pnl_amount: Decimal = DEFAULT_INITIAL_PNL,
) -> InputBundle:
    configured = underlyings or (underlying(),)
    config = BuildConfig(
        initial_market_date=initial_market_date,
        historical_start_date=initial_market_date,
        historical_end_date=historical_end_date,
        simulation_status=simulation_status,
    )
    return InputBundle(
        config=config,
        books=(BookConfig(book="BOOK1"),),
        underlyings=configured,
        calendar=calendar_rows(
            initial_market_date - timedelta(days=45),
            historical_end_date + timedelta(days=70),
            holidays,
        ),
        initial_exposure=initial_exposure,
        initial_pnl=(
            InitialPnl(
                initial_market_date=initial_market_date,
                book="BOOK1",
                amount=initial_pnl_amount,
                source_row_id="IPNL-1",
            ),
        ),
        trades=trades,
        operating_flows=operating_flows,
    )


def with_prices(
    bundle: InputBundle,
    *,
    fixing_value: Decimal | Callable[[date, str], Decimal] = DEFAULT_FIXING_PRICE,
    curve_value: Decimal | Callable[[date, str, date], Decimal] = DEFAULT_CURVE_PRICE,
) -> InputBundle:
    fixing_profiles = {
        row.fixing_price_underlying or row.source_underlying: (row.currency, row.unit)
        for row in bundle.underlyings
    }
    fixing_prices: list[FixingPrice] = []
    for calendar_day in bundle.calendar:
        for name, (currency, unit) in sorted(fixing_profiles.items()):
            value = (
                fixing_value(calendar_day.date, name) if callable(fixing_value) else fixing_value
            )
            fixing_prices.append(
                FixingPrice(
                    price_lookup_date=calendar_day.date,
                    underlying=name,
                    fixing_price=value,
                    currency=currency,
                    unit=unit,
                    source_id=f"FIX-{calendar_day.date}-{name}",
                    source_as_of=datetime(2026, 1, 1, tzinfo=UTC),
                )
            )

    delivery_months = {row.delivery_month for row in bundle.initial_exposure}
    for trade in bundle.trades:
        delivery_months.update(
            month for month, _, _ in delivery_month_slices(trade.start_date, trade.end_date)
        )
    if not delivery_months:
        delivery_months.add(bundle.config.initial_market_date.replace(day=1))
    delivery_months |= {add_months(month, 1) for month in tuple(delivery_months)}
    curve_profiles = {
        row.curve_underlying or row.canonical_underlying: (row.currency, row.unit)
        for row in bundle.underlyings
    }
    market_dates = {
        bundle.config.initial_market_date,
        *(row.date for row in bundle.calendar if row.is_market_day),
    }
    curve_prices: list[CurvePrice] = []
    for market_date in sorted(market_dates):
        for name, (currency, unit) in sorted(curve_profiles.items()):
            for delivery_month in sorted(delivery_months):
                value = (
                    curve_value(market_date, name, delivery_month)
                    if callable(curve_value)
                    else curve_value
                )
                curve_prices.append(
                    CurvePrice(
                        market_date=market_date,
                        underlying=name,
                        delivery_month=delivery_month,
                        curve_price=value,
                        currency=currency,
                        unit=unit,
                        source_id=f"CURVE-{market_date}-{name}-{delivery_month}",
                        source_as_of=datetime(2026, 1, 1, tzinfo=UTC),
                    )
                )
    currencies = {row.currency for row in bundle.underlyings if row.currency != "EUR"}
    fx_rates = tuple(
        FxRate(
            rate_date=calendar_day.date,
            currency=currency,
            currency_per_eur=D("1.2"),
            source_id=f"FX-{calendar_day.date}-{currency}",
        )
        for calendar_day in bundle.calendar
        if calendar_day.is_market_day
        for currency in sorted(currencies)
    )
    return bundle.model_copy(
        update={
            "fixing_prices": tuple(fixing_prices),
            "curve_prices": tuple(curve_prices),
            "fx_rates": fx_rates,
        }
    )
