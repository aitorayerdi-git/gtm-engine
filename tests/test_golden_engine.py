from datetime import date
from decimal import Decimal

import pytest

from gtm_engine.decimal_utils import VOLUME_TOLERANCE
from gtm_engine.models import (
    DeliveryElection,
    FixingMethod,
    InitialExposure,
    OperatingFlow,
    PriceDateBasis,
    Side,
    Trade,
    TradeSource,
)
from gtm_engine.pipeline import build

from .helpers import base_bundle, underlying, with_prices

D = Decimal


@pytest.mark.golden
@pytest.mark.parametrize("name", ["Brent Dated", "HH"])
def test_brent_and_hh_roll_regression_uses_next_curve_month(name: str) -> None:
    configured = underlying(
        name,
        FixingMethod.BRENT_HH,
        price_basis=PriceDateBasis.DELIVERY_DAY,
        current_month_uses_next_curve=True,
        unit="bbl" if name == "Brent Dated" else "MMBtu",
        currency="USD",
    )
    initial = InitialExposure(
        initial_market_date=date(2026, 7, 28),
        book="BOOK1",
        underlying=name,
        delivery_month=date(2026, 7, 1),
        exposure_volume=D("100"),
        source_row_id=f"INIT-{name}",
    )
    bundle = base_bundle(
        initial_market_date=date(2026, 7, 28),
        historical_end_date=date(2026, 7, 30),
        underlyings=(configured,),
        initial_exposure=(initial,),
    )

    def curve_value(market_date: date, _underlying: str, delivery_month: date) -> Decimal:
        if market_date == date(2026, 7, 28) and delivery_month == date(2026, 8, 1):
            return D("70")
        if market_date == date(2026, 7, 29) and delivery_month == date(2026, 8, 1):
            return D("72")
        return D("999")

    result = build(with_prices(bundle, curve_value=curve_value))
    assert result.manifest.status.value == "VERIFIED"
    july_29 = next(row for row in result.exposure if row.market_date == date(2026, 7, 29))
    assert july_29.curve_delivery_month == date(2026, 8, 1)
    assert july_29.curve_price == D("72")
    assert july_29.exposure_volume == D("50")
    assert july_29.exposure_mtm == D("3600")
    assert july_29.currency == "USD"
    first_pnl = next(row for row in result.pnl if row.market_date == date(2026, 7, 29))
    assert first_pnl.gross_delta_exposure_mtm == D("-3400") / D("1.2")


@pytest.mark.golden
def test_usd_fixing_uses_fixing_date_fx_and_prior_available_rate() -> None:
    configured = underlying(
        "HH",
        FixingMethod.BRENT_HH,
        price_basis=PriceDateBasis.DELIVERY_DAY,
        unit="MMBtu",
        currency="USD",
    )
    trade = Trade(
        source_row_id="T-HH-FX",
        trade_date=date(2026, 1, 2),
        book="BOOK1",
        underlying="HH",
        side=Side.BUY,
        start_date=date(2026, 1, 5),
        end_date=date(2026, 1, 5),
        daily_qty=D("10"),
        execution_price=D("4"),
        trade_source=TradeSource.ACTUAL,
    )
    result = build(
        with_prices(
            base_bundle(
                initial_market_date=date(2026, 1, 1),
                underlyings=(configured,),
                trades=(trade,),
            )
        )
    )
    assert result.manifest.status.value == "VERIFIED"
    fixing = next(row for row in result.fixings if row.source_underlying == "HH")
    assert fixing.price_lookup_date == date(2026, 1, 5)
    assert fixing.fixing_date == date(2026, 1, 2)
    assert fixing.fixing_amount == D("-200")
    pnl = next(row for row in result.pnl if row.fixing_amount)
    assert pnl.fixing_amount == D("200") / D("1.2")


@pytest.mark.golden
def test_month_ahead_prices_by_fixing_date_but_can_be_excluded_from_pnl() -> None:
    configured = underlying(
        "MA",
        FixingMethod.MONTH_AHEAD,
        include_fixing_in_pnl=False,
    )
    trade = Trade(
        source_row_id="T-MA-NO-PNL",
        trade_date=date(2026, 1, 15),
        book="BOOK1",
        underlying="MA",
        side=Side.BUY,
        start_date=date(2026, 2, 1),
        end_date=date(2026, 2, 28),
        daily_qty=D("2"),
        execution_price=D("18"),
        trade_source=TradeSource.ACTUAL,
    )
    bundle = base_bundle(
        historical_end_date=date(2026, 1, 30), underlyings=(configured,), trades=(trade,)
    )
    result = build(
        with_prices(
            bundle,
            fixing_value=lambda value_date, _name: D("30") + D(value_date.day) / D("100"),
        )
    )
    assert result.manifest.status.value == "VERIFIED"
    fixings = [row for row in result.fixings if row.source_underlying == "MA"]
    assert len({row.fixing_price for row in fixings}) > 1
    assert all(row.price_lookup_date == row.fixing_date for row in fixings)
    assert sum((row.fixing_amount for row in result.pnl), D("0")) == D("0")


@pytest.mark.golden
def test_delivery_election_splits_daily_volume_and_preserves_side() -> None:
    base = underlying(
        "TTFDA Heren",
        FixingMethod.HEREN,
        canonical="TTF DA",
        curve_underlying="TTF DA",
        include_fixing_in_pnl=False,
    )
    trade = Trade(
        source_row_id="T-TTF",
        trade_date=date(2026, 1, 2),
        book="BOOK1",
        underlying="TTFDA Heren",
        side=Side.BUY,
        start_date=date(2026, 2, 2),
        end_date=date(2026, 2, 2),
        daily_qty=D("20"),
        execution_price=D("18"),
        trade_source=TradeSource.ACTUAL,
    )
    election = DeliveryElection(
        decision_date=date(2026, 1, 30),
        book="BOOK1",
        underlying="TTFDA Heren",
        side=Side.BUY,
        start_date=date(2026, 2, 2),
        end_date=date(2026, 2, 2),
        delivery_daily_qty=D("10"),
        unit="MWh",
        source_row_id="DEL-TTF",
    )
    bundle = base_bundle(
        initial_market_date=date(2026, 1, 1),
        historical_end_date=date(2026, 2, 3),
        underlyings=(base,),
        trades=(trade,),
    ).model_copy(update={"delivery_elections": (election,)})
    result = build(with_prices(bundle))
    assert result.manifest.status.value == "VERIFIED"
    rows = [row for row in result.fixings if row.delivery_day == date(2026, 2, 2)]
    assert {(row.source_underlying, row.fixing_volume) for row in rows} == {
        ("TTFDA Heren", D("-10")),
        ("TTFDA Delivery", D("-10")),
    }
    assert sum((row.fixing_volume for row in rows), D("0")) == D("-20")
    assert sum((row.fixing_amount for row in result.pnl), D("0")) == D("200")


@pytest.mark.golden
def test_month_ahead_remaining_dates_conserve_volume() -> None:
    configured = underlying("MA", FixingMethod.MONTH_AHEAD)
    trade = Trade(
        source_row_id="T-MA",
        trade_date=date(2026, 1, 15),
        book="BOOK1",
        underlying="MA",
        side=Side.BUY,
        start_date=date(2026, 2, 1),
        end_date=date(2026, 2, 28),
        daily_qty=D("2"),
        execution_price=D("18"),
        trade_source=TradeSource.ACTUAL,
    )
    bundle = base_bundle(
        historical_end_date=date(2026, 1, 30), underlyings=(configured,), trades=(trade,)
    )
    result = build(with_prices(bundle))
    assert result.manifest.status.value == "VERIFIED"
    trade_fixing_volume = sum(
        (row.fixing_volume for row in result.fixings if row.trade_source is TradeSource.ACTUAL),
        D("0"),
    )
    assert abs(trade_fixing_volume - D("-56")) <= VOLUME_TOLERANCE
    assert min(row.fixing_date for row in result.fixings) >= trade.trade_date


@pytest.mark.golden
def test_zero_closure_is_emitted_once_without_requiring_a_zero_position_price() -> None:
    configured = underlying("WD", FixingMethod.WITHINDAY)
    initial = InitialExposure(
        initial_market_date=date(2026, 1, 29),
        book="BOOK1",
        underlying="WD",
        delivery_month=date(2026, 1, 1),
        exposure_volume=D("100"),
        source_row_id="INIT-WD",
    )
    bundle = base_bundle(
        initial_market_date=date(2026, 1, 29),
        historical_end_date=date(2026, 2, 4),
        underlyings=(configured,),
        initial_exposure=(initial,),
    )
    priced = with_prices(bundle)
    priced = priced.model_copy(
        update={
            "curve_prices": tuple(
                row for row in priced.curve_prices if row.market_date != date(2026, 2, 2)
            )
        }
    )
    result = build(priced)
    assert result.manifest.status.value == "VERIFIED"
    closures = [row for row in result.exposure if row.is_explicit_closure]
    assert len(closures) == 1
    assert closures[0].market_date == date(2026, 2, 2)
    assert closures[0].curve_price is None
    assert closures[0].exposure_mtm == D("0")
    assert not [row for row in result.exposure if row.market_date > date(2026, 2, 2)]


@pytest.mark.golden
def test_operating_flows_and_initial_pnl_bridge() -> None:
    flow = OperatingFlow(
        market_date=date(2026, 1, 5),
        book="BOOK1",
        logistics_source_amount=D("100"),
        fees_and_optimizations=D("10"),
        replication=D("5"),
        source_row_id="OPS-1",
    )
    result = build(with_prices(base_bundle(operating_flows=(flow,))))
    assert result.manifest.status.value == "VERIFIED"
    row = next(row for row in result.pnl if row.underlying == "TOTAL / BOOK LEVEL")
    assert row.logistical_costs == D("-100")
    assert row.total_pnl == D("-85")
    cumulative = next(row for row in result.cumulative_pnl if row.market_date == date(2026, 1, 5))
    assert cumulative.initial_pnl == D("1000")
    assert cumulative.daily_pnl == D("-85")
    assert cumulative.cumulative_pnl == D("915")


@pytest.mark.golden
def test_weekend_trade_is_applied_on_next_market_date() -> None:
    trade = Trade(
        source_row_id="T-WEEKEND",
        trade_date=date(2026, 1, 3),
        book="BOOK1",
        underlying="GAS",
        side=Side.BUY,
        start_date=date(2026, 1, 7),
        end_date=date(2026, 1, 7),
        daily_qty=D("100"),
        execution_price=D("20"),
        trade_source=TradeSource.ACTUAL,
    )
    result = build(with_prices(base_bundle(trades=(trade,))))
    assert result.manifest.status.value == "VERIFIED"
    trade_ledger = next(row for row in result.event_ledger if row.source_row_id == "T-WEEKEND")
    assert trade_ledger.economic_date == date(2026, 1, 3)
    assert trade_ledger.applied_market_date == date(2026, 1, 5)
    exposure = next(row for row in result.exposure if row.market_date == date(2026, 1, 5))
    assert exposure.exposure_volume == D("100")
