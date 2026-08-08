"""Executable cases from the approved GTM v2 Golden Regression Test Pack."""

from __future__ import annotations

from datetime import date
from decimal import Decimal

import pytest
from pydantic import ValidationError

from gtm_engine.models import (
    FixingMethod,
    InitialExposure,
    OperatingFlow,
    Side,
    Trade,
    TradeSource,
)
from gtm_engine.pipeline import build

from .helpers import base_bundle, underlying, with_prices

D = Decimal
pytestmark = pytest.mark.golden


def _trade(
    source_row_id: str,
    *,
    trade_date: date,
    side: Side = Side.BUY,
    start_date: date,
    end_date: date | None = None,
    daily_qty: str = "100",
    execution_price: str = "45",
    underlying_name: str = "TEST GAS",
) -> Trade:
    return Trade(
        source_row_id=source_row_id,
        trade_date=trade_date,
        book="BOOK1",
        underlying=underlying_name,
        side=side,
        start_date=start_date,
        end_date=end_date or start_date,
        daily_qty=D(daily_qty),
        execution_price=D(execution_price),
        trade_source=TradeSource.ACTUAL,
    )


def _test_gas(method: FixingMethod = FixingMethod.WITHINDAY):
    return underlying("TEST GAS", method)


def _curve_on_dates(values: dict[date, str], default: str = "47"):
    def curve_value(market_date: date, _underlying: str, _delivery_month: date) -> Decimal:
        return D(values.get(market_date, default))

    return curve_value


def _pnl_on(result: object, market_date: date):
    return next(row for row in result.pnl if row.market_date == market_date)


def test_pack_01_buy_sell_sign_and_quantity_validation() -> None:
    common = dict(
        trade_date=date(2026, 7, 1),
        start_date=date(2026, 9, 1),
        end_date=date(2026, 9, 1),
    )
    buy = _trade("PACK-01-BUY", side=Side.BUY, **common)
    sell = _trade("PACK-01-SELL", side=Side.SELL, **common)

    buy_result = build(
        with_prices(
            base_bundle(
                initial_market_date=date(2026, 6, 30),
                historical_end_date=date(2026, 7, 1),
                underlyings=(_test_gas(),),
                trades=(buy,),
            )
        )
    )
    sell_result = build(
        with_prices(
            base_bundle(
                initial_market_date=date(2026, 6, 30),
                historical_end_date=date(2026, 7, 1),
                underlyings=(_test_gas(),),
                trades=(sell,),
            )
        )
    )
    assert buy_result.manifest.status.value == "VERIFIED"
    assert sell_result.manifest.status.value == "VERIFIED"
    assert buy_result.exposure[0].exposure_volume == D("100")
    assert sell_result.exposure[0].exposure_volume == D("-100")

    negative_sell = sell.model_copy(update={"source_row_id": "PACK-01-NEG", "daily_qty": D("-100")})
    negative_result = build(
        base_bundle(
            initial_market_date=date(2026, 6, 30),
            historical_end_date=date(2026, 7, 1),
            underlyings=(_test_gas(),),
            trades=(negative_sell,),
        )
    )
    assert negative_result.manifest.status.value == "FAILED"
    assert "NEGATIVE_DAILY_QTY" in {row.code for row in negative_result.validation}
    assert not negative_result.exposure

    zero_trade = buy.model_copy(update={"source_row_id": "PACK-01-ZERO", "daily_qty": D("0")})
    zero_result = build(
        base_bundle(
            initial_market_date=date(2026, 6, 30),
            historical_end_date=date(2026, 7, 1),
            underlyings=(_test_gas(),),
            trades=(zero_trade,),
        )
    )
    assert zero_result.manifest.status.value == "VERIFIED"
    assert "ZERO_DAILY_QTY" in {row.code for row in zero_result.validation}
    assert not zero_result.event_ledger

    blank_payload = buy.model_dump(mode="python")
    blank_payload["daily_qty"] = ""
    with pytest.raises(ValidationError, match="daily_qty"):
        Trade.model_validate(blank_payload)


def test_pack_02_trade_entry_exposure_and_pnl_for_buy_and_sell() -> None:
    for side, expected_exposure, expected_mtm, expected_adjustment, expected_total in (
        (Side.BUY, D("100"), D("4700"), D("-4500"), D("200")),
        (Side.SELL, D("-100"), D("-4700"), D("4500"), D("-200")),
    ):
        trade = _trade(
            f"PACK-02-{side.value}",
            trade_date=date(2026, 7, 1),
            side=side,
            start_date=date(2026, 9, 1),
        )
        bundle = base_bundle(
            initial_market_date=date(2026, 6, 30),
            historical_end_date=date(2026, 7, 1),
            underlyings=(_test_gas(),),
            trades=(trade,),
            initial_pnl_amount=D("0"),
        )
        result = build(with_prices(bundle, curve_value=_curve_on_dates({date(2026, 7, 1): "47"})))
        assert result.manifest.status.value == "VERIFIED"
        assert result.exposure[0].exposure_volume == expected_exposure
        assert result.exposure[0].exposure_mtm == expected_mtm
        pnl = _pnl_on(result, date(2026, 7, 1))
        assert pnl.gross_delta_exposure_mtm == expected_mtm
        assert pnl.trade_entry_adjustment == expected_adjustment
        assert pnl.total_pnl == expected_total


def test_pack_03_pure_market_movement_after_trade_entry() -> None:
    trade = _trade(
        "PACK-03",
        trade_date=date(2026, 7, 1),
        start_date=date(2026, 9, 1),
    )
    result = build(
        with_prices(
            base_bundle(
                initial_market_date=date(2026, 6, 30),
                historical_end_date=date(2026, 7, 2),
                underlyings=(_test_gas(),),
                trades=(trade,),
                initial_pnl_amount=D("0"),
            ),
            curve_value=_curve_on_dates({date(2026, 7, 1): "47", date(2026, 7, 2): "49"}),
        )
    )
    july_1 = _pnl_on(result, date(2026, 7, 1))
    july_2 = _pnl_on(result, date(2026, 7, 2))
    assert july_1.total_pnl == D("200")
    assert july_2.exposure_mtm == D("4900")
    assert july_2.gross_delta_exposure_mtm == D("200")
    assert july_2.trade_entry_adjustment == D("0")
    assert july_2.fixing_amount == D("0")
    assert july_2.total_pnl == D("200")


def test_pack_04_full_fixing_closes_exposure_with_positive_economic_value() -> None:
    trade = _trade(
        "PACK-04",
        trade_date=date(2026, 7, 1),
        start_date=date(2026, 7, 2),
        execution_price="47",
    )
    result = build(
        with_prices(
            base_bundle(
                initial_market_date=date(2026, 6, 30),
                historical_end_date=date(2026, 7, 2),
                underlyings=(_test_gas(),),
                trades=(trade,),
                initial_pnl_amount=D("0"),
            ),
            fixing_value=D("48"),
            curve_value=D("47"),
        )
    )
    assert result.manifest.status.value == "VERIFIED"
    assert len(result.fixings) == 1
    assert result.fixings[0].fixing_volume == D("-100")
    assert result.fixings[0].fixing_amount == D("-4800")
    closure = next(row for row in result.exposure if row.market_date == date(2026, 7, 2))
    assert closure.exposure_volume == D("0")
    assert closure.exposure_mtm == D("0")
    assert closure.is_explicit_closure
    pnl = _pnl_on(result, date(2026, 7, 2))
    assert pnl.gross_delta_exposure_mtm == D("-4700")
    assert pnl.fixing_amount == D("4800")
    assert pnl.total_pnl == D("100")


def test_pack_05_same_day_fixing_is_eligible_in_fixings_and_exposure() -> None:
    trade = _trade(
        "PACK-05",
        trade_date=date(2026, 7, 14),
        start_date=date(2026, 7, 15),
    )
    result = build(
        with_prices(
            base_bundle(
                initial_market_date=date(2026, 7, 13),
                historical_end_date=date(2026, 7, 15),
                underlyings=(_test_gas(FixingMethod.DAY_AHEAD),),
                trades=(trade,),
                initial_pnl_amount=D("0"),
            ),
            fixing_value=D("48"),
        )
    )
    assert result.manifest.status.value == "VERIFIED"
    assert len(result.fixings) == 1
    assert result.fixings[0].fixing_date == trade.trade_date
    events = [row for row in result.event_ledger if row.source_row_id == trade.source_row_id]
    assert {row.economic_date for row in events} == {trade.trade_date}
    assert sum((row.signed_volume_change for row in events), D("0")) == D("0")
    assert not result.exposure


def test_pack_06_trade_after_all_month_ahead_fixing_opportunities_fails() -> None:
    trade = _trade(
        "PACK-06",
        trade_date=date(2026, 6, 10),
        start_date=date(2026, 6, 1),
        end_date=date(2026, 6, 30),
        daily_qty="1",
        underlying_name="MONTH AHEAD TEST",
    )
    result = build(
        base_bundle(
            initial_market_date=date(2026, 6, 1),
            historical_end_date=date(2026, 6, 12),
            underlyings=(underlying("MONTH AHEAD TEST", FixingMethod.MONTH_AHEAD),),
            trades=(trade,),
        )
    )
    assert result.manifest.status.value == "FAILED"
    assert "LATE_MONTH_AHEAD_TRADE" in {row.code for row in result.validation}
    assert not result.fixings
    assert not result.exposure


def test_pack_07_weekend_trade_is_applied_once_on_monday() -> None:
    trade = _trade(
        "PACK-07",
        trade_date=date(2026, 7, 11),
        start_date=date(2026, 9, 1),
    )
    result = build(
        with_prices(
            base_bundle(
                initial_market_date=date(2026, 7, 10),
                historical_end_date=date(2026, 7, 14),
                underlyings=(_test_gas(),),
                trades=(trade,),
                initial_pnl_amount=D("0"),
            )
        )
    )
    assert result.manifest.status.value == "VERIFIED"
    ledger = next(
        row
        for row in result.event_ledger
        if row.source_row_id == trade.source_row_id and row.event_type.value == "TRADE"
    )
    assert ledger.economic_date == date(2026, 7, 11)
    assert ledger.applied_market_date == date(2026, 7, 13)
    exposures = {row.market_date: row.exposure_volume for row in result.exposure}
    assert date(2026, 7, 11) not in exposures
    assert date(2026, 7, 12) not in exposures
    assert exposures[date(2026, 7, 13)] == D("100")
    assert exposures[date(2026, 7, 14)] == D("100")


def test_pack_08_closure_emits_one_zero_and_reopening_resumes_rows() -> None:
    opening_trade = _trade(
        "PACK-08-OPEN",
        trade_date=date(2026, 7, 1),
        start_date=date(2026, 7, 2),
        execution_price="47",
    )
    reopening_trade = _trade(
        "PACK-08-REOPEN",
        trade_date=date(2026, 7, 5),
        start_date=date(2026, 7, 10),
        daily_qty="50",
        execution_price="47",
    )
    bundle = base_bundle(
        initial_market_date=date(2026, 6, 30),
        historical_end_date=date(2026, 7, 6),
        underlyings=(_test_gas(),),
        trades=(opening_trade, reopening_trade),
        initial_pnl_amount=D("0"),
    )
    calendar = tuple(
        row.model_copy(update={"is_market_day": True})
        if date(2026, 7, 4) <= row.date <= date(2026, 7, 5)
        else row
        for row in bundle.calendar
    )
    result = build(with_prices(bundle.model_copy(update={"calendar": calendar})))
    assert result.manifest.status.value == "VERIFIED"
    exposure_by_date = {row.market_date: row for row in result.exposure}
    assert exposure_by_date[date(2026, 7, 1)].exposure_volume == D("100")
    assert exposure_by_date[date(2026, 7, 2)].exposure_volume == D("0")
    assert exposure_by_date[date(2026, 7, 2)].is_explicit_closure
    assert date(2026, 7, 3) not in exposure_by_date
    assert date(2026, 7, 4) not in exposure_by_date
    assert exposure_by_date[date(2026, 7, 5)].exposure_volume == D("50")
    assert len([row for row in result.exposure if row.is_explicit_closure]) == 1


def test_pack_09_missing_required_price_blocks_but_unused_absence_does_not() -> None:
    trade = _trade(
        "PACK-09",
        trade_date=date(2026, 7, 1),
        start_date=date(2026, 9, 1),
    )
    priced = with_prices(
        base_bundle(
            initial_market_date=date(2026, 6, 30),
            historical_end_date=date(2026, 7, 1),
            underlyings=(_test_gas(),),
            trades=(trade,),
        )
    )
    missing_key = (date(2026, 7, 1), "TEST GAS", date(2026, 9, 1))
    missing_curve = tuple(
        row
        for row in priced.curve_prices
        if (row.market_date, row.underlying, row.delivery_month) != missing_key
    )
    result = build(priced.model_copy(update={"curve_prices": missing_curve}))
    assert result.manifest.status.value == "FAILED"
    error = next(row for row in result.validation if row.code == "MISSING_CURVE_PRICE")
    assert "Market Date=2026-07-01" in error.message
    assert "Underlying=TEST GAS" in error.message
    assert "Delivery Month=2026-09-01" in error.message
    assert not result.exposure

    unused_absence = build(
        base_bundle(
            initial_market_date=date(2026, 6, 30),
            historical_end_date=date(2026, 7, 1),
            underlyings=(_test_gas(),),
        )
    )
    assert unused_absence.manifest.status.value == "VERIFIED"


def test_pack_10_operating_pnl_components_sum_to_minus_100() -> None:
    opening = InitialExposure(
        initial_market_date=date(2026, 6, 30),
        book="BOOK1",
        underlying="MARKET GAS",
        delivery_month=date(2026, 9, 1),
        exposure_volume=D("100"),
        source_row_id="PACK-10-OPEN",
    )
    fixing_trade = _trade(
        "PACK-10-FIX",
        trade_date=date(2026, 7, 1),
        start_date=date(2026, 7, 1),
        execution_price="0",
        underlying_name="FIX GAS",
    )
    flow = OperatingFlow(
        market_date=date(2026, 7, 1),
        book="BOOK1",
        logistics_source_amount=D("1000"),
        fees_and_optimizations=D("150"),
        replication=D("50"),
        source_row_id="PACK-10",
    )

    def curve_value(market_date: date, curve_underlying: str, _delivery_month: date) -> Decimal:
        if curve_underlying == "MARKET GAS":
            return D("45") if market_date == date(2026, 6, 30) else D("50")
        return D("1")

    def fixing_value(price_date: date, pricing_underlying: str) -> Decimal:
        if price_date == date(2026, 7, 1) and pricing_underlying == "FIX GAS":
            return D("2")
        return D("1")

    result = build(
        with_prices(
            base_bundle(
                initial_market_date=date(2026, 6, 30),
                historical_end_date=date(2026, 7, 1),
                underlyings=(
                    underlying("MARKET GAS", FixingMethod.WITHINDAY),
                    underlying("FIX GAS", FixingMethod.WITHINDAY),
                ),
                initial_exposure=(opening,),
                trades=(fixing_trade,),
                operating_flows=(flow,),
                initial_pnl_amount=D("0"),
            ),
            curve_value=curve_value,
            fixing_value=fixing_value,
        )
    )
    assert result.manifest.status.value == "VERIFIED"
    market_pnl = next(row for row in result.pnl if row.underlying == "MARKET GAS")
    fixing_pnl = next(row for row in result.pnl if row.underlying == "FIX GAS")
    operating_pnl = next(row for row in result.pnl if row.underlying == "TOTAL / BOOK LEVEL")
    assert market_pnl.gross_delta_exposure_mtm == D("500")
    assert fixing_pnl.fixing_amount == D("200")
    assert operating_pnl.logistical_costs == D("-1000")
    assert operating_pnl.fees_and_optimizations == D("150")
    assert operating_pnl.replication == D("50")
    assert sum((row.total_pnl for row in result.pnl), D("0")) == D("-100")


def test_pack_11_only_trades_after_initial_market_date_are_incremental() -> None:
    trade_a = _trade(
        "PACK-11-A",
        trade_date=date(2026, 6, 30),
        start_date=date(2026, 9, 1),
    )
    trade_b = _trade(
        "PACK-11-B",
        trade_date=date(2026, 7, 1),
        start_date=date(2026, 9, 1),
    )
    result = build(
        with_prices(
            base_bundle(
                initial_market_date=date(2026, 6, 30),
                historical_end_date=date(2026, 7, 1),
                underlyings=(_test_gas(),),
                trades=(trade_a, trade_b),
                initial_pnl_amount=D("0"),
            )
        )
    )
    source_ids = {row.source_row_id for row in result.event_ledger}
    assert result.manifest.status.value == "VERIFIED"
    assert "PACK-11-A" not in source_ids
    assert "PACK-11-B" in source_ids
    assert result.exposure[0].exposure_volume == D("100")


def test_pack_12_end_to_end_conservation_reconciles_to_eur_500() -> None:
    """Method-valid pipeline twin of the pack's three-day economic lifecycle.

    The pack labels the delivery Sep-26 but fixes it on 03/07, which no configured GTM method
    can produce. A one-day WITHINDAY delivery on 03/07 preserves every stated economic amount and
    exercises the complete production pipeline without introducing a test-only fixing method.
    """

    trade = _trade(
        "PACK-12",
        trade_date=date(2026, 7, 1),
        start_date=date(2026, 7, 3),
    )
    result = build(
        with_prices(
            base_bundle(
                initial_market_date=date(2026, 6, 30),
                historical_end_date=date(2026, 7, 3),
                underlyings=(_test_gas(),),
                trades=(trade,),
                initial_pnl_amount=D("0"),
            ),
            fixing_value=D("50"),
            curve_value=_curve_on_dates(
                {date(2026, 7, 1): "47", date(2026, 7, 2): "49"},
                default="49",
            ),
        )
    )
    assert result.manifest.status.value == "VERIFIED"
    july_1 = _pnl_on(result, date(2026, 7, 1))
    july_2 = _pnl_on(result, date(2026, 7, 2))
    july_3 = _pnl_on(result, date(2026, 7, 3))

    assert (july_1.exposure_mtm, july_1.gross_delta_exposure_mtm) == (D("4700"), D("4700"))
    assert july_1.trade_entry_adjustment == D("-4500")
    assert july_1.total_pnl == D("200")

    assert (july_2.exposure_mtm, july_2.gross_delta_exposure_mtm) == (D("4900"), D("200"))
    assert july_2.total_pnl == D("200")

    assert (july_3.exposure_mtm, july_3.gross_delta_exposure_mtm) == (D("0"), D("-4900"))
    assert july_3.fixing_amount == D("5000")
    assert july_3.total_pnl == D("100")

    actual_cumulative = sum((row.total_pnl for row in result.pnl), D("0"))
    independent_economics = D("100") * (D("50") - D("45"))
    assert actual_cumulative == independent_economics == D("500")
