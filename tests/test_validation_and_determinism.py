from datetime import date
from decimal import Decimal

from gtm_engine.models import InitialExposure, OperatingFlow, Side, Trade, TradeSource
from gtm_engine.pipeline import build

from .helpers import base_bundle, with_prices

D = Decimal


def test_missing_fixing_price_blocks_and_identifies_key() -> None:
    initial = InitialExposure(
        initial_market_date=date(2026, 1, 2),
        book="BOOK1",
        underlying="GAS",
        delivery_month=date(2026, 1, 1),
        exposure_volume=D("100"),
        source_row_id="INIT-1",
    )
    result = build(base_bundle(initial_exposure=(initial,)))
    assert result.manifest.status.value == "FAILED"
    error = next(item for item in result.validation if item.code == "MISSING_FIXING_PRICE")
    assert "Market Date=" in error.message
    assert "Underlying=GAS" in error.message
    assert "Delivery Month=2026-01-01" in error.message
    assert not result.fixings


def test_missing_curve_price_blocks_after_fixing_prices_are_complete() -> None:
    initial = InitialExposure(
        initial_market_date=date(2026, 1, 2),
        book="BOOK1",
        underlying="GAS",
        delivery_month=date(2026, 1, 1),
        exposure_volume=D("100"),
        source_row_id="INIT-1",
    )
    priced = with_prices(base_bundle(initial_exposure=(initial,)))
    no_curves = priced.model_copy(update={"curve_prices": ()})
    result = build(no_curves)
    assert result.manifest.status.value == "FAILED"
    assert "MISSING_CURVE_PRICE" in {item.code for item in result.validation}
    assert not result.exposure


def test_unused_price_absence_does_not_block_empty_portfolio() -> None:
    result = build(base_bundle())
    assert result.manifest.status.value == "VERIFIED"
    assert not result.fixings
    assert not result.exposure


def test_required_prices_must_match_configured_currency_and_unit() -> None:
    initial = InitialExposure(
        initial_market_date=date(2026, 1, 2),
        book="BOOK1",
        underlying="GAS",
        delivery_month=date(2026, 1, 1),
        exposure_volume=D("100"),
        source_row_id="INIT-CONTRACT",
    )
    priced = with_prices(base_bundle(initial_exposure=(initial,)))
    wrong_fixing_unit = tuple(
        row.model_copy(update={"unit": "therm"})
        if row.price_lookup_date == date(2026, 1, 4)
        else row
        for row in priced.fixing_prices
    )
    fixing_result = build(priced.model_copy(update={"fixing_prices": wrong_fixing_unit}))
    assert "FIXING_PRICE_UNIT_MISMATCH" in {item.code for item in fixing_result.validation}

    wrong_curve_currency = tuple(
        row.model_copy(update={"currency": "USD"})
        if row.market_date == date(2026, 1, 2) and row.delivery_month == date(2026, 1, 1)
        else row
        for row in priced.curve_prices
    )
    curve_result = build(priced.model_copy(update={"curve_prices": wrong_curve_currency}))
    assert "CURVE_PRICE_CURRENCY_MISMATCH" in {item.code for item in curve_result.validation}


def test_calendar_gap_is_a_blocking_error() -> None:
    bundle = base_bundle()
    missing_date = date(2026, 1, 6)
    gapped = bundle.model_copy(
        update={"calendar": tuple(row for row in bundle.calendar if row.date != missing_date)}
    )
    result = build(gapped)
    assert result.manifest.status.value == "FAILED"
    assert "CALENDAR_GAP" in {item.code for item in result.validation}


def test_duplicate_operating_flow_key_is_a_blocking_error() -> None:
    flow = OperatingFlow(
        market_date=date(2026, 1, 5),
        book="BOOK1",
        fees_and_optimizations=D("10"),
        source_row_id="OPS-1",
    )
    duplicate = flow.model_copy(update={"source_row_id": "OPS-2", "replication": D("3")})
    result = build(base_bundle(operating_flows=(flow, duplicate)))
    assert result.manifest.status.value == "FAILED"
    assert "DUPLICATE_OPERATING_FLOW" in {item.code for item in result.validation}


def test_unknown_initial_exposure_book_and_underlying_are_blocking_errors() -> None:
    initial = InitialExposure(
        initial_market_date=date(2026, 1, 2),
        book="UNKNOWN BOOK",
        underlying="UNKNOWN UNDERLYING",
        delivery_month=date(2026, 1, 1),
        exposure_volume=D("100"),
        source_row_id="INIT-UNKNOWN",
    )
    result = build(base_bundle(initial_exposure=(initial,)))
    codes = {item.code for item in result.validation}
    assert result.manifest.status.value == "FAILED"
    assert "UNKNOWN_BOOK" in codes
    assert "UNKNOWN_UNDERLYING" in codes


def test_explicit_zero_trade_warns_but_does_not_create_event() -> None:
    trade = Trade(
        source_row_id="T-ZERO",
        trade_date=date(2026, 1, 5),
        book="BOOK1",
        underlying="GAS",
        side=Side.SELL,
        start_date=date(2026, 1, 7),
        end_date=date(2026, 1, 7),
        daily_qty=D("0"),
        execution_price=D("20"),
        trade_source=TradeSource.ACTUAL,
    )
    result = build(base_bundle(trades=(trade,)))
    assert result.manifest.status.value == "VERIFIED"
    assert "ZERO_DAILY_QTY" in {item.code for item in result.validation}
    assert not result.fixings
    assert not result.exposure


def _economic_rows(rows: tuple) -> list[dict]:
    return [row.model_dump(mode="json", exclude={"build_id"}) for row in rows]


def test_input_order_does_not_change_economic_output_or_fingerprint() -> None:
    trade_1 = Trade(
        source_row_id="T-1",
        trade_date=date(2026, 1, 5),
        book="BOOK1",
        underlying="GAS",
        side=Side.BUY,
        start_date=date(2026, 1, 7),
        end_date=date(2026, 1, 8),
        daily_qty=D("10"),
        execution_price=D("20"),
        trade_source=TradeSource.ACTUAL,
    )
    trade_2 = trade_1.model_copy(
        update={"source_row_id": "T-2", "side": Side.SELL, "daily_qty": D("3")}
    )
    first = with_prices(base_bundle(trades=(trade_1, trade_2)))
    second = first.model_copy(update={"trades": (trade_2, trade_1)})
    result_1 = build(first)
    result_2 = build(second)
    assert result_1.manifest.status.value == result_2.manifest.status.value == "VERIFIED"
    assert result_1.manifest.calculation_fingerprint == result_2.manifest.calculation_fingerprint
    assert result_1.manifest.build_id == result_2.manifest.build_id
    assert result_1.manifest.output_hashes == result_2.manifest.output_hashes
    assert _economic_rows(result_1.fixings) == _economic_rows(result_2.fixings)
    assert _economic_rows(result_1.exposure) == _economic_rows(result_2.exposure)
    assert _economic_rows(result_1.pnl) == _economic_rows(result_2.pnl)
