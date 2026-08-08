from datetime import date
from decimal import Decimal

from gtm_engine.calendar import CalendarIndex, normal_fixing_date
from gtm_engine.canonicalize import normalize_text
from gtm_engine.models import (
    FixingMethod,
    Side,
    Trade,
    TradeSource,
)
from gtm_engine.pipeline import build

from .helpers import base_bundle, underlying, with_prices


def test_text_normalization_is_case_and_accent_insensitive() -> None:
    assert normalize_text("  Metodología   Heren ") == "METODOLOGIA HEREN"


def test_day_ahead_and_heren_weekend_rules_are_distinct() -> None:
    bundle = base_bundle()
    calendar = CalendarIndex(bundle.calendar, bundle.config)
    monday = date(2026, 1, 5)
    assert normal_fixing_date(monday, FixingMethod.DAY_AHEAD, calendar) == date(2026, 1, 4)
    assert normal_fixing_date(monday, FixingMethod.HEREN, calendar) == date(2026, 1, 2)


def test_same_day_fixing_is_valid() -> None:
    trade = Trade(
        source_row_id="T-SAME",
        trade_date=date(2026, 1, 5),
        book="BOOK1",
        underlying="GAS",
        side=Side.BUY,
        start_date=date(2026, 1, 6),
        end_date=date(2026, 1, 6),
        daily_qty=Decimal("100"),
        execution_price=Decimal("19"),
        trade_source=TradeSource.ACTUAL,
    )
    result = build(with_prices(base_bundle(trades=(trade,))))
    assert result.manifest.status.value == "VERIFIED"
    assert not {item.code for item in result.validation} & {"LATE_FIXING_OPPORTUNITY"}
    assert len(result.fixings) == 1
    assert result.fixings[0].fixing_date == trade.trade_date


def test_non_month_ahead_late_trade_is_rejected() -> None:
    trade = Trade(
        source_row_id="T-LATE",
        trade_date=date(2026, 1, 6),
        book="BOOK1",
        underlying="GAS",
        side=Side.BUY,
        start_date=date(2026, 1, 6),
        end_date=date(2026, 1, 7),
        daily_qty=Decimal("100"),
        execution_price=Decimal("19"),
        trade_source=TradeSource.ACTUAL,
    )
    result = build(with_prices(base_bundle(trades=(trade,))))
    assert result.manifest.status.value == "FAILED"
    assert "LATE_FIXING_OPPORTUNITY" in {item.code for item in result.validation}
    assert not result.fixings


def test_month_ahead_with_no_remaining_date_is_rejected() -> None:
    config = underlying("MA", FixingMethod.MONTH_AHEAD)
    trade = Trade(
        source_row_id="T-MA-LATE",
        trade_date=date(2026, 2, 2),
        book="BOOK1",
        underlying="MA",
        side=Side.BUY,
        start_date=date(2026, 2, 1),
        end_date=date(2026, 2, 28),
        daily_qty=Decimal("1"),
        execution_price=Decimal("18"),
        trade_source=TradeSource.ACTUAL,
    )
    bundle = base_bundle(
        historical_end_date=date(2026, 2, 10), underlyings=(config,), trades=(trade,)
    )
    result = build(with_prices(bundle))
    assert result.manifest.status.value == "FAILED"
    assert "LATE_MONTH_AHEAD_TRADE" in {item.code for item in result.validation}


def test_product_name_does_not_override_setup_method() -> None:
    configured = underlying("TTF DA", FixingMethod.HEREN)
    bundle = base_bundle(underlyings=(configured,))
    calendar = CalendarIndex(bundle.calendar, bundle.config)
    assert configured.fixing_method is FixingMethod.HEREN
    assert normal_fixing_date(date(2026, 1, 5), configured.fixing_method, calendar) == date(
        2026, 1, 2
    )
