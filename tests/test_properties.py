from datetime import date
from decimal import Decimal

from hypothesis import given, settings
from hypothesis import strategies as st

from gtm_engine.models import Side, Trade, TradeSource
from gtm_engine.pipeline import build

from .helpers import base_bundle, with_prices


def _totals(result: object) -> tuple[Decimal, Decimal, Decimal]:
    return (
        sum((row.fixing_volume for row in result.fixings), Decimal("0")),
        sum((row.exposure_volume for row in result.exposure), Decimal("0")),
        sum((row.total_pnl for row in result.pnl), Decimal("0")),
    )


@settings(max_examples=15, deadline=None)
@given(st.integers(min_value=1, max_value=1000))
def test_splitting_trade_does_not_change_aggregate_economics(quantity: int) -> None:
    full = Trade(
        source_row_id="FULL",
        trade_date=date(2026, 1, 5),
        book="BOOK1",
        underlying="GAS",
        side=Side.BUY,
        start_date=date(2026, 1, 7),
        end_date=date(2026, 1, 8),
        daily_qty=Decimal(quantity),
        execution_price=Decimal("20"),
        trade_source=TradeSource.ACTUAL,
    )
    half = Decimal(quantity) / Decimal("2")
    split_1 = full.model_copy(update={"source_row_id": "SPLIT-1", "daily_qty": half})
    split_2 = full.model_copy(update={"source_row_id": "SPLIT-2", "daily_qty": half})
    result_full = build(with_prices(base_bundle(trades=(full,))))
    result_split = build(with_prices(base_bundle(trades=(split_1, split_2))))
    assert result_full.manifest.status.value == result_split.manifest.status.value == "VERIFIED"
    assert _totals(result_full) == _totals(result_split)


@settings(max_examples=15, deadline=None)
@given(st.integers(min_value=1, max_value=1000))
def test_equal_buy_and_sell_net_to_zero_exposure(quantity: int) -> None:
    common = dict(
        trade_date=date(2026, 1, 5),
        book="BOOK1",
        underlying="GAS",
        start_date=date(2026, 1, 7),
        end_date=date(2026, 1, 8),
        daily_qty=Decimal(quantity),
        execution_price=Decimal("20"),
        trade_source=TradeSource.ACTUAL,
    )
    buy = Trade(source_row_id="BUY", side=Side.BUY, **common)
    sell = Trade(source_row_id="SELL", side=Side.SELL, **common)
    result = build(with_prices(base_bundle(trades=(buy, sell))))
    assert result.manifest.status.value == "VERIFIED"
    assert sum((row.exposure_volume for row in result.exposure), Decimal("0")) == 0
