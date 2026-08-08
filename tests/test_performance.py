from __future__ import annotations

from datetime import date
from time import perf_counter

import pytest

from gtm_engine.models import Side, Trade, TradeSource
from gtm_engine.pipeline import build

from .helpers import D, base_bundle, with_prices


@pytest.mark.performance
def test_synthetic_daily_build_avoids_quadratic_slowdown() -> None:
    trades = tuple(
        Trade(
            source_row_id=f"PERF-{number:04d}",
            trade_date=date(2026, 1, 5),
            book="BOOK1",
            underlying="GAS",
            side=Side.BUY if number % 2 else Side.SELL,
            start_date=date(2026, 2, 1),
            end_date=date(2026, 2, 28),
            daily_qty=D("10") + D(number % 7),
            execution_price=D("20"),
            trade_source=TradeSource.ACTUAL,
        )
        for number in range(476)
    )
    bundle = with_prices(base_bundle(historical_end_date=date(2026, 2, 27), trades=trades))

    started = perf_counter()
    result = build(bundle)
    elapsed = perf_counter() - started

    assert result.manifest.status.value == "VERIFIED"
    assert result.manifest.row_counts["event_ledger"] == 476 * 29
    assert elapsed < 5
    assert result.manifest.peak_memory_bytes < 2 * 1024**3
