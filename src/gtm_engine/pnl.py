"""Daily and cumulative P&L built from normalized engine outputs."""

from __future__ import annotations

from collections import defaultdict
from datetime import date
from decimal import Decimal

from .calendar import CalendarIndex
from .canonicalize import Registry
from .decimal_utils import ZERO
from .exposure import ExposureKey
from .models import (
    CumulativePnlRow,
    ExposureRow,
    FixingEvent,
    FixingRow,
    InputBundle,
    PnlRow,
    TradeEvent,
)


def trade_entry_adjustments(
    trade_events: tuple[TradeEvent, ...],
    fixing_events: tuple[FixingEvent, ...],
) -> dict[tuple[date, ExposureKey], Decimal]:
    same_day_fixing: dict[tuple[str, date, date], Decimal] = defaultdict(lambda: ZERO)
    for fixing_event in fixing_events:
        same_day_fixing[
            (
                fixing_event.source_row_id,
                fixing_event.delivery_month,
                fixing_event.fixing_date,
            )
        ] += fixing_event.fixing_volume

    adjustments: dict[tuple[date, ExposureKey], Decimal] = defaultdict(lambda: ZERO)
    for trade_event in trade_events:
        if trade_event.applied_market_date is None:
            continue
        open_volume = trade_event.signed_volume + same_day_fixing.get(
            (
                trade_event.source_row_id,
                trade_event.delivery_month,
                trade_event.economic_date,
            ),
            ZERO,
        )
        key: ExposureKey = (
            trade_event.book,
            trade_event.underlying,
            trade_event.delivery_month,
            trade_event.trade_source,
            trade_event.scenario,
        )
        adjustments[(trade_event.applied_market_date, key)] += (
            -open_volume * trade_event.execution_price
        )
    return dict(adjustments)


def build_pnl(
    bundle: InputBundle,
    registry: Registry,
    calendar: CalendarIndex,
    exposure: tuple[ExposureRow, ...],
    fixing_rows: tuple[FixingRow, ...],
    fixing_events: tuple[FixingEvent, ...],
    trade_events: tuple[TradeEvent, ...],
    initial_mtm: dict[ExposureKey, Decimal],
    build_id: str,
) -> tuple[tuple[PnlRow, ...], tuple[CumulativePnlRow, ...]]:
    exposure_by_date: dict[tuple[date, ExposureKey], ExposureRow] = {}
    keys_by_date: dict[date, set[ExposureKey]] = defaultdict(set)
    for exposure_row in exposure:
        key: ExposureKey = (
            exposure_row.book,
            exposure_row.underlying,
            exposure_row.delivery_month,
            exposure_row.trade_source,
            exposure_row.scenario,
        )
        exposure_by_date[(exposure_row.market_date, key)] = exposure_row
        keys_by_date[exposure_row.market_date].add(key)

    economic_fixing_amounts: dict[tuple[date, ExposureKey], Decimal] = defaultdict(lambda: ZERO)
    for fixing_row in fixing_rows:
        key = (
            fixing_row.book,
            fixing_row.underlying,
            fixing_row.delivery_month,
            fixing_row.trade_source,
            fixing_row.scenario,
        )
        # Fixing output preserves the signed exposure-closing settlement:
        # fixing_volume * fixing_price. P&L consumes the opposite cash-flow sign.
        economic_fixing_amounts[(fixing_row.applied_market_date, key)] -= fixing_row.fixing_amount
        keys_by_date[fixing_row.applied_market_date].add(key)

    adjustments = trade_entry_adjustments(trade_events, fixing_events)
    for adjustment_date, adjustment_key in adjustments:
        keys_by_date[adjustment_date].add(adjustment_key)
    prior_mtm = dict(initial_mtm)
    pnl_rows: list[PnlRow] = []

    for market_date in calendar.output_market_dates:
        keys = keys_by_date.get(market_date, set())
        previous_market_date = calendar.previous_output_market_day(market_date)
        for key in sorted(
            keys,
            key=lambda value: (value[0], value[1], value[2], value[3].value, value[4] or ""),
        ):
            current_exposure = exposure_by_date.get((market_date, key))
            current_mtm = current_exposure.exposure_mtm if current_exposure is not None else ZERO
            gross_delta = current_mtm - prior_mtm.get(key, ZERO)
            adjustment = adjustments.get((market_date, key), ZERO)
            delta_exposure = gross_delta + adjustment
            fixing_amount = economic_fixing_amounts.get((market_date, key), ZERO)
            total = delta_exposure + fixing_amount
            pnl_rows.append(
                PnlRow(
                    market_date=market_date,
                    previous_market_date=previous_market_date,
                    book=key[0],
                    underlying=key[1],
                    delivery_month=key[2],
                    exposure_mtm=current_mtm,
                    gross_delta_exposure_mtm=gross_delta,
                    trade_entry_adjustment=adjustment,
                    delta_exposure_mtm=delta_exposure,
                    fixing_amount=fixing_amount,
                    logistical_costs=ZERO,
                    fees_and_optimizations=ZERO,
                    replication=ZERO,
                    total_pnl=total,
                    trade_source=key[3],
                    scenario=key[4],
                    simulation_status=bundle.config.simulation_status,
                    build_id=build_id,
                )
            )
            prior_mtm[key] = current_mtm

    for flow in sorted(
        bundle.operating_flows,
        key=lambda row: (row.market_date, row.book.casefold(), row.source_row_id),
    ):
        if flow.market_date not in calendar.output_market_dates:
            continue
        book_config = registry.book(flow.book)
        if book_config is None:
            continue
        logistical = flow.logistics_source_amount * bundle.config.logistics_sign
        total = logistical + flow.fees_and_optimizations + flow.replication
        pnl_rows.append(
            PnlRow(
                market_date=flow.market_date,
                previous_market_date=calendar.previous_output_market_day(flow.market_date),
                book=book_config.book,
                underlying="TOTAL / BOOK LEVEL",
                delivery_month=None,
                exposure_mtm=ZERO,
                gross_delta_exposure_mtm=ZERO,
                trade_entry_adjustment=ZERO,
                delta_exposure_mtm=ZERO,
                fixing_amount=ZERO,
                logistical_costs=logistical,
                fees_and_optimizations=flow.fees_and_optimizations,
                replication=flow.replication,
                total_pnl=total,
                trade_source=None,
                scenario=None,
                simulation_status=bundle.config.simulation_status,
                build_id=build_id,
            )
        )

    pnl_rows.sort(
        key=lambda row: (
            row.market_date,
            row.book,
            row.underlying,
            row.delivery_month or date.min,
            row.trade_source.value if row.trade_source else "",
            row.scenario or "",
        )
    )

    initial_pnl: dict[str, Decimal] = defaultdict(lambda: ZERO)
    for opening_pnl in bundle.initial_pnl:
        book_config = registry.book(opening_pnl.book)
        if book_config is not None:
            initial_pnl[book_config.book] += opening_pnl.amount
    daily_by_book: dict[tuple[date, str], Decimal] = defaultdict(lambda: ZERO)
    for pnl_row in pnl_rows:
        daily_by_book[(pnl_row.market_date, pnl_row.book)] += pnl_row.total_pnl

    cumulative_rows: list[CumulativePnlRow] = []
    running = dict(initial_pnl)
    active_books = sorted(row.book for row in bundle.books if row.active)
    for market_date in calendar.output_market_dates:
        previous_market_date = calendar.previous_output_market_day(market_date)
        for book_name in active_books:
            daily = daily_by_book.get((market_date, book_name), ZERO)
            running[book_name] = running.get(book_name, ZERO) + daily
            cumulative_rows.append(
                CumulativePnlRow(
                    market_date=market_date,
                    previous_market_date=previous_market_date,
                    book=book_name,
                    initial_pnl=initial_pnl.get(book_name, ZERO),
                    daily_pnl=daily,
                    cumulative_pnl=running[book_name],
                    simulation_status=bundle.config.simulation_status,
                    build_id=build_id,
                )
            )
    return tuple(pnl_rows), tuple(cumulative_rows)
