"""Market-day axis and event-effective-date functions."""

from __future__ import annotations

from bisect import bisect_left
from calendar import monthrange
from dataclasses import dataclass, field
from datetime import date, timedelta

from .models import BuildConfig, FixingMethod, MarketCalendarDay


class CalendarCoverageError(ValueError):
    pass


def month_start(value: date) -> date:
    return value.replace(day=1)


def add_months(value: date, months: int) -> date:
    zero_based = value.year * 12 + value.month - 1 + months
    year, month_zero = divmod(zero_based, 12)
    return date(year, month_zero + 1, 1)


def month_end(value: date) -> date:
    first = month_start(value)
    return date(first.year, first.month, monthrange(first.year, first.month)[1])


def date_range(first: date, last: date) -> tuple[date, ...]:
    if last < first:
        return ()
    return tuple(first + timedelta(days=offset) for offset in range((last - first).days + 1))


@dataclass
class CalendarIndex:
    rows: tuple[MarketCalendarDay, ...]
    config: BuildConfig
    _ordered: tuple[MarketCalendarDay, ...] = field(init=False)
    _by_date: dict[date, bool] = field(init=False)
    _market_dates: tuple[date, ...] = field(init=False)

    def __post_init__(self) -> None:
        self._ordered = tuple(sorted(self.rows, key=lambda row: row.date))
        self._by_date = {row.date: row.is_market_day for row in self._ordered}
        self._market_dates = tuple(row.date for row in self._ordered if row.is_market_day)

    @property
    def all_dates(self) -> tuple[date, ...]:
        return tuple(row.date for row in self._ordered)

    @property
    def market_dates(self) -> tuple[date, ...]:
        return self._market_dates

    @property
    def output_market_dates(self) -> tuple[date, ...]:
        first = max(
            self.config.historical_start_date,
            self.config.initial_market_date + timedelta(days=1),
        )
        return tuple(
            value
            for value in self.market_dates
            if first <= value <= self.config.historical_end_date
        )

    def has_date(self, value: date) -> bool:
        return value in self._by_date

    def is_market_day(self, value: date) -> bool:
        try:
            return self._by_date[value]
        except KeyError as exc:
            raise CalendarCoverageError(f"Calendar has no row for {value.isoformat()}") from exc

    def previous_market_day(self, value: date) -> date:
        position = bisect_left(self.market_dates, value) - 1
        if position < 0:
            raise CalendarCoverageError(f"No configured Market Date before {value.isoformat()}")
        return self.market_dates[position]

    def first_output_market_on_or_after(self, value: date) -> date | None:
        output = self.output_market_dates
        position = bisect_left(output, value)
        return output[position] if position < len(output) else None

    def previous_output_market_day(self, value: date) -> date:
        position = bisect_left(self.market_dates, value) - 1
        if position < 0:
            raise CalendarCoverageError(
                f"No previous configured Market Date for {value.isoformat()}"
            )
        return self.market_dates[position]

    def market_dates_between(self, first: date, last: date) -> tuple[date, ...]:
        return tuple(value for value in self.market_dates if first <= value <= last)


def eligible_delivery_days(
    first: date,
    last: date,
    method: FixingMethod,
    calendar: CalendarIndex,
) -> tuple[date, ...]:
    days = date_range(first, last)
    if method is FixingMethod.BRENT_HH:
        return tuple(day for day in days if calendar.is_market_day(day))
    return days


def normal_fixing_date(
    delivery_day: date,
    method: FixingMethod,
    calendar: CalendarIndex,
) -> date:
    if method is FixingMethod.WITHINDAY:
        return delivery_day
    if method is FixingMethod.DAY_AHEAD:
        return delivery_day - timedelta(days=1)
    if method in (FixingMethod.HEREN, FixingMethod.BRENT_HH):
        return calendar.previous_market_day(delivery_day)
    raise ValueError(f"Normal daily fixing date does not apply to {method}")


def delivery_month_slices(first: date, last: date) -> tuple[tuple[date, date, date], ...]:
    slices: list[tuple[date, date, date]] = []
    cursor = month_start(first)
    final = month_start(last)
    while cursor <= final:
        overlap_first = max(first, cursor)
        overlap_last = min(last, month_end(cursor))
        slices.append((cursor, overlap_first, overlap_last))
        cursor = add_months(cursor, 1)
    return tuple(slices)
