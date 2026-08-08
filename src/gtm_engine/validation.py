"""Fail-closed structural and v0.3 policy validation."""

from __future__ import annotations

from collections import Counter
from collections.abc import Iterable
from datetime import date, timedelta
from decimal import Decimal
from itertools import pairwise

from .calendar import (
    CalendarCoverageError,
    CalendarIndex,
    add_months,
    delivery_month_slices,
    eligible_delivery_days,
    normal_fixing_date,
)
from .canonicalize import Registry, normalize_text
from .decimal_utils import is_material
from .models import (
    FixingMethod,
    InputBundle,
    Severity,
    SimulationStatus,
    Trade,
    TradeSource,
    ValidationItem,
)


def _item(
    build_id: str,
    code: str,
    message: str,
    *,
    severity: Severity = Severity.ERROR,
    table: str | None = None,
    source_row_id: str | None = None,
    actual: str | None = None,
    expected: str | None = None,
    remediation: str | None = None,
) -> ValidationItem:
    return ValidationItem(
        build_id=build_id,
        stage="Preflight",
        severity=severity,
        code=code,
        message=message,
        table=table,
        source_row_id=source_row_id,
        actual=actual,
        expected=expected,
        remediation=remediation,
    )


def included_trade(trade: Trade, bundle: InputBundle) -> bool:
    if trade.trade_date <= bundle.config.initial_market_date:
        return False
    if trade.trade_source is TradeSource.SIMULATION:
        return bundle.config.simulation_status is SimulationStatus.ON
    return trade.trade_source is TradeSource.ACTUAL


def _duplicates(values: Iterable[str]) -> set[str]:
    counts = Counter(values)
    return {value for value, count in counts.items() if count > 1}


def _validate_late_trade(
    trade: Trade,
    method: FixingMethod,
    calendar: CalendarIndex,
    build_id: str,
) -> list[ValidationItem]:
    issues: list[ValidationItem] = []
    try:
        if method is FixingMethod.MONTH_AHEAD:
            for delivery_month, _, _ in delivery_month_slices(trade.start_date, trade.end_date):
                previous_month = add_months(delivery_month, -1)
                remaining = tuple(
                    value
                    for value in calendar.market_dates_between(
                        previous_month, delivery_month - timedelta(days=1)
                    )
                    if value >= trade.trade_date
                )
                if not remaining:
                    issues.append(
                        _item(
                            build_id,
                            "LATE_MONTH_AHEAD_TRADE",
                            "No valid Month Ahead fixing opportunity remains.",
                            table="trades",
                            source_row_id=trade.source_row_id,
                            actual=trade.trade_date.isoformat(),
                            expected=(
                                "a configured fixing date on/after Trade Date for "
                                f"{delivery_month:%Y-%m}"
                            ),
                            remediation=(
                                "Correct or reject the trade; do not create a catch-up fixing."
                            ),
                        )
                    )
            return issues

        expired: list[tuple[date, date]] = []
        for delivery_day in eligible_delivery_days(
            trade.start_date, trade.end_date, method, calendar
        ):
            fixing_date = normal_fixing_date(delivery_day, method, calendar)
            if fixing_date < trade.trade_date:
                expired.append((delivery_day, fixing_date))
        if expired:
            delivery_day, fixing_date = expired[0]
            issues.append(
                _item(
                    build_id,
                    "LATE_FIXING_OPPORTUNITY",
                    "The trade contains quantity whose normal fixing opportunity has passed.",
                    table="trades",
                    source_row_id=trade.source_row_id,
                    actual=(
                        f"delivery_day={delivery_day.isoformat()}, "
                        f"fixing_date={fixing_date.isoformat()}, "
                        f"trade_date={trade.trade_date.isoformat()}"
                    ),
                    expected="Fixing Date >= Trade Date for every included delivery slice",
                    remediation=(
                        "Reject or correct the source trade; catch-up fixing is not allowed."
                    ),
                )
            )
    except CalendarCoverageError as exc:
        issues.append(
            _item(
                build_id,
                "CALENDAR_COVERAGE",
                str(exc),
                table="market_calendar",
                source_row_id=trade.source_row_id,
                remediation="Extend the configured calendar before rebuilding.",
            )
        )
    return issues


def validate_bundle(bundle: InputBundle, build_id: str) -> tuple[ValidationItem, ...]:
    issues: list[ValidationItem] = []
    config = bundle.config
    registry = Registry.from_bundle(bundle)

    if config.schema_version != "0.3.0" or config.policy_version != "0.3.0":
        issues.append(
            _item(
                build_id,
                "VERSION_INCOMPATIBLE",
                "Input schema and policy must match the v0.3 engine contract.",
                actual=f"schema={config.schema_version}, policy={config.policy_version}",
                expected="schema=0.3.0, policy=0.3.0",
            )
        )
    if config.historical_end_date < config.historical_start_date:
        issues.append(
            _item(
                build_id,
                "CONTROL_DATE_ORDER",
                "Historical End Date precedes Historical Start Date.",
            )
        )
    if config.historical_end_date <= config.initial_market_date:
        issues.append(
            _item(
                build_id,
                "CONTROL_DATE_ORDER",
                "Historical End Date must be after Initial Market Date.",
            )
        )

    book_names = [normalize_text(row.book) for row in bundle.books if row.active]
    for duplicate in sorted(_duplicates(book_names)):
        issues.append(_item(build_id, "DUPLICATE_BOOK", f"Duplicate active BOOK: {duplicate}"))
    underlying_names = [
        normalize_text(row.source_underlying) for row in bundle.underlyings if row.active
    ]
    for duplicate in sorted(_duplicates(underlying_names)):
        issues.append(
            _item(build_id, "DUPLICATE_UNDERLYING", f"Duplicate active Underlying: {duplicate}")
        )

    calendar_dates = [row.date for row in bundle.calendar]
    if calendar_dates != sorted(calendar_dates):
        issues.append(
            _item(build_id, "CALENDAR_NON_MONOTONIC", "Calendar rows are not strictly ordered.")
        )
    if len(set(calendar_dates)) != len(calendar_dates):
        issues.append(_item(build_id, "CALENDAR_DUPLICATE_DATE", "Calendar dates are not unique."))
    for previous, current in pairwise(calendar_dates):
        if current != previous + timedelta(days=1):
            issues.append(
                _item(
                    build_id,
                    "CALENDAR_GAP",
                    "Calendar must contain one row for every consecutive calendar date.",
                    table="market_calendar",
                    actual=f"{previous.isoformat()} -> {current.isoformat()}",
                    expected=(previous + timedelta(days=1)).isoformat(),
                    remediation="Add the missing calendar date rows and rebuild.",
                )
            )
            break
    calendar = CalendarIndex(bundle.calendar, config)
    if not calendar.output_market_dates:
        issues.append(
            _item(build_id, "MARKET_AXIS_EMPTY", "No output Market Dates exist in the model range.")
        )

    for table_name, rows in (
        ("initial_exposure", bundle.initial_exposure),
        ("initial_pnl", bundle.initial_pnl),
        ("trades", bundle.trades),
        ("operating_flows", bundle.operating_flows),
    ):
        source_ids = [row.source_row_id for row in rows]
        for duplicate in sorted(_duplicates(source_ids)):
            issues.append(
                _item(
                    build_id,
                    "DUPLICATE_SOURCE_ROW_ID",
                    f"Duplicate source_row_id {duplicate}.",
                    table=table_name,
                    source_row_id=duplicate,
                )
            )

    active_books = set(registry.books)
    pnl_books = [normalize_text(row.book) for row in bundle.initial_pnl]
    if Counter(pnl_books) != Counter(active_books):
        issues.append(
            _item(
                build_id,
                "INITIAL_PNL_BOOK_SET",
                "Initial P&L must contain exactly one row for every active BOOK.",
                table="initial_pnl",
                actual=", ".join(sorted(pnl_books)),
                expected=", ".join(sorted(active_books)),
            )
        )

    for initial_position in bundle.initial_exposure:
        if initial_position.initial_market_date != config.initial_market_date:
            issues.append(
                _item(
                    build_id,
                    "INITIAL_DATE_MISMATCH",
                    "Initial Exposure date differs from the bundle cut-off.",
                    table="initial_exposure",
                    source_row_id=initial_position.source_row_id,
                )
            )

    initial_keys = [
        (
            normalize_text(row.book),
            normalize_text(row.underlying),
            row.delivery_month,
        )
        for row in bundle.initial_exposure
    ]
    if len(initial_keys) != len(set(initial_keys)):
        issues.append(
            _item(
                build_id,
                "DUPLICATE_INITIAL_EXPOSURE_KEY",
                "Initial Exposure contains repeated economic keys; calculation aggregates them.",
                severity=Severity.WARNING,
                table="initial_exposure",
            )
        )

    for initial_position in bundle.initial_exposure:
        if registry.book(initial_position.book) is None:
            issues.append(
                _item(
                    build_id,
                    "UNKNOWN_BOOK",
                    f"Unknown or inactive BOOK: {initial_position.book}",
                    table="initial_exposure",
                    source_row_id=initial_position.source_row_id,
                )
            )
        if registry.underlying(initial_position.underlying) is None:
            issues.append(
                _item(
                    build_id,
                    "UNKNOWN_UNDERLYING",
                    f"Unknown or inactive Underlying: {initial_position.underlying}",
                    table="initial_exposure",
                    source_row_id=initial_position.source_row_id,
                )
            )

    for opening_pnl in bundle.initial_pnl:
        if opening_pnl.initial_market_date != config.initial_market_date:
            issues.append(
                _item(
                    build_id,
                    "INITIAL_DATE_MISMATCH",
                    "Initial P&L date differs from the bundle cut-off.",
                    table="initial_pnl",
                    source_row_id=opening_pnl.source_row_id,
                )
            )
        if registry.book(opening_pnl.book) is None:
            issues.append(
                _item(
                    build_id,
                    "UNKNOWN_BOOK",
                    f"Unknown or inactive Initial P&L BOOK: {opening_pnl.book}",
                    table="initial_pnl",
                    source_row_id=opening_pnl.source_row_id,
                )
            )

    for trade in bundle.trades:
        if not included_trade(trade, bundle):
            continue
        book = registry.book(trade.book)
        underlying = registry.underlying(trade.underlying)
        if book is None:
            issues.append(
                _item(
                    build_id,
                    "UNKNOWN_BOOK",
                    f"Unknown or inactive BOOK: {trade.book}",
                    table="trades",
                    source_row_id=trade.source_row_id,
                )
            )
        if underlying is None:
            issues.append(
                _item(
                    build_id,
                    "UNKNOWN_UNDERLYING",
                    f"Unknown or inactive Underlying: {trade.underlying}",
                    table="trades",
                    source_row_id=trade.source_row_id,
                )
            )
        if trade.end_date < trade.start_date:
            issues.append(
                _item(
                    build_id,
                    "INVALID_DELIVERY_RANGE",
                    "End Date precedes Start Date.",
                    table="trades",
                    source_row_id=trade.source_row_id,
                )
            )
        if trade.daily_qty < Decimal("0"):
            issues.append(
                _item(
                    build_id,
                    "NEGATIVE_DAILY_QTY",
                    "Daily Qty must be a non-negative magnitude; Side supplies sign.",
                    table="trades",
                    source_row_id=trade.source_row_id,
                )
            )
        elif not is_material(trade.daily_qty, Decimal("0")):
            issues.append(
                _item(
                    build_id,
                    "ZERO_DAILY_QTY",
                    "Explicit zero Daily Qty creates no economic event.",
                    severity=Severity.WARNING,
                    table="trades",
                    source_row_id=trade.source_row_id,
                )
            )
        if trade.trade_source is TradeSource.SIMULATION and not trade.scenario:
            issues.append(
                _item(
                    build_id,
                    "SIMULATION_SCENARIO_MISSING",
                    "Simulation trade requires a Scenario.",
                    table="trades",
                    source_row_id=trade.source_row_id,
                )
            )
        if trade.trade_source is TradeSource.ACTUAL and trade.scenario:
            issues.append(
                _item(
                    build_id,
                    "ACTUAL_SCENARIO_IGNORED",
                    "Actual trade should not carry a Scenario.",
                    severity=Severity.WARNING,
                    table="trades",
                    source_row_id=trade.source_row_id,
                )
            )
        if underlying is not None and trade.end_date >= trade.start_date:
            issues.extend(_validate_late_trade(trade, underlying.fixing_method, calendar, build_id))

    for flow in bundle.operating_flows:
        if registry.book(flow.book) is None:
            issues.append(
                _item(
                    build_id,
                    "UNKNOWN_BOOK",
                    f"Unknown or inactive operating-flow BOOK: {flow.book}",
                    table="operating_flows",
                    source_row_id=flow.source_row_id,
                )
            )

    operating_keys = [(row.market_date, normalize_text(row.book)) for row in bundle.operating_flows]
    if len(operating_keys) != len(set(operating_keys)):
        issues.append(
            _item(
                build_id,
                "DUPLICATE_OPERATING_FLOW",
                "Operating flows contain more than one row for a Market Date and BOOK.",
                table="operating_flows",
                remediation="Aggregate the direct source components to one BOOK/date row.",
            )
        )

    curve_keys = [
        (row.market_date, normalize_text(row.underlying), row.delivery_month)
        for row in bundle.curve_prices
    ]
    if len(set(curve_keys)) != len(curve_keys):
        issues.append(_item(build_id, "DUPLICATE_CURVE_PRICE", "Duplicate curve-price key."))
    fixing_keys = [
        (row.price_lookup_date, normalize_text(row.underlying)) for row in bundle.fixing_prices
    ]
    if len(set(fixing_keys)) != len(fixing_keys):
        issues.append(_item(build_id, "DUPLICATE_FIXING_PRICE", "Duplicate fixing-price key."))

    return tuple(issues)


def has_errors(items: Iterable[ValidationItem]) -> bool:
    return any(item.severity is Severity.ERROR for item in items)
