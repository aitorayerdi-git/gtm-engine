"""Small independent assertions run after each calculation stage."""

from __future__ import annotations

from collections import defaultdict
from datetime import date
from decimal import Decimal

from .decimal_utils import VOLUME_TOLERANCE, ZERO
from .models import (
    ExposureRow,
    FixingEvent,
    FixingRow,
    PnlRow,
    Severity,
    TradeEvent,
    ValidationItem,
)


def validate_fixing_conservation(
    fixing_events: tuple[FixingEvent, ...],
    trade_events: tuple[TradeEvent, ...],
    build_id: str,
) -> tuple[ValidationItem, ...]:
    scheduled: dict[tuple[str, date], Decimal] = defaultdict(lambda: ZERO)
    expected: dict[tuple[str, date], Decimal] = defaultdict(lambda: ZERO)
    for fixing_event in fixing_events:
        scheduled[(fixing_event.source_row_id, fixing_event.delivery_month)] += (
            fixing_event.fixing_volume
        )
    for trade_event in trade_events:
        expected[
            (trade_event.source_row_id, trade_event.delivery_month)
        ] += -trade_event.signed_volume

    issues: list[ValidationItem] = []
    for key, expected_volume in expected.items():
        actual = scheduled.get(key, ZERO)
        if abs(actual - expected_volume) > VOLUME_TOLERANCE:
            issues.append(
                ValidationItem(
                    build_id=build_id,
                    stage="Fixings",
                    severity=Severity.ERROR,
                    code="FIXING_VOLUME_NOT_CONSERVED",
                    message="Trade fixing schedule does not conserve signed monthly volume.",
                    source_row_id=key[0],
                    economic_key=f"{key[0]}|{key[1].isoformat()}",
                    actual=str(actual),
                    expected=str(expected_volume),
                    remediation="Correct the schedule allocation before publishing.",
                )
            )
    return tuple(issues)


def validate_output_invariants(
    fixings: tuple[FixingRow, ...],
    exposure: tuple[ExposureRow, ...],
    pnl: tuple[PnlRow, ...],
    build_id: str,
) -> tuple[ValidationItem, ...]:
    issues: list[ValidationItem] = []
    exposure_keys = [
        (
            row.market_date,
            row.book,
            row.underlying,
            row.delivery_month,
            row.trade_source,
            row.scenario,
        )
        for row in exposure
    ]
    if len(exposure_keys) != len(set(exposure_keys)):
        issues.append(
            ValidationItem(
                build_id=build_id,
                stage="Exposure",
                severity=Severity.ERROR,
                code="DUPLICATE_EXPOSURE_KEY",
                message="Exposure output contains a duplicate economic key.",
                remediation="Aggregate events by the declared output grain.",
            )
        )

    fixing_keys = [
        (
            row.fixing_date,
            row.applied_market_date,
            row.book,
            row.source_underlying,
            row.underlying,
            row.delivery_month,
            row.delivery_day,
            row.trade_source,
            row.scenario,
        )
        for row in fixings
    ]
    if len(fixing_keys) != len(set(fixing_keys)):
        issues.append(
            ValidationItem(
                build_id=build_id,
                stage="Fixings",
                severity=Severity.ERROR,
                code="DUPLICATE_FIXING_KEY",
                message="Fixings output contains a duplicate economic key.",
                remediation="Aggregate source events deterministically.",
            )
        )

    economic_fixing_by_key: dict[tuple[object, ...], Decimal] = defaultdict(lambda: ZERO)
    for fixing_row in fixings:
        key = (
            fixing_row.applied_market_date,
            fixing_row.book,
            fixing_row.underlying,
            fixing_row.delivery_month,
            fixing_row.trade_source,
            fixing_row.scenario,
        )
        economic_fixing_by_key[key] -= fixing_row.fixing_amount

    for pnl_row in pnl:
        pnl_key = (
            pnl_row.market_date,
            pnl_row.book,
            pnl_row.underlying,
            pnl_row.delivery_month,
            pnl_row.trade_source,
            pnl_row.scenario,
        )
        expected_fixing = economic_fixing_by_key.get(pnl_key, ZERO)
        if pnl_row.fixing_amount != expected_fixing:
            issues.append(
                ValidationItem(
                    build_id=build_id,
                    stage="P&L",
                    severity=Severity.ERROR,
                    code="PNL_FIXING_SIGN",
                    message=(
                        "P&L fixing contribution is not the inverse of the signed raw "
                        "fixing settlement."
                    ),
                    economic_key=(
                        f"{pnl_row.market_date.isoformat()}|{pnl_row.book}|"
                        f"{pnl_row.underlying}|{pnl_row.delivery_month or ''}"
                    ),
                    actual=str(pnl_row.fixing_amount),
                    expected=str(expected_fixing),
                    remediation="Apply the economic cash-flow sign exactly once in P&L.",
                )
            )

        expected_total = (
            pnl_row.delta_exposure_mtm
            + pnl_row.fixing_amount
            + pnl_row.logistical_costs
            + pnl_row.fees_and_optimizations
            + pnl_row.replication
        )
        if pnl_row.total_pnl != expected_total:
            issues.append(
                ValidationItem(
                    build_id=build_id,
                    stage="P&L",
                    severity=Severity.ERROR,
                    code="PNL_COMPONENT_SUM",
                    message="Total P&L does not equal the declared components.",
                    economic_key=(
                        f"{pnl_row.market_date.isoformat()}|{pnl_row.book}|"
                        f"{pnl_row.underlying}|{pnl_row.delivery_month or ''}"
                    ),
                    actual=str(pnl_row.total_pnl),
                    expected=str(expected_total),
                    remediation="Correct the component aggregation.",
                )
            )
    return tuple(issues)
