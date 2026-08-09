"""Fail-closed conversion of native currency amounts to EUR."""

from __future__ import annotations

from bisect import bisect_right
from collections import defaultdict
from datetime import date
from decimal import Decimal

from .canonicalize import normalize_text
from .models import FxRate, InputBundle, Severity, ValidationItem


class FxIndex:
    def __init__(self, rates: tuple[FxRate, ...]) -> None:
        grouped: dict[str, list[FxRate]] = defaultdict(list)
        for row in rates:
            grouped[normalize_text(row.currency)].append(row)
        self._rows = {
            currency: tuple(sorted(rows, key=lambda row: row.rate_date))
            for currency, rows in grouped.items()
        }

    def rate(self, currency: str, value_date: date) -> Decimal:
        if normalize_text(currency) == "EUR":
            return Decimal("1")
        rows = self._rows.get(normalize_text(currency), ())
        position = bisect_right([row.rate_date for row in rows], value_date) - 1
        if position < 0:
            raise KeyError(f"No {currency}/EUR rate on or before {value_date.isoformat()}")
        rate = rows[position].currency_per_eur
        if rate <= 0:
            raise ValueError(f"Non-positive {currency}/EUR rate for {value_date.isoformat()}")
        return rate

    def to_eur(self, amount: Decimal, currency: str, value_date: date) -> Decimal:
        return amount / self.rate(currency, value_date)


def validate_fx_requirements(
    bundle: InputBundle,
    requirements: set[tuple[date, str]],
    build_id: str,
) -> tuple[ValidationItem, ...]:
    index = FxIndex(bundle.fx_rates)
    issues: list[ValidationItem] = []
    for value_date, currency in sorted(requirements):
        if normalize_text(currency) == "EUR":
            continue
        try:
            index.rate(currency, value_date)
        except (KeyError, ValueError) as exc:
            issues.append(
                ValidationItem(
                    build_id=build_id,
                    stage="Preflight",
                    severity=Severity.ERROR,
                    code="MISSING_FX_RATE",
                    message=str(exc),
                    table="fx_rates",
                    economic_key=f"{value_date.isoformat()}|{currency}",
                    expected="one positive rate on the date or the latest prior date",
                    remediation="Supply the missing FX observation and rebuild.",
                )
            )
    return tuple(issues)
