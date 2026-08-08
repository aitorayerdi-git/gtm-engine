"""Exact decimal helpers used by every economic calculation."""

from __future__ import annotations

from decimal import Decimal, getcontext
from typing import Any

getcontext().prec = 34

ZERO = Decimal("0")
MATERIALITY = Decimal("0.0000001")
VOLUME_TOLERANCE = Decimal("0.000001")
PRICE_TOLERANCE = Decimal("0.00000001")
PNL_TOLERANCE = Decimal("0.01")


def to_decimal(value: Any) -> Decimal:
    """Convert input through text so binary floats do not leak into arithmetic."""
    if isinstance(value, Decimal):
        result = value
    elif isinstance(value, bool):
        raise ValueError("Boolean is not a valid economic number")
    else:
        result = Decimal(str(value).strip())
    if not result.is_finite():
        raise ValueError("Economic numbers must be finite")
    return result


def is_material(value: Decimal, threshold: Decimal = MATERIALITY) -> bool:
    return abs(value) > threshold


def decimal_text(value: Decimal) -> str:
    """Return a deterministic non-exponential decimal representation."""
    if value == ZERO:
        return "0"
    return format(value.normalize(), "f")
