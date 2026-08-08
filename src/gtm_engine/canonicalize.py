"""Canonical text and SETUP-driven mapping registry."""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

from .models import BookConfig, FixingMethod, InputBundle, UnderlyingConfig


def normalize_text(value: str) -> str:
    decomposed = unicodedata.normalize("NFKD", value.strip())
    unaccented = "".join(char for char in decomposed if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", unaccented).upper()


_METHOD_ALIASES: dict[str, FixingMethod] = {
    "WITHINDAY": FixingMethod.WITHINDAY,
    "WITHIN DAY": FixingMethod.WITHINDAY,
    "DAY AHEAD": FixingMethod.DAY_AHEAD,
    "DAY_AHEAD": FixingMethod.DAY_AHEAD,
    "METODOLOGIA HEREN": FixingMethod.HEREN,
    "HEREN": FixingMethod.HEREN,
    "MONTH AHEAD": FixingMethod.MONTH_AHEAD,
    "MONTH_AHEAD": FixingMethod.MONTH_AHEAD,
    "BRENT & HH": FixingMethod.BRENT_HH,
    "BRENT_HH": FixingMethod.BRENT_HH,
}


def canonical_method(value: str | FixingMethod) -> FixingMethod:
    if isinstance(value, FixingMethod):
        return value
    normalized = normalize_text(value)
    try:
        return _METHOD_ALIASES[normalized]
    except KeyError as exc:
        raise ValueError(f"Unsupported fixing method: {value}") from exc


@dataclass(frozen=True)
class Registry:
    books: dict[str, BookConfig]
    underlyings: dict[str, UnderlyingConfig]

    @classmethod
    def from_bundle(cls, bundle: InputBundle) -> Registry:
        books: dict[str, BookConfig] = {}
        underlyings: dict[str, UnderlyingConfig] = {}
        for book_row in bundle.books:
            if book_row.active:
                books.setdefault(normalize_text(book_row.book), book_row)
        for underlying_row in bundle.underlyings:
            if underlying_row.active:
                underlyings.setdefault(
                    normalize_text(underlying_row.source_underlying), underlying_row
                )
        return cls(books=books, underlyings=underlyings)

    def book(self, value: str) -> BookConfig | None:
        return self.books.get(normalize_text(value))

    def underlying(self, value: str) -> UnderlyingConfig | None:
        return self.underlyings.get(normalize_text(value))

    def canonical_book(self, value: str) -> str:
        row = self.book(value)
        if row is None:
            raise KeyError(value)
        return row.book

    def canonical_underlying(self, value: str) -> str:
        row = self.underlying(value)
        if row is None:
            raise KeyError(value)
        return row.canonical_underlying
