"""Versioned input, output, audit, and manifest contracts."""

from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from enum import StrEnum
from typing import Any

from pydantic import BaseModel, ConfigDict, Field, ValidationInfo, field_validator

from .decimal_utils import MATERIALITY

SCHEMA_VERSION = "0.3.0"
POLICY_VERSION = "0.3.0"
ENGINE_VERSION = "0.3.0"


class StrictModel(BaseModel):
    model_config = ConfigDict(
        extra="forbid",
        frozen=True,
        allow_inf_nan=False,
        str_strip_whitespace=True,
    )


class Side(StrEnum):
    BUY = "BUY"
    SELL = "SELL"


class TradeSource(StrEnum):
    INITIAL = "INITIAL"
    ACTUAL = "ACTUAL"
    SIMULATION = "SIMULATION"


class SimulationStatus(StrEnum):
    ON = "ON"
    OFF = "OFF"


class FixingMethod(StrEnum):
    WITHINDAY = "WITHINDAY"
    DAY_AHEAD = "DAY_AHEAD"
    HEREN = "HEREN"
    MONTH_AHEAD = "MONTH_AHEAD"
    BRENT_HH = "BRENT_HH"


class PriceDateBasis(StrEnum):
    FIXING_DATE = "FIXING_DATE"
    DELIVERY_DAY = "DELIVERY_DAY"


class Severity(StrEnum):
    ERROR = "ERROR"
    WARNING = "WARNING"
    INFO = "INFO"


class BuildStatus(StrEnum):
    CREATED = "CREATED"
    VALIDATED = "VALIDATED"
    CALCULATED = "CALCULATED"
    VERIFIED = "VERIFIED"
    PUBLISHED = "PUBLISHED"
    FAILED = "FAILED"


class EventType(StrEnum):
    INITIAL_OPEN = "INITIAL_OPEN"
    TRADE = "TRADE"
    FIXING = "FIXING"


class BuildConfig(StrictModel):
    model_id: str = "GTM"
    schema_version: str = SCHEMA_VERSION
    policy_version: str = POLICY_VERSION
    engine_version: str = ENGINE_VERSION
    initial_market_date: date
    historical_start_date: date
    historical_end_date: date
    simulation_status: SimulationStatus = SimulationStatus.OFF
    timezone: str = "Europe/Madrid"
    logistics_sign: Decimal = Decimal("-1")
    materiality: Decimal = MATERIALITY


class BookConfig(StrictModel):
    book: str
    active: bool = True


class UnderlyingConfig(StrictModel):
    source_underlying: str
    canonical_underlying: str
    fixing_method: FixingMethod
    unit: str
    currency: str = "EUR"
    curve_underlying: str | None = Field(default=None, validate_default=True)
    fixing_price_underlying: str | None = Field(default=None, validate_default=True)
    fixing_price_basis: PriceDateBasis = PriceDateBasis.FIXING_DATE
    include_fixing_in_pnl: bool = True
    active: bool = True
    current_month_uses_next_curve: bool = False

    @field_validator("curve_underlying", mode="before")
    @classmethod
    def populate_curve_name(cls, value: str | None, info: ValidationInfo) -> str:
        return value or str(info.data["canonical_underlying"])

    @field_validator("fixing_price_underlying", mode="before")
    @classmethod
    def populate_fixing_price_name(cls, value: str | None, info: ValidationInfo) -> str:
        return value or str(info.data["source_underlying"])


class MarketCalendarDay(StrictModel):
    date: date
    is_market_day: bool
    calendar_id: str = "GTM"


class InitialExposure(StrictModel):
    initial_market_date: date
    book: str
    underlying: str
    delivery_month: date
    exposure_volume: Decimal
    source_row_id: str

    @field_validator("delivery_month")
    @classmethod
    def month_start(cls, value: date) -> date:
        if value.day != 1:
            raise ValueError("delivery_month must be the first calendar day")
        return value


class InitialPnl(StrictModel):
    initial_market_date: date
    book: str
    amount: Decimal
    source_row_id: str
    comment: str | None = None


class Trade(StrictModel):
    source_row_id: str
    trade_date: date
    book: str
    underlying: str
    side: Side
    start_date: date
    end_date: date
    daily_qty: Decimal
    execution_price: Decimal
    trade_source: TradeSource
    scenario: str | None = None

    @property
    def signed_daily_qty(self) -> Decimal:
        return self.daily_qty if self.side is Side.BUY else -self.daily_qty


class DeliveryElection(StrictModel):
    decision_date: date
    book: str
    underlying: str
    side: Side
    start_date: date
    end_date: date
    delivery_daily_qty: Decimal
    unit: str
    trade_source: TradeSource = TradeSource.ACTUAL
    scenario: str | None = None
    source_row_id: str
    comment: str | None = None


class CurvePrice(StrictModel):
    market_date: date
    underlying: str
    delivery_month: date
    curve_price: Decimal
    currency: str
    unit: str
    source_id: str
    source_as_of: datetime | None = None

    @field_validator("delivery_month")
    @classmethod
    def month_start(cls, value: date) -> date:
        if value.day != 1:
            raise ValueError("delivery_month must be the first calendar day")
        return value


class FixingPrice(StrictModel):
    price_lookup_date: date
    underlying: str
    fixing_price: Decimal
    currency: str
    unit: str
    source_id: str
    source_as_of: datetime | None = None


class FxRate(StrictModel):
    rate_date: date
    currency: str
    currency_per_eur: Decimal
    source_id: str
    source_as_of: datetime | None = None


class OperatingFlow(StrictModel):
    market_date: date
    book: str
    logistics_source_amount: Decimal = Decimal("0")
    fees_and_optimizations: Decimal = Decimal("0")
    replication: Decimal = Decimal("0")
    source_row_id: str


class InputBundle(StrictModel):
    config: BuildConfig
    books: tuple[BookConfig, ...]
    underlyings: tuple[UnderlyingConfig, ...]
    calendar: tuple[MarketCalendarDay, ...]
    initial_exposure: tuple[InitialExposure, ...]
    initial_pnl: tuple[InitialPnl, ...]
    trades: tuple[Trade, ...] = ()
    delivery_elections: tuple[DeliveryElection, ...] = ()
    curve_prices: tuple[CurvePrice, ...] = ()
    fixing_prices: tuple[FixingPrice, ...] = ()
    fx_rates: tuple[FxRate, ...] = ()
    operating_flows: tuple[OperatingFlow, ...] = ()
    input_hashes: dict[str, str] = Field(default_factory=dict)


class ValidationItem(StrictModel):
    build_id: str
    stage: str
    severity: Severity
    code: str
    message: str
    table: str | None = None
    source_row_id: str | None = None
    economic_key: str | None = None
    actual: str | None = None
    expected: str | None = None
    remediation: str | None = None


class FixingEvent(StrictModel):
    event_id: str
    fixing_date: date
    applied_market_date: date | None
    price_lookup_date: date
    book: str
    source_underlying: str
    underlying: str
    pricing_underlying: str
    delivery_month: date
    delivery_day: date
    fixing_volume: Decimal
    trade_source: TradeSource
    scenario: str | None
    source_row_id: str


class TradeEvent(StrictModel):
    event_id: str
    economic_date: date
    applied_market_date: date | None
    book: str
    source_underlying: str
    underlying: str
    delivery_month: date
    signed_volume: Decimal
    trade_source: TradeSource
    scenario: str | None
    source_row_id: str
    execution_price: Decimal


class FixingRow(StrictModel):
    fixing_date: date
    applied_market_date: date
    price_lookup_date: date
    book: str
    source_underlying: str
    underlying: str
    pricing_underlying: str
    delivery_month: date
    delivery_day: date
    fixing_volume: Decimal
    fixing_price: Decimal
    fixing_amount: Decimal
    currency: str
    trade_source: TradeSource
    scenario: str | None
    simulation_status: SimulationStatus
    build_id: str


class ExposureRow(StrictModel):
    market_date: date
    previous_market_date: date
    book: str
    underlying: str
    delivery_month: date
    curve_delivery_month: date
    exposure_volume: Decimal
    curve_price: Decimal | None
    exposure_mtm: Decimal
    currency: str
    trade_source: TradeSource
    scenario: str | None
    simulation_status: SimulationStatus
    is_explicit_closure: bool
    build_id: str


class PnlRow(StrictModel):
    market_date: date
    previous_market_date: date
    book: str
    underlying: str
    delivery_month: date | None
    exposure_mtm: Decimal
    gross_delta_exposure_mtm: Decimal
    trade_entry_adjustment: Decimal
    delta_exposure_mtm: Decimal
    fixing_amount: Decimal
    logistical_costs: Decimal
    fees_and_optimizations: Decimal
    replication: Decimal
    total_pnl: Decimal
    trade_source: TradeSource | None
    scenario: str | None
    simulation_status: SimulationStatus
    build_id: str


class CumulativePnlRow(StrictModel):
    market_date: date
    previous_market_date: date
    book: str
    initial_pnl: Decimal
    daily_pnl: Decimal
    cumulative_pnl: Decimal
    simulation_status: SimulationStatus
    build_id: str


class EventLedgerRow(StrictModel):
    event_id: str
    event_type: EventType
    economic_date: date
    applied_market_date: date | None
    book: str
    underlying: str
    delivery_month: date
    signed_volume_change: Decimal
    trade_source: TradeSource
    scenario: str | None
    source_row_id: str


class BuildManifest(StrictModel):
    run_id: str
    build_id: str
    calculation_fingerprint: str
    status: BuildStatus
    started_at: datetime
    finished_at: datetime
    timezone: str
    engine_version: str
    schema_version: str
    policy_version: str
    simulation_status: SimulationStatus
    input_hashes: dict[str, str]
    input_row_counts: dict[str, int]
    output_hashes: dict[str, str]
    row_counts: dict[str, int]
    component_totals: dict[str, str]
    validation_counts: dict[str, int]
    elapsed_seconds: float
    peak_memory_bytes: int
    runtime_platform: str
    python_version: str
    machine_architecture: str
    failure_stage: str | None = None


class BuildResult(StrictModel):
    manifest: BuildManifest
    validation: tuple[ValidationItem, ...]
    fixings: tuple[FixingRow, ...] = ()
    exposure: tuple[ExposureRow, ...] = ()
    pnl: tuple[PnlRow, ...] = ()
    cumulative_pnl: tuple[CumulativePnlRow, ...] = ()
    event_ledger: tuple[EventLedgerRow, ...] = ()


def jsonable(value: StrictModel | list[Any] | tuple[Any, ...] | dict[str, Any]) -> Any:
    if isinstance(value, StrictModel):
        return value.model_dump(mode="json")
    if isinstance(value, (list, tuple)):
        return [jsonable(item) for item in value]
    if isinstance(value, dict):
        return {key: jsonable(item) for key, item in value.items()}
    return value
