from __future__ import annotations

import json
from datetime import date
from pathlib import Path

import pytest

from gtm_engine.cli import main
from gtm_engine.io import BundleLoadError, _write_models, load_bundle, publish_result
from gtm_engine.models import (
    BookConfig,
    BuildStatus,
    CurvePrice,
    FixingPrice,
    InitialExposure,
    InitialPnl,
    InputBundle,
    MarketCalendarDay,
    OperatingFlow,
    Side,
    Trade,
    TradeSource,
    UnderlyingConfig,
)
from gtm_engine.pipeline import build

from .helpers import D, base_bundle, with_prices


def _write_bundle(root: Path, bundle: InputBundle) -> None:
    root.mkdir()
    (root / "bundle.json").write_text(bundle.config.model_dump_json(indent=2), encoding="utf-8")
    tables = (
        ("books.csv", bundle.books, BookConfig),
        ("underlyings.csv", bundle.underlyings, UnderlyingConfig),
        ("market_calendar.csv", bundle.calendar, MarketCalendarDay),
        ("initial_exposure.csv", bundle.initial_exposure, InitialExposure),
        ("initial_pnl.csv", bundle.initial_pnl, InitialPnl),
        ("trades.csv", bundle.trades, Trade),
        ("curve_prices.csv", bundle.curve_prices, CurvePrice),
        ("fixing_prices.csv", bundle.fixing_prices, FixingPrice),
        ("operating_flows.csv", bundle.operating_flows, OperatingFlow),
    )
    for filename, rows, model in tables:
        _write_models(root / filename, rows, model)


def test_bundle_round_trip_and_atomic_publication(tmp_path: Path) -> None:
    input_root = tmp_path / "input"
    output_root = tmp_path / "output"
    original = with_prices(base_bundle())
    _write_bundle(input_root, original)

    loaded = load_bundle(input_root)
    assert loaded.config == original.config
    assert loaded.books == original.books
    assert set(loaded.input_hashes) == {
        "bundle.json",
        "books.csv",
        "underlyings.csv",
        "market_calendar.csv",
        "initial_exposure.csv",
        "initial_pnl.csv",
        "trades.csv",
        "curve_prices.csv",
        "fixing_prices.csv",
        "operating_flows.csv",
    }

    verified = build(loaded)
    published, destination = publish_result(verified, output_root)
    assert published.manifest.status is BuildStatus.PUBLISHED
    assert destination.parent == output_root / "runs"
    assert not any(path.name.startswith(".staging-") for path in destination.parent.iterdir())
    assert (output_root / "LATEST").read_text(encoding="utf-8").strip() == str(destination)
    assert {path.name for path in destination.iterdir()} == {
        "build_manifest.json",
        "validation.csv",
        "fixings.csv",
        "exposure.csv",
        "pnl.csv",
        "cumulative_pnl.csv",
        "event_ledger.csv",
    }
    manifest = json.loads((destination / "build_manifest.json").read_text(encoding="utf-8"))
    assert manifest["status"] == "PUBLISHED"
    assert manifest["input_row_counts"]["books"] == 1
    assert set(manifest["output_hashes"]) == {
        "validation",
        "fixings",
        "exposure",
        "pnl",
        "cumulative_pnl",
        "event_ledger",
    }


def test_failed_build_is_retained_but_never_becomes_latest(tmp_path: Path) -> None:
    failed = build(
        base_bundle(
            initial_exposure=(
                InitialExposure(
                    initial_market_date=date(2026, 1, 2),
                    book="BOOK1",
                    underlying="GAS",
                    delivery_month=date(2026, 1, 1),
                    exposure_volume=D("100"),
                    source_row_id="OPEN-1",
                ),
            )
        )
    )
    assert failed.manifest.status is BuildStatus.FAILED
    retained, destination = publish_result(failed, tmp_path)
    assert retained is failed
    assert destination.parent == tmp_path / "failed"
    assert not (tmp_path / "LATEST").exists()
    assert {path.name for path in destination.iterdir()} == {
        "build_manifest.json",
        "validation.csv",
    }


def test_load_bundle_reports_missing_and_invalid_inputs(tmp_path: Path) -> None:
    with pytest.raises(BundleLoadError, match=r"bundle\.json"):
        load_bundle(tmp_path)

    invalid_json = tmp_path / "invalid-json"
    invalid_json.mkdir()
    (invalid_json / "bundle.json").write_text("{", encoding="utf-8")
    with pytest.raises(BundleLoadError, match=r"Cannot load bundle\.json"):
        load_bundle(invalid_json)

    missing_table = tmp_path / "missing-table"
    _write_bundle(missing_table, with_prices(base_bundle()))
    (missing_table / "books.csv").unlink()
    with pytest.raises(BundleLoadError, match=r"books\.csv"):
        load_bundle(missing_table)

    invalid_table = tmp_path / "invalid-table"
    _write_bundle(invalid_table, with_prices(base_bundle()))
    (invalid_table / "books.csv").write_text("book,active\nBOOK1,not-a-bool\n", encoding="utf-8")
    with pytest.raises(BundleLoadError, match=r"Cannot load books\.csv"):
        load_bundle(invalid_table)


def test_blank_trade_quantity_reports_csv_row_and_source_id(tmp_path: Path) -> None:
    input_root = tmp_path / "blank-quantity"
    trade = Trade(
        source_row_id="BLANK-QTY-ROW",
        trade_date=date(2026, 1, 5),
        book="BOOK1",
        underlying="GAS",
        side=Side.BUY,
        start_date=date(2026, 1, 7),
        end_date=date(2026, 1, 7),
        daily_qty=D("1"),
        execution_price=D("20"),
        trade_source=TradeSource.ACTUAL,
        scenario=None,
    )
    _write_bundle(input_root, base_bundle(trades=(trade,)))
    (input_root / "trades.csv").write_text(
        "source_row_id,trade_date,book,underlying,side,start_date,end_date,daily_qty,"
        "execution_price,trade_source,scenario\n"
        "BLANK-QTY-ROW,2026-01-05,BOOK1,GAS,BUY,2026-01-07,2026-01-07,,20,ACTUAL,\n",
        encoding="utf-8",
    )

    with pytest.raises(BundleLoadError) as caught:
        load_bundle(input_root)
    message = str(caught.value)
    assert "trades.csv at CSV row 2" in message
    assert "source_row_id=BLANK-QTY-ROW" in message
    assert "daily_qty" in message


def test_cli_publishes_success_and_reports_load_failure(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    input_root = tmp_path / "input"
    output_root = tmp_path / "output"
    _write_bundle(input_root, with_prices(base_bundle()))

    assert main(["build", "--input", str(input_root), "--output", str(output_root)]) == 0
    success = json.loads(capsys.readouterr().out)
    assert success["status"] == "PUBLISHED"
    assert Path(success["output"]).is_dir()

    assert main(["build", "--input", str(tmp_path / "absent"), "--output", str(output_root)]) == 2
    failure = json.loads(capsys.readouterr().out)
    assert failure["status"] == "FAILED"
    assert failure["stage"] == "Load"
