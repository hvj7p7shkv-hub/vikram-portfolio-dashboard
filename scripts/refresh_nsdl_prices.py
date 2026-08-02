#!/usr/bin/env python3
"""
Refresh an NSDL holdings CSV with Yahoo Finance prices.

If Yahoo does not return a quote, the existing NSDL snapshot values are kept.
"""

from __future__ import annotations

import argparse
import contextlib
import datetime as dt
import io
import math
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

try:
    import yfinance as yf
except ImportError:  # pragma: no cover
    yf = None


NUMERIC_COLUMNS = [
    "Free Qty",
    "Pledged Qty",
    "Quantity",
    "LTP",
    "Day Change",
    "Day Change %",
    "Day P&L",
    "Current Value",
    "Weight %",
]


def clean_number(value: object) -> float:
    text = str(value).replace(",", "").replace("%", "").replace("+", "").strip()
    try:
        return float(text)
    except ValueError:
        return math.nan


def clean_numeric_columns(data: pd.DataFrame) -> pd.DataFrame:
    for column in NUMERIC_COLUMNS:
        if column in data.columns:
            data[column] = data[column].map(clean_number)
    return data


def numeric_series(frame: pd.DataFrame, column: str) -> pd.Series:
    if frame is None or frame.empty:
        return pd.Series(dtype=float)

    values: pd.Series | pd.DataFrame | None = None
    if isinstance(frame.columns, pd.MultiIndex):
        if column in frame.columns.get_level_values(0):
            values = frame.xs(column, axis=1, level=0)
        elif column in frame.columns.get_level_values(1):
            values = frame.xs(column, axis=1, level=1)
    elif column in frame.columns:
        values = frame[column]

    if values is None:
        return pd.Series(dtype=float)
    if isinstance(values, pd.DataFrame):
        for subcolumn in values.columns:
            series = pd.to_numeric(values[subcolumn], errors="coerce").dropna()
            if not series.empty:
                return series.astype(float)
        return pd.Series(dtype=float)
    return pd.to_numeric(values, errors="coerce").dropna().astype(float)


def quote_from_frame(frame: pd.DataFrame | None) -> dict[str, object] | None:
    if frame is None or frame.empty:
        return None
    close = numeric_series(frame, "Close")
    if close.empty:
        return None
    last = float(close.iloc[-1])
    previous = float(close.iloc[-2]) if len(close) >= 2 else last
    price_date = close.index[-1]
    if hasattr(price_date, "date"):
        price_date = price_date.date().isoformat()
    return {"last": last, "previous": previous, "date": str(price_date)}


def extract_frame(downloaded: pd.DataFrame, ticker: str, batch_size: int) -> pd.DataFrame | None:
    if downloaded is None or downloaded.empty:
        return None
    if isinstance(downloaded.columns, pd.MultiIndex):
        level_0 = set(downloaded.columns.get_level_values(0))
        level_1 = set(downloaded.columns.get_level_values(1))
        if ticker in level_0:
            return downloaded[ticker].dropna(how="all")
        if ticker in level_1:
            return downloaded.xs(ticker, axis=1, level=1).dropna(how="all")
        return None
    if batch_size == 1:
        return downloaded.dropna(how="all")
    return None


def fallback_symbols(symbol: str) -> list[str]:
    symbols = [symbol]
    if symbol.endswith(".NS"):
        symbols.append(symbol[:-3] + ".BO")
    return symbols


def download_batch(tickers: list[str]) -> dict[str, dict[str, object]]:
    if yf is None or not tickers:
        return {}
    try:
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            downloaded = yf.download(
                tickers=tickers,
                period="5d",
                interval="1d",
                group_by="ticker",
                auto_adjust=False,
                progress=False,
                threads=True,
                timeout=30,
            )
    except Exception:
        return {}

    quotes: dict[str, dict[str, object]] = {}
    for ticker in tickers:
        quote = quote_from_frame(extract_frame(downloaded, ticker, len(tickers)))
        if quote:
            quotes[ticker] = quote
    return quotes


def download_single(ticker: str) -> dict[str, object] | None:
    if yf is None:
        return None
    try:
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            downloaded = yf.download(
                tickers=ticker,
                period="5d",
                interval="1d",
                auto_adjust=False,
                progress=False,
                timeout=20,
            )
    except Exception:
        return None
    return quote_from_frame(downloaded)


def download_quotes(primary_symbols: list[str]) -> dict[str, dict[str, object]]:
    quotes: dict[str, dict[str, object]] = {}
    unique = sorted({symbol for symbol in primary_symbols if str(symbol).strip()})
    batch_size = 45
    for start in range(0, len(unique), batch_size):
        quotes.update(download_batch(unique[start : start + batch_size]))

    for symbol in unique:
        if symbol in quotes:
            continue
        for fallback in fallback_symbols(symbol):
            quote = download_single(fallback)
            if quote:
                quote["downloaded_ticker"] = fallback
                quotes[symbol] = quote
                break
    return quotes


def recalculate_row(row: pd.Series, quote: dict[str, object] | None) -> pd.Series:
    quantity = clean_number(row.get("Quantity"))
    current_ltp = clean_number(row.get("LTP"))
    current_day_change = clean_number(row.get("Day Change"))

    if quote:
        ltp = float(quote["last"])
        previous = float(quote["previous"])
        row["LTP"] = ltp
        row["Day Change"] = ltp - previous
        row["Day Change %"] = (ltp / previous - 1) * 100 if previous else math.nan
        row["Price Source"] = str(quote.get("downloaded_ticker") or row.get("Yahoo Ticker") or "")
        row["Price Date"] = str(quote.get("date") or "")
    else:
        ltp = current_ltp
        row["Price Source"] = "NSDL fallback"
        row["Price Date"] = ""

    day_change = clean_number(row.get("Day Change"))
    if math.isfinite(quantity) and math.isfinite(ltp):
        row["Current Value"] = quantity * ltp
    if math.isfinite(quantity) and math.isfinite(day_change):
        row["Day P&L"] = quantity * day_change
    elif math.isfinite(quantity) and math.isfinite(current_day_change):
        row["Day P&L"] = quantity * current_day_change
    return row


def refresh(input_path: Path, output_path: Path) -> tuple[int, int]:
    data = clean_numeric_columns(pd.read_csv(input_path))
    for column in ("Yahoo Ticker", "Price Source", "Price Date"):
        if column not in data.columns:
            data[column] = ""
        data[column] = data[column].astype("object")

    tickers = data["Yahoo Ticker"].astype(str).str.strip().tolist()
    quotes = download_quotes(tickers)
    live_count = 0
    fallback_count = 0
    for index, row in data.iterrows():
        symbol = str(row.get("Yahoo Ticker") or "").strip()
        quote = quotes.get(symbol)
        if quote:
            live_count += 1
        else:
            fallback_count += 1
        data.loc[index] = recalculate_row(row, quote)

    total_value = float(data["Current Value"].sum())
    data["Weight %"] = data["Current Value"] / total_value * 100 if total_value else 0
    data["Last Refreshed"] = dt.datetime.now(ZoneInfo("Asia/Kolkata")).isoformat(timespec="minutes")
    data.to_csv(output_path, index=False)
    return live_count, fallback_count


def main() -> int:
    parser = argparse.ArgumentParser(description="Refresh NSDL dashboard prices from Yahoo Finance.")
    parser.add_argument("--input", type=Path, default=Path("data/holdings.csv"))
    parser.add_argument("--output", type=Path, default=Path("data/holdings.csv"))
    args = parser.parse_args()

    live_count, fallback_count = refresh(args.input, args.output)
    print(f"Live Yahoo quotes: {live_count}")
    print(f"NSDL fallbacks: {fallback_count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
