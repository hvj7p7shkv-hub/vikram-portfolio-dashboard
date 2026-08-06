#!/usr/bin/env python3
"""
Refresh an NSDL holdings CSV with Yahoo Finance prices.

Intraday bars provide the displayed price while daily bars provide the prior
official close used for the day move. A refresh with insufficient Yahoo
coverage fails instead of publishing an old NSDL snapshot as fresh data.
"""

from __future__ import annotations

import argparse
import datetime as dt
import math
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

from yahoo_http import download_many


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
    "Previous Close",
    "Session Volume",
    "Volume vs 20D",
]

MARKET_TIMEZONE = ZoneInfo("Asia/Kolkata")
INTRADAY_INTERVAL = "5m"
INTRADAY_PERIOD = "5d"
DAILY_INTERVAL = "1d"
DAILY_PERIOD = "3mo"
MAX_QUOTE_STALE_DAYS = 3
MAX_DAILY_MOVE_PCT = 35.0


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


def market_timestamp(value: object) -> pd.Timestamp:
    timestamp = pd.Timestamp(value)
    if timestamp.tzinfo is None:
        return timestamp.tz_localize(MARKET_TIMEZONE)
    return timestamp.tz_convert(MARKET_TIMEZONE)


def prior_close(
    daily_close: pd.Series,
    intraday_close: pd.Series,
    quote_date: dt.date,
    last: float,
) -> float:
    daily_candidates = [
        float(value)
        for index, value in daily_close.items()
        if market_timestamp(index).date() < quote_date
    ]
    if daily_candidates:
        return daily_candidates[-1]

    intraday_candidates = [
        float(value)
        for index, value in intraday_close.items()
        if market_timestamp(index).date() < quote_date
    ]
    return intraday_candidates[-1] if intraday_candidates else last


def volume_context(
    intraday_frame: pd.DataFrame | None,
    daily_frame: pd.DataFrame | None,
    quote_date: dt.date,
    use_intraday: bool,
) -> tuple[float, float, float]:
    intraday_volume = numeric_series(intraday_frame, "Volume")
    daily_volume = numeric_series(daily_frame, "Volume")
    completed_daily = [
        float(value)
        for index, value in daily_volume.items()
        if market_timestamp(index).date() < quote_date and float(value) >= 0
    ]
    average_20d = float(pd.Series(completed_daily[-20:]).mean()) if completed_daily else math.nan
    if use_intraday:
        session_volume = sum(
            float(value)
            for index, value in intraday_volume.items()
            if market_timestamp(index).date() == quote_date and float(value) >= 0
        )
    else:
        same_day = [
            float(value)
            for index, value in daily_volume.items()
            if market_timestamp(index).date() == quote_date and float(value) >= 0
        ]
        session_volume = same_day[-1] if same_day else math.nan
    volume_ratio = session_volume / average_20d if average_20d and math.isfinite(session_volume) else math.nan
    return session_volume, average_20d, volume_ratio


def intraday_alert(move_pct: float, volume_ratio: float) -> str:
    move_alert = math.isfinite(move_pct) and abs(move_pct) >= 5
    volume_alert = math.isfinite(volume_ratio) and volume_ratio >= 2
    if move_alert and volume_alert:
        return f"Exceptional move ({move_pct:+.2f}%) on {volume_ratio:.1f}x 20-day volume"
    if move_alert:
        return f"Exceptional intraday move ({move_pct:+.2f}%)"
    if volume_alert:
        return f"Unusual volume ({volume_ratio:.1f}x 20-day average)"
    return ""


def quote_from_frames(
    intraday_frame: pd.DataFrame | None,
    daily_frame: pd.DataFrame | None,
) -> dict[str, object] | None:
    intraday_close = numeric_series(intraday_frame, "Close")
    daily_close = numeric_series(daily_frame, "Close")

    if not intraday_close.empty:
        timestamp = market_timestamp(intraday_close.index[-1])
        last = float(intraday_close.iloc[-1])
        previous = prior_close(daily_close, intraday_close, timestamp.date(), last)
        move_pct = (last / previous - 1) * 100 if previous else math.nan
        session_volume, average_20d, volume_ratio = volume_context(
            intraday_frame, daily_frame, timestamp.date(), True
        )
        return {
            "last": last,
            "previous": previous,
            "daily_move_pct": move_pct,
            "session_volume": session_volume,
            "average_volume_20d": average_20d,
            "volume_ratio_20d": volume_ratio,
            "intraday_alert": intraday_alert(move_pct, volume_ratio),
            "date": timestamp.date().isoformat(),
            "timestamp": timestamp.strftime("%Y-%m-%d %H:%M %Z"),
            "quote_type": "Yahoo 5-minute quote",
        }

    if daily_close.empty:
        return None
    timestamp = market_timestamp(daily_close.index[-1])
    last = float(daily_close.iloc[-1])
    previous = float(daily_close.iloc[-2]) if len(daily_close) >= 2 else last
    move_pct = (last / previous - 1) * 100 if previous else math.nan
    session_volume, average_20d, volume_ratio = volume_context(
        intraday_frame, daily_frame, timestamp.date(), False
    )
    return {
        "last": last,
        "previous": previous,
        "daily_move_pct": move_pct,
        "session_volume": session_volume,
        "average_volume_20d": average_20d,
        "volume_ratio_20d": volume_ratio,
        "intraday_alert": intraday_alert(move_pct, volume_ratio),
        "date": timestamp.date().isoformat(),
        "timestamp": f"{timestamp.date().isoformat()} market close",
        "quote_type": "Yahoo daily close",
    }


def validate_quote(quote: dict[str, object] | None) -> dict[str, object] | None:
    if not quote:
        return None

    quote = dict(quote)
    quote["valid_quote"] = True
    quote["quote_note"] = ""

    today = dt.datetime.now(MARKET_TIMEZONE).date()
    try:
        quote_date = dt.date.fromisoformat(str(quote.get("date") or ""))
    except ValueError:
        quote["valid_quote"] = False
        quote["quote_status"] = "Rejected quote - date unavailable"
        quote["quote_note"] = "Yahoo returned a quote without a usable date, so the previous dashboard value was retained."
        return quote

    stale_days = (today - quote_date).days
    quote["stale_days"] = stale_days
    if stale_days > MAX_QUOTE_STALE_DAYS:
        quote["valid_quote"] = False
        quote["quote_status"] = "Stale Yahoo quote - previous value retained"
        quote["quote_note"] = (
            f"Yahoo returned {quote_date.isoformat()}, which is {stale_days} calendar days old. "
            "The previous dashboard value was retained."
        )
        return quote

    move_pct = clean_number(quote.get("daily_move_pct"))
    if math.isfinite(move_pct) and abs(move_pct) > MAX_DAILY_MOVE_PCT:
        quote["valid_quote"] = False
        quote["quote_status"] = "Outlier Yahoo quote - previous value retained"
        quote["quote_note"] = (
            f"Yahoo implied a one-day move of {move_pct:.2f}%, which is outside the dashboard safety band. "
            "The previous dashboard value was retained for review."
        )
        return quote

    quote["quote_status"] = "Updated"
    return quote


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


def download_history(tickers: list[str], period: str, interval: str) -> pd.DataFrame:
    if not tickers:
        return pd.DataFrame()
    frames = download_many(
        tickers,
        range_period=period,
        interval=interval,
        timeout=30,
        max_workers=6,
    )
    if not frames:
        return pd.DataFrame()
    if len(tickers) == 1:
        return frames.get(tickers[0], pd.DataFrame())
    return pd.concat(frames, axis=1)


def download_batch(tickers: list[str]) -> dict[str, dict[str, object]]:
    if not tickers:
        return {}
    intraday = download_history(tickers, INTRADAY_PERIOD, INTRADAY_INTERVAL)
    daily = download_history(tickers, DAILY_PERIOD, DAILY_INTERVAL)

    quotes: dict[str, dict[str, object]] = {}
    for ticker in tickers:
        quote = quote_from_frames(
            extract_frame(intraday, ticker, len(tickers)),
            extract_frame(daily, ticker, len(tickers)),
        )
        if quote:
            quotes[ticker] = validate_quote(quote)
    return quotes


def download_single(ticker: str) -> dict[str, object] | None:
    intraday = download_history([ticker], INTRADAY_PERIOD, INTRADAY_INTERVAL)
    daily = download_history([ticker], DAILY_PERIOD, DAILY_INTERVAL)
    return validate_quote(quote_from_frames(
        extract_frame(intraday, ticker, 1),
        extract_frame(daily, ticker, 1),
    ))


def download_quotes(primary_symbols: list[str]) -> dict[str, dict[str, object]]:
    quotes: dict[str, dict[str, object]] = {}
    unique = sorted({symbol for symbol in primary_symbols if str(symbol).strip()})
    batch_size = 45
    for start in range(0, len(unique), batch_size):
        quotes.update(download_batch(unique[start : start + batch_size]))

    for symbol in unique:
        existing = quotes.get(symbol)
        if existing and existing.get("valid_quote", True):
            continue
        best_quote = existing
        for fallback in fallback_symbols(symbol):
            if fallback == symbol and existing:
                continue
            quote = download_single(fallback)
            if quote:
                quote["downloaded_ticker"] = fallback
                if quote.get("valid_quote", True):
                    best_quote = quote
                    break
                if best_quote is None:
                    best_quote = quote
        if best_quote:
            quotes[symbol] = best_quote
    return quotes


def recalculate_row(row: pd.Series, quote: dict[str, object] | None) -> pd.Series:
    quantity = clean_number(row.get("Quantity"))
    current_ltp = clean_number(row.get("LTP"))
    current_day_change = clean_number(row.get("Day Change"))

    if quote and quote.get("valid_quote", True):
        ltp = float(quote["last"])
        previous = float(quote["previous"])
        row["LTP"] = ltp
        row["Previous Close"] = previous
        row["Session Volume"] = clean_number(quote.get("session_volume"))
        row["Volume vs 20D"] = clean_number(quote.get("volume_ratio_20d"))
        row["Intraday Alert"] = str(quote.get("intraday_alert") or "")
        row["Day Change"] = ltp - previous
        row["Day Change %"] = (ltp / previous - 1) * 100 if previous else math.nan
        row["Price Source"] = str(quote.get("downloaded_ticker") or row.get("Yahoo Ticker") or "")
        row["Price Date"] = str(quote.get("date") or "")
        row["Price Timestamp"] = str(quote.get("timestamp") or row["Price Date"])
        row["Quote Type"] = str(quote.get("quote_type") or "Yahoo quote")
        row["Quote Status"] = str(quote.get("quote_status") or "Updated")
        row["Quote Note"] = ""
    elif quote:
        ltp = current_ltp
        row["Intraday Alert"] = ""
        row["Price Source"] = str(quote.get("downloaded_ticker") or row.get("Yahoo Ticker") or "")
        row["Quote Type"] = str(quote.get("quote_type") or "Yahoo quote")
        row["Quote Status"] = str(quote.get("quote_status") or "Rejected quote - previous value retained")
        row["Quote Note"] = str(quote.get("quote_note") or "")
    else:
        ltp = current_ltp
        row["Intraday Alert"] = ""
        row["Price Source"] = "NSDL fallback"
        row["Quote Status"] = "Source unavailable - previous value retained"
        row["Quote Note"] = "Yahoo did not return a usable quote in this refresh."

    day_change = clean_number(row.get("Day Change"))
    if math.isfinite(quantity) and math.isfinite(ltp):
        row["Current Value"] = quantity * ltp
    if math.isfinite(quantity) and math.isfinite(day_change):
        row["Day P&L"] = quantity * day_change
    elif math.isfinite(quantity) and math.isfinite(current_day_change):
        row["Day P&L"] = quantity * current_day_change
    return row


def refresh(input_path: Path, output_path: Path, min_coverage: float = 0.70) -> tuple[int, int, int, int]:
    data = clean_numeric_columns(pd.read_csv(input_path))
    for column in ("Previous Close", "Session Volume", "Volume vs 20D"):
        if column not in data.columns:
            data[column] = math.nan
    for column in (
        "Yahoo Ticker",
        "Price Source",
        "Price Date",
        "Price Timestamp",
        "Quote Type",
        "Quote Status",
        "Quote Note",
        "Intraday Alert",
    ):
        if column not in data.columns:
            data[column] = ""
        data[column] = data[column].astype("object")

    tickers = data["Yahoo Ticker"].astype(str).str.strip().tolist()
    quotes = download_quotes(tickers)

    expected_quotes = len({ticker for ticker in tickers if ticker})
    valid_quotes = {ticker: quote for ticker, quote in quotes.items() if quote and quote.get("valid_quote", True)}
    coverage = len(valid_quotes) / expected_quotes if expected_quotes else 0
    if expected_quotes and coverage < min_coverage:
        raise RuntimeError(
            f"Yahoo quote coverage was {coverage:.0%} ({len(valid_quotes)}/{expected_quotes}); "
            f"minimum required coverage is {min_coverage:.0%}. Previous dashboard was not overwritten."
        )

    live_count = 0
    fallback_count = 0
    intraday_count = 0
    rejected_count = 0
    for index, row in data.iterrows():
        symbol = str(row.get("Yahoo Ticker") or "").strip()
        quote = quotes.get(symbol)
        if quote and quote.get("valid_quote", True):
            live_count += 1
            intraday_count += int(quote.get("quote_type") == "Yahoo 5-minute quote")
        elif quote:
            rejected_count += 1
        else:
            fallback_count += 1
        data.loc[index] = recalculate_row(row, quote)

    total_value = float(data["Current Value"].sum())
    data["Weight %"] = data["Current Value"] / total_value * 100 if total_value else 0
    data["Last Refreshed"] = dt.datetime.now(MARKET_TIMEZONE).isoformat(timespec="minutes")
    data.to_csv(output_path, index=False)
    return live_count, fallback_count, intraday_count, rejected_count


def main() -> int:
    parser = argparse.ArgumentParser(description="Refresh NSDL dashboard prices from Yahoo Finance.")
    parser.add_argument("--input", type=Path, default=Path("data/holdings.csv"))
    parser.add_argument("--output", type=Path, default=Path("data/holdings.csv"))
    parser.add_argument("--min-coverage", type=float, default=0.70)
    args = parser.parse_args()

    live_count, fallback_count, intraday_count, rejected_count = refresh(
        args.input,
        args.output,
        min_coverage=args.min_coverage,
    )
    print(f"Live Yahoo quotes: {live_count}")
    print(f"Yahoo 5-minute quotes: {intraday_count}")
    print(f"Yahoo daily-close quotes: {live_count - intraday_count}")
    print(f"Rejected Yahoo quotes: {rejected_count}")
    print(f"NSDL fallbacks: {fallback_count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
