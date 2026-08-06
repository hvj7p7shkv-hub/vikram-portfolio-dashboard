#!/usr/bin/env python3
"""Small Yahoo chart-API client using only requests and pandas.

This module deliberately avoids yfinance and curl_cffi. It retrieves the same
OHLCV fields needed by the local dashboards from Yahoo's public chart endpoint.
"""

from __future__ import annotations

import math
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, Iterable, Optional
from urllib.parse import quote

import pandas as pd
import requests

BASE_URL = "https://query1.finance.yahoo.com/v8/finance/chart/{symbol}"
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/605.1.15 (KHTML, like Gecko) Safari/605.1.15"
)


def _as_float_list(values: object, length: int) -> list:
    if not isinstance(values, list):
        return [math.nan] * length
    output = []
    for value in values[:length]:
        try:
            output.append(float(value) if value is not None else math.nan)
        except (TypeError, ValueError):
            output.append(math.nan)
    if len(output) < length:
        output.extend([math.nan] * (length - len(output)))
    return output


def frame_from_chart(payload: dict) -> pd.DataFrame:
    chart = payload.get("chart") if isinstance(payload, dict) else None
    results = chart.get("result") if isinstance(chart, dict) else None
    if not results:
        return pd.DataFrame()

    result = results[0]
    timestamps = result.get("timestamp") or []
    indicators = result.get("indicators") or {}
    quotes = indicators.get("quote") or []
    if not timestamps or not quotes:
        return pd.DataFrame()

    quote_data = quotes[0] or {}
    count = len(timestamps)
    adjclose_groups = indicators.get("adjclose") or []
    adjclose_data = adjclose_groups[0] if adjclose_groups else {}

    index = pd.to_datetime(timestamps, unit="s", utc=True)
    timezone_name = str((result.get("meta") or {}).get("exchangeTimezoneName") or "Asia/Kolkata")
    try:
        index = index.tz_convert(timezone_name)
    except Exception:
        index = index.tz_convert("Asia/Kolkata")

    frame = pd.DataFrame(
        {
            "Open": _as_float_list(quote_data.get("open"), count),
            "High": _as_float_list(quote_data.get("high"), count),
            "Low": _as_float_list(quote_data.get("low"), count),
            "Close": _as_float_list(quote_data.get("close"), count),
            "Adj Close": _as_float_list(adjclose_data.get("adjclose"), count),
            "Volume": _as_float_list(quote_data.get("volume"), count),
        },
        index=index,
    )
    if frame["Adj Close"].isna().all():
        frame["Adj Close"] = frame["Close"]
    return frame.dropna(how="all").sort_index()


def download_one(
    ticker: str,
    range_period: str,
    interval: str,
    timeout: int = 30,
    attempts: int = 3,
) -> pd.DataFrame:
    symbol = quote(str(ticker).strip(), safe="")
    if not symbol:
        return pd.DataFrame()

    url = BASE_URL.format(symbol=symbol)
    params = {
        "range": range_period,
        "interval": interval,
        "includePrePost": "false",
        "events": "div,splits,capitalGains",
    }
    headers = {"User-Agent": USER_AGENT, "Accept": "application/json"}

    for attempt in range(max(1, attempts)):
        try:
            response = requests.get(url, params=params, headers=headers, timeout=timeout)
            if response.status_code in (429, 500, 502, 503, 504):
                if attempt + 1 < attempts:
                    time.sleep(1.5 * (attempt + 1))
                    continue
            response.raise_for_status()
            return frame_from_chart(response.json())
        except (requests.RequestException, ValueError):
            if attempt + 1 < attempts:
                time.sleep(1.5 * (attempt + 1))
                continue
            return pd.DataFrame()
    return pd.DataFrame()


def download_many(
    tickers: Iterable[str],
    range_period: str,
    interval: str,
    timeout: int = 30,
    max_workers: int = 6,
) -> Dict[str, pd.DataFrame]:
    unique = sorted({str(ticker).strip() for ticker in tickers if str(ticker).strip()})
    if not unique:
        return {}

    frames: Dict[str, pd.DataFrame] = {}
    workers = max(1, min(max_workers, len(unique)))
    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {
            executor.submit(download_one, ticker, range_period, interval, timeout): ticker
            for ticker in unique
        }
        for future in as_completed(futures):
            ticker = futures[future]
            try:
                frame = future.result()
            except Exception:
                frame = pd.DataFrame()
            if frame is not None and not frame.empty:
                frames[ticker] = frame
    return frames
