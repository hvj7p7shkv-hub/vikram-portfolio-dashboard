#!/usr/bin/env python3
"""
Add technical indicators to the client portfolio CSV.

The output remains a single holdings CSV so the GitHub Pages workflow can
refresh prices, update technicals, and rebuild the dashboard in one pass.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import math
import re
from dataclasses import dataclass
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

from yahoo_http import download_many, download_one


TECHNICAL_COLUMNS = [
    "Technical As Of",
    "Technical Downloaded",
    "Technical Status",
    "Technical Score",
    "Technical Note",
    "RS Trend",
    "RS vs 50D %",
    "RS 3M %",
    "RS Leader",
    "RSI 14",
    "P&F Signal",
    "Above 50DMA",
    "Above 200DMA",
    "50DMA Distance %",
    "200DMA Distance %",
    "52W High Distance %",
    "Technical Error",
]

MARKET_TIMEZONE = ZoneInfo("Asia/Kolkata")
MAX_HISTORY_STALE_DAYS = 3
MAX_DAILY_MOVE_PCT = 35.0


def clean_number(value: object) -> float:
    text = str(value).replace(",", "").replace("%", "").replace("+", "").strip()
    try:
        return float(text)
    except ValueError:
        return math.nan


def normalize_symbol(name: object) -> str:
    symbol = str(name).strip().upper()
    symbol = re.sub(r"\s+", "", symbol)
    symbol = re.sub(r"-(EQ|BE)$", "", symbol)
    if symbol.startswith("^") or "." in symbol:
        return symbol
    return f"{symbol}.NS"


def fallback_symbols(symbol: str) -> list[str]:
    symbols = [symbol]
    if symbol.endswith(".NS"):
        symbols.append(symbol[:-3] + ".BO")
    return symbols


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


def completed_eod_series(series: pd.Series, now: dt.datetime | None = None) -> pd.Series:
    """Return only completed Indian-market daily bars."""
    clean = series.dropna().sort_index()
    if clean.empty:
        return clean

    current = now or dt.datetime.now(MARKET_TIMEZONE)
    if current.tzinfo is None:
        current = current.replace(tzinfo=MARKET_TIMEZONE)
    else:
        current = current.astimezone(MARKET_TIMEZONE)
    latest_date = market_timestamp(clean.index[-1]).date()
    if latest_date == current.date() and current.time() < dt.time(15, 40):
        clean = clean.iloc[:-1]
    return clean


def close_series(frame: pd.DataFrame | None, now: dt.datetime | None = None) -> pd.Series:
    if frame is None or frame.empty:
        return pd.Series(dtype=float)
    # Technical indicators use adjusted history so corporate actions do not
    # appear as false RS, RSI, moving-average, or P&F moves.
    for column in ("Adj Close", "Close"):
        series = numeric_series(frame, column)
        if not series.empty:
            return completed_eod_series(series, now=now)
    return pd.Series(dtype=float)


def market_timestamp(value: object) -> pd.Timestamp:
    timestamp = pd.Timestamp(value)
    if timestamp.tzinfo is None:
        return timestamp.tz_localize(MARKET_TIMEZONE)
    return timestamp.tz_convert(MARKET_TIMEZONE)


def validate_close_history(close: pd.Series, label: str) -> None:
    clean = close.dropna().sort_index()
    if clean.empty:
        raise ValueError(f"No usable daily price history for {label}")

    latest_date = market_timestamp(clean.index[-1]).date()
    stale_days = (dt.datetime.now(MARKET_TIMEZONE).date() - latest_date).days
    if stale_days > MAX_HISTORY_STALE_DAYS:
        raise ValueError(
            f"Stale daily price history for {label}: latest bar is {latest_date.isoformat()} "
            f"({stale_days} calendar days old)"
        )

    if len(clean) >= 2:
        previous = float(clean.iloc[-2])
        latest = float(clean.iloc[-1])
        if previous:
            move_pct = (latest / previous - 1) * 100
            if math.isfinite(move_pct) and abs(move_pct) > MAX_DAILY_MOVE_PCT:
                raise ValueError(
                    f"Outlier daily price move for {label}: latest bar implies {move_pct:.2f}%"
                )


def rsi(series: pd.Series, length: int = 14) -> pd.Series:
    delta = series.diff()
    gain = delta.clip(lower=0).ewm(alpha=1 / length, adjust=False).mean()
    loss = (-delta.clip(upper=0)).ewm(alpha=1 / length, adjust=False).mean()
    rs = gain / loss.replace(0, math.nan)
    result = 100 - (100 / (1 + rs))
    result = result.mask((loss == 0) & (gain > 0), 100)
    result = result.mask((gain == 0) & (loss > 0), 0)
    return result


def pct_return(series: pd.Series, days: int) -> float | None:
    clean = series.dropna()
    if len(clean) <= days:
        return None
    return float((clean.iloc[-1] / clean.iloc[-days - 1] - 1) * 100)


@dataclass
class PnfColumn:
    kind: str
    boxes: list[int]


def build_pnf(close: pd.Series, box_size: float, reversal: int = 3) -> list[PnfColumn]:
    prices = close.dropna()
    if prices.empty or box_size <= 0:
        return []

    def box(price: float) -> int:
        return int(math.floor(price / box_size))

    columns: list[PnfColumn] = []
    current_kind: str | None = None
    current_boxes: list[int] = []
    first_box = box(float(prices.iloc[0]))

    for price_value in prices.iloc[1:]:
        price_box = box(float(price_value))
        if current_kind is None:
            if price_box > first_box:
                current_kind = "X"
                current_boxes = list(range(first_box + 1, price_box + 1))
            elif price_box < first_box:
                current_kind = "O"
                current_boxes = list(range(first_box - 1, price_box - 1, -1))
            continue

        high_box = max(current_boxes)
        low_box = min(current_boxes)
        if current_kind == "X":
            if price_box > high_box:
                current_boxes.extend(range(high_box + 1, price_box + 1))
            elif price_box <= high_box - reversal:
                columns.append(PnfColumn("X", sorted(set(current_boxes))))
                current_kind = "O"
                current_boxes = list(range(high_box - 1, price_box - 1, -1))
        elif price_box < low_box:
            current_boxes.extend(range(low_box - 1, price_box - 1, -1))
        elif price_box >= low_box + reversal:
            columns.append(PnfColumn("O", sorted(set(current_boxes))))
            current_kind = "X"
            current_boxes = list(range(low_box + 1, price_box + 1))

    if current_kind and current_boxes:
        columns.append(PnfColumn(current_kind, sorted(set(current_boxes))))
    return columns


def pnf_signal(columns: list[PnfColumn]) -> str:
    if len(columns) < 2:
        return "Insufficient P&F structure"
    latest = columns[-1]
    previous_same = [column for column in columns[:-1] if column.kind == latest.kind]
    if latest.kind == "X" and previous_same and max(latest.boxes) > max(previous_same[-1].boxes):
        return "Bullish P&F breakout"
    if latest.kind == "O" and previous_same and min(latest.boxes) < min(previous_same[-1].boxes):
        return "Bearish P&F breakdown"
    if latest.kind == "X":
        return "Rising P&F column, no fresh breakout yet"
    return "Falling P&F column, wait for reversal"


def download_batch(tickers: list[str], period: str) -> pd.DataFrame:
    if not tickers:
        return pd.DataFrame()
    frames = download_many(
        tickers,
        range_period=period,
        interval="1d",
        timeout=45,
        max_workers=6,
    )
    if not frames:
        return pd.DataFrame()
    if len(tickers) == 1:
        return frames.get(tickers[0], pd.DataFrame())
    return pd.concat(frames, axis=1)


def download_history(ticker: str, period: str) -> pd.DataFrame | None:
    history = download_one(ticker, range_period=period, interval="1d", timeout=45)
    if history is None or history.empty:
        return None
    return history.dropna(how="all")


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


def download_single(ticker: str, period: str) -> pd.DataFrame | None:
    downloaded = download_batch([ticker], period)
    frame = extract_frame(downloaded, ticker, 1)
    if frame is not None and not close_series(frame).empty:
        return frame
    return download_history(ticker, period)


def download_frames(primary_symbols: list[str], period: str) -> dict[str, pd.DataFrame]:
    frames: dict[str, pd.DataFrame] = {}
    unique = sorted(set(primary_symbols))
    batch_size = 35
    for start in range(0, len(unique), batch_size):
        batch = unique[start : start + batch_size]
        downloaded = download_batch(batch, period)
        for ticker in batch:
            frame = extract_frame(downloaded, ticker, len(batch))
            if frame is not None and not close_series(frame).empty:
                frames[ticker] = frame

    for symbol in unique:
        if symbol in frames:
            continue
        for fallback in fallback_symbols(symbol):
            frame = download_single(fallback, period)
            if frame is not None and not close_series(frame).empty:
                frames[symbol] = frame
                break
    return frames


def relative_strength(close: pd.Series, benchmark_close: pd.Series) -> tuple[float | None, float | None, bool | None]:
    aligned = pd.concat([close.rename("stock"), benchmark_close.rename("benchmark")], axis=1).dropna()
    aligned = aligned[aligned["benchmark"] > 0]
    if len(aligned) < 60:
        return None, None, None
    ratio = aligned["stock"] / aligned["benchmark"]
    ratio_ma_50 = ratio.rolling(50).mean()
    if pd.isna(ratio_ma_50.iloc[-1]) or ratio_ma_50.iloc[-1] == 0:
        rs_vs_50 = None
    else:
        rs_vs_50 = float((ratio.iloc[-1] / ratio_ma_50.iloc[-1] - 1) * 100)
    rs_3m = pct_return(ratio, 63)
    leader = bool(rs_vs_50 is not None and rs_vs_50 > 0)
    return rs_vs_50, rs_3m, leader


def relative_strength_trend(close: pd.Series, benchmark_close: pd.Series, days: int = 90) -> str | None:
    aligned = pd.concat([close.rename("stock"), benchmark_close.rename("benchmark")], axis=1).dropna()
    aligned = aligned[aligned["benchmark"] > 0]
    if len(aligned) < 20:
        return None
    ratio = (aligned["stock"] / aligned["benchmark"]).tail(days).dropna()
    if ratio.empty or ratio.iloc[0] == 0:
        return None
    normalized = (ratio / ratio.iloc[0] * 100).round(2).tolist()
    return json.dumps(normalized, separators=(",", ":"))


def score_and_status(metrics: dict[str, object]) -> tuple[int, str, str]:
    score = 0
    reasons: list[str] = []

    rs_vs_50 = metrics.get("RS vs 50D %")
    rsi_14 = metrics.get("RSI 14")
    above_50_value = metrics.get("Above 50DMA")
    above_200_value = metrics.get("Above 200DMA")
    above_50 = above_50_value is True
    above_200 = above_200_value is True
    has_50dma = isinstance(above_50_value, bool)
    has_200dma = isinstance(above_200_value, bool)
    pnf = str(metrics.get("P&F Signal") or "")

    if above_50:
        score += 20
        reasons.append("price is above the 50DMA")
    if above_200:
        score += 20
        reasons.append("price is above the 200DMA")
    if isinstance(rs_vs_50, (int, float)) and math.isfinite(float(rs_vs_50)):
        if float(rs_vs_50) > 0:
            score += 25
            reasons.append("relative strength is above its 50-day average")
        elif float(rs_vs_50) < -3:
            score -= 10
            reasons.append("relative strength is below its 50-day average")
    if isinstance(rsi_14, (int, float)) and math.isfinite(float(rsi_14)):
        if 45 <= float(rsi_14) <= 70:
            score += 15
            reasons.append("RSI is in a constructive momentum zone")
        elif float(rsi_14) > 75:
            score -= 5
            reasons.append("RSI is extended")
        elif float(rsi_14) < 35:
            score -= 10
            reasons.append("RSI is weak or oversold")
    if "Bullish" in pnf:
        score += 20
        reasons.append("P&F shows a bullish breakout")
    elif "Bearish" in pnf:
        score -= 25
        reasons.append("P&F shows a bearish breakdown")

    score = max(0, min(100, score))
    if score >= 75:
        status = "Leader / hold"
    elif score >= 60:
        status = "Constructive"
    elif score <= 25 or "Bearish" in pnf:
        status = "Risk review"
    elif has_50dma and has_200dma and not above_50 and not above_200:
        status = "Loss + weak structure"
    else:
        status = "Monitor"

    note = "; ".join(reasons) if reasons else "Technical data was not strong enough to classify clearly."
    return score, status, note.capitalize()


def analyse_frame(frame: pd.DataFrame, benchmark_close: pd.Series, symbol: str) -> dict[str, object]:
    close = close_series(frame)
    validate_close_history(close, symbol)
    if len(close) < 30:
        raise ValueError("Not enough daily price history")

    latest = float(close.iloc[-1])
    sma_50_value = close.rolling(50).mean().iloc[-1] if len(close) >= 50 else math.nan
    sma_200_value = close.rolling(200).mean().iloc[-1] if len(close) >= 200 else math.nan
    sma_50 = float(sma_50_value) if pd.notna(sma_50_value) else None
    sma_200 = float(sma_200_value) if pd.notna(sma_200_value) else None
    high_52w = float(close.tail(min(252, len(close))).max())
    rsi_14_series = rsi(close)
    rsi_14_value = rsi_14_series.iloc[-1] if not rsi_14_series.dropna().empty else math.nan
    rsi_14 = float(rsi_14_value) if pd.notna(rsi_14_value) else None
    rs_vs_50, rs_3m, rs_leader = relative_strength(close, benchmark_close)
    rs_trend = relative_strength_trend(close, benchmark_close)
    box_size = max(latest * 0.02, 0.01)
    pnf = pnf_signal(build_pnf(close, box_size=box_size, reversal=3))

    metrics: dict[str, object] = {
        "Technical As Of": market_timestamp(close.index[-1]).date().isoformat(),
        "RS Trend": rs_trend,
        "RS vs 50D %": rs_vs_50,
        "RS 3M %": rs_3m,
        "RS Leader": rs_leader,
        "RSI 14": round(rsi_14, 2) if rsi_14 is not None else None,
        "P&F Signal": pnf,
        "Above 50DMA": bool(latest > sma_50) if sma_50 else None,
        "Above 200DMA": bool(latest > sma_200) if sma_200 else None,
        "50DMA Distance %": (latest / sma_50 - 1) * 100 if sma_50 else None,
        "200DMA Distance %": (latest / sma_200 - 1) * 100 if sma_200 else None,
        "52W High Distance %": (latest / high_52w - 1) * 100 if high_52w else None,
    }
    score, status, note = score_and_status(metrics)
    if len(close) < 220:
        status = f"Partial: {status}"
        note = f"{note} History available for {len(close)} trading days, so long-term context is limited."
    metrics.update(
        {
            "Technical Downloaded": True,
            "Technical Status": status,
            "Technical Score": score,
            "Technical Note": note,
            "Technical Error": "",
        }
    )
    return metrics


def clear_technical_columns(data: pd.DataFrame) -> pd.DataFrame:
    for column in TECHNICAL_COLUMNS:
        if column not in data.columns:
            data[column] = None
    return data


def unavailable_technical_details(symbol: str, error: str) -> tuple[str, str]:
    normalized_symbol = symbol.strip().upper()
    normalized_error = error.strip().lower()

    if "not enough daily price history" in normalized_error:
        if normalized_symbol == "TMCV.NS":
            return (
                "Limited post-demerger history",
                "Tata Motors' commercial-vehicle listing has a shorter standalone history after the demerger. "
                "Current price and daily movement are available, but the full long-term technical ranking "
                "will remain limited until enough trading history accumulates.",
            )
        return (
            "Limited price history",
            "The listing does not yet have enough daily trading history for the complete technical set. "
            "Current price and daily movement are still available; long-term indicators will appear as "
            "more history accumulates.",
        )

    if not normalized_symbol or normalized_symbol in {"NAN", "NONE"}:
        return (
            "Ticker review required",
            "The market ticker has not been confirmed, so technical indicators are paused to avoid showing "
            "analysis for the wrong security. Current NSDL holding data remains available.",
        )

    if any(phrase in normalized_error for phrase in ("no price data", "no data", "possibly delisted", "empty")):
        return (
            "Price history unavailable",
            "The data provider returned no usable daily price history in the latest refresh. Current NSDL "
            "holding data remains visible, and the technical layer will retry automatically.",
        )

    return (
        "Technical refresh unavailable",
        "The technical layer could not be calculated in the latest refresh. Current holding value and daily "
        "movement remain available; the next refresh will retry the analysis.",
    )


def annualized_return(returns: pd.Series) -> float | None:
    clean = returns.dropna()
    if len(clean) < 20:
        return None
    compounded = float((1 + clean).prod())
    if compounded <= 0:
        return None
    return (compounded ** (252 / len(clean)) - 1) * 100


def trailing_return(returns: pd.Series, days: int) -> float | None:
    clean = returns.dropna()
    if len(clean) < min(days, 20):
        return None
    return float(((1 + clean.tail(days)).prod() - 1) * 100)


def portfolio_risk_snapshot(
    data: pd.DataFrame,
    frames: dict[str, pd.DataFrame],
    benchmark: str,
    benchmark_close: pd.Series,
) -> dict[str, object]:
    total_value = float(pd.to_numeric(data.get("Current Value"), errors="coerce").fillna(0).sum())
    if total_value <= 0 or benchmark_close.empty:
        return {"available": False, "error": "Portfolio value or benchmark history is unavailable."}

    weights: dict[str, float] = {}
    return_series: dict[str, pd.Series] = {}
    for _, row in data.iterrows():
        symbol = str(row.get("Yahoo Ticker") or "").strip()
        value = clean_number(row.get("Current Value"))
        frame = frames.get(symbol)
        close = close_series(frame)
        try:
            validate_close_history(close, symbol)
        except Exception:
            continue
        if not symbol or not math.isfinite(value) or value <= 0 or len(close) < 60:
            continue
        weights[symbol] = weights.get(symbol, 0.0) + value / total_value
        if symbol not in return_series:
            return_series[symbol] = close.sort_index().pct_change(fill_method=None)

    if not return_series:
        return {"available": False, "error": "No holding had sufficient price history for portfolio risk estimates."}

    returns = pd.DataFrame(return_series).sort_index()
    benchmark_returns = benchmark_close.sort_index().pct_change(fill_method=None).rename("benchmark")
    weight_series = pd.Series(weights, dtype=float)
    available_weight = returns.notna().mul(weight_series, axis=1).sum(axis=1)
    portfolio_returns = returns.mul(weight_series, axis=1).sum(axis=1) / available_weight.replace(0, math.nan)
    portfolio_returns = portfolio_returns.where(available_weight >= 0.5)
    aligned = pd.concat([portfolio_returns.rename("portfolio"), benchmark_returns], axis=1).dropna()
    if len(aligned) < 60:
        return {"available": False, "error": "Less than 60 overlapping sessions were available for risk estimates."}

    window = aligned.tail(min(252, len(aligned))).copy()
    benchmark_variance = float(window["benchmark"].var())
    beta = float(window["portfolio"].cov(window["benchmark"]) / benchmark_variance) if benchmark_variance else None
    correlation = float(window["portfolio"].corr(window["benchmark"]))
    volatility = float(window["portfolio"].std() * math.sqrt(252) * 100)
    benchmark_volatility = float(window["benchmark"].std() * math.sqrt(252) * 100)
    active = window["portfolio"] - window["benchmark"]
    tracking_error = float(active.std() * math.sqrt(252) * 100)
    active_annualized = float(active.mean() * 252 * 100)
    information_ratio = active_annualized / tracking_error if tracking_error else None

    portfolio_curve = (1 + window["portfolio"]).cumprod()
    benchmark_curve = (1 + window["benchmark"]).cumprod()
    drawdown = portfolio_curve / portfolio_curve.cummax() - 1
    max_drawdown = float(drawdown.min() * 100)
    portfolio_annualized = annualized_return(window["portfolio"])
    benchmark_annualized = annualized_return(window["benchmark"])
    alpha = (
        portfolio_annualized - beta * benchmark_annualized
        if portfolio_annualized is not None and benchmark_annualized is not None and beta is not None
        else None
    )

    up = window["benchmark"] > 0
    down = window["benchmark"] < 0
    up_benchmark_mean = float(window.loc[up, "benchmark"].mean()) if up.any() else None
    down_benchmark_mean = float(window.loc[down, "benchmark"].mean()) if down.any() else None
    upside_capture = (
        float(window.loc[up, "portfolio"].mean()) / up_benchmark_mean * 100 if up_benchmark_mean else None
    )
    downside_capture = (
        float(window.loc[down, "portfolio"].mean()) / down_benchmark_mean * 100 if down_benchmark_mean else None
    )

    normalized_portfolio = portfolio_curve / portfolio_curve.iloc[0] * 100
    normalized_benchmark = benchmark_curve / benchmark_curve.iloc[0] * 100
    relative_strength = normalized_portfolio / normalized_benchmark * 100
    chart = pd.DataFrame(
        {
            "portfolio": normalized_portfolio,
            "benchmark": normalized_benchmark,
            "relativeStrength": relative_strength,
        }
    ).tail(180)

    benchmark_label = "Nifty Midcap 50" if benchmark == "^NSEMDCP50" else benchmark
    return {
        "available": True,
        "asOf": dt.datetime.now().astimezone().isoformat(timespec="minutes"),
        "benchmark": benchmark,
        "benchmarkLabel": benchmark_label,
        "method": "Current-value-weighted daily return estimate; weights are re-normalized when a price is missing.",
        "coveragePct": round(sum(weights.values()) * 100, 2),
        "sessions": int(len(window)),
        "beta": round(beta, 3) if beta is not None else None,
        "alphaAnnualizedPct": round(alpha, 2) if alpha is not None else None,
        "volatilityAnnualizedPct": round(volatility, 2),
        "benchmarkVolatilityAnnualizedPct": round(benchmark_volatility, 2),
        "correlation": round(correlation, 3),
        "maxDrawdownPct": round(max_drawdown, 2),
        "trackingErrorPct": round(tracking_error, 2),
        "informationRatio": round(information_ratio, 2) if information_ratio is not None else None,
        "upsideCapturePct": round(upside_capture, 2) if upside_capture is not None else None,
        "downsideCapturePct": round(downside_capture, 2) if downside_capture is not None else None,
        "portfolioReturn3mPct": round(trailing_return(window["portfolio"], 63), 2)
        if trailing_return(window["portfolio"], 63) is not None
        else None,
        "benchmarkReturn3mPct": round(trailing_return(window["benchmark"], 63), 2)
        if trailing_return(window["benchmark"], 63) is not None
        else None,
        "portfolioReturn6mPct": round(trailing_return(window["portfolio"], 126), 2)
        if trailing_return(window["portfolio"], 126) is not None
        else None,
        "benchmarkReturn6mPct": round(trailing_return(window["benchmark"], 126), 2)
        if trailing_return(window["benchmark"], 126) is not None
        else None,
        "dates": [timestamp.strftime("%Y-%m-%d") for timestamp in chart.index],
        "portfolioPath": [round(float(value), 3) for value in chart["portfolio"]],
        "benchmarkPath": [round(float(value), 3) for value in chart["benchmark"]],
        "relativeStrengthPath": [round(float(value), 3) for value in chart["relativeStrength"]],
    }


def refresh(
    input_path: Path,
    output_path: Path,
    period: str,
    benchmark: str,
    portfolio_output: Path | None = None,
) -> tuple[int, int]:
    data = pd.read_csv(input_path)
    holdings_mask = data["Name"].astype(str).str.strip().str.lower() != "total"
    data = clear_technical_columns(data)
    data.loc[holdings_mask, "Yahoo Ticker"] = data.loc[holdings_mask].apply(
        lambda row: row["Yahoo Ticker"] if str(row.get("Yahoo Ticker", "")).strip() else normalize_symbol(row["Name"]),
        axis=1,
    )

    symbols = data.loc[holdings_mask, "Yahoo Ticker"].astype(str).tolist()
    frames = download_frames(symbols + [benchmark], period)
    benchmark_frame = frames.get(benchmark)
    benchmark_close = close_series(benchmark_frame)
    try:
        validate_close_history(benchmark_close, benchmark)
    except Exception:
        benchmark_close = pd.Series(dtype=float)

    downloaded_count = 0
    failed_count = 0
    for index, row in data.loc[holdings_mask].iterrows():
        symbol = str(row["Yahoo Ticker"])
        frame = frames.get(symbol)
        try:
            metrics = analyse_frame(frame, benchmark_close, symbol)
            downloaded_count += 1
        except Exception as exc:
            error = str(exc)
            previous_downloaded = row.get("Technical Downloaded") is True or str(
                row.get("Technical Downloaded")
            ).strip().lower() == "true"
            previous_rsi = clean_number(row.get("RSI 14"))
            if previous_downloaded and math.isfinite(previous_rsi):
                metrics = {column: row.get(column) for column in TECHNICAL_COLUMNS}
                previous_note = str(row.get("Technical Note") or "").strip()
                stale_message = "Latest refresh unavailable; showing the previous successful technical snapshot."
                if stale_message.lower() not in previous_note.lower():
                    metrics["Technical Note"] = f"{previous_note} {stale_message}".strip()
                metrics["Technical Error"] = f"{stale_message} {error}".strip()
            else:
                status, note = unavailable_technical_details(symbol, error)
                metrics = {
                    "Technical Downloaded": False,
                    "Technical Status": status,
                    "Technical Score": None,
                    "Technical Note": note,
                    "Technical Error": error,
                }
            failed_count += 1
        for key, value in metrics.items():
            data.loc[index, key] = value

    data.to_csv(output_path, index=False)
    if portfolio_output is not None:
        portfolio_output.parent.mkdir(parents=True, exist_ok=True)
        snapshot = portfolio_risk_snapshot(data.loc[holdings_mask], frames, benchmark, benchmark_close)
        if not snapshot.get("available") and portfolio_output.exists():
            try:
                previous_snapshot = json.loads(portfolio_output.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError):
                previous_snapshot = {}
            if previous_snapshot.get("available"):
                previous_snapshot["latestRefreshError"] = snapshot.get("error")
                previous_snapshot["dataFreshness"] = "Previous successful market-data snapshot"
                snapshot = previous_snapshot
        portfolio_output.write_text(json.dumps(snapshot, indent=2), encoding="utf-8")
    return downloaded_count, failed_count


def main() -> int:
    parser = argparse.ArgumentParser(description="Refresh client portfolio technical indicators.")
    parser.add_argument("--input", type=Path, default=Path("data/holdings.csv"))
    parser.add_argument("--output", type=Path, default=Path("data/holdings.csv"))
    parser.add_argument("--period", default="2y")
    parser.add_argument("--benchmark", default="^NSEMDCP50")
    parser.add_argument("--portfolio-output", type=Path, default=Path("data/portfolio_risk.json"))
    args = parser.parse_args()

    downloaded, failed = refresh(args.input, args.output, args.period, args.benchmark, args.portfolio_output)
    print(f"Technical indicators downloaded: {downloaded}")
    print(f"Technical fallbacks: {failed}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
