#!/usr/bin/env python3
"""
Build a client dashboard from an NSDL DP holdings report.

The NSDL file is a current holdings snapshot. It does not contain acquisition
price, so the dashboard focuses on current exposure, day movement, and the
technical layer when available.
"""

from __future__ import annotations

import argparse
import datetime as dt
import html
import json
import math
import re
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_XLSX = Path("/Users/anshumanomjhunjhunwala/Downloads/NSDLReport_22042401320260802133818.xlsx")
CLIENT_NAME = "Mr Vikram"

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

QUOTE_COLUMNS = [
    "Price Source",
    "Price Date",
    "Price Timestamp",
    "Quote Type",
    "Quote Status",
    "Quote Note",
    "Previous Close",
    "Session Volume",
    "Volume vs 20D",
    "Intraday Alert",
]

COST_COLUMNS = [
    "P&L Matched",
    "P&L Lots",
    "P&L Quantity",
    "Costed Quantity",
    "Uncosted Quantity",
    "Known Cost Value",
    "Average Recorded Cost",
    "Broker Unrealized P&L",
    "Return on Recorded Cost %",
    "Earliest Buy Date",
    "Latest Buy Date",
    "Cost Basis Coverage %",
    "Quantity Reconciliation %",
    "Cost Basis Status",
    "Cost Basis Note",
]

TICKER_MAP = {
    "ADITYA BIRLA LIFES BRAN L": "ABLBL.NS",
    "ASHOK LEYLAND LTD": "ASHOKLEY.NS",
    "ASTER DM QUALITY CARE LTD": "ASTERDM.NS",
    "ASTRAZENECA PHARMA IND LT": "ASTRAZEN.NS",
    "AVADH SUG & ENERGY LTD": "AVADHSUGAR.NS",
    "BSE LIMITED": "BSE.NS",
    "EASY TRIP PLANNERS LTD": "EASEMYTRIP.NS",
    "EICHER MOTORS LTD": "EICHERMOT.NS",
    "FREDUN PHARMACEUTICALS LIMITED": "FREDUN.NS",
    "GANGES SECURITIES LIMITED": "GANGESSECU.NS",
    "INGERSOLL-RAND INDIA LTD": "INGERRAND.NS",
    "INOX WIND LIMITED": "INOXWIND.NS",
    "LAURUS LABS LIMITED": "LAURUSLABS.NS",
    "MAHINDRA & MAHINDRA LTD": "M&M.NS",
    "MAGADH SUGAR & ENERGY LTD": "MAGADSUGAR.NS",
    "NIPPON L I A M LTD": "NAM-INDIA.NS",
    "NETWORK18 MEDIA & INV LTD": "NETWORK18.NS",
    "NEULAND LAB LTD.": "NEULANDLAB.NS",
    "ONESOURCE SPECL PHARMA L": "ONESOURCE.NS",
    "PEARL GLOBAL IND LIMITED": "PGIL.NS",
    "PTC INDUSTRIES LIMITED": "PTCIL.NS",
    "RAMCO INDUSTRIES LIMITED": "RAMCOIND.NS",
    "STRIDES PHARMA SCI LTD": "STAR.NS",
    "TATA STEEL LIMITED": "TATASTEEL.NS",
    "TATA MOTORS LIMITED": "TMCV.NS",
    "TATA MOTORS PASS VEH LTD": "TMPV.NS",
    "VIYASH SCIENTIFIC LIMITED": "VIYASH.NS",
    "VOLTAS LTD": "VOLTAS.NS",
    "VARDHMAN TEXTILES LIMITED": "VTL.NS",
}

DISPLAY_NAME_MAP = {
    "ADITYA BIRLA LIFES BRAN L": "Aditya Birla Lifestyle Brands",
    "ASHOK LEYLAND LTD": "Ashok Leyland",
    "ASTER DM QUALITY CARE LTD": "Aster DM Quality Care",
    "ASTRAZENECA PHARMA IND LT": "AstraZeneca Pharma India",
    "AVADH SUG & ENERGY LTD": "Avadh Sugar & Energy",
    "BSE LIMITED": "BSE",
    "EASY TRIP PLANNERS LTD": "Easy Trip Planners",
    "EICHER MOTORS LTD": "Eicher Motors",
    "FREDUN PHARMACEUTICALS LIMITED": "Fredun Pharmaceuticals",
    "GANGES SECURITIES LIMITED": "Ganges Securities",
    "INGERSOLL-RAND INDIA LTD": "Ingersoll-Rand India",
    "INOX WIND LIMITED": "Inox Wind",
    "LAURUS LABS LIMITED": "Laurus Labs",
    "MAHINDRA & MAHINDRA LTD": "Mahindra & Mahindra",
    "MAGADH SUGAR & ENERGY LTD": "Magadh Sugar & Energy",
    "NIPPON L I A M LTD": "Nippon Life India AMC",
    "NETWORK18 MEDIA & INV LTD": "Network18 Media & Investments",
    "NEULAND LAB LTD.": "Neuland Laboratories",
    "ONESOURCE SPECL PHARMA L": "OneSource Specialty Pharma",
    "PEARL GLOBAL IND LIMITED": "Pearl Global Industries",
    "PTC INDUSTRIES LIMITED": "PTC Industries",
    "RAMCO INDUSTRIES LIMITED": "Ramco Industries",
    "STRIDES PHARMA SCI LTD": "Strides Pharma Science",
    "TATA STEEL LIMITED": "Tata Steel",
    "TATA MOTORS LIMITED": "Tata Motors",
    "TATA MOTORS PASS VEH LTD": "Tata Motors Passenger Vehicles",
    "VIYASH SCIENTIFIC LIMITED": "Viyash Scientific",
    "VOLTAS LTD": "Voltas",
    "VARDHMAN TEXTILES LIMITED": "Vardhman Textiles",
}

THEME_MAP = {
    "ADITYA BIRLA LIFES BRAN L": "Consumption / Apparel",
    "ASHOK LEYLAND LTD": "Auto",
    "ASTER DM QUALITY CARE LTD": "Healthcare",
    "ASTRAZENECA PHARMA IND LT": "Pharma",
    "AVADH SUG & ENERGY LTD": "Sugar / Energy",
    "BSE LIMITED": "Market Infrastructure",
    "EASY TRIP PLANNERS LTD": "Travel",
    "EICHER MOTORS LTD": "Auto",
    "FREDUN PHARMACEUTICALS LIMITED": "Pharma",
    "GANGES SECURITIES LIMITED": "Financials",
    "INGERSOLL-RAND INDIA LTD": "Industrials",
    "INOX WIND LIMITED": "Renewables",
    "LAURUS LABS LIMITED": "Pharma",
    "MAHINDRA & MAHINDRA LTD": "Auto",
    "MAGADH SUGAR & ENERGY LTD": "Sugar / Energy",
    "NIPPON L I A M LTD": "Asset Management",
    "NETWORK18 MEDIA & INV LTD": "Media",
    "NEULAND LAB LTD.": "Pharma",
    "ONESOURCE SPECL PHARMA L": "Pharma",
    "PEARL GLOBAL IND LIMITED": "Textiles / Apparel",
    "PTC INDUSTRIES LIMITED": "Industrials",
    "RAMCO INDUSTRIES LIMITED": "Building Materials",
    "STRIDES PHARMA SCI LTD": "Pharma",
    "TATA STEEL LIMITED": "Metals",
    "TATA MOTORS LIMITED": "Auto",
    "TATA MOTORS PASS VEH LTD": "Auto",
    "VIYASH SCIENTIFIC LIMITED": "Pharma",
    "VOLTAS LTD": "Consumer Durables",
    "VARDHMAN TEXTILES LIMITED": "Textiles",
}


def clean_number(value: object) -> float:
    text = str(value).replace(",", "").replace("%", "").replace("+", "").strip()
    if text in {"", "nan", "None"}:
        return math.nan
    try:
        return float(text)
    except ValueError:
        return math.nan


def find_meta(raw: pd.DataFrame, label: str) -> str:
    needle = label.lower()
    for _, row in raw.iterrows():
        values = [str(value).strip() for value in row.tolist()]
        for idx, value in enumerate(values):
            if value.lower() == needle and idx + 1 < len(values):
                candidate = values[idx + 1].strip()
                if candidate and candidate.lower() != "nan":
                    return candidate
    return ""


def parse_nsdl_xlsx(path: Path) -> tuple[pd.DataFrame, dict[str, str]]:
    raw = pd.read_excel(path, sheet_name="NSDL Report", header=None)
    header_rows = raw.index[raw.iloc[:, 0].astype(str).str.strip().eq("Name")].tolist()
    if not header_rows:
        raise SystemExit("Could not find the equity holdings table in the NSDL workbook.")

    header_row = header_rows[0]
    header = raw.iloc[header_row].fillna("").astype(str).str.strip().tolist()
    body = raw.iloc[header_row + 1 :].copy()
    body.columns = header
    body = body[body["Name"].notna()].copy()
    body["Name"] = body["Name"].astype(str).str.strip()
    body = body[body["Name"].ne("")]
    total_rows = body.index[body["Name"].str.lower().eq("total")].tolist()
    if total_rows:
        body = body.loc[: total_rows[0] - 1].copy()
    body = body[body["Name"].str.upper().ne("MUTUAL FUNDS")]
    body = body[body["Name"].str.upper().ne("BONDS")]

    rename = {
        "Free Qty": "Free Qty",
        "Pledged Qty": "Pledged Qty",
        "LTP % Change": "Day Change %",
        "LTP": "LTP",
        "Total Qty": "Quantity",
        "LTP Change": "Day Change",
        "Market Value": "Current Value",
    }
    data = body[["Name", *rename.keys()]].rename(columns=rename).copy()
    for column in ["Free Qty", "Pledged Qty", "Day Change %", "LTP", "Quantity", "Day Change", "Current Value"]:
        data[column] = data[column].map(clean_number)

    meta = {
        "client_id": find_meta(raw, "Client ID"),
        "report_date": find_meta(raw, "Date"),
        "source_file": path.name,
    }
    return enrich_holdings(data), meta


def normalize_name(name: object) -> str:
    return re.sub(r"\s+", " ", str(name).strip().upper())


def priority(row: pd.Series) -> str:
    weight = float(row.get("Weight %") or 0)
    day_change = float(row.get("Day Change %") or 0)
    ticker_confidence = str(row.get("Ticker Confidence") or "")
    if weight >= 10 or abs(day_change) >= 5 or ticker_confidence == "Review":
        return "High"
    if weight >= 5 or abs(day_change) >= 2.5:
        return "Medium"
    return "Low"


def bucket(row: pd.Series) -> str:
    weight = float(row.get("Weight %") or 0)
    day_change = float(row.get("Day Change %") or 0)
    if weight >= 10:
        return "Anchor exposure"
    if weight >= 5:
        return "Core exposure"
    if day_change >= 4:
        return "Positive momentum"
    if day_change <= -3:
        return "Event watch"
    return "Monitor"


def suggested_discussion(row: pd.Series) -> str:
    weight = float(row.get("Weight %") or 0)
    day_change = float(row.get("Day Change %") or 0)
    if str(row.get("Ticker Confidence") or "") == "Review":
        return "Verify ticker before relying on automated technical readings."
    if weight >= 10:
        return "Large exposure; confirm that the thesis and risk limit are still appropriate."
    if day_change <= -3:
        return "Sharp daily weakness; check whether this is event-driven or part of wider sector pressure."
    if day_change >= 4:
        return "Strong daily move; watch whether it follows through or simply retraces."
    if weight >= 5:
        return "Meaningful exposure; keep the thesis and technical structure under review."
    return "No urgent action; monitor with the rest of the portfolio."


def enrich_holdings(data: pd.DataFrame) -> pd.DataFrame:
    data = data.copy()
    data["Source Name"] = data["Name"].map(normalize_name)
    data["Display Name"] = data["Source Name"].map(DISPLAY_NAME_MAP).fillna(data["Name"].astype(str).str.title())
    data["Symbol"] = data["Display Name"]
    data["Yahoo Ticker"] = data["Source Name"].map(TICKER_MAP).fillna("")
    data["Theme"] = data["Source Name"].map(THEME_MAP).fillna("Other")
    data["Ticker Confidence"] = data["Yahoo Ticker"].map(lambda value: "Mapped" if str(value).strip() else "Review")
    data["Ticker Note"] = data["Ticker Confidence"].map(
        lambda value: "Ticker mapped for Yahoo/NSE refresh." if value == "Mapped" else "Ticker needs manual verification."
    )
    data["Day P&L"] = data["Quantity"] * data["Day Change"]
    total_value = float(data["Current Value"].sum())
    data["Weight %"] = data["Current Value"] / total_value * 100 if total_value else 0
    data["Portfolio Bucket"] = data.apply(bucket, axis=1)
    data["Coordination Priority"] = data.apply(priority, axis=1)
    data["Suggested Discussion"] = data.apply(suggested_discussion, axis=1)
    data["Technical Downloaded"] = False
    data["Technical Status"] = "Technical refresh pending"
    data["Technical Score"] = None
    data["Technical Note"] = "Technical refresh has not been run yet."
    data["Technical Error"] = ""
    for column in TECHNICAL_COLUMNS:
        if column not in data.columns:
            data[column] = None
    return data


def read_holdings_csv(path: Path) -> tuple[pd.DataFrame, dict[str, str]]:
    data = pd.read_csv(path)
    for column in ["Free Qty", "Pledged Qty", "Day Change %", "LTP", "Quantity", "Day Change", "Current Value", "Day P&L", "Weight %", "Previous Close"]:
        if column in data.columns:
            data[column] = pd.to_numeric(data[column], errors="coerce")
    missing_derived = "Display Name" not in data.columns or "Portfolio Bucket" not in data.columns
    if missing_derived:
        data = enrich_holdings(data)
    else:
        if "Previous Close" in data.columns:
            valid_previous = data["Previous Close"].notna() & data["LTP"].notna() & data["Previous Close"].ne(0)
            data.loc[valid_previous, "Day Change"] = data.loc[valid_previous, "LTP"] - data.loc[valid_previous, "Previous Close"]
            data.loc[valid_previous, "Day Change %"] = (
                data.loc[valid_previous, "LTP"] / data.loc[valid_previous, "Previous Close"] - 1
            ) * 100
        valid_value = data["Quantity"].notna() & data["LTP"].notna()
        data.loc[valid_value, "Current Value"] = data.loc[valid_value, "Quantity"] * data.loc[valid_value, "LTP"]
        total_value = float(data["Current Value"].sum())
        data["Day P&L"] = data["Quantity"] * data["Day Change"]
        data["Weight %"] = data["Current Value"] / total_value * 100 if total_value else 0
        data["Portfolio Bucket"] = data.apply(bucket, axis=1)
        data["Coordination Priority"] = data.apply(priority, axis=1)
        data["Suggested Discussion"] = data.apply(suggested_discussion, axis=1)
        for column in TECHNICAL_COLUMNS:
            if column not in data.columns:
                data[column] = None
    for column in COST_COLUMNS:
        if column not in data.columns:
            data[column] = None
    data = add_signal_profiles(data)
    meta = {
        "client_id": "220424013",
        "report_date": "",
        "source_file": path.name,
    }
    return data, meta


def money(value: object) -> str:
    try:
        return f"Rs {float(value):,.0f}"
    except Exception:
        return "-"


def pct(value: object) -> str:
    try:
        return f"{float(value):.2f}%"
    except Exception:
        return "-"


def tone(value: object) -> str:
    try:
        number = float(value)
    except Exception:
        return "flat"
    if number > 0:
        return "positive"
    if number < 0:
        return "negative"
    return "flat"


def records(data: pd.DataFrame) -> list[dict[str, object]]:
    fields = [
        "Display Name",
        "Source Name",
        "Symbol",
        "Yahoo Ticker",
        "Theme",
        "Free Qty",
        "Pledged Qty",
        "Quantity",
        "LTP",
        "Day Change",
        "Day Change %",
        "Day P&L",
        "Current Value",
        "Weight %",
        "Portfolio Bucket",
        "Coordination Priority",
        "Suggested Discussion",
        "Ticker Confidence",
        "Ticker Note",
        "Evidence Quality",
        "Signal Agreement",
        "Signal Note",
        *QUOTE_COLUMNS,
        *TECHNICAL_COLUMNS,
        *COST_COLUMNS,
    ]
    for field in fields:
        if field not in data.columns:
            data[field] = None
    cleaned = data[fields].where(pd.notna(data[fields]), None)
    return json.loads(cleaned.to_json(orient="records"))


def weighted_metric(data: pd.DataFrame, column: str) -> float | None:
    if column not in data.columns or "Current Value" not in data.columns:
        return None
    values = pd.to_numeric(data[column], errors="coerce")
    weights = pd.to_numeric(data["Current Value"], errors="coerce")
    mask = values.notna() & weights.notna() & (weights > 0)
    if not mask.any() or float(weights[mask].sum()) == 0:
        return None
    return float((values[mask] * weights[mask]).sum() / weights[mask].sum())


def bool_weight(data: pd.DataFrame, column: str, expected: bool = True) -> float | None:
    if column not in data.columns:
        return None
    values = data[column].map(
        lambda value: True if value is True or value == "True" else False if value is False or value == "False" else None
    )
    weights = pd.to_numeric(data["Current Value"], errors="coerce")
    mask = values.notna() & weights.notna() & (weights > 0)
    if not mask.any() or float(weights[mask].sum()) == 0:
        return None
    matched = weights[mask & values.eq(expected)].sum()
    return float(matched / weights[mask].sum() * 100)


def portfolio_rs_estimate(data: pd.DataFrame) -> dict[str, object]:
    paths: list[tuple[float, list[float]]] = []
    for _, row in data.iterrows():
        raw = row.get("RS Trend")
        weight = clean_number(row.get("Current Value"))
        if not isinstance(raw, str) or not raw.strip() or not math.isfinite(weight) or weight <= 0:
            continue
        try:
            values = [float(value) for value in json.loads(raw) if math.isfinite(float(value))]
        except Exception:
            continue
        if len(values) >= 20:
            paths.append((weight, values))

    if not paths:
        return {"available": False, "path": [], "changePct": None, "vs50dPct": None, "coveragePct": 0}

    common_length = min(len(values) for _, values in paths)
    if common_length < 20:
        return {"available": False, "path": [], "changePct": None, "vs50dPct": None, "coveragePct": 0}
    total_weight = sum(weight for weight, _ in paths)
    portfolio_value = float(pd.to_numeric(data["Current Value"], errors="coerce").fillna(0).sum())
    combined = []
    for index in range(-common_length, 0):
        combined.append(sum(weight * values[index] for weight, values in paths) / total_weight)
    first = combined[0]
    normalized = [round(value / first * 100, 2) for value in combined] if first else combined
    latest = normalized[-1]
    ma_window = normalized[-min(50, len(normalized)) :]
    ma_50 = sum(ma_window) / len(ma_window) if ma_window else None
    return {
        "available": True,
        "path": normalized,
        "changePct": latest - 100,
        "vs50dPct": (latest / ma_50 - 1) * 100 if ma_50 else None,
        "coveragePct": total_weight / portfolio_value * 100 if portfolio_value else 0,
        "method": "Current-value-weighted blend of each holding's stock/benchmark ratio",
    }


def signal_profile(row: pd.Series) -> tuple[str, str, str]:
    if not (row.get("Technical Downloaded") is True or str(row.get("Technical Downloaded")) == "True"):
        return "Limited", "Data incomplete", "Technical history is not sufficient for a full signal-agreement reading."

    positive = 0
    negative = 0
    observed = 0
    for column in ("Above 50DMA", "Above 200DMA"):
        value = row.get(column)
        if value is True or value == "True":
            positive += 1
            observed += 1
        elif value is False or value == "False":
            negative += 1
            observed += 1

    rs = clean_number(row.get("RS vs 50D %"))
    if math.isfinite(rs):
        observed += 1
        if rs > 0:
            positive += 1
        elif rs < -3:
            negative += 1

    rsi_value = clean_number(row.get("RSI 14"))
    if math.isfinite(rsi_value):
        observed += 1
        if 45 <= rsi_value <= 70:
            positive += 1
        elif rsi_value < 35 or rsi_value > 78:
            negative += 1

    pnf = str(row.get("P&F Signal") or "")
    if pnf:
        observed += 1
        if "Bullish" in pnf:
            positive += 1
        elif "Bearish" in pnf:
            negative += 1

    quality = "High" if observed >= 5 else "Medium" if observed >= 3 else "Limited"
    if positive >= 4 and negative <= 1:
        state = "Broad positive confirmation"
        note = "Most measured trend and momentum signals agree positively."
    elif negative >= 3 and positive <= 1:
        state = "Broad weakness"
        note = "Several measured trend and momentum signals agree negatively."
    elif math.isfinite(rsi_value) and rsi_value < 35 and (row.get("Above 200DMA") is True or row.get("Above 200DMA") == "True"):
        state = "Mean-reversion watch"
        note = "Momentum is oversold while the longer-term trend remains above the 200DMA."
    else:
        state = "Mixed evidence"
        note = "The indicators do not yet agree strongly enough for a directional classification."
    return quality, state, note


def add_signal_profiles(data: pd.DataFrame) -> pd.DataFrame:
    data = data.copy()
    profiles = data.apply(signal_profile, axis=1)
    data["Evidence Quality"] = profiles.map(lambda value: value[0])
    data["Signal Agreement"] = profiles.map(lambda value: value[1])
    data["Signal Note"] = profiles.map(lambda value: value[2])
    return data


def summary(data: pd.DataFrame, meta: dict[str, str]) -> dict[str, object]:
    current = float(data["Current Value"].sum())
    day_pnl = float(data["Day P&L"].sum())
    base = current - day_pnl
    advancers = int((data["Day Change %"] > 0).sum())
    decliners = int((data["Day Change %"] < 0).sum())
    top_weight = data.sort_values("Weight %", ascending=False).iloc[0] if len(data) else None
    technical_status = data["Technical Status"].astype(str) if "Technical Status" in data.columns else pd.Series([], dtype=str)
    downloaded = int((data["Technical Downloaded"].astype(str) == "True").sum()) if "Technical Downloaded" in data.columns else 0
    known_cost = float(pd.to_numeric(data.get("Known Cost Value"), errors="coerce").fillna(0).sum()) if "Known Cost Value" in data.columns else 0
    unrealized = float(pd.to_numeric(data.get("Broker Unrealized P&L"), errors="coerce").fillna(0).sum()) if "Broker Unrealized P&L" in data.columns else 0
    quantity_series = pd.to_numeric(data.get("Quantity"), errors="coerce").replace(0, math.nan)
    costed_series = pd.to_numeric(data.get("Costed Quantity"), errors="coerce").fillna(0)
    value_series = pd.to_numeric(data.get("Current Value"), errors="coerce").fillna(0)
    cost_fraction = (costed_series / quantity_series).clip(lower=0, upper=1).fillna(0)
    cost_covered_value = float((value_series * cost_fraction).sum())
    normalized_weights = value_series / current if current else value_series * 0
    concentration = float((normalized_weights**2).sum())
    top_five_weight = float(normalized_weights.nlargest(5).sum() * 100)
    return {
        "clientId": meta.get("client_id") or "",
        "reportDate": meta.get("report_date") or "",
        "holdings": int(len(data)),
        "currentValue": current,
        "dayPnl": day_pnl,
        "dayMovePct": day_pnl / base * 100 if base else None,
        "advancers": advancers,
        "decliners": decliners,
        "highPriority": int((data["Coordination Priority"] == "High").sum()),
        "mappedTickers": int((data["Ticker Confidence"] == "Mapped").sum()),
        "technicalDownloaded": downloaded,
        "technicalLeaders": int(technical_status.isin(["Leader / hold", "Constructive"]).sum()),
        "technicalLaggards": int(technical_status.isin(["Risk review", "Loss + weak structure"]).sum()),
        "topHolding": None if top_weight is None else str(top_weight["Display Name"]),
        "topWeightPct": None if top_weight is None else float(top_weight["Weight %"]),
        "knownCostValue": known_cost or None,
        "matchedUnrealizedPnl": unrealized or None,
        "returnOnRecordedCostPct": unrealized / known_cost * 100 if known_cost else None,
        "costBasisValueCoveragePct": cost_covered_value / current * 100 if current else None,
        "weightedRsi": weighted_metric(data, "RSI 14"),
        "weightedRsVs50": weighted_metric(data, "RS vs 50D %"),
        "weightedRs3m": weighted_metric(data, "RS 3M %"),
        "weightAbove50dma": bool_weight(data, "Above 50DMA"),
        "weightAbove200dma": bool_weight(data, "Above 200DMA"),
        "topFiveWeightPct": top_five_weight,
        "effectiveHoldings": 1 / concentration if concentration else None,
    }


def grouped_rows(data: pd.DataFrame, group_col: str) -> list[dict[str, object]]:
    total = float(data["Current Value"].sum())
    rows = []
    for name, group in data.groupby(group_col, dropna=False):
        value = float(group["Current Value"].sum())
        rows.append(
            {
                "name": str(name),
                "count": int(len(group)),
                "value": value,
                "weightPct": value / total * 100 if total else 0,
                "dayMovePct": float(group["Day P&L"].sum() / (value - group["Day P&L"].sum()) * 100)
                if value != float(group["Day P&L"].sum())
                else None,
            }
        )
    return sorted(rows, key=lambda item: item["weightPct"], reverse=True)


def action_queue(data: pd.DataFrame) -> pd.DataFrame:
    rank = {"High": 0, "Medium": 1, "Low": 2}
    queue = data[data["Coordination Priority"].isin(["High", "Medium"])].copy()
    queue["rank"] = queue["Coordination Priority"].map(rank)
    return queue.sort_values(["rank", "Weight %", "Day Change %"], ascending=[True, False, True])[
        [
            "Display Name",
            "Yahoo Ticker",
            "Coordination Priority",
            "Portfolio Bucket",
            "Theme",
            "Weight %",
            "Day Change %",
            "Day P&L",
            "Current Value",
            "Suggested Discussion",
        ]
    ]


def load_json(path: Path | None) -> dict[str, object]:
    if path is None or not path.exists():
        return {}
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    return value if isinstance(value, dict) else {}


def dashboard_html(
    data: pd.DataFrame,
    meta: dict[str, str],
    source_name: str,
    portfolio_risk: dict[str, object],
    pnl_meta: dict[str, object],
) -> str:
    generated_at = dt.datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%Y-%m-%d %H:%M IST")
    price_times = (
        data.get("Price Timestamp", pd.Series(dtype=object)).dropna().astype(str).str.strip()
    )
    price_times = price_times[price_times != ""]
    price_dates = data.get("Price Date", pd.Series(dtype=object)).dropna().astype(str).str.strip()
    price_dates = price_dates[price_dates != ""]
    data_as_of = (
        sorted(price_times.tolist())[-1]
        if not price_times.empty
        else sorted(price_dates.tolist())[-1] if not price_dates.empty else "unavailable"
    )
    refresh_times = data.get("Last Refreshed", pd.Series(dtype=object)).dropna().astype(str).str.strip()
    refresh_times = refresh_times[refresh_times != ""]
    refreshed_at = sorted(refresh_times.tolist())[-1] if not refresh_times.empty else "unavailable"
    payload = {
        "title": f"{CLIENT_NAME} Holdings Dashboard",
        "generatedAt": generated_at,
        "dataAsOf": data_as_of,
        "refreshedAt": refreshed_at,
        "source": source_name,
        "summary": summary(data, meta),
        "portfolioRs": portfolio_rs_estimate(data),
        "portfolioRisk": portfolio_risk,
        "pnlMeta": pnl_meta,
        "themes": grouped_rows(data, "Theme"),
        "buckets": grouped_rows(data, "Portfolio Bucket"),
        "holdings": records(data),
    }
    payload_json = json.dumps(payload, ensure_ascii=True)
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html.escape(CLIENT_NAME)} Holdings Dashboard</title>
  <style>
    :root {{
      --bg: #f5f6f1;
      --panel: #fffefa;
      --ink: #171916;
      --muted: #626960;
      --line: #dce2d7;
      --green: #0e7259;
      --red: #b33e45;
      --amber: #936807;
      --blue: #235b8f;
    }}
    * {{ box-sizing: border-box; }}
    body {{ margin: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: var(--ink); background: var(--bg); }}
    header {{ background: #fbfcf8; border-bottom: 1px solid var(--line); }}
    .wrap {{ max-width: 1440px; margin: 0 auto; padding: 22px 24px; }}
    .topbar {{ display: flex; justify-content: space-between; gap: 16px; align-items: flex-start; }}
    h1 {{ margin: 0; font-size: 35px; line-height: 1.05; letter-spacing: 0; }}
    h2 {{ margin: 0; font-size: 22px; letter-spacing: 0; }}
    .muted {{ color: var(--muted); line-height: 1.45; }}
    .header-actions {{ display: flex; flex-direction: column; gap: 10px; align-items: flex-end; }}
    .badge {{ border: 1px solid var(--line); border-radius: 999px; padding: 8px 12px; background: #f8faf5; color: var(--muted); white-space: nowrap; font-weight: 700; }}
    .refresh-box {{ display: flex; gap: 8px; align-items: center; justify-content: flex-end; flex-wrap: wrap; }}
    .refresh-button {{ min-height: 38px; border: 1px solid var(--line); border-radius: 7px; padding: 8px 12px; background: #fff; color: var(--ink); font: inherit; font-weight: 750; cursor: pointer; }}
    .refresh-button:hover {{ background: #f7faf4; }}
    .refresh-status {{ color: var(--muted); font-size: 13px; white-space: nowrap; }}
    .metrics {{ display: grid; grid-template-columns: repeat(6, minmax(145px, 1fr)); gap: 12px; margin-top: 18px; }}
    .metric {{ min-height: 102px; padding: 15px; background: #fff; border: 1px solid var(--line); border-radius: 8px; }}
    .label {{ color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: .08em; font-weight: 760; }}
    .value {{ margin-top: 13px; font-size: 25px; font-weight: 780; line-height: 1.12; }}
    .positive {{ color: var(--green); }}
    .negative {{ color: var(--red); }}
    .flat {{ color: var(--ink); }}
    section {{ background: #fff; border: 1px solid var(--line); border-radius: 8px; margin: 16px 0; overflow: hidden; }}
    .section-head {{ padding: 17px 18px; display: flex; justify-content: space-between; gap: 12px; align-items: center; border-bottom: 1px solid var(--line); }}
    .controls {{ display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }}
    input, select {{ min-height: 40px; border: 1px solid var(--line); border-radius: 7px; padding: 9px 11px; font: inherit; background: #fbfcf8; }}
    table {{ width: 100%; border-collapse: collapse; }}
    th, td {{ padding: 11px 13px; border-bottom: 1px solid #edf0ea; text-align: left; vertical-align: top; }}
    th {{ color: var(--muted); background: #fbfcf8; font-size: 12px; letter-spacing: .06em; text-transform: uppercase; }}
    td.num {{ text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }}
    tbody tr {{ cursor: pointer; }}
    tbody tr:hover, tr.selected {{ background: #f7faf4; }}
    .small-table {{ overflow: auto; max-height: 620px; }}
    .pill {{ display: inline-flex; border: 1px solid var(--line); border-radius: 999px; padding: 4px 9px; white-space: nowrap; font-size: 13px; font-weight: 700; background: #f8f9f5; }}
    .priority-high, .bucket-event-watch {{ color: var(--red); border-color: #edc3c8; background: #fff3f4; }}
    .priority-medium, .bucket-core-exposure {{ color: var(--amber); border-color: #e8d7aa; background: #fff9ea; }}
    .priority-low, .bucket-monitor {{ color: var(--muted); }}
    .bucket-anchor-exposure {{ color: var(--blue); border-color: #c5d9ea; background: #edf5fc; }}
    .bucket-positive-momentum {{ color: var(--green); border-color: #bbdacc; background: #eef9f4; }}
    .technical-status-leader-hold {{ color: var(--green); border-color: #bbdacc; background: #eef9f4; }}
    .technical-status-constructive {{ color: var(--blue); border-color: #c5d9ea; background: #edf5fc; }}
    .technical-status-risk-review, .technical-status-loss-weak-structure {{ color: var(--red); border-color: #edc3c8; background: #fff3f4; }}
    .technical-status-monitor, .technical-status-technical-refresh-pending,
    .technical-status-technical-refresh-unavailable {{ color: var(--muted); }}
    .technical-status-limited-price-history,
    .technical-status-limited-post-demerger-history {{ color: var(--amber); border-color: #e8d7aa; background: #fff9ea; }}
    .technical-status-price-history-unavailable,
    .technical-status-ticker-review-required {{ color: var(--red); border-color: #edc3c8; background: #fff3f4; }}
    .technical-status-technical-data-review {{ color: var(--amber); border-color: #e8d7aa; background: #fff9ea; }}
    .leadership-leader {{ color: var(--green); border-color: #bbdacc; background: #eef9f4; }}
    .leadership-constructive, .leadership-partial-constructive {{ color: var(--blue); border-color: #c5d9ea; background: #edf5fc; }}
    .leadership-monitor {{ color: var(--amber); border-color: #e8d7aa; background: #fff9ea; }}
    .leadership-weak, .leadership-partial-weak, .leadership-weak-structure {{ color: var(--red); border-color: #edc3c8; background: #fff3f4; }}
    .leadership-data-unavailable {{ color: var(--muted); }}
    .risk-flag-no-active-flag {{ color: var(--green); border-color: #bbdacc; background: #eef9f4; }}
    .risk-flag-watch {{ color: var(--amber); border-color: #e8d7aa; background: #fff9ea; }}
    .risk-flag-risk-review {{ color: var(--red); border-color: #edc3c8; background: #fff3f4; }}
    .risk-flag-data-limited {{ color: var(--muted); }}
    .technical-context {{ margin-top: 6px; max-width: 230px; color: var(--muted); font-size: 12px; line-height: 1.35; }}
    .group-stack {{ padding: 4px 18px 18px; }}
    .group-row {{ display: grid; grid-template-columns: 210px 1fr 76px 92px; gap: 12px; align-items: center; padding: 10px 0; border-bottom: 1px solid #edf0ea; }}
    .group-row:last-child {{ border-bottom: 0; }}
    .bar {{ height: 12px; background: #edf0ea; border-radius: 999px; overflow: hidden; }}
    .bar span {{ display: block; height: 100%; background: var(--green); border-radius: inherit; }}
    .guide-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: 12px; padding: 18px; }}
    .guide-item {{ border: 1px solid var(--line); border-radius: 8px; padding: 14px; background: #fbfcf8; min-height: 132px; }}
    .guide-item strong {{ display: block; font-size: 16px; margin-bottom: 8px; }}
    .guide-item p {{ margin: 0; color: var(--muted); line-height: 1.45; }}
    .analytics-grid {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); }}
    .analytics-panel {{ min-width: 0; padding: 18px; }}
    .analytics-panel + .analytics-panel {{ border-left: 1px solid var(--line); }}
    .analytics-strip {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); border-bottom: 1px solid var(--line); }}
    .analytics-stat {{ min-height: 92px; padding: 16px 18px; border-right: 1px solid var(--line); }}
    .analytics-stat:nth-child(4n) {{ border-right: 0; }}
    .analytics-stat:nth-child(n+5) {{ border-top: 1px solid var(--line); }}
    .analytics-stat:last-child {{ border-right: 0; }}
    .analytics-stat .value {{ font-size: 22px; }}
    .chart-title {{ display: flex; justify-content: space-between; gap: 12px; align-items: baseline; margin-bottom: 10px; }}
    .chart-title strong {{ font-size: 16px; }}
    .line-chart {{ display: block; width: 100%; aspect-ratio: 2.8 / 1; min-height: 180px; overflow: visible; }}
    .line-chart .grid {{ stroke: #e2e7de; stroke-width: 1; }}
    .line-chart .axis {{ fill: var(--muted); font-size: 11px; }}
    .line-chart .series {{ fill: none; stroke-width: 2.4; vector-effect: non-scaling-stroke; }}
    .chart-legend {{ display: flex; gap: 16px; flex-wrap: wrap; margin-top: 8px; color: var(--muted); font-size: 13px; }}
    .chart-legend span::before {{ content: ''; display: inline-block; width: 16px; height: 3px; margin-right: 6px; vertical-align: middle; background: var(--legend); }}
    .analytics-note {{ padding: 14px 18px; border-top: 1px solid var(--line); color: var(--muted); line-height: 1.5; }}
    .analytics-note strong {{ color: var(--ink); }}
    .empty-state {{ min-height: 180px; display: grid; place-items: center; padding: 28px; text-align: center; color: var(--muted); background: #fbfcf8; border: 1px dashed var(--line); }}
    .spark-cell {{ min-width: 150px; }}
    .sparkline {{ display: block; width: 148px; height: 42px; overflow: visible; }}
    .sparkline path.line {{ fill: none; stroke-width: 2.4; }}
    .sparkline path.area {{ opacity: .14; }}
    .sparkline line {{ stroke: #d9dfd4; stroke-width: 1; stroke-dasharray: 3 3; }}
    .spark-label {{ display: block; margin-top: 4px; color: var(--muted); font-size: 12px; line-height: 1.25; white-space: nowrap; }}
    .holdings-list {{ width: 100%; }}
    .holding-item {{ border-bottom: 1px solid #e5e9e1; }}
    .holding-item:last-child {{ border-bottom: 0; }}
    .holding-item.expanded {{ background: #fbfcf8; }}
    .holding-summary {{ width: 100%; display: grid; grid-template-columns: minmax(210px, 1.5fr) minmax(120px, .85fr) minmax(120px, .8fr) minmax(110px, .75fr) 82px 120px 140px 28px; gap: 12px; align-items: center; padding: 15px 18px; border: 0; background: transparent; color: var(--ink); font: inherit; text-align: left; cursor: pointer; }}
    .holding-summary.day-summary {{ grid-template-columns: minmax(210px, 1.55fr) minmax(120px, .85fr) minmax(120px, .8fr) 82px 120px 150px 28px; }}
    .holding-summary.technical-summary {{ grid-template-columns: minmax(205px, 1.4fr) minmax(120px, .85fr) minmax(120px, .8fr) 100px 76px minmax(110px, .75fr) 68px 28px; }}
    .holding-summary:hover {{ background: #f7faf4; }}
    .holding-summary:focus-visible {{ outline: 3px solid #a8c9b7; outline-offset: -3px; }}
    .holding-identity strong {{ display: block; font-size: 16px; line-height: 1.25; }}
    .holding-identity .muted {{ margin-top: 4px; font-size: 13px; }}
    .identity-priority {{ display: flex; gap: 7px; flex-wrap: wrap; align-items: center; margin-top: 8px; }}
    .identity-priority .summary-label {{ margin: 0; }}
    .classification-summary .pill {{ max-width: 100%; white-space: normal; text-align: center; }}
    .detail-classification {{ display: inline-flex; gap: 7px; flex-wrap: wrap; align-items: center; }}
    .detail-classification .summary-label {{ margin: 0; }}
    .summary-metric {{ min-width: 0; }}
    .summary-label {{ display: block; margin-bottom: 5px; color: var(--muted); font-size: 11px; font-weight: 760; letter-spacing: .07em; text-transform: uppercase; }}
    .summary-value {{ display: block; font-size: 15px; font-weight: 760; overflow-wrap: anywhere; }}
    .holding-chevron {{ color: var(--muted); font-size: 22px; text-align: center; transform: rotate(0deg); transition: transform .16s ease; }}
    .holding-item.expanded .holding-chevron {{ transform: rotate(180deg); }}
    .inline-detail {{ padding: 18px; border-top: 1px solid var(--line); background: #fff; }}
    .inline-detail-head {{ display: flex; justify-content: space-between; gap: 16px; align-items: flex-start; }}
    .detail-close {{ width: 36px; height: 36px; flex: 0 0 36px; border: 1px solid var(--line); border-radius: 7px; background: #fff; color: var(--muted); font-size: 22px; line-height: 1; cursor: pointer; }}
    .detail-close:hover {{ background: #f7faf4; color: var(--ink); }}
    .detail-title strong {{ display: block; font-size: 28px; line-height: 1.1; }}
    .kv {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 10px; margin-top: 15px; }}
    .kv div {{ min-height: 76px; border: 1px solid var(--line); border-radius: 8px; padding: 12px; background: #fbfcf8; }}
    .kv .k {{ color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: .07em; font-weight: 760; }}
    .kv .v {{ display: block; margin-top: 8px; font-size: 19px; font-weight: 740; overflow-wrap: anywhere; }}
    .note {{ margin-top: 14px; padding-top: 14px; border-top: 1px solid var(--line); color: var(--muted); line-height: 1.5; }}
    .note strong {{ color: var(--ink); }}
    @media (max-width: 1120px) {{
      .metrics {{ grid-template-columns: repeat(3, minmax(145px, 1fr)); }}
      .guide-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
      .analytics-strip {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
      .analytics-stat:nth-child(2) {{ border-right: 0; }}
      .analytics-stat:nth-child(odd) {{ border-right: 1px solid var(--line); }}
      .analytics-stat:nth-child(even) {{ border-right: 0; }}
      .analytics-stat:nth-child(n+3) {{ border-top: 1px solid var(--line); }}
      .holding-summary {{ grid-template-columns: minmax(200px, 1.4fr) minmax(115px, .8fr) minmax(115px, .75fr) 88px 115px 135px 28px; }}
      .holding-summary.day-summary {{ grid-template-columns: minmax(200px, 1.45fr) minmax(115px, .8fr) minmax(115px, .75fr) 82px 112px 140px 28px; }}
      .holding-summary.technical-summary {{ grid-template-columns: minmax(195px, 1.35fr) minmax(115px, .8fr) minmax(115px, .75fr) 95px 72px 28px; }}
      .holding-summary .holding-signal {{ display: none; }}
      .holding-summary.technical-summary .tertiary-summary {{ display: none; }}
      .kv {{ grid-template-columns: repeat(3, minmax(0, 1fr)); }}
    }}
    @media (max-width: 760px) {{
      .wrap {{ padding: 18px 14px; }}
      .topbar, .section-head {{ display: block; }}
      .header-actions {{ align-items: flex-start; margin-top: 12px; }}
      .badge {{ display: inline-flex; margin-top: 12px; }}
      h1 {{ font-size: 30px; }}
      .metrics {{ grid-template-columns: repeat(2, minmax(130px, 1fr)); }}
      .guide-grid {{ grid-template-columns: 1fr; }}
      .analytics-grid, .analytics-strip {{ grid-template-columns: 1fr; }}
      .analytics-panel + .analytics-panel {{ border-left: 0; border-top: 1px solid var(--line); }}
      .analytics-stat {{ border-right: 0; border-bottom: 1px solid var(--line); }}
      .analytics-stat:nth-child(odd) {{ border-right: 0; }}
      .analytics-stat:last-child {{ border-bottom: 0; }}
      .controls {{ margin-top: 12px; }}
      table {{ min-width: 900px; }}
      .holding-summary {{ grid-template-columns: minmax(0, 1fr) minmax(110px, .55fr) 26px; gap: 10px; padding: 14px; }}
      .holding-summary.day-summary, .holding-summary.technical-summary {{ grid-template-columns: minmax(0, 1fr) minmax(110px, .55fr) 26px; }}
      .holding-summary .holding-weight, .holding-summary .holding-value, .holding-summary .holding-signal {{ display: none; }}
      .holding-summary .holding-day {{ grid-column: 2; grid-row: 1; text-align: right; }}
      .holding-summary .leadership-summary {{ grid-column: 1; grid-row: 2; }}
      .holding-summary .risk-summary {{ grid-column: 2; grid-row: 2; }}
      .compact-summary .secondary-summary, .compact-summary .tertiary-summary {{ display: none; }}
      .compact-summary .primary-summary {{ grid-column: 2; grid-row: 1; text-align: right; }}
      .holding-chevron {{ grid-column: 3; grid-row: 1; }}
      .inline-detail {{ padding: 16px 14px; }}
      .detail-title strong {{ font-size: 23px; }}
      .group-row, .kv {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
    }}
    @media (max-width: 480px) {{
      .kv {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>
  <header>
    <div class="wrap">
      <div class="topbar">
        <div>
          <h1>{html.escape(CLIENT_NAME)} Holdings Dashboard</h1>
          <div class="muted">NSDL client {html.escape(str(meta.get("client_id") or ""))} · Market data {html.escape(data_as_of)} · Refreshed {html.escape(refreshed_at)}</div>
        </div>
        <div class="header-actions">
          <div class="badge">Client view</div>
          <div class="refresh-box">
            <button class="refresh-button" type="button" onclick="refreshPage()">Check Latest Data</button>
            <span id="refreshStatus" class="refresh-status">Checks for a published update in 15:00</span>
          </div>
        </div>
      </div>
      <div id="metrics" class="metrics"></div>
    </div>
  </header>
  <main class="wrap">
    <section>
      <div class="section-head">
        <h2>Portfolio Risk &amp; Relative Strength</h2>
        <div class="muted">Portfolio-level behaviour versus the selected benchmark</div>
      </div>
      <div id="riskMetrics" class="analytics-strip"></div>
      <div class="analytics-grid">
        <div class="analytics-panel">
          <div class="chart-title"><strong>Portfolio vs Benchmark</strong><span id="performanceCaption" class="muted"></span></div>
          <div id="performanceChart"></div>
        </div>
        <div class="analytics-panel">
          <div class="chart-title"><strong>Portfolio Relative Strength</strong><span class="muted">Rising means outperformance</span></div>
          <div id="portfolioRsChart"></div>
        </div>
      </div>
      <div id="riskNote" class="analytics-note"></div>
    </section>

    <section>
      <div class="section-head">
        <h2>Cost Basis &amp; P&amp;L Reconciliation</h2>
        <div class="muted">Broker P&amp;L report matched to the current NSDL holdings</div>
      </div>
      <div id="costMetrics" class="analytics-strip"></div>
      <div id="costNote" class="analytics-note"></div>
    </section>

    <section>
      <div class="section-head">
        <h2>Stock Classification Guide</h2>
        <div class="muted">Two separate readings used in every stock list below</div>
      </div>
      <div class="guide-grid">
        <div class="guide-item">
          <strong>Trend / leadership</strong>
          <p>Measures whether the stock is participating in a durable trend using relative strength, moving averages, RSI, and Point &amp; Figure. <b>Leader</b> means strong technical alignment; it is not an instruction to hold or buy. <b>Partial</b> means shorter price history limits long-term confirmation.</p>
        </div>
        <div class="guide-item">
          <strong>Risk flag</strong>
          <p>Separately highlights weak structure, bearish P&amp;F, or a low technical score. A stock can rise strongly today and still show <b>Risk review</b> because a one-day gain does not by itself repair the broader trend.</p>
        </div>
      </div>
      <div class="analytics-note"><strong>Read the columns together:</strong> Day move describes what happened today. Trend / leadership describes the broader technical position. Risk flag identifies where closer review is warranted.</div>
    </section>

    <section>
          <div class="section-head">
            <h2>Best Day Moves</h2>
            <div class="muted">Largest positive moves in the current NSDL snapshot</div>
          </div>
          <div id="bestDayMovesList" class="holdings-list compact-holdings"></div>
        </section>

    <section>
          <div class="section-head">
            <h2>Worst Day Moves</h2>
            <div class="muted">Largest negative moves that may need attention</div>
          </div>
          <div id="worstDayMovesList" class="holdings-list compact-holdings"></div>
        </section>

    <section>
          <div class="section-head">
            <h2>Theme Exposure</h2>
            <div class="muted">Where the portfolio is concentrated</div>
          </div>
          <div id="themes" class="group-stack"></div>
        </section>

    <section>
          <div class="section-head">
            <h2>Portfolio Buckets</h2>
            <div class="muted">Review priority based on exposure and day movement</div>
          </div>
          <div id="buckets" class="group-stack"></div>
        </section>

    <section>
          <div class="section-head">
            <h2>Technical Reading Guide</h2>
            <div class="muted">What the technical layer is measuring</div>
          </div>
          <div class="guide-grid">
            <div class="guide-item">
              <strong>Relative Strength</strong>
              <p>Compares the stock with Nifty Midcap. Positive RS vs 50D means the stock is leading the broader market.</p>
            </div>
            <div class="guide-item">
              <strong>RSI 14</strong>
              <p>Momentum gauge. 40-60 is often a healthy reset, above 70 can be extended, below 30 can be oversold.</p>
            </div>
            <div class="guide-item">
              <strong>Moving Averages</strong>
              <p>50DMA tracks medium-term trend. 200DMA tracks long-term structure. Above both is usually stronger.</p>
            </div>
            <div class="guide-item">
              <strong>Point &amp; Figure</strong>
              <p>Filters noise and highlights breakouts or breakdowns. It is a structure check, not a standalone signal.</p>
            </div>
            <div class="guide-item">
              <strong>Signal Agreement</strong>
              <p>Summarises whether trend, relative strength, momentum, and P&amp;F point in the same direction. It is evidence confluence, not a calibrated probability or instruction.</p>
            </div>
          </div>
        </section>

    <section>
          <div class="section-head">
            <h2>Technical Leaders</h2>
            <div class="muted">Relative strength, RSI, moving averages, and P&amp;F structure</div>
          </div>
          <div id="technicalLeadersList" class="holdings-list compact-holdings"></div>
        </section>

    <section>
          <div class="section-head">
            <h2>Technical Laggards</h2>
            <div class="muted">Weak structure or weak relative strength that needs attention</div>
          </div>
          <div id="technicalLaggardsList" class="holdings-list compact-holdings"></div>
        </section>

    <section>
          <div class="section-head">
            <h2>Review Priority Guide</h2>
            <div class="muted">How to read the High, Medium, and Low labels used in All Holdings</div>
          </div>
          <div class="guide-grid">
            <div class="guide-item">
              <strong><span class="pill priority-high">High</span> Review first</strong>
              <p>Triggered when a holding is at least 10% of the portfolio, moves at least 5% in one day, or its market ticker needs confirmation.</p>
            </div>
            <div class="guide-item">
              <strong><span class="pill priority-medium">Medium</span> Review next</strong>
              <p>Triggered when a holding is 5% to under 10% of the portfolio, or moves 2.5% to under 5% in one day, provided no High trigger applies.</p>
            </div>
            <div class="guide-item">
              <strong><span class="pill priority-low">Low</span> Routine monitoring</strong>
              <p>No High or Medium review trigger is active. The holding remains visible for normal portfolio monitoring and can still be opened for full detail.</p>
            </div>
          </div>
          <div class="analytics-note"><strong>How to use it:</strong> Priority highlights where discussion may be most useful. It does not predict direction and does not recommend a transaction.</div>
        </section>

    <section>
          <div class="section-head">
            <h2>All Holdings</h2>
            <div class="controls">
              <input id="search" placeholder="Search stock" oninput="renderHoldings()">
              <select id="priority" onchange="renderHoldings()"><option value="">All priorities</option><option>High</option><option>Medium</option><option>Low</option></select>
              <select id="theme" onchange="renderHoldings()"><option value="">All themes</option></select>
              <select id="leadershipStatus" onchange="renderHoldings()"><option value="">All leadership states</option></select>
              <select id="riskFlag" onchange="renderHoldings()"><option value="">All risk flags</option></select>
              <select id="holdingSort" onchange="renderHoldings()">
                <option value="weight">Sort: Weight</option>
                <option value="alpha">Sort: A-Z</option>
                <option value="dayPnl">Sort: Day P&amp;L</option>
                <option value="overallPnl">Sort: Overall P&amp;L</option>
              </select>
            </div>
          </div>
          <div id="holdingsList" class="holdings-list"></div>
    </section>
  </main>

  <script>
    const DATA = {payload_json};
    const AUTO_REFRESH_MS = 15 * 60 * 1000;
    const autoRefreshStartedAt = Date.now();
    let expandedHoldingKey = null;

    const money = value => Number.isFinite(Number(value)) ? 'Rs ' + Number(value).toLocaleString('en-IN', {{ maximumFractionDigits: 0 }}) : '-';
    const pct = value => Number.isFinite(Number(value)) ? Number(value).toFixed(2) + '%' : '-';
    const num = value => Number.isFinite(Number(value)) ? Number(value).toLocaleString('en-IN', {{ maximumFractionDigits: 2 }}) : '-';
    const safe = value => value === null || value === undefined || value === '' ? '-' : String(value);
    const tone = value => Number(value) > 0 ? 'positive' : Number(value) < 0 ? 'negative' : 'flat';
    const slug = value => String(value || 'unknown').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'unknown';
    const pill = (type, value) => `<span class="pill ${{type}}-${{slug(value)}}">${{safe(value)}}</span>`;
    const boolText = value => value === true || value === 'True' ? 'Yes' : value === false || value === 'False' ? 'No' : '-';

    function metric(label, value, toneClass = 'flat') {{
      return `<div class="metric"><div class="label">${{label}}</div><div class="value ${{toneClass}}">${{value}}</div></div>`;
    }}

    function renderMetrics() {{
      const s = DATA.summary;
      document.getElementById('metrics').innerHTML = [
        metric('Current Value', money(s.currentValue)),
        metric('Day P&L', money(s.dayPnl), tone(s.dayPnl)),
        metric('Day Move', pct(s.dayMovePct), tone(s.dayMovePct)),
        metric('A/D', `${{s.advancers}} / ${{s.decliners}}`),
        metric('Holdings', safe(s.holdings)),
        metric('Top Weight', `${{safe(s.topHolding)}}<br>${{pct(s.topWeightPct)}}`)
      ].join('');
    }}

    function analyticsStat(label, value, toneClass = 'flat') {{
      return `<div class="analytics-stat"><div class="label">${{label}}</div><div class="value ${{toneClass}}">${{value}}</div></div>`;
    }}

    function lineChart(series, labels = []) {{
      const usable = series.map(item => ({{
        ...item,
        values: Array.isArray(item.values) ? item.values.map(Number) : []
      }})).filter(item => item.values.length > 2 && item.values.every(Number.isFinite));
      if (!usable.length) return '<div class="empty-state">Run the local technical refresh to calculate this chart.</div>';
      const all = usable.flatMap(item => item.values);
      const min = Math.min(...all);
      const max = Math.max(...all);
      const span = max - min || 1;
      const width = 640, height = 230, left = 48, right = 14, top = 16, bottom = 30;
      const plotWidth = width - left - right;
      const plotHeight = height - top - bottom;
      const y = value => top + (max - value) / span * plotHeight;
      const grid = [0, .25, .5, .75, 1].map(fraction => {{
        const value = max - span * fraction;
        const yy = top + plotHeight * fraction;
        return `<line class="grid" x1="${{left}}" y1="${{yy}}" x2="${{width - right}}" y2="${{yy}}"></line><text class="axis" x="4" y="${{yy + 4}}">${{value.toFixed(1)}}</text>`;
      }}).join('');
      const paths = usable.map(item => {{
        const count = item.values.length;
        const path = item.values.map((value, index) => {{
          const xx = left + index * (plotWidth / Math.max(1, count - 1));
          return `${{index ? 'L' : 'M'}}${{xx.toFixed(1)}},${{y(value).toFixed(1)}}`;
        }}).join(' ');
        return `<path class="series" d="${{path}}" stroke="${{item.color}}"></path>`;
      }}).join('');
      const firstLabel = labels.length ? safe(labels[0]) : '';
      const lastLabel = labels.length ? safe(labels[labels.length - 1]) : '';
      const legend = usable.map(item => `<span style="--legend:${{item.color}}">${{safe(item.name)}}</span>`).join('');
      return `<svg class="line-chart" viewBox="0 0 ${{width}} ${{height}}" role="img" aria-label="${{usable.map(item => item.name).join(' and ')}} chart">${{grid}}${{paths}}<text class="axis" x="${{left}}" y="${{height - 7}}">${{firstLabel}}</text><text class="axis" text-anchor="end" x="${{width - right}}" y="${{height - 7}}">${{lastLabel}}</text></svg><div class="chart-legend">${{legend}}</div>`;
    }}

    function renderPortfolioAnalytics() {{
      const risk = DATA.portfolioRisk || {{}};
      const s = DATA.summary;
      if (risk.available) {{
        document.getElementById('riskMetrics').innerHTML = [
          analyticsStat('Beta', num(risk.beta)),
          analyticsStat('Annualized Volatility', pct(risk.volatilityAnnualizedPct)),
          analyticsStat('Maximum Drawdown', pct(risk.maxDrawdownPct), tone(risk.maxDrawdownPct)),
          analyticsStat('Tracking Error', pct(risk.trackingErrorPct)),
          analyticsStat('Annualized Alpha Est.', pct(risk.alphaAnnualizedPct), tone(risk.alphaAnnualizedPct)),
          analyticsStat('Information Ratio', num(risk.informationRatio), tone(risk.informationRatio)),
          analyticsStat('Top 5 Concentration', pct(s.topFiveWeightPct)),
          analyticsStat('Effective Holdings', num(s.effectiveHoldings))
        ].join('');
        document.getElementById('performanceChart').innerHTML = lineChart([
          {{ name: 'Portfolio', values: risk.portfolioPath, color: 'var(--green)' }},
          {{ name: risk.benchmarkLabel || 'Benchmark', values: risk.benchmarkPath, color: 'var(--blue)' }}
        ], risk.dates || []);
        document.getElementById('portfolioRsChart').innerHTML = lineChart([
          {{ name: 'Portfolio relative strength', values: risk.relativeStrengthPath, color: 'var(--amber)' }}
        ], risk.dates || []);
        document.getElementById('performanceCaption').textContent = `${{safe(risk.sessions)}} sessions`;
        document.getElementById('riskNote').innerHTML = `<strong>Interpretation:</strong> Beta describes market sensitivity; volatility and drawdown describe the observed risk path. Correlation is ${{num(risk.correlation)}}; upside/downside capture are ${{pct(risk.upsideCapturePct)}} / ${{pct(risk.downsideCapturePct)}}. Three-month portfolio/benchmark returns are ${{pct(risk.portfolioReturn3mPct)}} / ${{pct(risk.benchmarkReturn3mPct)}}; six-month returns are ${{pct(risk.portfolioReturn6mPct)}} / ${{pct(risk.benchmarkReturn6mPct)}}. ${{safe(risk.method)}} Coverage is ${{pct(risk.coveragePct)}}. These are historical estimates, not forecasts.`;
        return;
      }}

      document.getElementById('riskMetrics').innerHTML = [
        analyticsStat('Weighted RSI 14', num(s.weightedRsi)),
        analyticsStat('Weighted RS vs 50D', pct(s.weightedRsVs50), tone(s.weightedRsVs50)),
        analyticsStat('Weight Above 50DMA', pct(s.weightAbove50dma)),
        analyticsStat('Weight Above 200DMA', pct(s.weightAbove200dma)),
        analyticsStat('Top 5 Concentration', pct(s.topFiveWeightPct)),
        analyticsStat('Effective Holdings', num(s.effectiveHoldings)),
        analyticsStat('Technical Coverage', `${{safe(s.technicalDownloaded)}} / ${{safe(s.holdings)}}`),
        analyticsStat('Portfolio RS vs 50D', pct((DATA.portfolioRs || {{}}).vs50dPct), tone((DATA.portfolioRs || {{}}).vs50dPct))
      ].join('');
      document.getElementById('performanceChart').innerHTML = '<div class="empty-state">Beta, volatility, drawdown, and benchmark performance are awaiting a local market-data refresh.</div>';
      const estimate = DATA.portfolioRs || {{}};
      document.getElementById('portfolioRsChart').innerHTML = estimate.available
        ? lineChart([{{ name: 'Estimated portfolio relative strength', values: estimate.path, color: 'var(--amber)' }}])
        : '<div class="empty-state">Portfolio relative-strength history is not yet available.</div>';
      document.getElementById('performanceCaption').textContent = 'Awaiting refresh';
      document.getElementById('riskNote').innerHTML = `<strong>Current state:</strong> ${{safe(risk.error || 'Portfolio risk history has not been calculated yet.')}} The holding-level indicators above remain available. Signal agreement is evidence confluence, not a calibrated probability or an instruction to transact.`;
    }}

    function renderCostBasis() {{
      const s = DATA.summary;
      const p = DATA.pnlMeta || {{}};
      document.getElementById('costMetrics').innerHTML = [
        analyticsStat('Known Recorded Cost', money(s.knownCostValue)),
        analyticsStat('Matched Unrealized P&L', money(s.matchedUnrealizedPnl), tone(s.matchedUnrealizedPnl)),
        analyticsStat('Return on Recorded Cost', pct(s.returnOnRecordedCostPct), tone(s.returnOnRecordedCostPct)),
        analyticsStat('Value with Cost Coverage', pct(s.costBasisValueCoveragePct))
      ].join('');
      const names = Array.isArray(p.unmatchedNames) && p.unmatchedNames.length ? p.unmatchedNames.join(', ') : 'none';
      const period = p.fromDate && p.toDate ? `${{safe(p.fromDate)}} to ${{safe(p.toDate)}}` : 'not stated';
      document.getElementById('costNote').innerHTML = `<strong>Report period:</strong> ${{period}}. <strong>Broker-reported overall P&amp;L:</strong> ${{money(p.reportedOverallPnl)}}. <strong>Manual review:</strong> ${{safe(p.unmatchedLots || 0)}} legacy/corporate-action lots (${{names}}), carrying ${{money(p.unmatchedKnownCostValue)}} of historical cost, were deliberately left unassigned. This prevents false average prices after restructurings or demergers.`;
    }}

    function renderGroupRows(id, rows) {{
      const max = Math.max(...rows.map(row => Number(row.weightPct) || 0), 1);
      document.getElementById(id).innerHTML = rows.map(row => `
        <div class="group-row">
          <div><strong>${{safe(row.name)}}</strong><div class="muted">${{row.count}} holdings</div></div>
          <div class="bar"><span style="width:${{Math.max(4, Number(row.weightPct) / max * 100)}}%"></span></div>
          <div class="${{tone(row.dayMovePct)}}">${{pct(row.dayMovePct)}}</div>
          <div class="num">${{pct(row.weightPct)}}</div>
        </div>
      `).join('');
    }}

    function holdingKey(section, name) {{
      return `${{section}}::${{name}}`;
    }}

    function holdingListItem(row, section, summaryClass, summaryMarkup) {{
      const key = holdingKey(section, row['Display Name']);
      const expanded = expandedHoldingKey === key;
      const dimensions = technicalDimensions(row);
      const showPriority = section === 'all-holdings';
      const escapedName = safe(row['Display Name']).replace(/'/g, "\\\\'");
      return `
        <article id="${{section}}-${{slug(row['Display Name'])}}" class="holding-item ${{expanded ? 'expanded' : ''}}">
          <button class="holding-summary ${{summaryClass}}" type="button" aria-expanded="${{expanded ? 'true' : 'false'}}" onclick="toggleHolding('${{escapedName}}', '${{section}}')">
            <span class="holding-identity">
              <strong>${{safe(row['Display Name'])}}</strong>
              <span class="muted">${{safe(row.Theme)}} · ${{safe(row['Yahoo Ticker'])}}</span>
              ${{showPriority ? `<span class="identity-priority"><span class="summary-label">Review priority</span>${{pill('priority', row['Coordination Priority'])}}</span>` : ''}}
            </span>
            <span class="summary-metric classification-summary leadership-summary">
              <span class="summary-label">Trend / leadership</span>
              ${{pill('leadership', dimensions.leadership)}}
            </span>
            <span class="summary-metric classification-summary risk-summary">
              <span class="summary-label">Risk flag</span>
              ${{pill('risk-flag', dimensions.risk)}}
            </span>
            ${{summaryMarkup}}
            <span class="holding-chevron" aria-hidden="true">⌄</span>
          </button>
          ${{expanded ? `<div class="inline-detail">${{holdingDetailMarkup(row, section)}}</div>` : ''}}
        </article>
      `;
    }}

    function renderDayMoves() {{
      const positiveRows = [...DATA.holdings]
        .filter(row => (Number(row['Day Change %']) || 0) > 0)
        .sort((a, b) => (Number(b['Day Change %']) || 0) - (Number(a['Day Change %']) || 0))
        .slice(0, 6);
      const negativeRows = [...DATA.holdings]
        .filter(row => (Number(row['Day Change %']) || 0) < 0)
        .sort((a, b) => (Number(a['Day Change %']) || 0) - (Number(b['Day Change %']) || 0))
        .slice(0, 6);

      const renderList = (id, rows, section, moveLabel, emptyMessage) => {{
        document.getElementById(id).innerHTML = rows.length
          ? rows.map(row => holdingListItem(
              row,
              section,
              'compact-summary day-summary',
              `
                <span class="summary-metric secondary-summary"><span class="summary-label">Weight</span><span class="summary-value">${{pct(row['Weight %'])}}</span></span>
                <span class="summary-metric tertiary-summary"><span class="summary-label">Value</span><span class="summary-value">${{money(row['Current Value'])}}</span></span>
                <span class="summary-metric primary-summary"><span class="summary-label">${{moveLabel}}</span><span class="summary-value ${{tone(row['Day Change %'])}}">${{pct(row['Day Change %'])}} · ${{money(row['Day P&L'])}}</span></span>
              `
            )).join('')
          : `<div class="empty-state">${{emptyMessage}}</div>`;
      }};

      renderList('bestDayMovesList', positiveRows, 'best-day-moves', 'Gain', 'No positive day moves are present in this snapshot.');
      renderList('worstDayMovesList', negativeRows, 'worst-day-moves', 'Decline', 'No negative day moves are present in this snapshot.');
    }}

    function sparkline(valuesRaw) {{
      if (!valuesRaw) return '-';
      let values;
      try {{ values = typeof valuesRaw === 'string' ? JSON.parse(valuesRaw) : valuesRaw; }} catch (err) {{ return '-'; }}
      if (!Array.isArray(values) || values.length < 3) return '-';
      const width = 148, height = 42, pad = 3;
      const nums = values.map(Number).filter(Number.isFinite);
      if (nums.length < 3) return '-';
      const min = Math.min(...nums), max = Math.max(...nums);
      const span = max - min || 1;
      const points = nums.map((value, idx) => {{
        const x = pad + idx * ((width - pad * 2) / Math.max(1, nums.length - 1));
        const y = height - pad - ((value - min) / span) * (height - pad * 2);
        return [x, y];
      }});
      const line = points.map((point, idx) => `${{idx ? 'L' : 'M'}}${{point[0].toFixed(1)}},${{point[1].toFixed(1)}}`).join(' ');
      const area = `${{line}} L${{points.at(-1)[0].toFixed(1)}},${{height - pad}} L${{points[0][0].toFixed(1)}},${{height - pad}} Z`;
      const latest = nums.at(-1);
      const first = nums[0];
      const color = latest >= first ? 'var(--green)' : 'var(--red)';
      return `<svg class="sparkline" viewBox="0 0 ${{width}} ${{height}}" aria-hidden="true">
        <line x1="0" y1="${{height / 2}}" x2="${{width}}" y2="${{height / 2}}"></line>
        <path class="area" d="${{area}}" fill="${{color}}"></path>
        <path class="line" d="${{line}}" stroke="${{color}}"></path>
      </svg><span class="spark-label">${{pct(((latest / first) - 1) * 100)}} over trend window</span>`;
    }}

    function technicalDownloaded(row) {{
      return row['Technical Downloaded'] === true || row['Technical Downloaded'] === 'True';
    }}

    function technicalAvailability(row) {{
      if (technicalDownloaded(row)) {{
        return {{
          label: safe(row['Technical Status']),
          short: '',
          detail: safe(row['Technical Note'])
        }};
      }}

      const ticker = String(row['Yahoo Ticker'] || '').trim().toUpperCase();
      const error = String(row['Technical Error'] || '').trim().toLowerCase();
      const storedStatus = String(row['Technical Status'] || '').trim();
      const storedNote = String(row['Technical Note'] || '').trim();

      if (ticker === 'TMCV.NS' && error.includes('not enough daily price history')) {{
        return {{
          label: 'Limited post-demerger history',
          short: 'Standalone listing history is still building.',
          detail: "Tata Motors' commercial-vehicle listing has a shorter standalone history after the demerger. Current price and daily movement are available, but the full long-term technical ranking will remain limited until enough trading history accumulates."
        }};
      }}
      if (error.includes('not enough daily price history')) {{
        return {{
          label: 'Limited price history',
          short: 'Not enough history for the complete indicator set.',
          detail: 'The listing does not yet have enough daily trading history for the complete technical set. Current price and daily movement are still available; long-term indicators will appear as more history accumulates.'
        }};
      }}
      if (!ticker || ticker === 'NAN' || ticker === 'NONE') {{
        return {{
          label: 'Ticker review required',
          short: 'Ticker must be confirmed before analysis.',
          detail: 'The market ticker has not been confirmed, so technical indicators are paused to avoid showing analysis for the wrong security. Current NSDL holding data remains available.'
        }};
      }}
      if (['no price data', 'no data', 'possibly delisted', 'empty'].some(phrase => error.includes(phrase))) {{
        return {{
          label: 'Price history unavailable',
          short: 'The provider returned no usable price history.',
          detail: 'The data provider returned no usable daily price history in the latest refresh. Current NSDL holding data remains visible, and the technical layer will retry automatically.'
        }};
      }}
      if (storedStatus && storedStatus !== 'Not downloaded') {{
        return {{
          label: storedStatus,
          short: storedNote && storedNote !== 'Technical refresh has not been run yet.' ? storedNote : '',
          detail: storedNote || 'The technical layer will retry during the next scheduled refresh.'
        }};
      }}
      return {{
        label: 'Technical refresh unavailable',
        short: 'The latest indicator refresh could not be completed.',
        detail: 'The technical layer could not be calculated in the latest refresh. Current holding value and daily movement remain available; the next refresh will retry the analysis.'
      }};
    }}

    function technicalDimensions(row) {{
      const coverage = technicalAvailability(row);
      if (!technicalDownloaded(row)) {{
        return {{ leadership: 'Data unavailable', risk: 'Data limited', coverage }};
      }}

      const rawStatus = String(row['Technical Status'] || '').trim();
      const partial = rawStatus.startsWith('Partial:');
      const status = rawStatus.replace(/^Partial:\s*/, '');
      let leadership = 'Monitor';
      let risk = 'Watch';

      if (status === 'Leader / hold') {{
        leadership = 'Leader';
        risk = 'No active flag';
      }} else if (status === 'Constructive') {{
        leadership = 'Constructive';
        risk = 'No active flag';
      }} else if (status === 'Risk review') {{
        leadership = 'Weak';
        risk = 'Risk review';
      }} else if (status === 'Loss + weak structure') {{
        leadership = 'Weak structure';
        risk = 'Risk review';
      }}

      if (partial) leadership = `Partial: ${{leadership}}`;
      return {{ leadership, risk, coverage }};
    }}

    function renderTechnicalTables() {{
      const downloaded = DATA.holdings.filter(technicalDownloaded);
      const leaders = [...downloaded].sort((a, b) => (Number(b['Technical Score']) || 0) - (Number(a['Technical Score']) || 0)).slice(0, 10);
      const laggards = [...downloaded].sort((a, b) => (Number(a['Technical Score']) || 999) - (Number(b['Technical Score']) || 999)).slice(0, 10);
      const leaderRows = leaders.length ? leaders : DATA.holdings.slice(0, 8);
      const laggardRows = laggards.length ? laggards : DATA.holdings.slice(-8);

      document.getElementById('technicalLeadersList').innerHTML = leaderRows.map(row => holdingListItem(
        row,
        'technical-leaders',
        'compact-summary technical-summary',
        `
          <span class="summary-metric primary-summary"><span class="summary-label">RS vs 50D</span><span class="summary-value ${{tone(row['RS vs 50D %'])}}">${{pct(row['RS vs 50D %'])}}</span></span>
          <span class="summary-metric secondary-summary"><span class="summary-label">RSI 14</span><span class="summary-value">${{num(row['RSI 14'])}}</span></span>
          <span class="summary-metric tertiary-summary"><span class="summary-label">P&amp;F</span><span class="summary-value">${{safe(row['P&F Signal'])}}</span></span>
          <span class="summary-metric tertiary-summary"><span class="summary-label">Score</span><span class="summary-value">${{safe(row['Technical Score'])}}</span></span>
        `
      )).join('');

      document.getElementById('technicalLaggardsList').innerHTML = laggardRows.map(row => holdingListItem(
        row,
        'technical-laggards',
        'compact-summary technical-summary',
        `
          <span class="summary-metric tertiary-summary spark-cell"><span class="summary-label">RS trend</span>${{sparkline(row['RS Trend'])}}</span>
          <span class="summary-metric primary-summary"><span class="summary-label">RS vs 50D</span><span class="summary-value ${{tone(row['RS vs 50D %'])}}">${{pct(row['RS vs 50D %'])}}</span></span>
          <span class="summary-metric secondary-summary"><span class="summary-label">RSI 14</span><span class="summary-value">${{num(row['RSI 14'])}}</span></span>
          <span class="summary-metric tertiary-summary"><span class="summary-label">Score</span><span class="summary-value">${{safe(row['Technical Score'])}}</span></span>
        `
      )).join('');
    }}

    function populateFilters() {{
      const themes = [...new Set(DATA.holdings.map(row => row.Theme).filter(Boolean))].sort();
      const leadershipStates = [...new Set(DATA.holdings.map(row => technicalDimensions(row).leadership).filter(Boolean))].sort();
      const riskFlags = [...new Set(DATA.holdings.map(row => technicalDimensions(row).risk).filter(Boolean))].sort();
      document.getElementById('theme').innerHTML += themes.map(value => `<option>${{safe(value)}}</option>`).join('');
      document.getElementById('leadershipStatus').innerHTML += leadershipStates.map(value => `<option>${{safe(value)}}</option>`).join('');
      document.getElementById('riskFlag').innerHTML += riskFlags.map(value => `<option>${{safe(value)}}</option>`).join('');
    }}

    function filteredHoldings() {{
      const search = document.getElementById('search').value.trim().toLowerCase();
      const priority = document.getElementById('priority').value;
      const theme = document.getElementById('theme').value;
      const leadershipStatus = document.getElementById('leadershipStatus').value;
      const riskFlag = document.getElementById('riskFlag').value;
      const sortBy = document.getElementById('holdingSort').value;
      const rows = DATA.holdings.filter(row => {{
        const dimensions = technicalDimensions(row);
        const text = `${{row['Display Name']}} ${{row['Yahoo Ticker']}} ${{row.Theme}} ${{row['Signal Agreement']}}`.toLowerCase();
        if (search && !text.includes(search)) return false;
        if (priority && row['Coordination Priority'] !== priority) return false;
        if (theme && row.Theme !== theme) return false;
        if (leadershipStatus && dimensions.leadership !== leadershipStatus) return false;
        if (riskFlag && dimensions.risk !== riskFlag) return false;
        return true;
      }});
      const byName = (a, b) => String(a['Display Name'] || '').localeCompare(String(b['Display Name'] || ''));
      const byNumberDesc = key => (a, b) => {{
        const left = Number(a[key]);
        const right = Number(b[key]);
        const aValid = Number.isFinite(left);
        const bValid = Number.isFinite(right);
        if (aValid && bValid && right !== left) return right - left;
        if (aValid !== bValid) return aValid ? -1 : 1;
        return byName(a, b);
      }};
      const sorters = {{
        alpha: byName,
        dayPnl: byNumberDesc('Day P&L'),
        overallPnl: byNumberDesc('Broker Unrealized P&L'),
        weight: byNumberDesc('Weight %'),
      }};
      return rows.sort(sorters[sortBy] || sorters.weight);
    }}

    function renderHoldings() {{
      const rows = filteredHoldings();
      const list = document.getElementById('holdingsList');
      if (!rows.length) {{
        list.innerHTML = '<div class="empty-state">No holdings match the selected filters.</div>';
        return;
      }}
      list.innerHTML = rows.map(row => holdingListItem(
        row,
        'all-holdings',
        '',
        `
          <span class="summary-metric holding-signal"><span class="summary-label">Signal agreement</span><span class="summary-value">${{safe(row['Signal Agreement'])}}</span></span>
          <span class="summary-metric holding-weight"><span class="summary-label">Weight</span><span class="summary-value">${{pct(row['Weight %'])}}</span></span>
          <span class="summary-metric holding-value"><span class="summary-label">Value</span><span class="summary-value">${{money(row['Current Value'])}}</span></span>
          <span class="summary-metric holding-day"><span class="summary-label">Day move</span><span class="summary-value ${{tone(row['Day Change %'])}}">${{pct(row['Day Change %'])}} · ${{money(row['Day P&L'])}}</span></span>
        `
      )).join('');
    }}

    function renderAllStockSections() {{
      renderDayMoves();
      renderTechnicalTables();
      renderHoldings();
    }}

    function toggleHolding(name, section) {{
      const key = holdingKey(section, name);
      expandedHoldingKey = expandedHoldingKey === key ? null : key;
      renderAllStockSections();
    }}

    function closeExpandedHolding() {{
      expandedHoldingKey = null;
      renderAllStockSections();
    }}

    function holdingDetailMarkup(row, section) {{
      const coverage = technicalAvailability(row);
      const dimensions = technicalDimensions(row);
      const showPriority = section === 'all-holdings';
      const quoteStatus = String(row['Quote Status'] || '').trim();
      const quoteNote = String(row['Quote Note'] || '').trim();
      const priceAsOf = row['Price Timestamp'] || row['Price Date'] || '';
      const technicalAsOf = row['Technical As Of'] || '';
      const intradayAlert = String(row['Intraday Alert'] || '').trim();
      const priceDataNote = quoteStatus && quoteStatus !== 'Updated'
        ? `<div class="note"><strong>Price data:</strong> ${{safe(quoteStatus)}}. ${{safe(quoteNote || 'The dashboard retained the previous value and will retry on the next refresh.')}}</div>`
        : '';
      return `
        <div class="inline-detail-head">
          <div class="detail-title">
            <strong>${{safe(row['Display Name'])}}</strong>
            <div class="muted">${{safe(row.Theme)}} · ${{safe(row['Yahoo Ticker'])}}</div>
            <div style="margin-top:10px; display:flex; gap:8px; flex-wrap:wrap;">
              ${{showPriority ? pill('priority', row['Coordination Priority']) : ''}}
              <span class="detail-classification"><span class="summary-label">Trend / leadership</span>${{pill('leadership', dimensions.leadership)}}</span>
              <span class="detail-classification"><span class="summary-label">Risk flag</span>${{pill('risk-flag', dimensions.risk)}}</span>
              ${{pill('bucket', row['Portfolio Bucket'])}}
              ${{pill('technical-status', row['Evidence Quality'] + ' evidence')}}
            </div>
          </div>
          <button class="detail-close" type="button" aria-label="Close holding details" title="Close details" onclick="closeExpandedHolding()">×</button>
        </div>
        <div class="kv">
          <div><span class="k">Current Value</span><span class="v">${{money(row['Current Value'])}}</span></div>
          <div><span class="k">Weight</span><span class="v">${{pct(row['Weight %'])}}</span></div>
          <div><span class="k">Day P&L</span><span class="v ${{tone(row['Day P&L'])}}">${{money(row['Day P&L'])}}</span></div>
          <div><span class="k">Day Move</span><span class="v ${{tone(row['Day Change %'])}}">${{pct(row['Day Change %'])}}</span></div>
          <div><span class="k">Quantity / Pledged</span><span class="v">${{num(row.Quantity)}} / ${{num(row['Pledged Qty'])}}</span></div>
          <div><span class="k">LTP</span><span class="v">${{num(row.LTP)}}</span></div>
          <div><span class="k">Previous Close</span><span class="v">${{num(row['Previous Close'])}}</span></div>
          <div><span class="k">Volume vs 20D</span><span class="v">${{Number.isFinite(Number(row['Volume vs 20D'])) ? Number(row['Volume vs 20D']).toFixed(2) + 'x' : '-'}}</span></div>
          <div><span class="k">Price Status</span><span class="v">${{safe(quoteStatus || 'Updated')}}</span></div>
          <div><span class="k">Live Quote As Of</span><span class="v">${{safe(priceAsOf || 'Unavailable')}}</span></div>
          <div><span class="k">EOD Technicals As Of</span><span class="v">${{safe(technicalAsOf || 'Unavailable')}}</span></div>
          <div><span class="k">RS vs 50D</span><span class="v ${{tone(row['RS vs 50D %'])}}">${{pct(row['RS vs 50D %'])}}</span></div>
          <div><span class="k">RSI 14</span><span class="v">${{num(row['RSI 14'])}}</span></div>
          <div><span class="k">50DMA / 200DMA</span><span class="v">${{boolText(row['Above 50DMA'])}} / ${{boolText(row['Above 200DMA'])}}</span></div>
          <div><span class="k">P&F</span><span class="v">${{safe(row['P&F Signal'])}}</span></div>
          <div><span class="k">Signal Agreement</span><span class="v">${{safe(row['Signal Agreement'])}}</span></div>
          <div><span class="k">Recorded Cost</span><span class="v">${{money(row['Known Cost Value'])}}</span></div>
          <div><span class="k">Average Recorded Cost</span><span class="v">${{money(row['Average Recorded Cost'])}}</span></div>
          <div><span class="k">Broker Unrealized P&L</span><span class="v ${{tone(row['Broker Unrealized P&L'])}}">${{money(row['Broker Unrealized P&L'])}}</span></div>
          <div><span class="k">Return on Recorded Cost</span><span class="v ${{tone(row['Return on Recorded Cost %'])}}">${{pct(row['Return on Recorded Cost %'])}}</span></div>
          <div><span class="k">Cost Coverage</span><span class="v">${{pct(row['Cost Basis Coverage %'])}}</span></div>
          <div><span class="k">Recorded Buy Dates</span><span class="v">${{safe(row['Earliest Buy Date'])}} to ${{safe(row['Latest Buy Date'])}}</span></div>
        </div>
        ${{intradayAlert ? `<div class="note"><strong>Intraday exception:</strong> ${{safe(intradayAlert)}}. This is a live alert and does not alter the EOD technical ranking.</div>` : ''}}
        <div class="note"><strong>Discussion:</strong> ${{safe(row['Suggested Discussion'])}}</div>
        <div class="note"><strong>Signal reading:</strong> ${{safe(row['Signal Note'])}} This classification describes indicator agreement; it is not a probability or a transaction instruction.</div>
        <div class="note"><strong>Technical coverage:</strong> ${{safe(coverage.detail)}}</div>
        ${{priceDataNote}}
        <div class="note"><strong>Cost basis:</strong> ${{safe(row['Cost Basis Status'])}}. ${{safe(row['Cost Basis Note'])}}</div>
        <div class="note"><strong>Data note:</strong> NSDL provides current quantity, LTP, day movement, and market value. Recorded acquisition cost and unrealized P&L are added from the supplied broker P&L report where the security and quantity can be reconciled safely.</div>
      `;
    }}

    function refreshPage() {{
      const url = new URL(window.location.href);
      url.searchParams.set('refresh', String(Date.now()));
      window.location.replace(url.toString());
    }}

    function updateCountdown() {{
      const elapsed = Date.now() - autoRefreshStartedAt;
      const remaining = Math.max(0, AUTO_REFRESH_MS - elapsed);
      const minutes = String(Math.floor(remaining / 60000)).padStart(2, '0');
      const seconds = String(Math.floor((remaining % 60000) / 1000)).padStart(2, '0');
      document.getElementById('refreshStatus').textContent = `Checks for a published update in ${{minutes}}:${{seconds}}`;
      if (remaining <= 0) refreshPage();
    }}

    renderMetrics();
    renderPortfolioAnalytics();
    renderCostBasis();
    renderGroupRows('themes', DATA.themes);
    renderGroupRows('buckets', DATA.buckets);
    renderDayMoves();
    populateFilters();
    renderTechnicalTables();
    renderHoldings();
    updateCountdown();
    setInterval(updateCountdown, 1000);
  </script>
</body>
</html>
"""


def write_outputs(
    data: pd.DataFrame,
    meta: dict[str, str],
    output_dir: Path,
    source_name: str,
    write_csv: bool,
    portfolio_risk: dict[str, object],
    pnl_meta: dict[str, object],
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "data").mkdir(exist_ok=True)
    if write_csv:
        data.to_csv(output_dir / "data" / "holdings.csv", index=False)
    action_queue(data).to_csv(output_dir / "data" / "action_queue.csv", index=False)
    (output_dir / "index.html").write_text(
        dashboard_html(data, meta, source_name, portfolio_risk, pnl_meta), encoding="utf-8"
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Build NSDL client holdings dashboard.")
    parser.add_argument("--input-xlsx", type=Path, default=DEFAULT_XLSX)
    parser.add_argument("--input-csv", type=Path)
    parser.add_argument("--output-dir", type=Path, default=ROOT)
    parser.add_argument("--no-write-csv", action="store_true")
    parser.add_argument("--portfolio-risk", type=Path, default=ROOT / "data" / "portfolio_risk.json")
    parser.add_argument("--pnl-summary", type=Path, default=ROOT / "data" / "pnl_summary.json")
    args = parser.parse_args()

    if args.input_csv:
        data, meta = read_holdings_csv(args.input_csv)
        source_name = args.input_csv.name
    else:
        data, meta = parse_nsdl_xlsx(args.input_xlsx)
        source_name = args.input_xlsx.name

    write_outputs(
        data,
        meta,
        args.output_dir,
        source_name,
        not args.no_write_csv,
        load_json(args.portfolio_risk),
        load_json(args.pnl_summary),
    )
    print(f"Built dashboard: {args.output_dir / 'index.html'}")
    print(f"Holdings: {len(data)}")
    print(f"Market value: {money(data['Current Value'].sum())}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
