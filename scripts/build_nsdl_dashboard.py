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

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_XLSX = Path("/Users/anshumanomjhunjhunwala/Downloads/NSDLReport_22042401320260802133818.xlsx")
CLIENT_NAME = "Mr Vikram"

TECHNICAL_COLUMNS = [
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
    data["Technical Status"] = "Not downloaded"
    data["Technical Score"] = None
    data["Technical Note"] = "Technical refresh has not been run yet."
    data["Technical Error"] = ""
    for column in TECHNICAL_COLUMNS:
        if column not in data.columns:
            data[column] = None
    return data


def read_holdings_csv(path: Path) -> tuple[pd.DataFrame, dict[str, str]]:
    data = pd.read_csv(path)
    for column in ["Free Qty", "Pledged Qty", "Day Change %", "LTP", "Quantity", "Day Change", "Current Value", "Day P&L", "Weight %"]:
        if column in data.columns:
            data[column] = pd.to_numeric(data[column], errors="coerce")
    missing_derived = "Display Name" not in data.columns or "Portfolio Bucket" not in data.columns
    if missing_derived:
        data = enrich_holdings(data)
    else:
        total_value = float(data["Current Value"].sum())
        data["Day P&L"] = data["Quantity"] * data["Day Change"]
        data["Weight %"] = data["Current Value"] / total_value * 100 if total_value else 0
        data["Portfolio Bucket"] = data.apply(bucket, axis=1)
        data["Coordination Priority"] = data.apply(priority, axis=1)
        data["Suggested Discussion"] = data.apply(suggested_discussion, axis=1)
        for column in TECHNICAL_COLUMNS:
            if column not in data.columns:
                data[column] = None
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
        *TECHNICAL_COLUMNS,
    ]
    for field in fields:
        if field not in data.columns:
            data[field] = None
    cleaned = data[fields].where(pd.notna(data[fields]), None)
    return json.loads(cleaned.to_json(orient="records"))


def summary(data: pd.DataFrame, meta: dict[str, str]) -> dict[str, object]:
    current = float(data["Current Value"].sum())
    day_pnl = float(data["Day P&L"].sum())
    base = current - day_pnl
    advancers = int((data["Day Change %"] > 0).sum())
    decliners = int((data["Day Change %"] < 0).sum())
    top_weight = data.sort_values("Weight %", ascending=False).iloc[0] if len(data) else None
    technical_status = data["Technical Status"].astype(str) if "Technical Status" in data.columns else pd.Series([], dtype=str)
    downloaded = int((data["Technical Downloaded"].astype(str) == "True").sum()) if "Technical Downloaded" in data.columns else 0
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


def dashboard_html(data: pd.DataFrame, meta: dict[str, str], source_name: str) -> str:
    generated_at = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    payload = {
        "title": f"{CLIENT_NAME} Holdings Dashboard",
        "generatedAt": generated_at,
        "source": source_name,
        "summary": summary(data, meta),
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
    .layout {{ display: grid; grid-template-columns: minmax(0, 1.6fr) minmax(370px, .86fr); gap: 16px; align-items: start; }}
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
    .technical-status-monitor, .technical-status-not-downloaded {{ color: var(--muted); }}
    .group-stack {{ padding: 4px 18px 18px; }}
    .group-row {{ display: grid; grid-template-columns: 210px 1fr 76px 92px; gap: 12px; align-items: center; padding: 10px 0; border-bottom: 1px solid #edf0ea; }}
    .group-row:last-child {{ border-bottom: 0; }}
    .bar {{ height: 12px; background: #edf0ea; border-radius: 999px; overflow: hidden; }}
    .bar span {{ display: block; height: 100%; background: var(--green); border-radius: inherit; }}
    .guide-grid {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; padding: 18px; }}
    .guide-item {{ border: 1px solid var(--line); border-radius: 8px; padding: 14px; background: #fbfcf8; min-height: 132px; }}
    .guide-item strong {{ display: block; font-size: 16px; margin-bottom: 8px; }}
    .guide-item p {{ margin: 0; color: var(--muted); line-height: 1.45; }}
    .spark-cell {{ min-width: 150px; }}
    .sparkline {{ display: block; width: 148px; height: 42px; overflow: visible; }}
    .sparkline path.line {{ fill: none; stroke-width: 2.4; }}
    .sparkline path.area {{ opacity: .14; }}
    .sparkline line {{ stroke: #d9dfd4; stroke-width: 1; stroke-dasharray: 3 3; }}
    .spark-label {{ display: block; margin-top: 4px; color: var(--muted); font-size: 12px; line-height: 1.25; white-space: nowrap; }}
    .detail {{ position: sticky; top: 14px; }}
    .detail-body {{ padding: 18px; }}
    .detail-title strong {{ display: block; font-size: 28px; line-height: 1.1; }}
    .kv {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px; margin-top: 15px; }}
    .kv div {{ min-height: 76px; border: 1px solid var(--line); border-radius: 8px; padding: 12px; background: #fbfcf8; }}
    .kv .k {{ color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: .07em; font-weight: 760; }}
    .kv .v {{ display: block; margin-top: 8px; font-size: 19px; font-weight: 740; overflow-wrap: anywhere; }}
    .note {{ margin-top: 14px; padding-top: 14px; border-top: 1px solid var(--line); color: var(--muted); line-height: 1.5; }}
    .note strong {{ color: var(--ink); }}
    @media (max-width: 1120px) {{
      .metrics {{ grid-template-columns: repeat(3, minmax(145px, 1fr)); }}
      .guide-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
      .layout {{ grid-template-columns: 1fr; }}
      .detail {{ position: static; }}
    }}
    @media (max-width: 760px) {{
      .wrap {{ padding: 18px 14px; }}
      .topbar, .section-head {{ display: block; }}
      .header-actions {{ align-items: flex-start; margin-top: 12px; }}
      .badge {{ display: inline-flex; margin-top: 12px; }}
      h1 {{ font-size: 30px; }}
      .metrics {{ grid-template-columns: repeat(2, minmax(130px, 1fr)); }}
      .guide-grid {{ grid-template-columns: 1fr; }}
      .controls {{ margin-top: 12px; }}
      table {{ min-width: 1260px; }}
      .group-row, .kv {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>
  <header>
    <div class="wrap">
      <div class="topbar">
        <div>
          <h1>{html.escape(CLIENT_NAME)} Holdings Dashboard</h1>
          <div class="muted">NSDL client {html.escape(str(meta.get("client_id") or ""))} · Report date {html.escape(str(meta.get("report_date") or "not shown"))} · Generated {html.escape(generated_at)}</div>
        </div>
        <div class="header-actions">
          <div class="badge">Client view</div>
          <div class="refresh-box">
            <button class="refresh-button" type="button" onclick="refreshPage()">Refresh Page</button>
            <span id="refreshStatus" class="refresh-status">Auto refresh in 30:00</span>
          </div>
        </div>
      </div>
      <div id="metrics" class="metrics"></div>
    </div>
  </header>
  <main class="wrap">
    <div class="layout">
      <div>
        <section>
          <div class="section-head">
            <h2>Day Leaders</h2>
            <div class="muted">Best and worst moves in the current NSDL snapshot</div>
          </div>
          <div class="small-table">
            <table id="leadersTable">
              <thead><tr><th>Stock</th><th>Theme</th><th>Weight</th><th>Day Move</th><th>Day P&L</th><th>Value</th></tr></thead>
              <tbody></tbody>
            </table>
          </div>
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
          </div>
        </section>

        <section>
          <div class="section-head">
            <h2>Technical Leaders</h2>
            <div class="muted">Relative strength, RSI, moving averages, and P&amp;F structure</div>
          </div>
          <div class="small-table">
            <table id="technicalLeadersTable">
              <thead><tr><th>Stock</th><th>Status</th><th>RS vs 50D</th><th>RSI</th><th>P&amp;F</th><th>50DMA</th><th>200DMA</th><th>Score</th></tr></thead>
              <tbody></tbody>
            </table>
          </div>
        </section>

        <section>
          <div class="section-head">
            <h2>Technical Laggards</h2>
            <div class="muted">Weak structure or weak relative strength that needs attention</div>
          </div>
          <div class="small-table">
            <table id="technicalLaggardsTable">
              <thead><tr><th>Stock</th><th>Status</th><th>RS Trend</th><th>RS vs 50D</th><th>RS 3M</th><th>RSI</th><th>P&amp;F</th><th>Score</th></tr></thead>
              <tbody></tbody>
            </table>
          </div>
        </section>

        <section>
          <div class="section-head">
            <h2>All Holdings</h2>
            <div class="controls">
              <input id="search" placeholder="Search stock" oninput="renderHoldings()">
              <select id="priority" onchange="renderHoldings()"><option value="">All priorities</option><option>High</option><option>Medium</option><option>Low</option></select>
              <select id="theme" onchange="renderHoldings()"><option value="">All themes</option></select>
              <select id="technicalStatus" onchange="renderHoldings()"><option value="">All technical statuses</option></select>
            </div>
          </div>
          <div class="small-table">
            <table id="holdingsTable">
              <thead>
                <tr>
                  <th>Stock</th>
                  <th>Priority</th>
                  <th>Bucket</th>
                  <th>Theme</th>
                  <th>Technical</th>
                  <th>Weight</th>
                  <th>Value</th>
                  <th>Day P&L</th>
                  <th>Day %</th>
                  <th>Qty</th>
                  <th>LTP</th>
                  <th>Ticker</th>
                </tr>
              </thead>
              <tbody></tbody>
            </table>
          </div>
        </section>
      </div>

      <section class="detail">
        <div class="section-head">
          <h2>Holding Detail</h2>
          <div class="muted">Click a holding</div>
        </div>
        <div id="detail" class="detail-body"></div>
      </section>
    </div>
  </main>

  <script>
    const DATA = {payload_json};
    const AUTO_REFRESH_MS = 30 * 60 * 1000;
    const autoRefreshStartedAt = Date.now();
    let selected = null;

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

    function sortByDay(direction) {{
      return [...DATA.holdings].sort((a, b) => direction * ((Number(b['Day Change %']) || 0) - (Number(a['Day Change %']) || 0))).slice(0, 6);
    }}

    function renderDayLeaders() {{
      const rows = [...sortByDay(1), ...sortByDay(-1)];
      const seen = new Set();
      document.querySelector('#leadersTable tbody').innerHTML = rows.filter(row => {{
        if (seen.has(row['Display Name'])) return false;
        seen.add(row['Display Name']);
        return true;
      }}).map(row => `
        <tr onclick="selectHolding('${{safe(row['Display Name']).replace(/'/g, "\\\\'")}}')">
          <td><strong>${{safe(row['Display Name'])}}</strong><div class="muted">${{safe(row['Yahoo Ticker'])}}</div></td>
          <td>${{safe(row.Theme)}}</td>
          <td class="num">${{pct(row['Weight %'])}}</td>
          <td class="num ${{tone(row['Day Change %'])}}">${{pct(row['Day Change %'])}}</td>
          <td class="num ${{tone(row['Day P&L'])}}">${{money(row['Day P&L'])}}</td>
          <td class="num">${{money(row['Current Value'])}}</td>
        </tr>
      `).join('');
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

    function renderTechnicalTables() {{
      const downloaded = DATA.holdings.filter(technicalDownloaded);
      const leaders = [...downloaded].sort((a, b) => (Number(b['Technical Score']) || 0) - (Number(a['Technical Score']) || 0)).slice(0, 10);
      const laggards = [...downloaded].sort((a, b) => (Number(a['Technical Score']) || 999) - (Number(b['Technical Score']) || 999)).slice(0, 10);
      const leaderRows = leaders.length ? leaders : DATA.holdings.slice(0, 8);
      const laggardRows = laggards.length ? laggards : DATA.holdings.slice(-8);

      document.querySelector('#technicalLeadersTable tbody').innerHTML = leaderRows.map(row => `
        <tr onclick="selectHolding('${{safe(row['Display Name']).replace(/'/g, "\\\\'")}}')">
          <td><strong>${{safe(row['Display Name'])}}</strong></td>
          <td>${{pill('technical-status', row['Technical Status'])}}</td>
          <td class="num ${{tone(row['RS vs 50D %'])}}">${{pct(row['RS vs 50D %'])}}</td>
          <td class="num">${{num(row['RSI 14'])}}</td>
          <td>${{safe(row['P&F Signal'])}}</td>
          <td class="num">${{boolText(row['Above 50DMA'])}}</td>
          <td class="num">${{boolText(row['Above 200DMA'])}}</td>
          <td class="num">${{safe(row['Technical Score'])}}</td>
        </tr>
      `).join('');

      document.querySelector('#technicalLaggardsTable tbody').innerHTML = laggardRows.map(row => `
        <tr onclick="selectHolding('${{safe(row['Display Name']).replace(/'/g, "\\\\'")}}')">
          <td><strong>${{safe(row['Display Name'])}}</strong></td>
          <td>${{pill('technical-status', row['Technical Status'])}}</td>
          <td class="spark-cell">${{sparkline(row['RS Trend'])}}</td>
          <td class="num ${{tone(row['RS vs 50D %'])}}">${{pct(row['RS vs 50D %'])}}</td>
          <td class="num ${{tone(row['RS 3M %'])}}">${{pct(row['RS 3M %'])}}</td>
          <td class="num">${{num(row['RSI 14'])}}</td>
          <td>${{safe(row['P&F Signal'])}}</td>
          <td class="num">${{safe(row['Technical Score'])}}</td>
        </tr>
      `).join('');
    }}

    function populateFilters() {{
      const themes = [...new Set(DATA.holdings.map(row => row.Theme).filter(Boolean))].sort();
      const statuses = [...new Set(DATA.holdings.map(row => row['Technical Status']).filter(Boolean))].sort();
      document.getElementById('theme').innerHTML += themes.map(value => `<option>${{safe(value)}}</option>`).join('');
      document.getElementById('technicalStatus').innerHTML += statuses.map(value => `<option>${{safe(value)}}</option>`).join('');
    }}

    function filteredHoldings() {{
      const search = document.getElementById('search').value.trim().toLowerCase();
      const priority = document.getElementById('priority').value;
      const theme = document.getElementById('theme').value;
      const status = document.getElementById('technicalStatus').value;
      return DATA.holdings.filter(row => {{
        const text = `${{row['Display Name']}} ${{row['Yahoo Ticker']}} ${{row.Theme}}`.toLowerCase();
        if (search && !text.includes(search)) return false;
        if (priority && row['Coordination Priority'] !== priority) return false;
        if (theme && row.Theme !== theme) return false;
        if (status && row['Technical Status'] !== status) return false;
        return true;
      }}).sort((a, b) => (Number(b['Weight %']) || 0) - (Number(a['Weight %']) || 0));
    }}

    function renderHoldings() {{
      const rows = filteredHoldings();
      document.querySelector('#holdingsTable tbody').innerHTML = rows.map(row => `
        <tr class="${{selected && selected['Display Name'] === row['Display Name'] ? 'selected' : ''}}" onclick="selectHolding('${{safe(row['Display Name']).replace(/'/g, "\\\\'")}}')">
          <td><strong>${{safe(row['Display Name'])}}</strong><div class="muted">${{safe(row['Source Name'])}}</div></td>
          <td>${{pill('priority', row['Coordination Priority'])}}</td>
          <td>${{pill('bucket', row['Portfolio Bucket'])}}</td>
          <td>${{safe(row.Theme)}}</td>
          <td>${{pill('technical-status', row['Technical Status'])}}</td>
          <td class="num">${{pct(row['Weight %'])}}</td>
          <td class="num">${{money(row['Current Value'])}}</td>
          <td class="num ${{tone(row['Day P&L'])}}">${{money(row['Day P&L'])}}</td>
          <td class="num ${{tone(row['Day Change %'])}}">${{pct(row['Day Change %'])}}</td>
          <td class="num">${{num(row.Quantity)}}</td>
          <td class="num">${{num(row.LTP)}}</td>
          <td>${{safe(row['Yahoo Ticker'])}}</td>
        </tr>
      `).join('');
    }}

    function selectHolding(name) {{
      selected = DATA.holdings.find(row => row['Display Name'] === name) || DATA.holdings[0];
      renderHoldings();
      renderDetail();
    }}

    function renderDetail() {{
      const row = selected || DATA.holdings[0];
      if (!row) return;
      document.getElementById('detail').innerHTML = `
        <div class="detail-title">
          <strong>${{safe(row['Display Name'])}}</strong>
          <div class="muted">${{safe(row.Theme)}} · ${{safe(row['Yahoo Ticker'])}}</div>
          <div style="margin-top:10px; display:flex; gap:8px; flex-wrap:wrap;">
            ${{pill('priority', row['Coordination Priority'])}}
            ${{pill('bucket', row['Portfolio Bucket'])}}
            ${{pill('technical-status', row['Technical Status'])}}
          </div>
        </div>
        <div class="kv">
          <div><span class="k">Current Value</span><span class="v">${{money(row['Current Value'])}}</span></div>
          <div><span class="k">Weight</span><span class="v">${{pct(row['Weight %'])}}</span></div>
          <div><span class="k">Day P&L</span><span class="v ${{tone(row['Day P&L'])}}">${{money(row['Day P&L'])}}</span></div>
          <div><span class="k">Day Move</span><span class="v ${{tone(row['Day Change %'])}}">${{pct(row['Day Change %'])}}</span></div>
          <div><span class="k">Quantity / Pledged</span><span class="v">${{num(row.Quantity)}} / ${{num(row['Pledged Qty'])}}</span></div>
          <div><span class="k">LTP</span><span class="v">${{num(row.LTP)}}</span></div>
          <div><span class="k">RS vs 50D</span><span class="v ${{tone(row['RS vs 50D %'])}}">${{pct(row['RS vs 50D %'])}}</span></div>
          <div><span class="k">RSI 14</span><span class="v">${{num(row['RSI 14'])}}</span></div>
          <div><span class="k">50DMA / 200DMA</span><span class="v">${{boolText(row['Above 50DMA'])}} / ${{boolText(row['Above 200DMA'])}}</span></div>
          <div><span class="k">P&F</span><span class="v">${{safe(row['P&F Signal'])}}</span></div>
        </div>
        <div class="note"><strong>Discussion:</strong> ${{safe(row['Suggested Discussion'])}}</div>
        <div class="note"><strong>Technical note:</strong> ${{safe(row['Technical Note'])}}</div>
        <div class="note"><strong>Data note:</strong> NSDL has provided current quantity, LTP, day movement, and market value. Buy price and total P&L are not present in this file.</div>
      `;
    }}

    function refreshPage() {{
      window.location.reload();
    }}

    function updateCountdown() {{
      const elapsed = Date.now() - autoRefreshStartedAt;
      const remaining = Math.max(0, AUTO_REFRESH_MS - elapsed);
      const minutes = String(Math.floor(remaining / 60000)).padStart(2, '0');
      const seconds = String(Math.floor((remaining % 60000) / 1000)).padStart(2, '0');
      document.getElementById('refreshStatus').textContent = `Auto refresh in ${{minutes}}:${{seconds}}`;
      if (remaining <= 0) refreshPage();
    }}

    renderMetrics();
    renderGroupRows('themes', DATA.themes);
    renderGroupRows('buckets', DATA.buckets);
    renderDayLeaders();
    populateFilters();
    renderTechnicalTables();
    selected = DATA.holdings[0];
    renderHoldings();
    renderDetail();
    updateCountdown();
    setInterval(updateCountdown, 1000);
  </script>
</body>
</html>
"""


def write_outputs(data: pd.DataFrame, meta: dict[str, str], output_dir: Path, source_name: str, write_csv: bool) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "data").mkdir(exist_ok=True)
    if write_csv:
        data.to_csv(output_dir / "data" / "holdings.csv", index=False)
    action_queue(data).to_csv(output_dir / "data" / "action_queue.csv", index=False)
    (output_dir / "index.html").write_text(dashboard_html(data, meta, source_name), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Build NSDL client holdings dashboard.")
    parser.add_argument("--input-xlsx", type=Path, default=DEFAULT_XLSX)
    parser.add_argument("--input-csv", type=Path)
    parser.add_argument("--output-dir", type=Path, default=ROOT)
    parser.add_argument("--no-write-csv", action="store_true")
    args = parser.parse_args()

    if args.input_csv:
        data, meta = read_holdings_csv(args.input_csv)
        source_name = args.input_csv.name
    else:
        data, meta = parse_nsdl_xlsx(args.input_xlsx)
        source_name = args.input_xlsx.name

    write_outputs(data, meta, args.output_dir, source_name, not args.no_write_csv)
    print(f"Built dashboard: {args.output_dir / 'index.html'}")
    print(f"Holdings: {len(data)}")
    print(f"Market value: {money(data['Current Value'].sum())}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
