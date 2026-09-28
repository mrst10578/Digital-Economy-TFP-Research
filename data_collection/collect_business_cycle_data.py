# -*- coding: utf-8 -*-
"""
Phase 1 data collector for:
Digital Economy, Artificial Intelligence, and Business Cycle Dynamics

This collector is based on the project files supplied by the client.
It gathers public/official series and writes them into one Excel workbook.
Sources that require API keys are skipped cleanly when the key is absent.
"""

from __future__ import annotations

import io
import os
import time
from datetime import datetime

import pandas as pd
import requests

OUTPUT_FILE = "AI_Digital_Economy_BusinessCycle_Data.xlsx"
START_YEAR = 2000
END_YEAR = datetime.now().year
REQUEST_TIMEOUT = 45
SLEEP_BETWEEN_CALLS = 0.5

FRED_API_KEY = os.environ.get("FRED_API_KEY", "")
BEA_API_KEY = os.environ.get("BEA_API_KEY", "")
BLS_API_KEY = os.environ.get("BLS_API_KEY", "")

_log_rows: list[dict] = []


def log(name: str, source: str, status: str, detail: str = "") -> None:
    _log_rows.append({
        "Variable": name,
        "Source": source,
        "Status": status,
        "Detail": detail,
    })
    suffix = f" ({detail})" if detail else ""
    print(f"[{status}] {source} :: {name}{suffix}")


def safe_sheet_name(name: str) -> str:
    for ch in "[]:*?/\":
        name = name.replace(ch, "")
    return name[:31]


def write_sheet(writer, df: pd.DataFrame | None, name: str, source: str) -> None:
    if df is None or df.empty:
        log(name, source, "EMPTY")
        return
    df.to_excel(writer, sheet_name=safe_sheet_name(name), index=False)
    log(name, source, "OK", f"{len(df)} rows")


FRED_SERIES = {
    "Real_GDP_Level": "GDPC1",
    "Real_GDP_Growth_Rate": "A191RL1Q225SBEA",
    "Industrial_Production": "INDPRO",
    "NBER_Recession_Indicator": "USREC",
    "Sahm_Rule_Indicator": "SAHMREALTIME",
    "Yield_Curve_10Y_minus_2Y": "T10Y2Y",
    "Nonfarm_Payrolls": "PAYEMS",
    "Unemployment_Rate": "UNRATE",
    "VIX_Volatility_Index": "VIXCLS",
    "OECD_CLI_United_States": "USALOLITONOSTSAM",
    "Info_Processing_Equip_Software_Inv": "A679RC1Q027SBEA",
    "Computer_Systems_Design_Employment": "CES6054150001",
    "Information_Sector_Employment": "USINFO",
    "JOLTS_Job_Openings_Total": "JTSJOL",
    "JOLTS_Prof_Business_Services": "JTS540099JOL",
    "Labor_Productivity_NonfarmBiz": "OPHNFB",
    "Total_Factor_Productivity": "MFPPBS",
    "CPI_Headline": "CPIAUCSL",
    "CPI_Core": "CPILFESL",
    "PCE_Price_Index": "PCEPI",
}


def collect_fred(writer) -> None:
    print("\n--- FRED ---")
    if not FRED_API_KEY:
        log("All FRED series", "FRED", "SKIPPED", "FRED_API_KEY not set")
        return

    url = "https://api.stlouisfed.org/fred/series/observations"
    for name, series_id in FRED_SERIES.items():
        try:
            r = requests.get(
                url,
                params={
                    "series_id": series_id,
                    "api_key": FRED_API_KEY,
                    "file_type": "json",
                    "observation_start": f"{START_YEAR}-01-01",
                },
                timeout=REQUEST_TIMEOUT,
            )
            r.raise_for_status()
            obs = r.json().get("observations", [])
            df = pd.DataFrame(obs)
            if not df.empty:
                df = df[["date", "value"]].rename(
                    columns={"date": "Date", "value": series_id}
                )
                df[series_id] = pd.to_numeric(df[series_id], errors="coerce")
            write_sheet(writer, df, name, "FRED")
        except Exception as exc:
            log(name, "FRED", "ERROR", str(exc))
        time.sleep(SLEEP_BETWEEN_CALLS)


WORLDBANK_INDICATORS = {
    "WB_Internet_Users_Percent": "IT.NET.USER.ZS",
    "WB_GDP_Growth_Annual_Percent": "NY.GDP.MKTP.KD.ZG",
    "WB_Unemployment_Percent": "SL.UEM.TOTL.ZS",
    "WB_Inflation_CPI_Percent": "FP.CPI.TOTL.ZG",
    "WB_ICT_Service_Exports_Percent": "BX.GSR.CCIS.ZS",
}


def collect_worldbank(writer) -> None:
    print("\n--- World Bank ---")
    for name, code in WORLDBANK_INDICATORS.items():
        try:
            r = requests.get(
                f"https://api.worldbank.org/v2/country/all/indicator/{code}",
                params={
                    "format": "json",
                    "per_page": 20000,
                    "date": f"{START_YEAR}:{END_YEAR}",
                },
                timeout=REQUEST_TIMEOUT,
            )
            r.raise_for_status()
            payload = r.json()
            df = pd.DataFrame()
            if isinstance(payload, list) and len(payload) >= 2 and payload[1]:
                df = pd.json_normalize(payload[1])
                keep = [
                    c for c in
                    ["country.value", "countryiso3code", "date", "value"]
                    if c in df.columns
                ]
                df = df[keep].rename(columns={
                    "country.value": "Country",
                    "countryiso3code": "ISO3",
                    "date": "Year",
                    "value": code,
                })
            write_sheet(writer, df, name, "World Bank")
        except Exception as exc:
            log(name, "World Bank", "ERROR", str(exc))
        time.sleep(SLEEP_BETWEEN_CALLS)


BEA_TABLES = {
    "BEA_RealGDP_PercentChange_T10101": "T10101",
}


def collect_bea(writer) -> None:
    print("\n--- BEA ---")
    if not BEA_API_KEY:
        log("All BEA tables", "BEA", "SKIPPED", "BEA_API_KEY not set")
        return

    for name, table in BEA_TABLES.items():
        try:
            r = requests.get(
                "https://apps.bea.gov/api/data",
                params={
                    "UserID": BEA_API_KEY,
                    "method": "GetData",
                    "DatasetName": "NIPA",
                    "TableName": table,
                    "Frequency": "Q",
                    "Year": "ALL",
                    "ResultFormat": "JSON",
                },
                timeout=REQUEST_TIMEOUT,
            )
            r.raise_for_status()
            data = r.json().get("BEAAPI", {}).get("Results", {}).get("Data", [])
            write_sheet(writer, pd.DataFrame(data), name, "BEA")
        except Exception as exc:
            log(name, "BEA", "ERROR", str(exc))
        time.sleep(SLEEP_BETWEEN_CALLS)


BLS_SERIES = {
    "BLS_Unemployment_Rate": "LNS14000000",
    "BLS_Total_Nonfarm_Payrolls": "CES0000000001",
    "BLS_CPI_All_Items": "CUUR0000SA0",
    "BLS_Computer_Systems_Design_Emp": "CES6054150001",
}


def collect_bls(writer) -> None:
    print("\n--- BLS ---")
    span = 20 if BLS_API_KEY else 10
    start_year = max(START_YEAR, END_YEAR - span + 1)

    body = {
        "seriesid": list(BLS_SERIES.values()),
        "startyear": str(start_year),
        "endyear": str(END_YEAR),
    }
    if BLS_API_KEY:
        body["registrationKey"] = BLS_API_KEY

    try:
        r = requests.post(
            "https://api.bls.gov/publicAPI/v2/timeseries/data/",
            json=body,
            headers={"Content-type": "application/json"},
            timeout=REQUEST_TIMEOUT,
        )
        r.raise_for_status()
        payload = r.json()
        if payload.get("status") != "REQUEST_SUCCEEDED":
            log("All BLS series", "BLS", "ERROR", str(payload.get("message", "")))
            return

        id_to_name = {v: k for k, v in BLS_SERIES.items()}
        for series in payload.get("Results", {}).get("series", []):
            sid = series.get("seriesID", "")
            name = id_to_name.get(sid, sid)
            write_sheet(
                writer,
                pd.DataFrame(series.get("data", [])),
                name,
                "BLS",
            )
    except Exception as exc:
        log("All BLS series", "BLS", "ERROR", str(exc))


def collect_oecd(writer) -> None:
    print("\n--- OECD ---")
    try:
        url = (
            "https://sdmx.oecd.org/public/rest/data/"
            "OECD.SDD.STES,DSD_STES@DF_CLI/.M.LI...AA...H"
        )
        r = requests.get(
            url,
            params={
                "startPeriod": "2000-01",
                "dimensionAtObservation": "AllDimensions",
                "format": "csvfilewithlabels",
            },
            timeout=REQUEST_TIMEOUT,
        )
        r.raise_for_status()
        df = pd.read_csv(io.StringIO(r.text))
        write_sheet(writer, df, "OECD_Composite_Leading_Indicator", "OECD")
    except Exception as exc:
        log("OECD_Composite_Leading_Indicator", "OECD", "ERROR", str(exc))


IMF_INDICATORS = {
    "IMF_AI_Preparedness_Index": "AIPI",
    "IMF_Real_GDP_Growth": "NGDP_RPCH",
    "IMF_Inflation_Percent": "PCPIPCH",
}


def collect_imf(writer) -> None:
    print("\n--- IMF ---")
    for name, code in IMF_INDICATORS.items():
        try:
            r = requests.get(
                f"https://www.imf.org/external/datamapper/api/v1/{code}",
                timeout=REQUEST_TIMEOUT,
            )
            r.raise_for_status()
            values = r.json().get("values", {}).get(code, {})
            records = [
                {"Country_Code": country, "Year": year, code: value}
                for country, year_map in values.items()
                for year, value in year_map.items()
            ]
            write_sheet(writer, pd.DataFrame(records), name, "IMF")
        except Exception as exc:
            log(name, "IMF", "ERROR", str(exc))
        time.sleep(SLEEP_BETWEEN_CALLS)


DIRECT_DOWNLOAD_FILES = {
    "BEA_Digital_Economy_Satellite_Acct": (
        "https://www.bea.gov/system/files/2019-04/"
        "digital-economy-tables-april-2019.xlsx"
    )
}


def collect_direct_downloads(writer) -> None:
    print("\n--- Direct downloads ---")
    for name, url in DIRECT_DOWNLOAD_FILES.items():
        try:
            r = requests.get(url, timeout=REQUEST_TIMEOUT)
            r.raise_for_status()
            sheets = pd.read_excel(io.BytesIO(r.content), sheet_name=None)
            df = next(iter(sheets.values()))
            write_sheet(writer, df, name, "Direct download")
        except Exception as exc:
            log(name, "Direct download", "ERROR", str(exc))


MANUAL_SOURCES = [
    {
        "Variable": "Census BTOS – AI use by businesses",
        "URL": "https://www.census.gov/programs-surveys/btos.html",
        "Note": "Biweekly CSV files; direct file path changes by release.",
    },
    {
        "Variable": "World Bank – Digital Adoption Index",
        "URL": "https://www.worldbank.org/en/publication/wdr2016/Digital-Adoption-Index",
        "Note": "Download XLSX manually; only 2014 and 2016.",
    },
    {
        "Variable": "Stanford HAI – AI Index",
        "URL": "https://hai.stanford.edu/ai-index/2026-ai-index-report/economy",
        "Note": "Report/public data; not a stable API.",
    },
]


def main() -> None:
    print("=" * 72)
    print("Starting data collection")
    print("=" * 72)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        collect_fred(writer)
        collect_worldbank(writer)
        collect_bea(writer)
        collect_bls(writer)
        collect_oecd(writer)
        collect_imf(writer)
        collect_direct_downloads(writer)

        pd.DataFrame(MANUAL_SOURCES).to_excel(
            writer,
            sheet_name="Manual Sources",
            index=False,
        )
        pd.DataFrame(_log_rows).to_excel(
            writer,
            sheet_name="Collection Status",
            index=False,
        )

    print("=" * 72)
    print(f"Finished: {OUTPUT_FILE}")
    print("=" * 72)


if __name__ == "__main__":
    main()
