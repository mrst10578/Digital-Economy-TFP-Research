# -*- coding: utf-8 -*-
"""
Phase 1 data collector for:
Digital Economy, Artificial Intelligence, and Business Cycle Dynamics

Revision 2 (2026-09-28)
- keeps Iran in cross-country auditing;
- removes the artificial World Bank 2000 start-year floor;
- paginates World Bank responses and labels countries vs aggregates;
- adds missing WDI digital indicators from the research document;
- uses an official FRED CSV fallback when no FRED API key is configured;
- chunks BLS history instead of discarding older observations;
- uses IMF AI Preparedness indicator AI_PI rather than dataset name AIPI;
- uses BEA's latest Digital Economy 2017-2022 workbook;
- writes coverage, Iran data, source dictionary, and collection status to Excel.

This script collects and audits data only. It does not run econometric models.
"""

from __future__ import annotations

import io
import os
import re
import time
from datetime import datetime

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

OUTPUT_FILE = "AI_Digital_Economy_BusinessCycle_Data.xlsx"
TEMP_OUTPUT_FILE = OUTPUT_FILE + ".tmp.xlsx"
END_YEAR = datetime.now().year
REQUEST_TIMEOUT = 60
SLEEP_BETWEEN_CALLS = 0.25
RETRIEVED_AT = datetime.utcnow().replace(microsecond=0).isoformat() + "Z"

FRED_API_KEY = os.environ.get("FRED_API_KEY", "")
BEA_API_KEY = os.environ.get("BEA_API_KEY", "")
BLS_API_KEY = os.environ.get("BLS_API_KEY", "")

_log_rows: list[dict] = []
_coverage_rows: list[dict] = []
_iran_rows: list[dict] = []
_dictionary_rows: list[dict] = []

retry = Retry(
    total=3,
    connect=3,
    read=3,
    backoff_factor=1.0,
    status_forcelist=(429, 500, 502, 503, 504),
    allowed_methods=frozenset({"GET", "POST"}),
)
SESSION = requests.Session()
SESSION.mount("https://", HTTPAdapter(max_retries=retry))


def log(name: str, source: str, status: str, detail: str = "", rows: int | None = None) -> None:
    _log_rows.append({
        "Variable": name,
        "Source": source,
        "Status": status,
        "Rows": rows,
        "Detail": clean_excel_text(str(detail)) if detail is not None else "",
        "Retrieved_At": RETRIEVED_AT,
    })
    suffix = f" ({detail})" if detail else ""
    print(f"[{status}] {source} :: {name}{suffix}")


def register(
    variable_name: str,
    source: str,
    series_id: str,
    geography: str,
    frequency: str = "",
    requirement_role: str = "",
    note: str = "",
) -> None:
    _dictionary_rows.append({
        "Variable": variable_name,
        "Source": source,
        "Series_ID": series_id,
        "Geography": geography,
        "Frequency_Hint": frequency,
        "Requirement_Role": requirement_role,
        "Note": note,
    })


def safe_sheet_name(name: str) -> str:
    for ch in '[]:*?/\\':
        name = name.replace(ch, "")
    return name[:31]


_ILLEGAL_EXCEL_CHARS = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")


def clean_excel_text(value):
    if isinstance(value, str):
        return _ILLEGAL_EXCEL_CHARS.sub("", value)[:32767]
    return value


def excel_safe_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in out.columns:
        if pd.api.types.is_object_dtype(out[col].dtype) or pd.api.types.is_string_dtype(out[col].dtype):
            out[col] = out[col].map(clean_excel_text)
    return out


def write_sheet(writer, df: pd.DataFrame | None, name: str, source: str) -> None:
    if df is None or df.empty:
        log(name, source, "EMPTY")
        return
    excel_safe_dataframe(df).to_excel(writer, sheet_name=safe_sheet_name(name), index=False)
    log(name, source, "OK", f"{len(df)} rows", len(df))


def numeric_valid(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").notna()


def add_iran_long_rows(
    df: pd.DataFrame,
    source: str,
    series_id: str,
    period_col: str,
    value_col: str,
    country_col: str,
    observation_status: str,
) -> None:
    if country_col not in df.columns:
        return
    subset = df[df[country_col].astype(str) == "IRN"].copy()
    for _, row in subset.iterrows():
        _iran_rows.append({
            "source": source,
            "series_id": series_id,
            "country_iso3": "IRN",
            "period": row.get(period_col),
            "value": row.get(value_col),
            "observation_status": observation_status,
            "retrieved_at": RETRIEVED_AT,
        })


# ---------------------------------------------------------------------------
# FRED
# ---------------------------------------------------------------------------

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
    "Info_Processing_Inv_Contribution": "A679RZ2Q224SBEA",
    "Computer_Systems_Design_Employment": "CES6054150001",
    "Information_Sector_Employment": "USINFO",
    "JOLTS_Job_Openings_Total": "JTSJOL",
    "JOLTS_Prof_Business_Services": "JTS540099JOL",
    "JOLTS_Hires_Total": "JTSHIL",
    "JOLTS_Quits_Total": "JTSQUL",
    "JOLTS_Total_Separations": "JTSTSL",
    "Labor_Productivity_NonfarmBiz": "OPHNFB",
    "Unit_Labor_Cost_NonfarmBiz": "ULCNFB",
    "Total_Factor_Productivity": "MFPPBS",
    "CPI_Headline": "CPIAUCSL",
    "CPI_Core": "CPILFESL",
    "PCE_Price_Index": "PCEPI",
    "PCE_Core": "PCEPILFE",
    "Economic_Policy_Uncertainty_Monthly": "USEPUINDXM",
    "Economic_Policy_Uncertainty_Daily": "USEPUINDXD",
    "National_Financial_Conditions_Index": "NFCI",
    "Effective_Federal_Funds_Rate_Candidate": "FEDFUNDS",
}

for _name, _sid in FRED_SERIES.items():
    _note = ""
    if _sid == "A679RZ2Q224SBEA":
        _note = "Official definition: contribution to percent change in real private fixed investment, not contribution to total GDP growth."
    if _sid == "FEDFUNDS":
        _note = "Candidate policy-rate proxy for U.S.; exact policy-rate specification remains a research-design choice."
    register(_name, "FRED", _sid, "US", requirement_role="research-document or supporting U.S. series", note=_note)


def fetch_fred_series(series_id: str) -> pd.DataFrame:
    if FRED_API_KEY:
        r = SESSION.get(
            "https://api.stlouisfed.org/fred/series/observations",
            params={
                "series_id": series_id,
                "api_key": FRED_API_KEY,
                "file_type": "json",
            },
            timeout=REQUEST_TIMEOUT,
        )
        r.raise_for_status()
        obs = r.json().get("observations", [])
        df = pd.DataFrame(obs)
        if df.empty:
            return df
        df = df[["date", "value"]].rename(columns={"date": "Date", "value": series_id})
    else:
        # Official FRED graph CSV endpoint does not require an API key.
        r = SESSION.get(
            "https://fred.stlouisfed.org/graph/fredgraph.csv",
            params={"id": series_id},
            timeout=REQUEST_TIMEOUT,
        )
        r.raise_for_status()
        df = pd.read_csv(io.StringIO(r.text))
        if "DATE" in df.columns:
            df = df.rename(columns={"DATE": "Date"})
        elif "observation_date" in df.columns:
            df = df.rename(columns={"observation_date": "Date"})
        elif "date" in df.columns:
            df = df.rename(columns={"date": "Date"})
        else:
            raise ValueError(f"FRED CSV did not contain a recognized date column: {list(df.columns)}")
        if series_id not in df.columns:
            raise ValueError(f"FRED CSV did not contain expected column {series_id}")
        df = df[["Date", series_id]]
    df[series_id] = pd.to_numeric(df[series_id], errors="coerce")
    return df


def collect_fred(writer) -> None:
    print("\n--- FRED ---")
    mode = "API" if FRED_API_KEY else "official keyless CSV fallback"
    for name, series_id in FRED_SERIES.items():
        try:
            df = fetch_fred_series(series_id)
            write_sheet(writer, df, name, f"FRED ({mode})")
        except Exception as exc:
            log(name, "FRED", "ERROR", str(exc))
        time.sleep(SLEEP_BETWEEN_CALLS)


# ---------------------------------------------------------------------------
# World Bank WDI
# ---------------------------------------------------------------------------

WORLDBANK_INDICATORS = {
    "WB_Internet_Users_Percent": "IT.NET.USER.ZS",
    "WB_Fixed_Broadband_Per_100": "IT.NET.BBND.P2",
    "WB_R_and_D_Percent_GDP": "GB.XPD.RSDV.GD.ZS",
    "WB_HighTech_Manufactured_Exports": "TX.VAL.TECH.MF.ZS",
    "WB_GDP_Growth_Annual_Percent": "NY.GDP.MKTP.KD.ZG",
    "WB_Unemployment_Percent": "SL.UEM.TOTL.ZS",
    "WB_Inflation_CPI_Percent": "FP.CPI.TOTL.ZG",
    "WB_ICT_Service_Exports_Percent": "BX.GSR.CCIS.ZS",
}

for _name, _sid in WORLDBANK_INDICATORS.items():
    _note = ""
    if _sid == "BX.GSR.CCIS.ZS":
        _note = "Optional separate metric. It is not a substitute for high-technology manufactured export share."
    register(_name, "World Bank WDI", _sid, "Cross-country", "Annual", "research-document / supporting cross-country series", _note)


def fetch_worldbank_entity_metadata() -> tuple[pd.DataFrame, dict[str, dict]]:
    r = SESSION.get(
        "https://api.worldbank.org/v2/country",
        params={"format": "json", "per_page": 400},
        timeout=REQUEST_TIMEOUT,
    )
    r.raise_for_status()
    payload = r.json()
    records = payload[1] if isinstance(payload, list) and len(payload) >= 2 and payload[1] else []
    rows = []
    mapping: dict[str, dict] = {}
    for item in records:
        iso3 = item.get("id", "")
        region = (item.get("region") or {})
        income = (item.get("incomeLevel") or {})
        lending = (item.get("lendingType") or {})
        entity_type = "aggregate" if region.get("id") == "NA" else "country"
        row = {
            "ISO3": iso3,
            "ISO2": item.get("iso2Code", ""),
            "Entity_Name": item.get("name", ""),
            "Entity_Type": entity_type,
            "Region_ID": region.get("id", ""),
            "Region": region.get("value", ""),
            "Income_Level": income.get("value", ""),
            "Lending_Type": lending.get("value", ""),
        }
        rows.append(row)
        mapping[iso3] = row
    return pd.DataFrame(rows), mapping


def fetch_worldbank_indicator(code: str, entity_meta: dict[str, dict]) -> pd.DataFrame:
    page = 1
    rows = []
    while True:
        r = SESSION.get(
            f"https://api.worldbank.org/v2/country/all/indicator/{code}",
            params={"format": "json", "per_page": 20000, "page": page},
            timeout=REQUEST_TIMEOUT,
        )
        r.raise_for_status()
        payload = r.json()
        if not isinstance(payload, list) or len(payload) < 2:
            raise ValueError("Unexpected World Bank response structure")
        meta = payload[0] or {}
        data = payload[1] or []
        rows.extend(data)
        pages = int(meta.get("pages", 1) or 1)
        if page >= pages:
            break
        page += 1
        time.sleep(SLEEP_BETWEEN_CALLS)

    df = pd.json_normalize(rows)
    if df.empty:
        return df
    keep = [
        c for c in ["country.value", "countryiso3code", "date", "value"]
        if c in df.columns
    ]
    df = df[keep].rename(columns={
        "country.value": "Country",
        "countryiso3code": "ISO3",
        "date": "Year",
        "value": code,
    })
    df["Year"] = pd.to_numeric(df["Year"], errors="coerce").astype("Int64")
    df[code] = pd.to_numeric(df[code], errors="coerce")
    df["Entity_Type"] = df["ISO3"].map(lambda x: entity_meta.get(str(x), {}).get("Entity_Type", "unknown"))
    df["Region"] = df["ISO3"].map(lambda x: entity_meta.get(str(x), {}).get("Region", ""))
    df["Income_Level"] = df["ISO3"].map(lambda x: entity_meta.get(str(x), {}).get("Income_Level", ""))
    return df


def add_worldbank_coverage(df: pd.DataFrame, name: str, code: str) -> None:
    if df.empty:
        return
    for iso3, group in df.groupby("ISO3", dropna=False):
        valid = group[numeric_valid(group[code])].copy()
        entity_type = str(group["Entity_Type"].iloc[0]) if "Entity_Type" in group.columns else "unknown"
        if valid.empty:
            _coverage_rows.append({
                "variable_id": code,
                "variable_name": name,
                "country_iso3": iso3,
                "entity_type": entity_type,
                "source": "World Bank WDI",
                "series_id": code,
                "frequency": "annual",
                "first_valid_period": None,
                "last_valid_period": None,
                "valid_observations": 0,
                "expected_periods": None,
                "missing_periods": None,
                "coverage_percent": None,
                "observation_status": "no numeric value returned",
                "availability_status": "missing in returned source data",
                "notes": "",
                "retrieved_at": RETRIEVED_AT,
            })
            continue
        years = pd.to_numeric(valid["Year"], errors="coerce").dropna().astype(int)
        first = int(years.min())
        last = int(years.max())
        expected = last - first + 1
        valid_count = int(years.nunique())
        missing = max(expected - valid_count, 0)
        _coverage_rows.append({
            "variable_id": code,
            "variable_name": name,
            "country_iso3": iso3,
            "entity_type": entity_type,
            "source": "World Bank WDI",
            "series_id": code,
            "frequency": "annual",
            "first_valid_period": first,
            "last_valid_period": last,
            "valid_observations": valid_count,
            "expected_periods": expected,
            "missing_periods": missing,
            "coverage_percent": valid_count / expected if expected else None,
            "observation_status": "observed source values",
            "availability_status": "available",
            "notes": "Coverage is based on numeric values, not merely first/last rows.",
            "retrieved_at": RETRIEVED_AT,
        })


def collect_worldbank(writer) -> None:
    print("\n--- World Bank ---")
    entity_df, entity_meta = fetch_worldbank_entity_metadata()
    write_sheet(writer, entity_df, "WorldBank_Entity_Metadata", "World Bank")
    for name, code in WORLDBANK_INDICATORS.items():
        try:
            df = fetch_worldbank_indicator(code, entity_meta)
            write_sheet(writer, df, name, "World Bank")
            add_worldbank_coverage(df, name, code)
            add_iran_long_rows(df, "World Bank WDI", code, "Year", code, "ISO3", "observed source value")
        except Exception as exc:
            log(name, "World Bank", "ERROR", str(exc))
        time.sleep(SLEEP_BETWEEN_CALLS)


# ---------------------------------------------------------------------------
# BEA NIPA (optional API key)
# ---------------------------------------------------------------------------

BEA_TABLES = {
    "BEA_RealGDP_PercentChange_T10101": "T10101",
}
register("BEA_RealGDP_PercentChange_T10101", "BEA NIPA", "T10101", "US", "Quarterly", "supporting U.S. GDP data")


def collect_bea(writer) -> None:
    print("\n--- BEA NIPA ---")
    if not BEA_API_KEY:
        log("All BEA NIPA tables", "BEA", "SKIPPED", "BEA_API_KEY not set; FRED/other official sources remain available")
        return
    for name, table in BEA_TABLES.items():
        try:
            r = SESSION.get(
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
            result = r.json().get("BEAAPI", {}).get("Results", {})
            if "Error" in result:
                raise ValueError(str(result["Error"]))
            data = result.get("Data", [])
            write_sheet(writer, pd.DataFrame(data), name, "BEA")
        except Exception as exc:
            log(name, "BEA", "ERROR", str(exc))
        time.sleep(SLEEP_BETWEEN_CALLS)


# ---------------------------------------------------------------------------
# BLS - chunked history
# ---------------------------------------------------------------------------

BLS_SERIES = {
    "BLS_Unemployment_Rate": "LNS14000000",
    "BLS_Total_Nonfarm_Payrolls": "CES0000000001",
    "BLS_CPI_All_Items": "CUUR0000SA0",
    "BLS_Computer_Systems_Design_Emp": "CES6054150001",
}
for _name, _sid in BLS_SERIES.items():
    register(_name, "BLS", _sid, "US", "Monthly", "research-document / supporting labor-price series")

# 1913 is the earliest start among the currently configured BLS series set (CPI).
# It is a source-specific floor, not a project-wide research start year.
BLS_HISTORY_START = 1913


def collect_bls(writer) -> None:
    print("\n--- BLS ---")
    span = 20 if BLS_API_KEY else 10
    by_series: dict[str, list[dict]] = {sid: [] for sid in BLS_SERIES.values()}

    chunk_start = BLS_HISTORY_START
    while chunk_start <= END_YEAR:
        chunk_end = min(chunk_start + span - 1, END_YEAR)
        body = {
            "seriesid": list(BLS_SERIES.values()),
            "startyear": str(chunk_start),
            "endyear": str(chunk_end),
        }
        if BLS_API_KEY:
            body["registrationKey"] = BLS_API_KEY

        try:
            r = SESSION.post(
                "https://api.bls.gov/publicAPI/v2/timeseries/data/",
                json=body,
                headers={"Content-type": "application/json"},
                timeout=REQUEST_TIMEOUT,
            )
            r.raise_for_status()
            payload = r.json()
            if payload.get("status") != "REQUEST_SUCCEEDED":
                raise ValueError(str(payload.get("message", "")))
            for series in payload.get("Results", {}).get("series", []):
                sid = series.get("seriesID", "")
                if sid in by_series:
                    by_series[sid].extend(series.get("data", []))
        except Exception as exc:
            log(f"BLS chunk {chunk_start}-{chunk_end}", "BLS", "ERROR", str(exc))
        chunk_start = chunk_end + 1
        time.sleep(SLEEP_BETWEEN_CALLS)

    id_to_name = {v: k for k, v in BLS_SERIES.items()}
    for sid, rows in by_series.items():
        name = id_to_name[sid]
        df = pd.DataFrame(rows)
        if not df.empty:
            if "year" in df.columns and "period" in df.columns:
                df = df.drop_duplicates(subset=["year", "period"], keep="first")
                df = df.sort_values(["year", "period"], ascending=[False, False])
            if "value" in df.columns:
                df["value_numeric"] = pd.to_numeric(df["value"], errors="coerce")
        write_sheet(writer, df, name, "BLS")


# ---------------------------------------------------------------------------
# OECD supplementary CLI
# ---------------------------------------------------------------------------

register(
    "OECD_Composite_Leading_Indicator",
    "OECD",
    "OECD.SDD.STES,DSD_STES@DF_CLI",
    "OECD/selected areas",
    "Monthly",
    "supplementary business-cycle indicator; NOT a substitute for Going Digital/OECD.AI metrics",
)


def collect_oecd(writer) -> None:
    print("\n--- OECD ---")
    try:
        url = (
            "https://sdmx.oecd.org/public/rest/data/"
            "OECD.SDD.STES,DSD_STES@DF_CLI/.M.LI...AA...H"
        )
        r = SESSION.get(
            url,
            params={
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


# ---------------------------------------------------------------------------
# IMF DataMapper
# ---------------------------------------------------------------------------

IMF_INDICATORS = {
    "IMF_AI_Preparedness_Index": "AI_PI",
    "IMF_Real_GDP_Growth": "NGDP_RPCH",
    "IMF_Inflation_Percent": "PCPIPCH",
}
register("IMF_AI_Preparedness_Index", "IMF DataMapper / AIPI", "AI_PI", "Cross-country", "2023 snapshot", "AI preparedness; not AI investment/adoption")
register("IMF_Real_GDP_Growth", "IMF DataMapper / WEO", "NGDP_RPCH", "Cross-country", "Annual", "macro control/outcome")
register("IMF_Inflation_Percent", "IMF DataMapper / WEO", "PCPIPCH", "Cross-country", "Annual", "macro price control")


def add_imf_coverage(df: pd.DataFrame, name: str, code: str) -> None:
    if df.empty:
        return
    for country, group in df.groupby("Country_Code", dropna=False):
        valid = group[numeric_valid(group[code])]
        if valid.empty:
            continue
        years = pd.to_numeric(valid["Year"], errors="coerce").dropna().astype(int)
        returned_first = int(years.min())
        returned_last = int(years.max())
        # DataMapper output here does not carry a reliable actual/estimate/forecast flag.
        # To avoid calling future values observed history, leave last_valid_period blank.
        _coverage_rows.append({
            "variable_id": code,
            "variable_name": name,
            "country_iso3": country,
            "entity_type": "country_or_imf_entity",
            "source": "IMF DataMapper",
            "series_id": code,
            "frequency": "annual_or_snapshot",
            "first_valid_period": returned_first,
            "last_valid_period": None,
            "valid_observations": int(len(valid)),
            "expected_periods": None,
            "missing_periods": None,
            "coverage_percent": None,
            "observation_status": "returned periods; actual/estimate/forecast status not retained",
            "availability_status": "available but observed endpoint unresolved",
            "notes": f"Returned through {returned_last}; do not treat returned endpoint as last observed year.",
            "retrieved_at": RETRIEVED_AT,
        })


def collect_imf(writer) -> None:
    print("\n--- IMF ---")
    for name, code in IMF_INDICATORS.items():
        try:
            r = SESSION.get(
                f"https://www.imf.org/external/datamapper/api/v1/{code}",
                timeout=REQUEST_TIMEOUT,
            )
            r.raise_for_status()
            values = r.json().get("values", {}).get(code, {})
            records = []
            for country, year_map in values.items():
                for year, value in year_map.items():
                    try:
                        y = int(year)
                    except Exception:
                        y = year
                    period_flag = (
                        "future_year_returned"
                        if isinstance(y, int) and y > END_YEAR
                        else "current_or_historical_year_returned"
                    )
                    records.append({
                        "Country_Code": country,
                        "Year": y,
                        code: value,
                        "Period_Flag": period_flag,
                    })
            df = pd.DataFrame(records)
            if not df.empty:
                df[code] = pd.to_numeric(df[code], errors="coerce")
            write_sheet(writer, df, name, "IMF")
            add_imf_coverage(df, name, code)
            add_iran_long_rows(
                df,
                "IMF DataMapper",
                code,
                "Year",
                code,
                "Country_Code",
                "returned; actual/estimate/forecast classification not retained",
            )
        except Exception as exc:
            log(name, "IMF", "ERROR", str(exc))
        time.sleep(SLEEP_BETWEEN_CALLS)


# ---------------------------------------------------------------------------
# BEA Digital Economy latest official workbook
# ---------------------------------------------------------------------------

BEA_DIGITAL_ECONOMY_URL = (
    "https://bea.gov/sites/default/files/2023-12/"
    "DigitalEconomy_2017-2022.xlsx"
)
register(
    "BEA_Digital_Economy_Satellite_Account",
    "BEA Digital Economy",
    "2017-2022 official workbook",
    "US",
    "Annual",
    "digital economy share/output/value added; account discontinued after 2023 release",
)


def collect_bea_digital_economy(writer) -> None:
    print("\n--- BEA Digital Economy ---")
    try:
        r = SESSION.get(BEA_DIGITAL_ECONOMY_URL, timeout=REQUEST_TIMEOUT)
        r.raise_for_status()
        sheets = pd.read_excel(io.BytesIO(r.content), sheet_name=None)
        wrote = 0
        for idx, (sheet_name, df) in enumerate(sheets.items(), start=1):
            if df is None or df.empty:
                continue
            out_name = f"BEA_DE_{idx:02d}_{sheet_name}"
            write_sheet(writer, df, out_name, "BEA Digital Economy 2017-2022")
            wrote += 1
        if wrote == 0:
            log("BEA_Digital_Economy", "BEA Digital Economy", "EMPTY", "Workbook contained no non-empty sheets")
    except Exception as exc:
        log("BEA_Digital_Economy", "BEA Digital Economy", "ERROR", str(exc))


# ---------------------------------------------------------------------------
# Epoch AI model-compute data
# ---------------------------------------------------------------------------

EPOCH_DATASETS = {
    "Epoch_Frontier_AI_Models": "https://epoch.ai/data/frontier_ai_models.csv",
    "Epoch_Notable_AI_Models": "https://epoch.ai/data/notable_ai_models.csv",
}
register(
    "Epoch_Frontier_AI_Models",
    "Epoch AI",
    "frontier_ai_models.csv",
    "Model-level / organization / country fields as published",
    "Irregular event-level",
    "frontier-model training compute and related model characteristics",
    "Epoch defines frontier models as models in the top five of training compute at release.",
)
register(
    "Epoch_Notable_AI_Models",
    "Epoch AI",
    "notable_ai_models.csv",
    "Model-level / organization / country fields as published",
    "Irregular event-level",
    "supporting AI model history",
)


def collect_epoch_ai(writer) -> None:
    print("\n--- Epoch AI ---")
    for name, url in EPOCH_DATASETS.items():
        try:
            r = SESSION.get(url, timeout=REQUEST_TIMEOUT)
            r.raise_for_status()
            df = pd.read_csv(io.StringIO(r.text))
            write_sheet(writer, df, name, "Epoch AI")
        except Exception as exc:
            log(name, "Epoch AI", "ERROR", str(exc))
        time.sleep(SLEEP_BETWEEN_CALLS)


# ---------------------------------------------------------------------------
# AIOE / AIIE occupational and industry exposure data
# ---------------------------------------------------------------------------

AIOE_FILES = {
    "AIOE_Base": "https://raw.githubusercontent.com/AIOE-Data/AIOE/main/AIOE_DataAppendix.xlsx",
    "AIOE_GenAI_Language": "https://raw.githubusercontent.com/AIOE-Data/AIOE/main/Language%20Modeling%20AIOE%20and%20AIIE.xlsx",
    "AIOE_GenAI_Image": "https://raw.githubusercontent.com/AIOE-Data/AIOE/main/Image%20Generation%20AIOE%20and%20AIIE.xlsx",
}
register(
    "AIOE_AIIE_AIGE",
    "Felten-Raj-Seamans / AIOE-Data",
    "AIOE_DataAppendix.xlsx",
    "US occupation / industry / county",
    "Cross-sectional exposure index",
    "H2 occupational, industry, and geographic AI exposure",
    "Base workbook contains AIOE by SOC, AIIE by NAICS, and AIGE by FIPS.",
)
register(
    "AIOE_Generative_AI_Extensions",
    "AIOE-Data",
    "Language Modeling / Image Generation AIOE and AIIE",
    "US occupation / industry",
    "Cross-sectional exposure index",
    "H2 generative-AI exposure extensions",
)


def collect_aioe(writer) -> None:
    print("\n--- AIOE / AIIE ---")
    for file_label, url in AIOE_FILES.items():
        try:
            r = SESSION.get(url, timeout=REQUEST_TIMEOUT)
            r.raise_for_status()
            book = pd.read_excel(io.BytesIO(r.content), sheet_name=None)
            wrote = 0
            for idx, (sheet_name, df) in enumerate(book.items(), start=1):
                if df is None or df.empty:
                    continue
                write_sheet(
                    writer,
                    df,
                    f"{file_label}_{idx:02d}_{sheet_name}",
                    "AIOE-Data",
                )
                wrote += 1
            if wrote == 0:
                log(file_label, "AIOE-Data", "EMPTY", "No non-empty sheets found")
        except Exception as exc:
            log(file_label, "AIOE-Data", "ERROR", str(exc))
        time.sleep(SLEEP_BETWEEN_CALLS)


# ---------------------------------------------------------------------------
# CSET / ETO Country AI Activity Metrics
# ---------------------------------------------------------------------------

CSET_ETO_FILES = {
    "CSET_AI_Patent_Applications": "patents_yearly_applications.csv",
    "CSET_AI_Investment_Disclosed": "companies_yearly_disclosed.csv",
    "CSET_AI_Investment_Estimated": "companies_yearly_estimated.csv",
}
CSET_ETO_RECORD = "https://zenodo.org/records/22772306/files/{filename}?download=1"
register(
    "CSET_AI_Patent_Applications",
    "CSET / ETO Country AI Activity Metrics",
    "patents_yearly_applications.csv",
    "Cross-country",
    "Annual",
    "AI patent filings",
    "Recent patent years can be materially incomplete; retain the source complete flag.",
)
register(
    "CSET_AI_Investment_Disclosed",
    "CSET / ETO Country AI Activity Metrics",
    "companies_yearly_disclosed.csv",
    "Cross-country",
    "Annual",
    "private-market AI investment",
    "Millions USD; equity investment into privately held AI-related companies.",
)
register(
    "CSET_AI_Investment_Estimated",
    "CSET / ETO Country AI Activity Metrics",
    "companies_yearly_estimated.csv",
    "Cross-country",
    "Annual",
    "estimated private-market AI investment",
    "Millions USD; preserve the source complete flag and distinguish estimated from disclosed values.",
)


def collect_cset_eto(writer) -> None:
    print("\n--- CSET / ETO Country AI Activity ---")
    for name, filename in CSET_ETO_FILES.items():
        try:
            url = CSET_ETO_RECORD.format(filename=filename)
            r = SESSION.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True)
            r.raise_for_status()
            df = pd.read_csv(io.BytesIO(r.content))
            write_sheet(writer, df, name, "CSET / ETO")
        except Exception as exc:
            log(name, "CSET / ETO", "ERROR", str(exc))
        time.sleep(SLEEP_BETWEEN_CALLS)


# ---------------------------------------------------------------------------
# Manual / not-yet-automated sources
# ---------------------------------------------------------------------------

MANUAL_SOURCES = [
    {
        "Variable": "Census BTOS – AI use by businesses",
        "URL": "https://www.census.gov/programs-surveys/btos.html",
        "Status": "data acquisition still required",
        "Note": "Biweekly AI-use supplement; preserve question/version, sector, geography and sample metadata.",
    },
    {
        "Variable": "World Bank – Digital Adoption Index",
        "URL": "https://www.worldbank.org/en/publication/wdr2016/Digital-Adoption-Index",
        "Status": "data acquisition still required",
        "Note": "Sparse cross-section (not an annual panel); do not interpolate into fake yearly observations.",
    },
    {
        "Variable": "Stanford HAI – AI Index economy data",
        "URL": "https://hai.stanford.edu/ai-index/2026-ai-index-report",
        "Status": "data acquisition still required",
        "Note": "Select exact public tables and preserve the underlying provider/methodology.",
    },
    {
        "Variable": "Epoch AI – model aggregation into country-time panels",
        "URL": "https://epoch.ai/data/ai-models",
        "Status": "raw Frontier/Notable model files automated; aggregation methodology still required",
        "Note": "Keep model-level observations distinct from any later country/year aggregation.",
    },
    {
        "Variable": "CSET / ETO Country Activity Tracker",
        "URL": "https://cat.eto.tech/",
        "Status": "data acquisition still required",
        "Note": "Candidate source for AI patents/research/investment; verify exact table and lag.",
    },
    {
        "Variable": "AIOE / AIIE / AIGE employment matching",
        "URL": "https://github.com/AIOE-Data/AIOE",
        "Status": "raw exposure workbooks automated; compatible employment histories/crosswalk still required",
        "Note": "Do not treat exposure scores alone as H2-ready without matched SOC/NAICS employment/output data.",
    },
    {
        "Variable": "OECD Going Digital / OECD.AI",
        "URL": "https://goingdigital.oecd.org/indicators",
        "Status": "data acquisition still required",
        "Note": "Current OECD CLI collection is supplementary and does not satisfy this requirement.",
    },
]


RESEARCH_REQUIREMENTS = [
    ["BC01","Business cycle","Real GDP level","US high-frequency + country equivalents","FRED GDPC1; national accounts","US collected; Iran equivalent not yet collected","partial"],
    ["BC02","Business cycle","Real GDP growth","Core outcome/control","WDI NY.GDP.MKTP.KD.ZG; IMF NGDP_RPCH; FRED A191RL1Q225SBEA","Iran + cross-country + US collected","available"],
    ["BC03","Business cycle","Industrial production index","Monthly activity control","FRED INDPRO; national equivalents","US collected; Iran equivalent not yet collected","partial"],
    ["BC04","Business cycle","Unemployment rate","Labor control","WDI SL.UEM.TOTL.ZS; BLS/FRED","Iran annual + cross-country + US monthly collected","available"],
    ["BC05","Business cycle","Nonfarm payroll employment","US labor control","FRED PAYEMS / BLS CES0000000001","US collected; not an Iran series","available US-only"],
    ["BC06","Business cycle","Recession indicator/dating","Regime variable","FRED USREC / NBER","US collected; not an Iran series","available US-only"],
    ["BC07","Business cycle","Policy interest rate","Monetary-policy interaction","FRED FEDFUNDS candidate; national central banks","US candidate collected; exact country-equivalent definition pending","partial / definition pending"],
    ["DE01","Digital economy","Internet penetration","Digital intensity","WDI IT.NET.USER.ZS","Iran + cross-country collected","available"],
    ["DE02","Digital economy","Fixed broadband subscriptions per 100","Digital intensity","WDI IT.NET.BBND.P2","Iran + cross-country collected","available"],
    ["DE03","Digital economy","High-technology exports (% manufactured exports)","Digital trade/intensity","WDI TX.VAL.TECH.MF.ZS","Iran + cross-country collected","available"],
    ["DE04","Digital economy","R&D expenditure (% GDP)","Innovation intensity","WDI GB.XPD.RSDV.GD.ZS","Iran + cross-country collected","available"],
    ["DE05","Digital economy","Digital economy share of GDP","National-account digital share","BEA Digital Economy + national equivalents","US BEA 2017-2022 collected; Iran equivalent not yet found","partial"],
    ["DE06","Digital economy","Composite digital-transformation index","Cross-sectional/panel digital intensity","World Bank DAI; OECD Going Digital; DESI","sources verified; exact raw tables not yet automated","missing raw data"],
    ["AI01","AI-specific","Private investment in AI","AI capital intensity","CSET/ETO; Stanford HAI; OECD.AI","CSET/ETO collection attempted in current run; see Collection Status","source-dependent"],
    ["AI02","AI-specific","AI patent filings","Innovation proxy","CSET/ETO; OECD.AI","CSET/ETO collection attempted in current run; preserve completeness flags","source-dependent"],
    ["AI03","AI-specific","Frontier-model training compute","Physical AI-capital proxy","Epoch AI","Frontier + notable model raw data collected","available model-level"],
    ["AI04","AI-specific","AI venture capital/private-market investment","AI investment proxy","CSET/ETO; Stanford HAI; OECD.AI","CSET/ETO collection attempted in current run","source-dependent"],
    ["AI05","AI-specific","Information-processing equipment & software investment","Long-run digital-capex proxy","FRED A679RC1Q027SBEA","US collected","available US-only"],
    ["AI06","AI-specific","Business AI use/adoption","High-frequency adoption","US Census BTOS","official download verified; exact raw AI supplement file not yet automated","missing raw data"],
    ["LB01","Labor/structure","Information/high-tech industry employment","Structural labor measure","FRED USINFO; BLS CES6054150001","US collected","available US-only"],
    ["LB02","Labor/structure","Occupational/industry/geographic AI exposure","H2 exposure variable","AIOE/AIIE/AIGE","base + generative-language + image exposure workbooks collected","available exposure data"],
    ["LB03","Labor/structure","Labor productivity","H1 outcome","FRED OPHNFB","US collected","available US-only"],
    ["LB04","Labor/structure","Total factor productivity","H1 outcome","FRED MFPPBS","US collected","available US-only"],
    ["LB05","Labor/structure","Unit labor costs","Cost/price mechanism","FRED ULCNFB","US collected","available US-only"],
    ["LB06","Labor/structure","Job openings","Matching efficiency","FRED JTSJOL / JTS540099JOL","US collected","available US-only"],
    ["LB07","Labor/structure","Labor turnover","Matching efficiency","FRED JTSHIL / JTSQUL / JTSTSL","US collected","available US-only"],
    ["PR01","Prices","Headline CPI/inflation","Price control/outcome","WDI FP.CPI.TOTL.ZG; FRED CPIAUCSL; BLS CUUR0000SA0","Iran + cross-country + US collected","available"],
    ["PR02","Prices","Core CPI / core PCE","Underlying inflation","FRED CPILFESL / PCEPILFE","US collected","available US-only"],
    ["PR03","Prices","Sectoral CPI/PPI by industry","H3 sectoral persistence","BLS sector price series","sector universe/mapping not approved yet","definition pending / missing"],
    ["PR04","Prices","Inflation volatility","Derived rolling standard deviation","Derived from price series","raw inputs partly collected; window/frequency not approved","derived methodology pending"],
    ["FN01","Finance/uncertainty","Economic Policy Uncertainty","H6 uncertainty measure","FRED USEPUINDXM / USEPUINDXD","US monthly + daily collected","available US-only"],
    ["FN02","Finance/uncertainty","VIX","Financial interaction","FRED VIXCLS","US market collected","available US-only"],
    ["FN03","Finance/uncertainty","Financial conditions","Robustness control","FRED NFCI","US collected","available US-only"],
    ["FN04","Finance/uncertainty","Multiple asset returns","H4 cross-asset correlation","financial-market source TBD","asset universe/frequency not specified in research document","definition pending / missing"],
    ["INV01","Investment","Information-processing investment contribution","H5","FRED A679RZ2Q224SBEA","US collected; official definition is contribution to real private fixed-investment growth, NOT total GDP growth","available with definition correction"],
    ["H6A","Events","Major AI release calendar","H6 event alignment","verified event sources TBD","not collected; event definition/calendar must be fixed first","definition pending / missing"],
    ["H6B","Investment","Non-AI investment","H6 comparison outcome","BEA/FRED component series TBD","not collected; exact exclusion rule for AI/digital investment is undefined","definition pending / missing"],
]

HYPOTHESIS_DATA_MATRIX = [
    ["H1","Digital/AI-capex growth vs productivity with lags","Digital capex; labor productivity; TFP; long history","US raw capex/productivity/TFP collected","partially data-ready","Cross-country capex/productivity equivalents still needed; no estimation authorized"],
    ["H2","Employment volatility/output volatility by AI exposure","AIOE/AIIE; occupation/industry employment; output; crosswalk","AIOE/AIIE raw exposure collected","partially data-ready","Compatible historical employment/output crosswalk still needed"],
    ["H3","Sectoral price persistence vs digital intensity","Sectoral CPI/PPI; sector digital intensity; persistence definition","Aggregate prices collected","not data-ready","Sector universe, sectoral price data and digital-intensity mapping still missing"],
    ["H4","Cross-asset correlation/volatility during AI-capex growth","Multiple asset returns; AI/digital capex; window definition","VIX and capex collected","not data-ready","VIX alone is insufficient; asset universe/frequency not defined"],
    ["H5","Post-2023 volatility of information-processing investment contribution","Contribution series; comparator components; long history","A679RZ2Q224SBEA collected","partially data-ready","Research document mislabels this as total-GDP contribution; comparator components must be selected"],
    ["H6","EPU around major AI releases followed by weaker non-AI investment","EPU; release calendar; non-AI investment","Daily/monthly EPU collected","not data-ready","Verified event calendar and non-AI investment definition/data missing"],
]


def write_requirement_matrices(writer) -> None:
    req_cols = ["ID","Family","Variable","Research_Role","Candidate_Source","Current_Evidence","Availability_Status"]
    hyp_cols = ["Hypothesis","Question","Required_Data","Current_Evidence","Data_Readiness","Remaining_Gap"]
    excel_safe_dataframe(pd.DataFrame(RESEARCH_REQUIREMENTS, columns=req_cols)).to_excel(
        writer, sheet_name="Requirements Matrix", index=False
    )
    excel_safe_dataframe(pd.DataFrame(HYPOTHESIS_DATA_MATRIX, columns=hyp_cols)).to_excel(
        writer, sheet_name="Hypothesis Matrix", index=False
    )


def main() -> None:
    print("=" * 78)
    print("Starting revised data collection and coverage audit")
    print("=" * 78)

    try:
        with pd.ExcelWriter(TEMP_OUTPUT_FILE, engine="openpyxl") as writer:
            collect_fred(writer)
            collect_worldbank(writer)
            collect_bea(writer)
            collect_bls(writer)
            collect_oecd(writer)
            collect_imf(writer)
            collect_bea_digital_economy(writer)
            collect_epoch_ai(writer)
            collect_aioe(writer)
            collect_cset_eto(writer)
            write_requirement_matrices(writer)

            excel_safe_dataframe(pd.DataFrame(MANUAL_SOURCES)).to_excel(
                writer, sheet_name="Manual Sources", index=False
            )

            coverage_df = pd.DataFrame(_coverage_rows)
            if not coverage_df.empty:
                excel_safe_dataframe(coverage_df).to_excel(writer, sheet_name="Coverage", index=False)
                iran_cov = coverage_df[
                    coverage_df["country_iso3"].astype(str) == "IRN"
                ].copy()
                excel_safe_dataframe(iran_cov).to_excel(writer, sheet_name="Iran Coverage", index=False)

            iran_df = pd.DataFrame(_iran_rows)
            if not iran_df.empty:
                excel_safe_dataframe(iran_df).to_excel(writer, sheet_name="Iran Data", index=False)

            excel_safe_dataframe(pd.DataFrame(_dictionary_rows).drop_duplicates()).to_excel(
                writer, sheet_name="Variable Dictionary", index=False
            )
            excel_safe_dataframe(pd.DataFrame(_log_rows)).to_excel(
                writer, sheet_name="Collection Status", index=False
            )

            audit_notes = pd.DataFrame([
                {
                    "Rule": "Iran mandatory",
                    "Implementation": "IRN is audited explicitly in Iran Data / Iran Coverage; U.S.-only series are never relabeled as Iran.",
                },
                {
                    "Rule": "No fixed project-wide start year",
                    "Implementation": "World Bank and FRED request full returned history; BLS uses a source-specific 1913 floor covering the currently configured series set.",
                },
                {
                    "Rule": "Forecasts do not extend observed history",
                    "Implementation": "IMF returned future periods are retained with flags, but last_valid_period is intentionally not certified without source status metadata.",
                },
                {
                    "Rule": "No synthetic completeness",
                    "Implementation": "Missing values remain missing; no interpolation or zero filling is performed.",
                },
                {
                    "Rule": "Econometrics deferred",
                    "Implementation": "No hypothesis test or model estimation is executed by this collector.",
                },
            ])
            excel_safe_dataframe(audit_notes).to_excel(writer, sheet_name="Audit Notes", index=False)

        os.replace(TEMP_OUTPUT_FILE, OUTPUT_FILE)
    finally:
        if os.path.exists(TEMP_OUTPUT_FILE):
            try:
                os.remove(TEMP_OUTPUT_FILE)
            except OSError:
                pass

    print("=" * 78)
    print(f"Finished: {OUTPUT_FILE}")
    print("=" * 78)


if __name__ == "__main__":
    main()
