"""Shared logic for talking to the Hutch RAN KPI portal API.

Used by both kpi_auto_app.py (desktop Tkinter app) and streamlit_app.py
(web app) so the fetch/parsing logic only lives in one place.
"""
import os
from datetime import datetime, timezone

import pandas as pd

BASE_URL = "https://randorsai.hutch.lk:8002"

# name -> (districtId, provinceId), from the KPI Analyzer Geo tree
DISTRICTS = {
    "Central": {
        "Badulla": (3, 8),
        "Kandy": (11, 1),
        "Matale": (16, 1),
        "Nuwara Eliya": (20, 1),
    },
    "South": {
        "Galle": (6, 7),
        "Matara": (17, 7),
        "Hambantota": (8, 7),
        "Monaragala": (18, 8),
        "Ratnapura": (23, 6),
    },
    "North East": {
        "Jaffna": (9, 5),
        "Kilinochchi": (13, 5),
        "Mannar": (15, 5),
        "Mullaitivu": (19, 5),
        "Vavuniya": (25, 5),
        "Ampara": (1, 2),
        "Batticaloa": (4, 2),
        "Trincomalee": (24, 2),
    },
    "Gampaha": {
        "Gampaha": (7, 9),
    },
    "Kegalle-Puttalama": {
        "Kegalle": (12, 6),
        "Puttalam": (22, 4),
    },
    "Anuradhapura-Kurunegala": {
        "Anuradhapura": (2, 3),
        "Kurunegala": (14, 4),
    },
    "Colombo-Kaluthara": {
        "Colombo": (5, 9),
        "Kalutara": (10, 9),
    },
}

DISTRICT_TEMPLATE_ID = 40  # "D. District - Default" (Geo tab)

# BSS tab: network-wide totals are fetched per-technology, not in one call.
# (technologyTypeId -> RANKPITemplateId) for each "NE (xG) - Default" template.
NETWORK_TECH_TEMPLATES = {
    1: 29,  # 2G
    2: 30,  # 3G
    3: 31,  # 4G
}

SHEET_SERIES = {
    ("2G", "Voice"): ["TCH HR (Erl)", "TCH FR (Erl)"],
    ("3G", "Voice"): ["CS AMR (Erl)"],
    ("3G", "Data"): ["PS Traffic Volume (GByte)"],
    ("4G", "Voice"): ["VoLTE Traffic (Erl)"],
    ("4G", "Data"): ["Traffic Volume (GByte)"],
}


def login(session, email, password):
    resp = session.post(f"{BASE_URL}/api/auth/signin", json={"email": email, "password": password})
    resp.raise_for_status()
    data = resp.json()
    if data.get("Status") != "Success":
        raise RuntimeError(data.get("ErrorMessage") or "Login failed")
    session.headers["Authorization"] = f"Bearer {data['User']['Token']}"
    return data["User"]


def fetch_district_data(session, district_id, province_id, start, end):
    payload = {
        "RANKPIAggregationLevel": "GeographicalSitesCluster",
        "SelectedNodes": {
            "CurrentNode": {"NodeId": district_id, "RANKPIHierarchyLevel": "District", "TechnologyTypeId": None},
            "ParentNodes": [
                {"NodeId": province_id, "RANKPIHierarchyLevel": "Province", "TechnologyTypeId": None},
                {"NodeId": district_id, "RANKPIHierarchyLevel": "District", "TechnologyTypeId": None},
            ],
        },
        "RANKPITemplateId": DISTRICT_TEMPLATE_ID,
        "RANOSSKPIInputTimeGranularityId": 1,
        "StartDateTime": start,
        "EndDateTime": end,
    }
    resp = session.post(f"{BASE_URL}/api/RANKPIAnalyzerChart/GetChartData", json=payload)
    resp.raise_for_status()
    return resp.json()


def fetch_network_tech_data(session, technology_type_id, template_id, start, end):
    payload = {
        "RANKPIAggregationLevel": "BSSSubSystem",
        "SelectedNodes": {
            "CurrentNode": {"NodeId": technology_type_id, "RANKPIHierarchyLevel": "Technology", "TechnologyTypeId": technology_type_id},
            "ParentNodes": [
                {"NodeId": -1, "RANKPIHierarchyLevel": "Network", "TechnologyTypeId": None},
                {"NodeId": technology_type_id, "RANKPIHierarchyLevel": "Technology", "TechnologyTypeId": technology_type_id},
            ],
        },
        "RANKPITemplateId": template_id,
        "RANOSSKPIInputTimeGranularityId": 1,
        "StartDateTime": start,
        "EndDateTime": end,
    }
    resp = session.post(f"{BASE_URL}/api/RANKPIAnalyzerChart/GetChartData", json=payload)
    resp.raise_for_status()
    return resp.json()


def fetch_network_data(session, start, end):
    """Network-wide totals live on the BSS tab, fetched per-technology (2G/3G/4G
    each have their own template), then merged into one combined chart_data dict
    so it can be handled by write_location_csvs() like a district response."""
    combined_sections = []
    for tech_id, template_id in NETWORK_TECH_TEMPLATES.items():
        data = fetch_network_tech_data(session, tech_id, template_id, start, end)
        combined_sections.extend(data.get("TemplateSections", []))
    return {"TemplateSections": combined_sections}


def extract_series(chart_data, series_name):
    for section in chart_data.get("TemplateSections", []):
        for chart in section.get("Charts", []):
            for series in chart.get("SeriesList", []):
                if series.get("SeriesName") == series_name:
                    return series.get("SeriesData", [])
    return None


def series_to_df(series_data, column_name):
    if not series_data:
        return None
    dates = [datetime.fromtimestamp(pt["X"] / 1000, tz=timezone.utc).strftime("%Y-%m-%d") for pt in series_data]
    values = [pt["Y"] if pt["Y"] is not None else 0 for pt in series_data]
    return pd.DataFrame({"DateTime": dates, column_name: values})


def build_location_df(chart_data, series_names, log):
    df = None
    for name in series_names:
        s = series_to_df(extract_series(chart_data, name), name)
        if s is None:
            log(f"    series '{name}' not found")
            continue
        df = s if df is None else pd.merge(df, s, on="DateTime", how="outer")
    return df


def write_location_csvs(chart_data, base_folder, location_name, log):
    for (tech, traffic_type), series_names in SHEET_SERIES.items():
        df = build_location_df(chart_data, series_names, log)
        if df is None or df.empty:
            continue
        folder = os.path.join(base_folder, f"{tech} {traffic_type}")
        os.makedirs(folder, exist_ok=True)
        out_path = os.path.join(folder, f"{tech}_{traffic_type}_{location_name}.csv")
        df.to_csv(out_path, index=False)
        log(f"    wrote {tech} {traffic_type} / {location_name}.csv ({len(df)} rows)")


def fetch_region_report(session, region, start, end, target_dir, include_network, log):
    """Fetches every district (and optionally Network) for a region and writes
    the CSVs into target_dir/Sales Report/<region>/... . Caller then calls
    report_full_1.build_excel_report(target_dir, region) to produce the xlsx."""
    base_folder = os.path.join(target_dir, "Sales Report", region)
    districts = DISTRICTS[region]

    for name, (district_id, province_id) in districts.items():
        log(f"Fetching {name}...")
        chart_data = fetch_district_data(session, district_id, province_id, start, end)
        if not chart_data.get("TemplateSections"):
            log(f"  WARNING: no data returned for {name} ({chart_data.get('InfoBanner')})")
            continue
        write_location_csvs(chart_data, base_folder, name, log)

    if include_network:
        log("Fetching Network (BSS tab)...")
        try:
            network_data = fetch_network_data(session, start, end)
            if network_data.get("TemplateSections"):
                write_location_csvs(network_data, base_folder, "Network", log)
            else:
                log("  WARNING: Network fetch returned no data — skipping Network column.")
        except Exception as e:
            log(f"  WARNING: Network fetch failed ({e}) — skipping Network column.")
