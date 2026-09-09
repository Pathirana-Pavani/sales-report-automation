"""Regenerates the monthly review presentation from its existing template,
swapping only the chart data (categories + series values) for the current
month's numbers. Layout, colors, titles, and every non-chart slide are left
untouched — those slides (capacity upgrades, relocations, new sites, Q&A)
are edited by hand each month, same as before.

Data source is the already-generated Excel report(s) (build_excel_report's
output), not the raw fetch, so the presentation always matches exactly what
is in the accompanying spreadsheet(s).

Two chart "dialects" are supported:

1. Full-name dialect (South, North East): each chart's series are named
   exactly after a sheet's columns (e.g. "Galle", "Network") — matched
   directly against one sheet.
2. Code-name / combined-metric dialect (Kegalle-Puttalama+Gampaha combo
   deck, Anuradhapura-Kurunegala deck): series use short district codes
   ("GAM", "KEG", "ANU"...) and some charts show a metric combined across
   technologies (e.g. "2G+3G Voice" as one line) that isn't a single sheet
   on its own — resolved via _build_combo_df() instead.

A chart is tried against dialect 1 first; if its series don't all resolve
there, dialect 2 is tried; otherwise it's left untouched (matches the
existing "unmatched charts are skipped" behavior for slides like the
site-count chart or per-site case-study dashboards).
"""
import re
from datetime import date

import pandas as pd
from pptx import Presentation
from pptx.oxml.ns import qn

EXCEL_EPOCH = date(1899, 12, 30)


def _to_excel_serial(d):
    return (d - EXCEL_EPOCH).days

# (substring checked in this order; first match wins) -> Excel sheet name
TITLE_TO_SHEET = [
    ("total voice", "Total Voice"),
    ("total data", "Total Data"),
    ("2g voice", "2G Voice"),
    ("3g voice", "3G Voice"),
    ("4g data", "4G Data"),
    ("3g data", "3G Data"),
    ("volte", None),  # resolved below: "network" in title -> Network sheet, else Cluster sheet
]

# Short district codes used in the Kegalle-Puttalama/Gampaha and
# Anuradhapura-Kurunegala combo decks, to their full district names
# (matching the column names in the Excel report).
DISTRICT_CODE_TO_NAME = {
    "GAM": "Gampaha",
    "KEG": "Kegalle",
    "PUT": "Puttalam",
    "ANU": "Anuradhapura",
    "POL": "Polonnaruwa",
    "KUR": "Kurunegala",
}


def _resolve_sheet_name(title, available_sheets):
    title_lower = title.lower()
    for keyword, sheet_name in TITLE_TO_SHEET:
        if keyword not in title_lower:
            continue
        if keyword == "volte":
            sheet_name = "4G Voice - Network" if "network" in title_lower else "4G Voice - Cluster"
        if sheet_name in available_sheets:
            return sheet_name
        return None
    return None


def _build_combo_df(title, sheets):
    """Builds a DataFrame for the code-name/combined-metric dialect. The
    reference/total series (whatever it's called in the deck - "Total
    Network", "Total(TB)", etc.) is always exposed as column '__TOTAL__'."""
    t = title.lower()

    if "2g" in t and "3g" in t and "voice" in t:
        s2, s3 = sheets.get("2G Voice"), sheets.get("3G Voice")
        if s2 is None or s3 is None:
            return None
        merged = pd.merge(s2, s3, on="Date", how="outer", suffixes=("_2g", "_3g")).fillna(0)
        out = pd.DataFrame({"Date": merged["Date"]})
        for full_name in DISTRICT_CODE_TO_NAME.values():
            c2, c3 = f"{full_name}_2g", f"{full_name}_3g"
            if c2 in merged.columns and c3 in merged.columns:
                out[full_name] = merged[c2] + merged[c3]
        if "Network_2g" in merged.columns and "Network_3g" in merged.columns:
            out["__TOTAL__"] = merged["Network_2g"] + merged["Network_3g"]
        return out

    if "3g" in t and "4g" in t and "data" in t:
        src = sheets.get("Total Data")
        if src is None:
            return None
        df = src.copy()
        total_col = next((c for c in df.columns if "total network" in c.lower()), None)
        if total_col:
            df = df.rename(columns={total_col: "__TOTAL__"})
        return df

    if "volte" in t:
        src = sheets.get("4G Voice - Cluster")
        if src is None:
            return None
        df = src.copy()
        if "Network" in df.columns:
            df = df.rename(columns={"Network": "__TOTAL__"})
        return df

    if "3g voice" in t:
        src = sheets.get("3G Voice")
        if src is None:
            return None
        df = src.copy()
        if "Network" in df.columns:
            df = df.rename(columns={"Network": "__TOTAL__"})
        return df

    return None


def _combo_column_resolver(series_name, df):
    name = series_name.strip() if series_name else series_name
    if name in df.columns:
        return name
    full = DISTRICT_CODE_TO_NAME.get(name)
    if full and full in df.columns:
        return full
    if name and "total" in name.lower() and "__TOTAL__" in df.columns:
        return "__TOTAL__"
    return None


def _exact_column_resolver(series_name, df):
    name = series_name.strip() if series_name else series_name
    return name if name in df.columns else None


def _load_sheets(excel_path):
    xl = pd.ExcelFile(excel_path)
    sheets = {}
    for name in xl.sheet_names:
        df = xl.parse(name)
        df["Date"] = pd.to_datetime(df["Date"]).dt.date
        df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
        sheets[name] = df
    return sheets


def _load_and_merge_sheets(excel_paths):
    """Loads one or more region Excel reports and, for sheets that appear
    in more than one (e.g. two regions both have a "2G Voice" sheet),
    merges their district columns together on Date. Network/Total-Network
    columns are assumed identical across regions (same portal-wide value)
    and are only kept once."""
    combined = {}
    for path in excel_paths:
        for name, df in _load_sheets(path).items():
            if name not in combined:
                combined[name] = df
                continue
            existing = combined[name]
            new_cols = [c for c in df.columns if c not in existing.columns]
            if new_cols:
                combined[name] = pd.merge(existing, df[["Date"] + new_cols], on="Date", how="outer")
    return combined


def _chart_title(chart):
    try:
        return chart.chart_title.text_frame.text if chart.has_title else ""
    except Exception:
        return ""


def _find(el, *tags):
    return el.find("/".join(qn(t) for t in tags))


def _rewrite_num_cache(num_cache_el, values, is_date=False):
    for child in num_cache_el.findall(qn("c:pt")):
        num_cache_el.remove(child)
    existing_count = num_cache_el.find(qn("c:ptCount"))
    if existing_count is not None:
        num_cache_el.remove(existing_count)

    ptCount_el = num_cache_el.makeelement(qn("c:ptCount"), {"val": str(len(values))})
    num_cache_el.append(ptCount_el)
    for idx, v in enumerate(values):
        v = _to_excel_serial(v) if is_date else v
        pt_el = num_cache_el.makeelement(qn("c:pt"), {"idx": str(idx)})
        v_el = pt_el.makeelement(qn("c:v"), {})
        v_el.text = str(v)
        pt_el.append(v_el)
        num_cache_el.append(pt_el)


def _bump_formula_range(ref_el, new_count):
    f_el = ref_el.find(qn("c:f"))
    if f_el is None or not f_el.text:
        return
    m = re.match(r"^(.*\$[A-Z]+\$)\d+(\$?)$", f_el.text)
    if m:
        f_el.text = f"{m.group(1)}{1 + new_count}{m.group(2)}"


def _all_series_resolve(chart, df, resolver):
    return all(resolver(s.name, df) is not None for s in chart.series)


def update_chart_data(chart, df, log, resolver=_exact_column_resolver):
    dates = list(df["Date"])

    series_cols = {}
    missing = []
    for series in chart.series:
        col = resolver(series.name, df)
        if col is None:
            missing.append(series.name)
        else:
            series_cols[series] = col

    if missing:
        raise ValueError(f"series {missing} not found in sheet columns {list(df.columns)}")

    for series, col in series_cols.items():
        ser_el = series._element
        cat_numRef = _find(ser_el, "c:cat", "c:numRef")
        val_numRef = _find(ser_el, "c:val", "c:numRef")
        cat_cache = _find(ser_el, "c:cat", "c:numRef", "c:numCache")
        val_cache = _find(ser_el, "c:val", "c:numRef", "c:numCache")

        _rewrite_num_cache(cat_cache, dates, is_date=True)
        _rewrite_num_cache(val_cache, df[col].fillna(0).tolist())
        _bump_formula_range(cat_numRef, len(dates))
        _bump_formula_range(val_numRef, len(dates))

    log(f"    updated ({len(chart.series)} series x {len(dates)} dates)")


def build_pptx_report(template_path, excel_paths, output_path, log=print):
    if isinstance(excel_paths, str):
        excel_paths = [excel_paths]
    sheets = _load_and_merge_sheets(excel_paths)
    prs = Presentation(template_path)

    updated, skipped = 0, 0
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if not getattr(shape, "has_chart", False):
                continue
            chart = shape.chart
            title = _chart_title(chart)

            sheet_name = _resolve_sheet_name(title, sheets)
            if sheet_name is not None and _all_series_resolve(chart, sheets[sheet_name], _exact_column_resolver):
                log(f"slide {i+1}: chart {title!r} -> sheet {sheet_name!r}")
                update_chart_data(chart, sheets[sheet_name], log, _exact_column_resolver)
                updated += 1
                continue

            combo_df = _build_combo_df(title, sheets)
            if combo_df is not None and _all_series_resolve(chart, combo_df, _combo_column_resolver):
                log(f"slide {i+1}: chart {title!r} -> combo data")
                update_chart_data(chart, combo_df, log, _combo_column_resolver)
                updated += 1
                continue

            log(f"slide {i+1}: chart {title!r} -> no matching data, left as-is")
            skipped += 1

    prs.save(output_path)
    log(f"Done. {updated} charts updated, {skipped} left untouched. Saved to {output_path}")
    return output_path
