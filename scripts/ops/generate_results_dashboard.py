"""
Generate a single standalone HTML dashboard from scenario result workbooks.

The script scans scenario folders for `Master_Plan.xlsx`, extracts key datasets,
and writes one responsive HTML file with embedded JSON + JavaScript logic.
"""

import argparse
import json
import math
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple

import pandas as pd


DEFAULT_DISCOUNT_RATE = 0.08
MAX_STACK_SERIES = 8
MAX_LINE_SERIES = 8


def get_repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def normalize_key(value: Any) -> str:
    return "".join(ch.lower() if ch.isalnum() else "_" for ch in str(value)).strip("_")


def pretty_label(name: str) -> str:
    return name.replace("_", " ").replace("##", " ").strip()


def sanitize_filename(value: str) -> str:
    out = []
    for ch in value:
        if ch.isalnum() or ch in ("-", "_"):
            out.append(ch)
        elif ch.isspace():
            out.append("_")
    safe = "".join(out).strip("_")
    return safe or "chart"


def find_column(df: pd.DataFrame, candidate_tokens: Iterable[str]) -> str:
    normalized_map = {col: normalize_key(col) for col in df.columns}
    for token in candidate_tokens:
        target = normalize_key(token)
        for col, norm in normalized_map.items():
            if target and target in norm:
                return col
    return ""


def clean_numeric(value: Any) -> float:
    if pd.isna(value):
        return 0.0
    try:
        fval = float(value)
    except (TypeError, ValueError):
        return 0.0
    return fval if math.isfinite(fval) else 0.0


def year_axis(series: pd.Series) -> List[Any]:
    years = []
    for value in series:
        if pd.isna(value):
            years.append(None)
            continue
        try:
            years.append(int(float(value)))
        except (TypeError, ValueError):
            years.append(str(value))
    return years


def to_float_list(series: pd.Series, scale: float = 1.0) -> List[float]:
    values = []
    for value in series:
        fv = clean_numeric(value) / scale
        values.append(round(fv, 6))
    return values


def non_zero_columns(df: pd.DataFrame, columns: Iterable[str], threshold: float = 1e-9) -> List[str]:
    active = []
    for col in columns:
        if col in df.columns:
            active.append(col)
    return active


def split_major_minor(
    df: pd.DataFrame,
    columns: List[str],
    top_n: int,
    other_label: str,
) -> Tuple[pd.DataFrame, List[str]]:
    if len(columns) <= top_n:
        return df, columns

    totals = {
        col: pd.to_numeric(df[col], errors="coerce").fillna(0.0).abs().sum()
        for col in columns
    }
    ordered = sorted(columns, key=lambda col: totals[col], reverse=True)
    major = ordered[:top_n]
    minor = ordered[top_n:]

    df_out = df.copy()
    label = other_label
    while label in df_out.columns:
        label += "_"
    df_out[label] = pd.DataFrame({c: pd.to_numeric(df_out[c], errors="coerce").fillna(0.0) for c in minor}).sum(axis=1)
    major.append(label)
    return df_out, major


def placeholder_figure(title: str, message: str) -> Dict[str, Any]:
    return {
        "data": [],
        "layout": {
            "title": {"text": title, "x": 0.02, "font": {"size": 20}},
            "annotations": [
                {
                    "text": message,
                    "xref": "paper",
                    "yref": "paper",
                    "x": 0.5,
                    "y": 0.5,
                    "showarrow": False,
                    "font": {"size": 15, "color": "#334155"},
                }
            ],
            "xaxis": {"visible": False},
            "yaxis": {"visible": False},
            "paper_bgcolor": "rgba(255,255,255,0)",
            "plot_bgcolor": "rgba(255,255,255,0)",
            "margin": {"l": 40, "r": 20, "t": 80, "b": 40},
        },
    }


def base_layout(title: str, y_title: str, years: List[Any], barmode: str = "stack", is_x_years: bool = True) -> Dict[str, Any]:
    if is_x_years:
        tickvals = []
        for y in years:
            if y is None: continue
            try:
                if int(float(y)) % 5 == 0:
                    tickvals.append(y)
            except (ValueError, TypeError):
                pass
    else:
        tickvals = years
        
    return {
        "title": {"text": title, "x": 0.02, "font": {"size": 20}},
        "xaxis": {
            "title": "Year" if is_x_years else "Project",
            "type": "category",
            "tickvals": tickvals,
            "automargin": True,
            "gridcolor": "rgba(37,99,235,0.08)",
        },
        "yaxis": {
            "title": y_title,
            "automargin": True,
            "gridcolor": "rgba(37,99,235,0.12)",
            "zerolinecolor": "rgba(15,23,42,0.3)",
        },
        "barmode": barmode,
        "legend": {
            "orientation": "h",
            "yanchor": "bottom",
            "y": 1.02,
            "xanchor": "left",
            "x": 0.0,
            "font": {"size": 11},
        },
        "hovermode": "x unified",
        "paper_bgcolor": "rgba(255,255,255,0)",
        "plot_bgcolor": "rgba(255,255,255,0)",
        "margin": {"l": 66, "r": 40, "t": 86, "b": 70},
    }


def build_capex_opex_graph(df_costs: pd.DataFrame) -> Tuple[Dict[str, Any], str]:
    title = "CAPEX / OPEX Breakdown"
    if df_costs.empty or "Year" not in df_costs.columns:
        return (
            placeholder_figure(title, "No Technology_Costs data is available for this scenario."),
            "This graph requires the Technology_Costs sheet with a Year column.",
        )

    years = year_axis(df_costs["Year"])
    numeric_cols = []
    for col in df_costs.columns:
        if col == "Year":
            continue
        numeric = pd.to_numeric(df_costs[col], errors="coerce")
        if numeric.notna().any():
            numeric_cols.append(col)

    excluded = {
        "Year",
        "Budget_Limit",
        "Total_Limit",
        "DAC_Opex",
        "Credit_Cost",
        "Financing Interests",
    }
    capex_cols = [
        col
        for col in numeric_cols
        if col not in excluded and not col.startswith("Aid_") and not col.endswith("_labels")
    ]
    capex_cols = non_zero_columns(df_costs, capex_cols)

    opex_cols = [col for col in ["DAC_Opex", "Credit_Cost", "Financing Interests"] if col in numeric_cols]
    opex_cols = non_zero_columns(df_costs, opex_cols)

    aid_cols = [col for col in numeric_cols if col.startswith("Aid_")]
    aid_cols = non_zero_columns(df_costs, aid_cols)

    if not capex_cols and not opex_cols and not aid_cols:
        return (
            placeholder_figure(title, "All CAPEX/OPEX/Aid values are zero for this scenario."),
            "The script detected the required sheet but all tracked cost columns sum to zero.",
        )

    df_plot, capex_cols = split_major_minor(df_costs, capex_cols, MAX_STACK_SERIES, "Other_CAPEX")

    capex_palette = [
        "#1D4ED8",
        "#0284C7",
        "#0EA5E9",
        "#2563EB",
        "#0F766E",
        "#16A34A",
        "#4F46E5",
        "#0369A1",
        "#334155",
    ]

    traces = []
    for idx, col in enumerate(capex_cols):
        series = pd.to_numeric(df_plot[col], errors="coerce").fillna(0.0)
        traces.append(
            {
                "type": "bar",
                "name": f"CAPEX - {pretty_label(col)}",
                "x": years,
                "y": to_float_list(series, scale=1_000_000.0),
                "marker": {"color": capex_palette[idx % len(capex_palette)]},
                "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
            }
        )

    if opex_cols:
        total_opex = pd.DataFrame(
            {col: pd.to_numeric(df_plot[col], errors="coerce").fillna(0.0) for col in opex_cols}
        ).sum(axis=1)
        traces.append(
            {
                "type": "scatter",
                "mode": "lines+markers",
                "name": "OPEX proxy (DAC + Credits + Interest)",
                "x": years,
                "y": to_float_list(total_opex, scale=1_000_000.0),
                "line": {"width": 3, "color": "#14B8A6"},
                "marker": {"size": 6, "color": "#14B8A6"},
                "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
            }
        )

    if aid_cols:
        total_aids = pd.DataFrame(
            {col: pd.to_numeric(df_plot[col], errors="coerce").fillna(0.0) for col in aid_cols}
        ).sum(axis=1)
        traces.append(
            {
                "type": "scatter",
                "mode": "lines+markers",
                "name": "Public aids (GRANT + CCFD)",
                "x": years,
                "y": to_float_list(total_aids, scale=1_000_000.0),
                "line": {"width": 2, "dash": "dot", "color": "#0F172A"},
                "marker": {"size": 5, "color": "#0F172A"},
                "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
            }
        )

    layout = base_layout(title, "Cost (MEUR)", years, barmode="stack")
    description = (
        "This chart is built from the Technology_Costs sheet. CAPEX bars aggregate yearly technology investment "
        "columns (excluding Aid_* and control fields), converted from EUR to MEUR. The OPEX proxy line sums "
        "DAC_Opex, Credit_Cost, and Financing Interests when available. Public aids are aggregated from Aid_GRANT_* "
        "and Aid_CCFD_* columns as a separate reference line."
    )
    return {"data": traces, "layout": layout}, description


def build_hydrogen_flow_graph(df_energy: pd.DataFrame) -> Tuple[Dict[str, Any], str]:
    title = "Hydrogen Flow Over Time"
    if df_energy.empty or "Year" not in df_energy.columns:
        return (
            placeholder_figure(title, "No Energy_Mix data is available for this scenario."),
            "This graph requires the Energy_Mix sheet with a Year column.",
        )

    years = year_axis(df_energy["Year"])
    h2_cols = [col for col in df_energy.columns if col != "Year" and "H2" in str(col).upper()]
    h2_cols = non_zero_columns(df_energy, h2_cols)

    if not h2_cols:
        return (
            placeholder_figure(title, "No hydrogen-related columns were found (or values are zero)."),
            "Hydrogen flows are detected by columns containing 'H2' in the Energy_Mix sheet.",
        )

    df_plot, h2_cols = split_major_minor(df_energy, h2_cols, MAX_LINE_SERIES, "Other_H2")

    palette = [
        "#16A34A",
        "#22C55E",
        "#10B981",
        "#2DD4BF",
        "#0284C7",
        "#0EA5E9",
        "#1D4ED8",
        "#334155",
    ]
    traces = []
    for idx, col in enumerate(h2_cols):
        series = pd.to_numeric(df_plot[col], errors="coerce").fillna(0.0)
        traces.append(
            {
                "type": "scatter",
                "mode": "lines+markers",
                "name": pretty_label(col),
                "x": years,
                "y": to_float_list(series),
                "line": {"width": 2.5, "color": palette[idx % len(palette)]},
                "marker": {"size": 5, "color": palette[idx % len(palette)]},
                "hovertemplate": "%{y:,.2f}<extra>%{fullData.name}</extra>",
            }
        )

    layout = base_layout(title, "Hydrogen volume (native units)", years, barmode="group")
    description = (
        "This chart reads the Energy_Mix sheet and keeps all non-zero columns whose names contain 'H2'. "
        "Each line represents a hydrogen source/sink stream over time (for example purchased colors or on-site production). "
        "If many streams exist, smaller ones are grouped into 'Other_H2' to preserve readability."
    )
    return {"data": traces, "layout": layout}, description


def build_utility_graph(df_energy: pd.DataFrame) -> Tuple[Dict[str, Any], str]:
    title = "Utility Consumption"
    if df_energy.empty or "Year" not in df_energy.columns:
        return (
            placeholder_figure(title, "No Energy_Mix data is available for this scenario."),
            "This graph requires the Energy_Mix sheet with a Year column.",
        )

    years = year_axis(df_energy["Year"])
    utility_cols = [col for col in df_energy.columns if col != "Year" and "H2" not in str(col).upper()]
    utility_cols = non_zero_columns(df_energy, utility_cols)

    if not utility_cols:
        return (
            placeholder_figure(title, "No non-hydrogen utility columns were found (or values are zero)."),
            "Utility series are all Energy_Mix columns excluding hydrogen-tagged ones.",
        )

    df_plot, utility_cols = split_major_minor(df_energy, utility_cols, MAX_LINE_SERIES, "Other_Utilities")

    palette = [
        "#1D4ED8",
        "#2563EB",
        "#0284C7",
        "#0EA5E9",
        "#0F766E",
        "#14B8A6",
        "#4F46E5",
        "#334155",
    ]

    traces = []
    for idx, col in enumerate(utility_cols):
        series = pd.to_numeric(df_plot[col], errors="coerce").fillna(0.0)
        traces.append(
            {
                "type": "scatter",
                "mode": "lines+markers",
                "name": pretty_label(col),
                "x": years,
                "y": to_float_list(series),
                "line": {"width": 2.5, "color": palette[idx % len(palette)]},
                "marker": {"size": 4.5, "color": palette[idx % len(palette)]},
                "hovertemplate": "%{y:,.2f}<extra>%{fullData.name}</extra>",
            }
        )

    layout = base_layout(title, "Utility consumption (native units)", years, barmode="group")
    description = (
        "This graph uses Energy_Mix and plots non-hydrogen utility streams through time. "
        "It is useful to track electrical/fuel/water burdens while decarbonization technologies are activated. "
        "When many utility series exist, smaller series are grouped in 'Other_Utilities'."
    )
    return {"data": traces, "layout": layout}, description


def build_financial_npv_graph(
    df_financing: pd.DataFrame,
    df_costs: pd.DataFrame,
    df_co2: pd.DataFrame,
    discount_rate: float,
) -> Tuple[Dict[str, Any], str]:
    title = "Financial NPV"
    if df_financing.empty or "Year" not in df_financing.columns:
        return (
            placeholder_figure(title, "No Financing data is available for this scenario."),
            "This graph requires the Financing sheet with a Year column.",
        )

    years = year_axis(df_financing["Year"])

    col_oop = find_column(df_financing, ["out_of_pocket_capex", "out of pocket capex"])
    col_total_capex = find_column(df_financing, ["total_capex"]) 
    col_annuity = find_column(df_financing, ["total_annuity"])

    base_cash_out = pd.Series([0.0] * len(df_financing))
    if col_oop:
        base_cash_out += pd.to_numeric(df_financing[col_oop], errors="coerce").fillna(0.0)
    elif col_total_capex:
        base_cash_out += pd.to_numeric(df_financing[col_total_capex], errors="coerce").fillna(0.0)

    if col_annuity:
        base_cash_out += pd.to_numeric(df_financing[col_annuity], errors="coerce").fillna(0.0)

    aids_meur = pd.Series([0.0] * len(df_financing))
    opex_meur = pd.Series([0.0] * len(df_financing))

    if not df_costs.empty and "Year" in df_costs.columns:
        aid_cols = [col for col in df_costs.columns if str(col).startswith("Aid_")]
        aid_cols = non_zero_columns(df_costs, aid_cols)
        if aid_cols:
            aids_by_year = pd.DataFrame(
                {col: pd.to_numeric(df_costs[col], errors="coerce").fillna(0.0) for col in aid_cols}
            ).sum(axis=1)
            by_year = dict(zip(year_axis(df_costs["Year"]), to_float_list(aids_by_year, scale=1_000_000.0)))
            aids_meur = pd.Series([by_year.get(y, 0.0) for y in years])

        opex_cols = [col for col in ["DAC_Opex", "Credit_Cost"] if col in df_costs.columns]
        opex_cols = non_zero_columns(df_costs, opex_cols)
        if opex_cols:
            opex_by_year = pd.DataFrame(
                {col: pd.to_numeric(df_costs[col], errors="coerce").fillna(0.0) for col in opex_cols}
            ).sum(axis=1)
            by_year = dict(zip(year_axis(df_costs["Year"]), to_float_list(opex_by_year, scale=1_000_000.0)))
            opex_meur = pd.Series([by_year.get(y, 0.0) for y in years])

    tax_meur = pd.Series([0.0] * len(df_financing))
    if not df_co2.empty and "Year" in df_co2.columns:
        tax_col = find_column(df_co2, ["net_tax_cost_meuros", "tax_cost_meuros"]) 
        if tax_col:
            tax_by_year = dict(
                zip(
                    year_axis(df_co2["Year"]),
                    to_float_list(pd.to_numeric(df_co2[tax_col], errors="coerce").fillna(0.0)),
                )
            )
            tax_meur = pd.Series([tax_by_year.get(y, 0.0) for y in years])

    # Keep every term in MEUR before discounting to avoid unit drift.
    annual_cash_flow_meur = -(base_cash_out + opex_meur + tax_meur)
    annual_cash_flow_meur += aids_meur

    discounted_cf = []
    cumulative = []
    run = 0.0
    for idx, cash_flow in enumerate(annual_cash_flow_meur):
        disc = cash_flow / ((1.0 + discount_rate) ** idx)
        discounted_cf.append(round(disc, 6))
        run += disc
        cumulative.append(round(run, 6))

    traces = [
        {
            "type": "bar",
            "name": "Annual net cash flow",
            "x": years,
            "y": [round(v, 6) for v in annual_cash_flow_meur.tolist()],
            "marker": {"color": "#0EA5E9"},
            "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
            "yaxis": "y",
        },
        {
            "type": "scatter",
            "mode": "lines+markers",
            "name": "Discounted cash flow",
            "x": years,
            "y": discounted_cf,
            "line": {"width": 2, "dash": "dot", "color": "#0F766E"},
            "marker": {"size": 5, "color": "#0F766E"},
            "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
            "yaxis": "y",
        },
        {
            "type": "scatter",
            "mode": "lines+markers",
            "name": "Cumulative NPV",
            "x": years,
            "y": cumulative,
            "line": {"width": 3, "color": "#1D4ED8"},
            "marker": {"size": 6, "color": "#1D4ED8"},
            "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
            "yaxis": "y2",
        },
    ]

    layout = base_layout(title, "Annual values (MEUR)", years, barmode="relative")
    layout["yaxis2"] = {
        "title": "Cumulative NPV (MEUR)",
        "overlaying": "y",
        "side": "right",
        "automargin": True,
        "gridcolor": "rgba(0,0,0,0)",
    }
    layout["barmode"] = "relative"

    description = (
        "The annual net cash flow is approximated as: -(Out_of_Pocket_CAPEX + Total_Annuity + DAC/Credit OPEX + carbon tax) "
        "+ public aids. Cumulative NPV is then computed from discounted cash flows using the configured discount rate. "
        "This provides a scenario-level financial trajectory from model outputs without external files."
    )
    return {"data": traces, "layout": layout}, description


def _series_map_by_year(df: pd.DataFrame, value_col: str, scale: float = 1.0) -> Dict[Any, float]:
    if df.empty or "Year" not in df.columns or value_col not in df.columns:
        return {}
    years = year_axis(df["Year"])
    values = to_float_list(pd.to_numeric(df[value_col], errors="coerce").fillna(0.0), scale=scale)
    return {year: value for year, value in zip(years, values)}


def _aligned_from_map(target_years: List[Any], value_map: Dict[Any, float]) -> List[float]:
    return [round(float(value_map.get(y, 0.0)), 6) for y in target_years]


def build_carbon_price_graph(df_co2: pd.DataFrame) -> Tuple[Dict[str, Any], str]:
    title = "CARBON PRICE"
    if df_co2.empty or "Year" not in df_co2.columns or "Tax_Price" not in df_co2.columns:
        return (
            placeholder_figure(title, "No carbon price data is available in CO2_Trajectory."),
            "This graph uses the Tax_Price column from the CO2_Trajectory sheet.",
        )

    years = year_axis(df_co2["Year"])
    price = to_float_list(pd.to_numeric(df_co2["Tax_Price"], errors="coerce").fillna(0.0))
    traces = [
        {
            "type": "scatter",
            "mode": "lines+markers",
            "name": "Carbon price",
            "x": years,
            "y": price,
            "line": {"width": 3, "color": "#1D4ED8"},
            "marker": {"size": 6, "color": "#1D4ED8"},
            "hovertemplate": "%{y:,.2f} EUR/tCO2<extra>%{fullData.name}</extra>",
        }
    ]
    layout = base_layout(title, "EUR / tCO2", years, barmode="group")
    description = (
        "This curve is directly read from CO2_Trajectory.Tax_Price. "
        "It represents the simulated carbon market unit price applied each year."
    )
    return {"data": traces, "layout": layout}, description


def build_carbon_tax_graph(df_co2: pd.DataFrame) -> Tuple[Dict[str, Any], str]:
    title = "CARBON TAX"
    if df_co2.empty or "Year" not in df_co2.columns:
        return (
            placeholder_figure(title, "No CO2_Trajectory data is available for carbon tax."),
            "This graph requires Tax_Cost_MEuros / Net_Tax_Cost_MEuros columns in CO2_Trajectory.",
        )

    years = year_axis(df_co2["Year"])
    gross_col = "Tax_Cost_MEuros" if "Tax_Cost_MEuros" in df_co2.columns else ""
    net_col = "Net_Tax_Cost_MEuros" if "Net_Tax_Cost_MEuros" in df_co2.columns else ""
    refund_col = "CCfD_Refund_MEuros" if "CCfD_Refund_MEuros" in df_co2.columns else ""

    if not gross_col and not net_col and not refund_col:
        return (
            placeholder_figure(title, "Tax columns are missing in CO2_Trajectory."),
            "No carbon-tax cost columns were found in the result workbook.",
        )

    traces = []
    if gross_col:
        traces.append(
            {
                "type": "bar",
                "name": "Gross carbon tax",
                "x": years,
                "y": to_float_list(pd.to_numeric(df_co2[gross_col], errors="coerce").fillna(0.0)),
                "marker": {"color": "#0369A1"},
                "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
            }
        )
    if refund_col:
        traces.append(
            {
                "type": "bar",
                "name": "CCfD refund",
                "x": years,
                "y": to_float_list(pd.to_numeric(df_co2[refund_col], errors="coerce").fillna(0.0)),
                "marker": {"color": "#10B981"},
                "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
            }
        )
    if net_col:
        traces.append(
            {
                "type": "scatter",
                "mode": "lines+markers",
                "name": "Net carbon tax",
                "x": years,
                "y": to_float_list(pd.to_numeric(df_co2[net_col], errors="coerce").fillna(0.0)),
                "line": {"width": 3, "color": "#0F172A"},
                "marker": {"size": 6, "color": "#0F172A"},
                "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
            }
        )

    layout = base_layout(title, "MEUR", years, barmode="relative")
    description = (
        "Gross tax is sourced from Tax_Cost_MEuros, refunds from CCfD_Refund_MEuros, and net tax from "
        "Net_Tax_Cost_MEuros when available."
    )
    return {"data": traces, "layout": layout}, description


def build_co2_trajectory_full_graph(df_co2: pd.DataFrame) -> Tuple[Dict[str, Any], str]:
    title = "CO2 TRAJECTORY"
    if df_co2.empty or "Year" not in df_co2.columns:
        return (
            placeholder_figure(title, "No CO2_Trajectory sheet is available."),
            "This graph requires Direct_CO2 / Indirect_CO2 / Total_CO2 columns.",
        )

    years = year_axis(df_co2["Year"])
    direct_kt = to_float_list(pd.to_numeric(df_co2.get("Direct_CO2", 0.0), errors="coerce").fillna(0.0), scale=1000.0)
    indirect_kt = to_float_list(pd.to_numeric(df_co2.get("Indirect_CO2", 0.0), errors="coerce").fillna(0.0), scale=1000.0)
    total_kt = to_float_list(pd.to_numeric(df_co2.get("Total_CO2", 0.0), errors="coerce").fillna(0.0), scale=1000.0)
    free_quota_kt = to_float_list(pd.to_numeric(df_co2.get("Free_Quota", 0.0), errors="coerce").fillna(0.0), scale=1000.0)
    taxed_kt = to_float_list(pd.to_numeric(df_co2.get("Taxed_CO2", 0.0), errors="coerce").fillna(0.0), scale=1000.0)

    traces = [
        {
            "type": "scatter",
            "mode": "lines+markers",
            "name": "Direct CO2",
            "x": years,
            "y": direct_kt,
            "line": {"width": 3, "color": "#111827"},
            "marker": {"size": 5, "color": "#111827"},
            "hovertemplate": "%{y:,.2f} ktCO2<extra>%{fullData.name}</extra>",
        },
        {
            "type": "scatter",
            "mode": "lines+markers",
            "name": "Indirect CO2",
            "x": years,
            "y": indirect_kt,
            "line": {"width": 2.5, "dash": "dot", "color": "#0284C7"},
            "marker": {"size": 5, "color": "#0284C7"},
            "hovertemplate": "%{y:,.2f} ktCO2<extra>%{fullData.name}</extra>",
        },
        {
            "type": "scatter",
            "mode": "lines+markers",
            "name": "Total CO2",
            "x": years,
            "y": total_kt,
            "line": {"width": 3, "color": "#DC2626"},
            "marker": {"size": 5, "color": "#DC2626"},
            "hovertemplate": "%{y:,.2f} ktCO2<extra>%{fullData.name}</extra>",
        },
        {
            "type": "bar",
            "name": "Free quota",
            "x": years,
            "y": free_quota_kt,
            "marker": {"color": "rgba(16,185,129,0.35)"},
            "hovertemplate": "%{y:,.2f} ktCO2<extra>%{fullData.name}</extra>",
        },
        {
            "type": "bar",
            "name": "Taxed emissions",
            "x": years,
            "y": taxed_kt,
            "marker": {"color": "rgba(71,85,105,0.45)"},
            "hovertemplate": "%{y:,.2f} ktCO2<extra>%{fullData.name}</extra>",
        },
    ]
    layout = base_layout(title, "ktCO2", years, barmode="stack")
    description = (
        "The trajectory combines direct, indirect, and total emissions with the free-quota and taxed-emissions "
        "decomposition from CO2_Trajectory. Emission values are displayed in ktCO2 for readability."
    )
    return {"data": traces, "layout": layout}, description


def build_energy_mix_full_graph(df_energy: pd.DataFrame) -> Tuple[Dict[str, Any], str]:
    title = "ENERGY MIX"
    if df_energy.empty or "Year" not in df_energy.columns:
        return (
            placeholder_figure(title, "No Energy_Mix data is available."),
            "This graph requires the Energy_Mix sheet with yearly resource columns.",
        )

    years = year_axis(df_energy["Year"])
    cols = [col for col in df_energy.columns if col != "Year"]
    cols = non_zero_columns(df_energy, cols)
    if not cols:
        return (
            placeholder_figure(title, "All Energy_Mix series are zero."),
            "No non-zero resource stream was found in Energy_Mix.",
        )

    df_plot, cols = split_major_minor(df_energy, cols, MAX_STACK_SERIES, "Other_Resources")
    palette = [
        "#1D4ED8",
        "#0284C7",
        "#0EA5E9",
        "#14B8A6",
        "#16A34A",
        "#22C55E",
        "#4F46E5",
        "#334155",
        "#EA580C",
    ]

    traces = []
    for idx, col in enumerate(cols):
        traces.append(
            {
                "type": "scatter",
                "mode": "lines",
                "name": pretty_label(col),
                "x": years,
                "y": to_float_list(pd.to_numeric(df_plot[col], errors="coerce").fillna(0.0)),
                "stackgroup": "one",
                "line": {"width": 1.5, "color": palette[idx % len(palette)]},
                "hovertemplate": "%{y:,.2f}<extra>%{fullData.name}</extra>",
            }
        )

    layout = base_layout(title, "Mixed units (as exported)", years, barmode="stack")
    description = (
        "Energy_Mix is rendered as a stacked area across all non-zero resource columns. "
        "Series are grouped into 'Other_Resources' when needed to keep the legend readable."
    )
    return {"data": traces, "layout": layout}, description


def build_external_financing_graph(df_financing: pd.DataFrame, df_costs: pd.DataFrame) -> Tuple[Dict[str, Any], str]:
    title = "FINANCING"
    if df_financing.empty or "Year" not in df_financing.columns:
        return (
            placeholder_figure(title, "No Financing sheet is available."),
            "This graph requires financing columns such as Loan_Principal_Taken (M€) and Total_Annuity (M€).",
        )

    years = year_axis(df_financing["Year"])
    col_loan = find_column(df_financing, ["loan_principal_taken"])
    col_oop = find_column(df_financing, ["out_of_pocket_capex"])
    col_principal = find_column(df_financing, ["principal_repayment"])
    col_interest = find_column(df_financing, ["interest_paid"])
    col_annuity = find_column(df_financing, ["total_annuity"])

    traces = []
    for col, label, color in [
        (col_loan, "Loan principal taken", "#0284C7"),
        (col_oop, "Out-of-pocket CAPEX", "#1D4ED8"),
        (col_principal, "Principal repayment", "#0F766E"),
        (col_interest, "Interest paid", "#EA580C"),
    ]:
        if col:
            traces.append(
                {
                    "type": "bar",
                    "name": label,
                    "x": years,
                    "y": to_float_list(pd.to_numeric(df_financing[col], errors="coerce").fillna(0.0)),
                    "marker": {"color": color},
                    "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
                }
            )

    if col_annuity:
        traces.append(
            {
                "type": "scatter",
                "mode": "lines+markers",
                "name": "Total annuity",
                "x": years,
                "y": to_float_list(pd.to_numeric(df_financing[col_annuity], errors="coerce").fillna(0.0)),
                "line": {"width": 3, "dash": "dot", "color": "#111827"},
                "marker": {"size": 6, "color": "#111827"},
                "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
            }
        )

    if not df_costs.empty and "Year" in df_costs.columns:
        budget_map = _series_map_by_year(df_costs, "Budget_Limit", scale=1_000_000.0)
        total_limit_map = _series_map_by_year(df_costs, "Total_Limit", scale=1_000_000.0)
        if budget_map:
            traces.append(
                {
                    "type": "scatter",
                    "mode": "lines",
                    "name": "Budget limit",
                    "x": years,
                    "y": _aligned_from_map(years, budget_map),
                    "line": {"width": 2, "color": "#16A34A"},
                    "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
                }
            )
        if total_limit_map:
            traces.append(
                {
                    "type": "scatter",
                    "mode": "lines",
                    "name": "Total investment limit",
                    "x": years,
                    "y": _aligned_from_map(years, total_limit_map),
                    "line": {"width": 2, "dash": "dash", "color": "#166534"},
                    "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
                }
            )

    if not traces:
        return (
            placeholder_figure(title, "No usable financing columns were found."),
            "The workbook does not contain financing metrics needed for this chart.",
        )

    layout = base_layout(title, "MEUR", years, barmode="relative")
    description = (
        "External financing combines loan intake, out-of-pocket CAPEX, debt service, and yearly annuities from the Financing sheet. "
        "When available, budget constraints are overlaid from Technology_Costs (Budget_Limit / Total_Limit)."
    )
    return {"data": traces, "layout": layout}, description


def build_indirect_emissions_graph(df_indir: pd.DataFrame) -> Tuple[Dict[str, Any], str]:
    title = "INDIRECT EMISSIONS"
    if df_indir.empty or "Year" not in df_indir.columns:
        return (
            placeholder_figure(title, "No Indirect_Emissions sheet is available."),
            "This chart needs the Indirect_Emissions sheet with Year and contribution columns.",
        )

    years = year_axis(df_indir["Year"])
    cols = [col for col in df_indir.columns if col != "Year"]
    cols = non_zero_columns(df_indir, cols)
    if not cols:
        return (
            placeholder_figure(title, "All indirect-emission series are zero."),
            "No non-zero indirect-emission contributor was found.",
        )

    df_plot, cols = split_major_minor(df_indir, cols, MAX_STACK_SERIES, "Other_Indirect")
    palette = [
        "#0EA5E9",
        "#0284C7",
        "#2563EB",
        "#4F46E5",
        "#14B8A6",
        "#16A34A",
        "#475569",
        "#EA580C",
    ]

    traces = []
    for idx, col in enumerate(cols):
        traces.append(
            {
                "type": "bar",
                "name": pretty_label(col),
                "x": years,
                "y": to_float_list(pd.to_numeric(df_plot[col], errors="coerce").fillna(0.0), scale=1000.0),
                "marker": {"color": palette[idx % len(palette)]},
                "hovertemplate": "%{y:,.2f} ktCO2<extra>%{fullData.name}</extra>",
            }
        )

    layout = base_layout(title, "ktCO2", years, barmode="stack")
    description = (
        "Indirect emissions are decomposed by source from the Indirect_Emissions sheet. "
        "Values are converted from tCO2 to ktCO2."
    )
    return {"data": traces, "layout": layout}, description


def build_investment_plan_graph(df_investments: pd.DataFrame) -> Tuple[Dict[str, Any], str]:
    title = "INVESTMENT PLAN"
    required_cols = {"Year", "Technology", "Capex_Euros"}
    if df_investments.empty or not required_cols.issubset(set(df_investments.columns)):
        return (
            placeholder_figure(title, "No valid Investments sheet data is available."),
            "This chart requires Year, Technology, and Capex_Euros in Investments.",
        )

    work = df_investments.copy()
    work["Capex_Euros"] = pd.to_numeric(work["Capex_Euros"], errors="coerce").fillna(0.0)
    pivot = (
        work.groupby(["Year", "Technology"], dropna=False)["Capex_Euros"]
        .sum()
        .unstack(fill_value=0.0)
        .reset_index()
    )
    if pivot.empty:
        return (
            placeholder_figure(title, "No investments were recorded in this scenario."),
            "Investments sheet exists but contains no CAPEX events.",
        )

    years = year_axis(pivot["Year"])
    tech_cols = [col for col in pivot.columns if col != "Year"]
    tech_cols = non_zero_columns(pivot, tech_cols)
    if not tech_cols:
        return (
            placeholder_figure(title, "All investment CAPEX values are zero."),
            "No non-zero CAPEX by technology was found in Investments.",
        )

    pivot, tech_cols = split_major_minor(pivot, tech_cols, MAX_STACK_SERIES, "Other_Technologies")
    palette = [
        "#1D4ED8",
        "#2563EB",
        "#0EA5E9",
        "#10B981",
        "#22C55E",
        "#4F46E5",
        "#EA580C",
        "#475569",
    ]

    traces = []
    for idx, col in enumerate(tech_cols):
        traces.append(
            {
                "type": "bar",
                "name": pretty_label(col),
                "x": years,
                "y": to_float_list(pd.to_numeric(pivot[col], errors="coerce").fillna(0.0), scale=1_000_000.0),
                "marker": {"color": palette[idx % len(palette)]},
                "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
            }
        )

    layout = base_layout(title, "CAPEX (MEUR)", years, barmode="stack")
    description = (
        "CAPEX from Investments is aggregated by Year and Technology (sum of Capex_Euros) and rendered as stacked bars in MEUR."
    )
    return {"data": traces, "layout": layout}, description


def build_resources_opex_graph(df_costs: pd.DataFrame) -> Tuple[Dict[str, Any], str]:
    title = "RESSOURCES OPEX"
    if df_costs.empty or "Year" not in df_costs.columns:
        return (
            placeholder_figure(title, "No Technology_Costs data is available."),
            "This graph uses OPEX-like columns from Technology_Costs.",
        )

    years = year_axis(df_costs["Year"])
    opex_candidates = [col for col in ["DAC_Opex", "Credit_Cost", "Financing Interests"] if col in df_costs.columns]
    opex_candidates = non_zero_columns(df_costs, opex_candidates)
    if not opex_candidates:
        return (
            placeholder_figure(title, "No OPEX-like columns were found in Technology_Costs."),
            "Expected one of DAC_Opex, Credit_Cost, or Financing Interests.",
        )

    color_map = {
        "DAC_Opex": "#0284C7",
        "Credit_Cost": "#EA580C",
        "Financing Interests": "#475569",
    }

    traces = []
    for col in opex_candidates:
        traces.append(
            {
                "type": "bar",
                "name": pretty_label(col),
                "x": years,
                "y": to_float_list(pd.to_numeric(df_costs[col], errors="coerce").fillna(0.0), scale=1_000_000.0),
                "marker": {"color": color_map.get(col, "#1D4ED8")},
                "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
            }
        )

    total = pd.DataFrame({col: pd.to_numeric(df_costs[col], errors="coerce").fillna(0.0) for col in opex_candidates}).sum(axis=1)
    traces.append(
        {
            "type": "scatter",
            "mode": "lines+markers",
            "name": "Total ressources OPEX",
            "x": years,
            "y": to_float_list(total, scale=1_000_000.0),
            "line": {"width": 3, "color": "#0F172A"},
            "marker": {"size": 6, "color": "#0F172A"},
            "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
        }
    )

    layout = base_layout(title, "MEUR", years, barmode="stack")
    description = (
        "Resource OPEX is reconstructed from OPEX-related outputs available in Technology_Costs: DAC_Opex, Credit_Cost, "
        "and Financing Interests."
    )
    return {"data": traces, "layout": layout}, description


def build_data_used_graph(df_data_used: pd.DataFrame) -> Tuple[Dict[str, Any], str]:
    title = "DATA USED"
    description = "Displays the price and CO2 emissions for each selected resource over time."

    if df_data_used.empty or "Resource" not in df_data_used.columns or "Year" not in df_data_used.columns:
        return (
            placeholder_figure(title, "No data"),
            "Data_Used sheet is missing or empty.",
        )

    # Sort years for consistent x-axis
    years = year_axis(sorted(df_data_used["Year"].unique()))
    resources = sorted(df_data_used["Resource"].dropna().unique())

    if not resources:
        return (
            placeholder_figure(title, "No data"),
            "No resources data available.",
        )

    traces = []
    buttons = []
    traces_per_res = 2  # Price and CO2 Emissions

    for i, res in enumerate(resources):
        df_res = df_data_used[df_data_used["Resource"] == res].sort_values("Year")
        # Ensure we map all years correctly even if missing
        res_years = year_axis(df_res["Year"])
        res_price = to_float_list(pd.to_numeric(df_res["Price"], errors="coerce").fillna(0.0))
        res_co2 = to_float_list(pd.to_numeric(df_res["CO2_Emissions"], errors="coerce").fillna(0.0))

        is_first = (i == 0)

        # Trace 1: Price (left axis)
        traces.append({
            "type": "scatter",
            "mode": "lines+markers",
            "name": f"Price",
            "x": res_years,
            "y": res_price,
            "line": {"width": 3, "color": "#1D4ED8"},
            "marker": {"size": 6, "color": "#1D4ED8"},
            "hovertemplate": "%{y:,.2f} EUR/unit<extra>%{fullData.name}</extra>",
            "yaxis": "y",
            "visible": is_first
        })

        # Trace 2: CO2 Emissions (right axis)
        traces.append({
            "type": "scatter",
            "mode": "lines+markers",
            "name": f"CO2 Emissions",
            "x": res_years,
            "y": res_co2,
            "line": {"width": 2.5, "dash": "dot", "color": "#E11D48"},
            "marker": {"size": 5, "color": "#E11D48"},
            "hovertemplate": "%{y:,.4f} tCO2/unit<extra>%{fullData.name}</extra>",
            "yaxis": "y2",
            "visible": is_first
        })

        # Dropdown button logic
        vis_array = [False] * (len(resources) * traces_per_res)
        vis_array[i * traces_per_res] = True
        vis_array[i * traces_per_res + 1] = True

        buttons.append({
            "method": "update",
            "label": res,
            "args": [
                {"visible": vis_array},
                {"title": f"DATA USED - {res}"}
            ]
        })

    layout = base_layout(f"DATA USED - {resources[0]}", "Price (EUR)", years, barmode="group")
    
    # Update layout with dropdown
    layout["updatemenus"] = [{
        "active": 0,
        "buttons": buttons,
        "direction": "down",
        "showactive": True,
        "x": 0.5,
        "xanchor": "center",
        "y": 1.15,
        "yanchor": "bottom",
        "bgcolor": "white",
        "bordercolor": "#ccc",
        "font": {"color": "#333"}
    }]

    # Add secondary Y-axis
    layout["yaxis2"] = {
        "title": "CO2 Emissions (tCO2/unit)",
        "overlaying": "y",
        "side": "right",
        "showgrid": False,
        "zeroline": False
    }

    return {"data": traces, "layout": layout}, description


def build_transition_cost_graph(df_financing: pd.DataFrame, df_costs: pd.DataFrame, df_co2: pd.DataFrame, df_transition_balance: pd.DataFrame) -> Tuple[Dict[str, Any], str]:
    title = "TRANSITION COST"
    if df_financing.empty or "Year" not in df_financing.columns:
        return (
            placeholder_figure(title, "No Financing data is available."),
            "This graph requires financing + cost sheets to reconstruct transition effort.",
        )

    years = year_axis(df_financing["Year"])
    col_oop = find_column(df_financing, ["out_of_pocket_capex"])
    col_annuity = find_column(df_financing, ["total_annuity"])

    self_funded = to_float_list(pd.to_numeric(df_financing[col_oop], errors="coerce").fillna(0.0)) if col_oop else [0.0] * len(years)
    loan_service = to_float_list(pd.to_numeric(df_financing[col_annuity], errors="coerce").fillna(0.0)) if col_annuity else [0.0] * len(years)

    tech_opex_map = {}
    credits_map = {}
    if not df_costs.empty and "Year" in df_costs.columns:
        if "DAC_Opex" in df_costs.columns:
            tech_opex_map = _series_map_by_year(df_costs, "DAC_Opex", scale=1_000_000.0)
        if "Credit_Cost" in df_costs.columns:
            credits_map = _series_map_by_year(df_costs, "Credit_Cost", scale=1_000_000.0)
        if "Financing Interests" in df_costs.columns:
            interest_map = _series_map_by_year(df_costs, "Financing Interests", scale=1_000_000.0)
            for yr, val in interest_map.items():
                tech_opex_map[yr] = round(tech_opex_map.get(yr, 0.0) + val, 6)

    net_tax_map = _series_map_by_year(df_co2, "Net_Tax_Cost_MEuros", scale=1.0)

    tech_opex = _aligned_from_map(years, tech_opex_map)
    credits = _aligned_from_map(years, credits_map)
    net_tax = _aligned_from_map(years, net_tax_map)

    traces = [
        {
            "type": "bar",
            "name": "Self-funded CAPEX",
            "x": years,
            "y": self_funded,
            "marker": {"color": "#1D4ED8"},
            "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
        },
        {
            "type": "bar",
            "name": "Bank loan service",
            "x": years,
            "y": loan_service,
            "marker": {"color": "#0284C7"},
            "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
        },
        {
            "type": "bar",
            "name": "Tech & DAC OPEX",
            "x": years,
            "y": tech_opex,
            "marker": {"color": "#14B8A6"},
            "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
        },
        {
            "type": "bar",
            "name": "Voluntary carbon credits",
            "x": years,
            "y": credits,
            "marker": {"color": "#EA580C"},
            "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
        },
        {
            "type": "bar",
            "name": "Net carbon tax",
            "x": years,
            "y": net_tax,
            "marker": {"color": "#475569"},
            "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
        },
    ]

    total = [round(a + b + c + d + e, 6) for a, b, c, d, e in zip(self_funded, loan_service, tech_opex, credits, net_tax)]

    # Add extra new metrics if available
    if not df_transition_balance.empty and "Year" in df_transition_balance.columns:
        # Align by Year index
        temp_tb = df_transition_balance.copy()
        temp_tb['Year'] = temp_tb['Year'].astype(str)
        temp_tb.set_index('Year', inplace=True)
        aligned_rmc = []
        aligned_avoid = []
        
        for y in years:
            y_str = str(y)
            if y_str in temp_tb.index:
                aligned_rmc.append(float(temp_tb.loc[y_str].get("Resource Mix Change", 0.0)))
                aligned_avoid.append(float(temp_tb.loc[y_str].get("Avoided Carbon Tax", 0.0)))
            else:
                aligned_rmc.append(0.0)
                aligned_avoid.append(0.0)

        res_plus = [round(max(0, x), 6) for x in aligned_rmc]
        res_moins = [round(min(0, x), 6) for x in aligned_rmc]
        
        traces.append({
            "type": "bar",
            "name": "Ressources en plus",
            "x": years,
            "y": res_plus,
            "marker": {"color": "#8B5CF6"},
            "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
        })
        traces.append({
            "type": "bar",
            "name": "Ressources en moins",
            "x": years,
            "y": res_moins,
            "marker": {"color": "#10B981"},
            "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
        })
        traces.append({
            "type": "bar",
            "name": "Taxe carbone évitée",
            "x": years,
            "y": [round(x, 6) for x in aligned_avoid],
            "marker": {"color": "#F59E0B"},
            "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
        })
        
        total = [round(t + rp + rm + a, 6) for t, rp, rm, a in zip(total, res_plus, res_moins, aligned_avoid)]

    traces.append(
        {
            "type": "scatter",
            "mode": "lines+markers",
            "name": "Total transition cost",
            "x": years,
            "y": total,
            "line": {"width": 3, "color": "#0F172A"},
            "marker": {"size": 6, "color": "#0F172A"},
            "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
        }
    )

    layout = base_layout(title, "MEUR", years, barmode="stack")
    description = (
        "Transition cost stacks self-funded CAPEX, debt service, technology OPEX proxies, voluntary carbon-credit cost, "
        "and net carbon-tax burden, then overlays the yearly total effort."
    )
    return {"data": traces, "layout": layout}, description


def build_total_annual_opex_graph(df_costs: pd.DataFrame, df_co2: pd.DataFrame) -> Tuple[Dict[str, Any], str]:
    title = "TOTAL ANNUAL OPEX"
    if df_costs.empty or "Year" not in df_costs.columns:
        return (
            placeholder_figure(title, "No Technology_Costs data is available."),
            "This graph combines OPEX-like columns from Technology_Costs and net carbon tax from CO2_Trajectory.",
        )

    years = year_axis(df_costs["Year"])
    dac = to_float_list(pd.to_numeric(df_costs.get("DAC_Opex", 0.0), errors="coerce").fillna(0.0), scale=1_000_000.0)
    credits = to_float_list(pd.to_numeric(df_costs.get("Credit_Cost", 0.0), errors="coerce").fillna(0.0), scale=1_000_000.0)
    interests = to_float_list(pd.to_numeric(df_costs.get("Financing Interests", 0.0), errors="coerce").fillna(0.0), scale=1_000_000.0)
    net_tax = _aligned_from_map(years, _series_map_by_year(df_co2, "Net_Tax_Cost_MEuros", scale=1.0))

    total = [round(a + b + c + d, 6) for a, b, c, d in zip(dac, credits, interests, net_tax)]
    traces = [
        {
            "type": "bar",
            "name": "DAC OPEX",
            "x": years,
            "y": dac,
            "marker": {"color": "#0284C7"},
            "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
        },
        {
            "type": "bar",
            "name": "Credit cost",
            "x": years,
            "y": credits,
            "marker": {"color": "#EA580C"},
            "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
        },
        {
            "type": "bar",
            "name": "Financing interests",
            "x": years,
            "y": interests,
            "marker": {"color": "#475569"},
            "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
        },
        {
            "type": "bar",
            "name": "Net carbon tax",
            "x": years,
            "y": net_tax,
            "marker": {"color": "#0F766E"},
            "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
        },
        {
            "type": "scatter",
            "mode": "lines+markers",
            "name": "Total annual OPEX",
            "x": years,
            "y": total,
            "line": {"width": 3, "color": "#111827"},
            "marker": {"size": 6, "color": "#111827"},
            "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>",
        },
    ]

    layout = base_layout(title, "MEUR", years, barmode="stack")
    description = (
        "Total annual OPEX is reconstructed as DAC_Opex + Credit_Cost + Financing Interests + Net_Tax_Cost_MEuros. "
        "It provides a yearly operating-cost burden proxy in MEUR."
    )
    return {"data": traces, "layout": layout}, description


def build_co2_abatement_graph(df_mac: pd.DataFrame) -> Tuple[Dict[str, Any], str]:
    title = "CO2 ABATEMENT"
    if df_mac.empty:
        return (
            placeholder_figure(title, "No MAC table was found in the Charts sheet."),
            "This graph needs a MAC table in Charts containing Project and MAC columns.",
        )

    label_col = "Project" if "Project" in df_mac.columns else "Display Label"
    mac_col = "MAC (€/tCO2)" if "MAC (€/tCO2)" in df_mac.columns else ""
    capex_col = "MAC CAPEX (€/tCO2)" if "MAC CAPEX (€/tCO2)" in df_mac.columns else ""
    opex_col = "MAC OPEX (€/tCO2)" if "MAC OPEX (€/tCO2)" in df_mac.columns else ""
    abated_col = "Total Abated (tCO2)" if "Total Abated (tCO2)" in df_mac.columns else ""

    if label_col not in df_mac.columns or not mac_col:
        return (
            placeholder_figure(title, "MAC table exists but required columns are missing."),
            "Expected columns: Project and MAC (€/tCO2).",
        )

    data = df_mac.copy()
    data[label_col] = data[label_col].astype(str)
    data[mac_col] = pd.to_numeric(data[mac_col], errors="coerce").fillna(0.0)
    if abated_col:
        data = data.sort_values(by=abated_col, ascending=False)
    else:
        data = data.sort_values(by=mac_col, ascending=True)
    data = data.head(15)

    labels = data[label_col].tolist()
    traces: List[Dict[str, Any]] = []
    if capex_col and opex_col:
        traces.append(
            {
                "type": "bar",
                "name": "MAC CAPEX",
                "x": labels,
                "y": to_float_list(pd.to_numeric(data[capex_col], errors="coerce").fillna(0.0)),
                "marker": {"color": "#0284C7"},
                "hovertemplate": "%{y:,.2f} EUR/tCO2<extra>%{fullData.name}</extra>",
            }
        )
        traces.append(
            {
                "type": "bar",
                "name": "MAC OPEX",
                "x": labels,
                "y": to_float_list(pd.to_numeric(data[opex_col], errors="coerce").fillna(0.0)),
                "marker": {"color": "#14B8A6"},
                "hovertemplate": "%{y:,.2f} EUR/tCO2<extra>%{fullData.name}</extra>",
            }
        )

    traces.append(
        {
            "type": "scatter",
            "mode": "lines+markers",
            "name": "Total MAC",
            "x": labels,
            "y": to_float_list(data[mac_col]),
            "line": {"width": 3, "color": "#111827"},
            "marker": {"size": 7, "color": "#111827"},
            "hovertemplate": "%{y:,.2f} EUR/tCO2<extra>%{fullData.name}</extra>",
        }
    )

    layout = base_layout(title, "EUR / tCO2", labels, barmode="stack", is_x_years=False)
    layout["xaxis"]["tickangle"] = -30
    description = (
        "CO2 abatement cost is built from the MAC table exported in the Charts sheet. "
        "It shows CAPEX/OPEX decomposition (when available) and total MAC by project."
    )
    return {"data": traces, "layout": layout}, description


def sanitize_payload(value: Any) -> Any:
    if isinstance(value, dict):
        return {k: sanitize_payload(v) for k, v in value.items()}
    if isinstance(value, list):
        return [sanitize_payload(v) for v in value]
    if isinstance(value, tuple):
        return [sanitize_payload(v) for v in value]
    if isinstance(value, float):
        if not math.isfinite(value):
            return None
        return round(value, 6)
    if pd.isna(value):
        return None
    return value


def discover_workbooks(results_root: Path) -> Dict[str, Path]:
    roots = [results_root, results_root / "Results"]
    workbooks: Dict[str, Path] = {}

    for root in roots:
        if not root.exists() or not root.is_dir():
            continue
        for child in sorted(root.iterdir(), key=lambda p: p.name.lower()):
            if not child.is_dir():
                continue
            if child.name.lower() == "results":
                continue
            workbook = child / "Master_Plan.xlsx"
            if workbook.exists() and child.name not in workbooks:
                workbooks[child.name] = workbook

    return workbooks


def load_sheet(workbook: Path, sheet: str) -> pd.DataFrame:
    try:
        return pd.read_excel(workbook, sheet_name=sheet)
    except ValueError:
        return pd.DataFrame()


def load_transition_balance_table(workbook: Path) -> pd.DataFrame:
    try:
        raw = pd.read_excel(workbook, sheet_name="Charts", header=None)
    except ValueError:
        return pd.DataFrame()

    header_row = None
    for idx, row in raw.iterrows():
        values = [str(v).strip() for v in row.tolist() if pd.notna(v)]
        if any("Self-funded CAPEX" in v for v in values) and any("Net_Cumulative_Cost" in v for v in values):
            header_row = idx
            break

    if header_row is None:
        return pd.DataFrame()

    df = pd.read_excel(workbook, sheet_name="Charts", header=header_row)
    df = df.dropna(how="all")
    unnamed = [col for col in df.columns if str(col).startswith("Unnamed:")]
    if unnamed:
        df = df.drop(columns=unnamed)
    if "Year" not in df.columns:
        # Fallback if index name was not exported correctly
        if "Net_Cumulative_Cost" in df.columns:
            df = df.rename(columns={df.columns[0]: "Year"})
    return df.reset_index(drop=True)

def load_mac_table(workbook: Path) -> pd.DataFrame:
    try:
        raw = pd.read_excel(workbook, sheet_name="Charts", header=None)
    except ValueError:
        return pd.DataFrame()

    header_row = None
    for idx, row in raw.iterrows():
        values = [str(v).strip() for v in row.tolist() if pd.notna(v)]
        if any(v == "Project" for v in values) and any("MAC" in v for v in values):
            header_row = idx
            break

    if header_row is None:
        return pd.DataFrame()

    df = pd.read_excel(workbook, sheet_name="Charts", header=header_row)
    df = df.dropna(how="all")
    unnamed = [col for col in df.columns if str(col).startswith("Unnamed:")]
    if unnamed:
        df = df.drop(columns=unnamed)
    if "Project" not in df.columns:
        return pd.DataFrame()
    return df.reset_index(drop=True)


def build_dashboard_data(workbooks: Dict[str, Path], discount_rate: float) -> Dict[str, Any]:
    latest_ts = max(path.stat().st_mtime for path in workbooks.values())
    generation_date = datetime.fromtimestamp(latest_ts).strftime("%Y-%m-%d %H:%M:%S")

    scenarios: Dict[str, Any] = {}
    for scenario_name, workbook in workbooks.items():
        df_energy = load_sheet(workbook, "Energy_Mix")
        df_costs = load_sheet(workbook, "Technology_Costs")
        df_financing = load_sheet(workbook, "Financing")
        df_co2 = load_sheet(workbook, "CO2_Trajectory")
        df_indir = load_sheet(workbook, "Indirect_Emissions")
        df_invest = load_sheet(workbook, "Investments")
        df_mac = load_mac_table(workbook)
        df_transition_balance = load_transition_balance_table(workbook)
        df_data_used = load_sheet(workbook, "Data_Used")

        carbon_price_fig, carbon_price_desc = build_carbon_price_graph(df_co2)
        carbon_tax_fig, carbon_tax_desc = build_carbon_tax_graph(df_co2)
        co2_traj_fig, co2_traj_desc = build_co2_trajectory_full_graph(df_co2)
        energy_mix_fig, energy_mix_desc = build_energy_mix_full_graph(df_energy)
        ext_finance_fig, ext_finance_desc = build_external_financing_graph(df_financing, df_costs)
        indirect_fig, indirect_desc = build_indirect_emissions_graph(df_indir)
        invest_fig, invest_desc = build_investment_plan_graph(df_invest)
        resources_opex_fig, resources_opex_desc = build_resources_opex_graph(df_costs)
        data_used_fig, data_used_desc = build_data_used_graph(df_data_used)
        transition_fig, transition_desc = build_transition_cost_graph(df_financing, df_costs, df_co2, df_transition_balance)
        total_opex_fig, total_opex_desc = build_total_annual_opex_graph(df_costs, df_co2)
        co2_abatement_fig, co2_abatement_desc = build_co2_abatement_graph(df_mac)

        scenarios[scenario_name] = {
            "displayName": scenario_name,
            "sourceWorkbook": str(workbook),
            "graphs": {
                "carbon_price": {
                    "label": "CARBON PRICE",
                    "title": "CARBON PRICE",
                    "description": carbon_price_desc,
                    "downloadName": sanitize_filename(f"{scenario_name}_carbon_price"),
                    "figure": carbon_price_fig,
                },
                "carbon_tax": {
                    "label": "CARBON TAX",
                    "title": "CARBON TAX",
                    "description": carbon_tax_desc,
                    "downloadName": sanitize_filename(f"{scenario_name}_carbon_tax"),
                    "figure": carbon_tax_fig,
                },
                "co2_trajectory": {
                    "label": "CO2 TRAJECTORY",
                    "title": "CO2 TRAJECTORY",
                    "description": co2_traj_desc,
                    "downloadName": sanitize_filename(f"{scenario_name}_co2_trajectory"),
                    "figure": co2_traj_fig,
                },
                "energy_mix": {
                    "label": "ENERGY MIX",
                    "title": "ENERGY MIX",
                    "description": energy_mix_desc,
                    "downloadName": sanitize_filename(f"{scenario_name}_energy_mix"),
                    "figure": energy_mix_fig,
                },
                "external_financing": {
                    "label": "FINANCING",
                    "title": "FINANCING",
                    "description": ext_finance_desc,
                    "downloadName": sanitize_filename(f"{scenario_name}_external_financing"),
                    "figure": ext_finance_fig,
                },
                "indirect_emissions": {
                    "label": "INDIRECT EMISSIONS",
                    "title": "INDIRECT EMISSIONS",
                    "description": indirect_desc,
                    "downloadName": sanitize_filename(f"{scenario_name}_indirect_emissions"),
                    "figure": indirect_fig,
                },
                "investment_plan": {
                    "label": "INVESTMENT PLAN",
                    "title": "INVESTMENT PLAN",
                    "description": invest_desc,
                    "downloadName": sanitize_filename(f"{scenario_name}_investment_plan"),
                    "figure": invest_fig,
                },
                "ressources_opex": {
                    "label": "RESSOURCES OPEX",
                    "title": "RESSOURCES OPEX",
                    "description": resources_opex_desc,
                    "downloadName": sanitize_filename(f"{scenario_name}_ressources_opex"),
                    "figure": resources_opex_fig,
                },
                "data_used": {
                    "label": "DATA USED",
                    "title": "DATA USED",
                    "description": data_used_desc,
                    "downloadName": sanitize_filename(f"{scenario_name}_data_used"),
                    "figure": data_used_fig,
                },
                "transition_cost": {
                    "label": "TRANSITION COST",
                    "title": "TRANSITION COST",
                    "description": transition_desc,
                    "downloadName": sanitize_filename(f"{scenario_name}_transition_cost"),
                    "figure": transition_fig,
                },
                "total_annual_opex": {
                    "label": "TOTAL ANNUAL OPEX",
                    "title": "TOTAL ANNUAL OPEX",
                    "description": total_opex_desc,
                    "downloadName": sanitize_filename(f"{scenario_name}_total_annual_opex"),
                    "figure": total_opex_fig,
                },
                "co2_abatement": {
                    "label": "CO2 ABATEMENT",
                    "title": "CO2 ABATEMENT",
                    "description": co2_abatement_desc,
                    "downloadName": sanitize_filename(f"{scenario_name}_co2_abatement"),
                    "figure": co2_abatement_fig,
                },
            },
        }

    payload = {
        "projectTitle": "Plant-Optimization-PathWay - Results Dashboard",
        "generationDate": generation_date,
        "discountRate": discount_rate,
        "scenarios": scenarios,
    }
    return sanitize_payload(payload)


def build_html(payload: Dict[str, Any]) -> str:
    payload_json = json.dumps(payload, ensure_ascii=True)

    template = """<!DOCTYPE html>
<html lang=\"en\">
<head>
  <meta charset=\"UTF-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\" />
  <title>Plant-Optimization-PathWay Dashboard</title>
  <script src=\"https://cdn.tailwindcss.com\"></script>
  <script>
    tailwind.config = {
      theme: {
        extend: {
          fontFamily: {
            heading: ['\"Sora\"', 'sans-serif'],
            body: ['\"Manrope\"', 'sans-serif']
          },
          boxShadow: {
            glass: '0 20px 50px rgba(15, 23, 42, 0.24)',
          }
        }
      }
    }
  </script>
  <link rel=\"preconnect\" href=\"https://fonts.googleapis.com\" />
  <link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin />
  <link href=\"https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700&family=Sora:wght@500;600;700&display=swap\" rel=\"stylesheet\" />
  <script src=\"https://cdn.plot.ly/plotly-2.35.2.min.js\"></script>
  <link rel=\"stylesheet\" href=\"https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css\" integrity=\"sha512-SnH5WK+bZxgPHs44uWix+LLJAJ9/2PkPKZ5QiAj6Ta86w+fsb2TkR4j8f5Z5gDQL4x0XSLwWf2fQJKfG8d8gQw==\" crossorigin=\"anonymous\" referrerpolicy=\"no-referrer\" />
  <style>
    :root {
      --bg-1: #0b1e3d;
      --bg-2: #11385f;
      --bg-3: #0f766e;
      --glass: rgba(255, 255, 255, 0.68);
      --glass-strong: rgba(255, 255, 255, 0.82);
      --text-main: #0f172a;
      --text-muted: #334155;
      --border-soft: rgba(255, 255, 255, 0.45);
    }

    html, body {
      min-height: 100%;
    }

    body {
      margin: 0;
      font-family: 'Manrope', sans-serif;
      color: var(--text-main);
      background:
        radial-gradient(circle at 10% 20%, rgba(14, 165, 233, 0.22) 0%, rgba(14, 165, 233, 0) 36%),
        radial-gradient(circle at 85% 12%, rgba(16, 185, 129, 0.25) 0%, rgba(16, 185, 129, 0) 34%),
        linear-gradient(125deg, var(--bg-1), var(--bg-2) 54%, var(--bg-3));
      background-attachment: fixed;
    }

    .glass-card {
      background: linear-gradient(155deg, var(--glass) 0%, var(--glass-strong) 100%);
      backdrop-filter: blur(18px);
      -webkit-backdrop-filter: blur(18px);
      border: 1px solid var(--border-soft);
      box-shadow: 0 20px 50px rgba(15, 23, 42, 0.24);
      transition: transform 260ms ease, box-shadow 260ms ease;
    }

    .glass-card:hover {
      transform: translateY(-2px);
      box-shadow: 0 24px 58px rgba(15, 23, 42, 0.28);
    }

    .subtle-pill {
      background: rgba(15, 23, 42, 0.08);
      color: #0f172a;
      border: 1px solid rgba(15, 23, 42, 0.12);
    }

    .selector {
      width: 100%;
      border: 1px solid rgba(15, 23, 42, 0.2);
      border-radius: 0.9rem;
      padding: 0.75rem 0.9rem;
      background: rgba(255, 255, 255, 0.85);
      color: #0f172a;
      font-size: 0.95rem;
      outline: none;
      transition: border-color 180ms ease, box-shadow 180ms ease;
    }

    .selector:focus {
      border-color: #0ea5e9;
      box-shadow: 0 0 0 3px rgba(14, 165, 233, 0.2);
    }

    .primary-btn {
      background: linear-gradient(130deg, #0284c7, #0369a1);
      color: #f8fafc;
      border: 0;
      border-radius: 0.85rem;
      padding: 0.72rem 1rem;
      font-weight: 700;
      letter-spacing: 0.01em;
      transition: transform 160ms ease, filter 160ms ease;
    }

    .primary-btn:hover {
      transform: translateY(-1px);
      filter: brightness(1.05);
    }

    .primary-btn:active {
      transform: translateY(0);
    }

    #chart {
      width: 100%;
      height: 530px;
    }

    @media (max-width: 768px) {
      #chart {
        height: 430px;
      }
    }

    .fade-in {
      animation: enter 480ms ease both;
    }

    @keyframes enter {
      from {
        opacity: 0;
        transform: translateY(12px);
      }
      to {
        opacity: 1;
        transform: translateY(0);
      }
    }
  </style>
</head>
<body>
  <main class=\"max-w-7xl mx-auto px-4 md:px-8 py-7 md:py-10\">
    <section class=\"glass-card rounded-3xl p-6 md:p-8 mb-6 fade-in\">
      <div class=\"flex flex-col md:flex-row md:items-end md:justify-between gap-4\">
        <div>
          <p class=\"subtle-pill inline-flex items-center gap-2 rounded-full px-3 py-1 text-xs font-semibold uppercase tracking-wide\">
            <i class=\"fa-solid fa-industry\"></i>
            Optimization Analytics
          </p>
          <h1 id=\"dashboardTitle\" class=\"font-heading text-2xl md:text-4xl font-semibold tracking-tight mt-3\"></h1>
        </div>
        <div class=\"rounded-2xl subtle-pill px-4 py-3 text-sm md:text-base\">
          <span class=\"font-semibold\">Generation Date:</span>
          <span id=\"generationDate\" class=\"font-medium\"></span>
        </div>
      </div>
    </section>

    <section class=\"glass-card rounded-3xl p-5 md:p-6 mb-6 fade-in\">
      <div class=\"grid grid-cols-1 md:grid-cols-2 gap-4\">
        <label class=\"block\">
          <span class=\"text-sm font-semibold text-slate-700\">Scenario Selector</span>
          <select id=\"scenarioSelect\" class=\"selector mt-2\"></select>
        </label>

        <label class=\"block\">
          <span class=\"text-sm font-semibold text-slate-700\">Graph Selector</span>
          <select id=\"graphSelect\" class=\"selector mt-2\"></select>
        </label>
      </div>
    </section>

    <section class=\"glass-card rounded-3xl p-4 md:p-6 mb-6 fade-in\">
      <div class=\"flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 mb-3\">
        <h2 id=\"graphTitle\" class=\"font-heading text-lg md:text-2xl font-semibold text-slate-900\"></h2>
        <button id=\"downloadBtn\" class=\"primary-btn inline-flex items-center justify-center gap-2\">
          <i class=\"fa-solid fa-download\"></i>
          Download Chart as Image
        </button>
      </div>
      <div id=\"chart\"></div>
    </section>

    <section class=\"glass-card rounded-3xl p-5 md:p-6 fade-in\">
      <h3 class=\"font-heading text-lg md:text-xl font-semibold mb-2 text-slate-900\">How this graph is constructed</h3>
      <p id=\"graphMethod\" class=\"text-slate-700 leading-relaxed\"></p>
    </section>
  </main>

  <script>
    const dashboardData = __DASHBOARD_DATA__;

    const scenarioSelect = document.getElementById('scenarioSelect');
    const graphSelect = document.getElementById('graphSelect');
    const graphTitle = document.getElementById('graphTitle');
    const graphMethod = document.getElementById('graphMethod');
    const generationDateEl = document.getElementById('generationDate');
    const titleEl = document.getElementById('dashboardTitle');
    const downloadBtn = document.getElementById('downloadBtn');
    const chartNode = document.getElementById('chart');

    const plotConfig = {
      responsive: true,
      displaylogo: false,
      modeBarButtonsToRemove: ['lasso2d', 'select2d', 'autoScale2d'],
      toImageButtonOptions: {
        format: 'png',
        filename: 'chart',
        scale: 2,
      },
    };

    function scenarioKeys() {
      return Object.keys(dashboardData.scenarios || {});
    }

    function graphKeysForScenario(scenarioKey) {
      const scenario = dashboardData.scenarios[scenarioKey] || {};
      return Object.keys(scenario.graphs || {});
    }

    function fillScenarioSelect() {
      const keys = scenarioKeys();
      scenarioSelect.innerHTML = '';
      keys.forEach((key) => {
        const option = document.createElement('option');
        option.value = key;
        option.textContent = dashboardData.scenarios[key].displayName || key;
        scenarioSelect.appendChild(option);
      });
    }

    function fillGraphSelect(selectedScenario) {
      const keys = graphKeysForScenario(selectedScenario);
      graphSelect.innerHTML = '';
      keys.forEach((key) => {
        const graph = dashboardData.scenarios[selectedScenario].graphs[key];
        const option = document.createElement('option');
        option.value = key;
        option.textContent = graph.label || key;
        graphSelect.appendChild(option);
      });
    }

    function currentGraphPayload() {
      const scenarioKey = scenarioSelect.value;
      const graphKey = graphSelect.value;
      const scenario = dashboardData.scenarios[scenarioKey] || { graphs: {} };
      return scenario.graphs[graphKey] || null;
    }

    async function renderGraph() {
      const payload = currentGraphPayload();
      if (!payload) {
        graphTitle.textContent = 'No graph available';
        graphMethod.textContent = 'No graph payload is available for this selection.';
        await Plotly.newPlot(chartNode, [], { paper_bgcolor: 'rgba(255,255,255,0)' }, plotConfig);
        return;
      }

      graphTitle.textContent = payload.title || payload.label || 'Chart';
      graphMethod.textContent = payload.description || '';

      const fig = payload.figure || { data: [], layout: {} };
      fig.layout = fig.layout || {};
      fig.layout.font = { family: 'Manrope, sans-serif', size: 13, color: '#0f172a' };

      await Plotly.newPlot(chartNode, fig.data || [], fig.layout, plotConfig);
    }

    async function handleScenarioChange() {
      fillGraphSelect(scenarioSelect.value);
      await renderGraph();
    }

    async function handleGraphChange() {
      await renderGraph();
    }

    function handleDownload() {
      const scenarioKey = scenarioSelect.value;
      const graphKey = graphSelect.value;
      const scenario = dashboardData.scenarios[scenarioKey] || {};
      const graph = (scenario.graphs || {})[graphKey] || {};
      const fallback = [scenarioKey, graphKey].filter(Boolean).join('_') || 'chart';
      const filename = graph.downloadName || fallback;

      Plotly.downloadImage(chartNode, {
        format: 'png',
        filename,
        scale: 2,
        width: 1600,
        height: 900,
      });
    }

    async function init() {
      titleEl.textContent = dashboardData.projectTitle || 'Plant-Optimization-PathWay - Results Dashboard';
      generationDateEl.textContent = dashboardData.generationDate || 'N/A';

      const keys = scenarioKeys();
      if (!keys.length) {
        graphTitle.textContent = 'No scenarios found';
        graphMethod.textContent = 'No scenario workbook was discovered. Please generate results first.';
        await Plotly.newPlot(chartNode, [], { paper_bgcolor: 'rgba(255,255,255,0)' }, plotConfig);
        scenarioSelect.disabled = true;
        graphSelect.disabled = true;
        downloadBtn.disabled = true;
        return;
      }

      fillScenarioSelect();
      await handleScenarioChange();
    }

    scenarioSelect.addEventListener('change', handleScenarioChange);
    graphSelect.addEventListener('change', handleGraphChange);
    downloadBtn.addEventListener('click', handleDownload);
    window.addEventListener('resize', () => {
      if (chartNode && chartNode.data) {
        Plotly.Plots.resize(chartNode);
      }
    });

    init();
  </script>
</body>
</html>
"""

    return template.replace("__DASHBOARD_DATA__", payload_json)


def write_dashboard_html(output_path: Path, html: str) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html, encoding="utf-8")


def parse_args() -> argparse.Namespace:
    repo_root = get_repo_root()
    default_results_root = repo_root / "artifacts" / "reports"
    default_output_dir = default_results_root / "Results"

    parser = argparse.ArgumentParser(
        description="Generate a standalone HTML dashboard from PathWay scenario results."
    )
    parser.add_argument(
        "--results-root",
        type=Path,
        default=default_results_root,
        help="Directory containing scenario folders with Master_Plan.xlsx",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=default_output_dir,
        help="Directory where the standalone HTML file is written",
    )
    parser.add_argument(
        "--output-name",
        type=str,
        default="results_dashboard.html",
        help="Name of the generated HTML file",
    )
    parser.add_argument(
        "--discount-rate",
        type=float,
        default=DEFAULT_DISCOUNT_RATE,
        help="Discount rate for NPV computation (example: 0.08)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if args.discount_rate <= -1.0:
        raise ValueError("discount-rate must be greater than -1.0")

    results_root = args.results_root
    if not results_root.is_absolute():
        results_root = get_repo_root() / results_root

    output_dir = args.output_dir
    if not output_dir.is_absolute():
        output_dir = get_repo_root() / output_dir

    workbooks = discover_workbooks(results_root)
    if not workbooks:
        raise FileNotFoundError(
            f"No scenario workbooks found in {results_root}. Expected <scenario>/Master_Plan.xlsx."
        )

    payload = build_dashboard_data(workbooks, discount_rate=args.discount_rate)
    html = build_html(payload)

    output_path = output_dir / args.output_name
    write_dashboard_html(output_path, html)

    print(f"Dashboard generated: {output_path}")
    print(f"Scenarios loaded: {len(workbooks)}")
    print(
        "Charts available per scenario: CARBON PRICE, CARBON TAX, CO2 TRAJECTORY, ENERGY MIX, "
        "FINANCING, INDIRECT EMISSIONS, INVESTMENT PLAN, RESSOURCES OPEX, DATA USED, "
        "TRANSITION COST, TOTAL ANNUAL OPEX, CO2 ABATEMENT"
    )


if __name__ == "__main__":
    main()