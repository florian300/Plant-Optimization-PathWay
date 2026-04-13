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
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
import plotly.graph_objects as go
import sys
import os

# Ensure we can import from src/
repo_root = str(Path(__file__).resolve().parents[2])
if repo_root not in sys.path:
    sys.path.append(os.path.join(repo_root, "src"))

from pathway.core.plots.financial import build_transition_cost_figure
from pathway.core.plots.carbon import build_carbon_price_figure
from pathway.core.plots.carbon_tax import build_carbon_tax_figure
from pathway.core.plots.energy_mix import build_energy_mix_figure
from pathway.core.plots.investment import build_investment_plan_figure


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


def to_nullable_float_list(series: pd.Series, scale: float = 1.0) -> List[Optional[float]]:
    values = []
    for value in series:
        if pd.isna(value):
            values.append(None)
            continue
        try:
            fv = float(value) / scale
            values.append(round(fv, 6))
        except (TypeError, ValueError):
            values.append(None)
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


def fig_to_dict(fig: go.Figure) -> Dict[str, Any]:
    """Converts a Plotly Figure to the {data, layout} format expected by the dashboard."""
    d = fig.to_dict()
    return {
        "data": d.get("data", []),
        "layout": d.get("layout", {})
    }


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
                    "font": {"size": 15, "color": "#475569"},
                }
            ],
            "xaxis": {"visible": False},
            "yaxis": {"visible": False},
            "paper_bgcolor": "rgba(255,255,255,0)",
            "plot_bgcolor": "rgba(255,255,255,0)",
            "font": {"family": "Manrope, sans-serif", "color": "#1e293b"},
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
            "gridcolor": "#f1f5f9",
            "linecolor": "#e2e8f0",
            "tickfont": {"color": "#64748b"},
        },
        "yaxis": {
            "title": y_title,
            "automargin": True,
            "gridcolor": "#f1f5f9",
            "zerolinecolor": "#cbd5e1",
            "linecolor": "#e2e8f0",
            "tickfont": {"color": "#64748b"},
        },
        "barmode": barmode,
        "legend": {
            "orientation": "h",
            "yanchor": "bottom",
            "y": 1.02,
            "xanchor": "left",
            "x": 0.0,
            "font": {"size": 11, "color": "#475569"},
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

    # 1. Gather positive efforts
    col_oop = find_column(df_financing, ["out_of_pocket_capex", "out of pocket capex"])
    col_principal = find_column(df_financing, ["principal_repayment"])
    col_interest = find_column(df_financing, ["interest_paid"])
    
    capex_effort = pd.Series([0.0] * len(years))
    if col_oop: capex_effort += pd.to_numeric(df_financing[col_oop], errors="coerce").fillna(0.0)
    if col_principal: capex_effort += pd.to_numeric(df_financing[col_principal], errors="coerce").fillna(0.0)
    
    interest_effort = pd.Series([0.0] * len(years))
    if col_interest: interest_effort += pd.to_numeric(df_financing[col_interest], errors="coerce").fillna(0.0)

    opex_effort = pd.Series([0.0] * len(years))
    tax_effort = pd.Series([0.0] * len(years))
    aids_saving = pd.Series([0.0] * len(years))

    # 3. Build traces (matching CO2 TRAJECTORY area style)
    colors_efforts = ['#1a5276', '#5499c7', '#8e44ad', '#5dade2', '#aed6f1']
    colors_savings = ['#1e8449', '#58d68d', '#f39c12', '#2ecc71']
    traces = []
    
    # Positive efforts (Area Stack)
    if capex_effort.sum() > 0.01:
        traces.append({
            "type": "scatter", "mode": "lines", "name": "CAPEX & Repayment", 
            "x": years, "y": capex_effort.tolist(), 
            "stackgroup": "pos", "fill": "tonexty",
            "line": {"width": 0.5, "color": "white"},
            "fillcolor": colors_efforts[0],
            "marker": {"color": colors_efforts[0]}, "yaxis": "y"
        })
    if interest_effort.sum() > 0.01:
        traces.append({
            "type": "scatter", "mode": "lines", "name": "Loan Interests", 
            "x": years, "y": interest_effort.tolist(), 
            "stackgroup": "pos", "fill": "tonexty",
            "line": {"width": 0.5, "color": "white"},
            "fillcolor": colors_efforts[1],
            "marker": {"color": colors_efforts[1]}, "yaxis": "y"
        })
    if opex_effort.sum() > 0.01:
        traces.append({
            "type": "scatter", "mode": "lines", "name": "Operational Costs", 
            "x": years, "y": opex_effort.tolist(), 
            "stackgroup": "pos", "fill": "tonexty",
            "line": {"width": 0.5, "color": "white"},
            "fillcolor": colors_efforts[2],
            "marker": {"color": colors_efforts[2]}, "yaxis": "y"
        })
    if tax_effort.sum() > 0.01:
        traces.append({
            "type": "scatter", "mode": "lines", "name": "Carbon Tax (Actual)", 
            "x": years, "y": tax_effort.tolist(), 
            "stackgroup": "pos", "fill": "tonexty",
            "line": {"width": 0.5, "color": "white"},
            "fillcolor": colors_efforts[3],
            "marker": {"color": colors_efforts[3]}, "yaxis": "y"
        })
    
    # Negative savings (Area Stack)
    if aids_saving.sum() > 0.01:
        traces.append({
            "type": "scatter", "mode": "lines", "name": "Public Aids", 
            "x": years, "y": [-v for v in aids_saving.tolist()], 
            "stackgroup": "neg", "fill": "tonexty",
            "line": {"width": 0.5, "color": "white"},
            "fillcolor": colors_savings[0],
            "marker": {"color": colors_savings[0]}, "yaxis": "y"
        })

    # Cumulative NPV line
    traces.append({
        "type": "scatter", "mode": "lines+markers", "name": "Cumulative NPV / Balance",
        "x": years, "y": [round(v, 4) for v in cumulative_npv],
        "line": {"width": 4, "color": "#e74c3c"},
        "marker": {"size": 8, "color": "#e74c3c", "line": {"marker_color": "white", "width": 1.5}},
        "yaxis": "y2",
        "hovertemplate": "%{y:,.2f} MEUR<extra>%{fullData.name}</extra>"
    })

    layout = base_layout(title, "Annual Impact (MEUR)", years, barmode="relative")
    layout["template"] = "plotly_white"
    layout["yaxis"]["gridcolor"] = "#eeeeee"
    layout["xaxis"]["gridcolor"] = "#eeeeee"
    layout["yaxis2"] = {
        "title": {"text": "Cumulative NPV (MEUR)", "font": {"color": "#e74c3c"}},
        "overlaying": "y",
        "side": "right",
        "automargin": True,
        "tickfont": {"color": "#e74c3c"},
        "gridcolor": "rgba(0,0,0,0)",
    }
    layout["barmode"] = "relative"

    description = (
        "This chart displays the annual investment efforts (CAPEX, OPEX, Taxes) vs Savings (Aids). "
        "The red line tracks the Cumulative Net Present Value (NPV) of the transition strategy."
    )
    return {"data": traces, "layout": layout}, description
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
        return ({}, "No carbon price data available.")

    years = year_axis(df_co2["Year"])
    
    # Extract market prices
    market_prices = to_float_list(pd.to_numeric(df_co2["Tax_Price"], errors="coerce").fillna(0.0))
    
    # Extract penalties if available
    penalty_col = find_column(df_co2, ["Penalty_Factor", "carbon_penalty"])
    penalties = to_float_list(pd.to_numeric(df_co2[penalty_col], errors="coerce").fillna(0.0)) if penalty_col else None
    
    # Extract effective prices if available
    eff_col = find_column(df_co2, ["Effective_Price", "eff_tax_price"])
    effective_prices = to_float_list(pd.to_numeric(df_co2[eff_col], errors="coerce").fillna(0.0)) if eff_col else None
    
    # Extract strike prices if available
    strike_col = find_column(df_co2, ["Strike_Price", "strike_p"])
    raw_strikes = to_float_list(pd.to_numeric(df_co2[strike_col], errors="coerce").fillna(0.0)) if strike_col else []
    
    # Convert simplified strike prices back to the format expected by the module
    strike_prices = []
    if raw_strikes and any(s > 0 for s in raw_strikes):
        # We find contiguous blocks of the same non-zero strike price
        current_val = 0
        current_years = []
        for i, val in enumerate(raw_strikes):
            y = years[i]
            if val > 0:
                if abs(val - current_val) < 1e-4:
                    current_years.append(y)
                else:
                    if current_val > 0:
                        strike_prices.append({'name': 'CCfD Strike', 'val': current_val, 'years': current_years})
                    current_val = val
                    current_years = [y]
            else:
                if current_val > 0:
                    strike_prices.append({'name': 'CCfD Strike', 'val': current_val, 'years': current_years})
                current_val = 0
                current_years = []
        if current_val > 0:
            strike_prices.append({'name': 'CCfD Strike', 'val': current_val, 'years': current_years})

    fig = build_carbon_price_figure(
        years=years,
        market_prices=market_prices,
        effective_prices=effective_prices,
        penalties=penalties,
        strike_prices=strike_prices if strike_prices else None,
        title=title
    )
    
    description = (
        "Evolution of the market carbon price, combined with potential penalties "
        "and active CCfD strike prices over the simulation period."
    )
    
    return (fig.to_dict(), description)


def build_carbon_tax_graph(df_co2: pd.DataFrame) -> Tuple[Dict[str, Any], str]:
    title = "CARBON TAX"
    if df_co2.empty or "Year" not in df_co2.columns:
        return (
            placeholder_figure(title, "No CO2_Trajectory data is available for carbon tax."),
            "This graph requires Tax_Cost_MEuros columns in CO2_Trajectory.",
        )

    years = year_axis(df_co2["Year"])
    
    # 1. Primary Cost Columns (Differentiated if possible)
    std_tax_col = find_column(df_co2, ["Standard_Tax_Cost_MEuros", "std_tax"])
    pen_tax_col = find_column(df_co2, ["Penalty_Cost_MEuros", "penalty_tax"])
    gross_col = find_column(df_co2, ["Tax_Cost_MEuros", "gross_tax"])
    
    standard_tax = to_float_list(pd.to_numeric(df_co2[std_tax_col], errors="coerce").fillna(0.0)) if std_tax_col else []
    penalties = to_float_list(pd.to_numeric(df_co2[pen_tax_col], errors="coerce").fillna(0.0)) if pen_tax_col else []
    
    # Fallback if differentiated columns are missing (old results)
    if not standard_tax and gross_col:
        standard_tax = to_float_list(pd.to_numeric(df_co2[gross_col], errors="coerce").fillna(0.0))
        penalties = [0.0] * len(years)

    # 2. Avoided Costs
    avoided_red_col = find_column(df_co2, ["Really_Avoided_CO2_kt", "avoided_red"])
    avoided_cap_col = find_column(df_co2, ["Captured_CO2_kt", "avoided_cap"])
    tax_p_col = find_column(df_co2, ["Tax_Price", "carbon_p"])
    
    tax_prices = to_float_list(pd.to_numeric(df_co2[tax_p_col], errors="coerce").fillna(0.0)) if tax_p_col else [0.0]*len(years)
    
    avoided_reduced = []
    if avoided_red_col:
        red_kt = to_float_list(pd.to_numeric(df_co2[avoided_red_col], errors="coerce").fillna(0.0))
        avoided_reduced = [rt * 1000 * tp / 1_000_000.0 for rt, tp in zip(red_kt, tax_prices)]
    else:
        avoided_reduced = [0.0] * len(years)
        
    avoided_captured = []
    if avoided_cap_col:
        cap_kt = to_float_list(pd.to_numeric(df_co2[avoided_cap_col], errors="coerce").fillna(0.0))
        avoided_captured = [ct * 1000 * tp / 1_000_000.0 for ct, tp in zip(cap_kt, tax_prices)]
    else:
        avoided_captured = [0.0] * len(years)

    # 3. Refunds
    refund_col = find_column(df_co2, ["CCfD_Refund_MEuros", "ccfd_ref"])
    ccfd_refunds = to_float_list(pd.to_numeric(df_co2[refund_col], errors="coerce").fillna(0.0)) if refund_col else None

    # Build Figure using shared module
    fig = build_carbon_tax_figure(
        years=years,
        standard_tax=standard_tax,
        penalties=penalties,
        avoided_reduced=avoided_reduced,
        avoided_captured=avoided_captured,
        ccfd_refunds=ccfd_refunds,
        title=title
    )
    
    # Adapt to dashboard layout
    fig.update_layout(
        font=dict(family="Manrope, sans-serif", size=13, color="#1e293b"),
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="rgba(255,255,255,0)",
    )
    
    description = (
        "This chart differentiates between standard carbon tax and performance penalties. "
        "Savings from emission reduction and capture are shown as negative values."
    )
    return fig_to_dict(fig), description


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
    
    dac_kt = to_float_list(pd.to_numeric(df_co2.get("DAC_Captured_kt", 0.0), errors="coerce").fillna(0.0), scale=1.0)
    credits_kt = to_float_list(pd.to_numeric(df_co2.get("Credits_Purchased_kt", 0.0), errors="coerce").fillna(0.0), scale=1.0)

    # Net Direct Balance calculation: Direct - DAC - Credits
    net_direct_bal = [round(d - dc - c, 6) for d, dc, c in zip(direct_kt, dac_kt, credits_kt)]

    traces = [
        # --- 1. AIRES DE RÉFÉRENCE (Remplissages continus) ---
        
        # Ombre sous Net Direct
        {
            "type": "scatter",
            "mode": "none",
            "name": "Net Direct Shade",
            "x": years,
            "y": net_direct_bal,
            "fill": "tozeroy",
            "fillcolor": "rgba(52, 152, 219, 0.1)",
            "showlegend": False,
        },
        # DAC et Crédits (Aires Négatives)
        {
            "type": "scatter",
            "mode": "lines",
            "name": "DAC Captured (ktCO2)",
            "x": years,
            "y": [round(-abs(x), 6) for x in dac_kt],
            "fill": "tozeroy",
            "fillcolor": "rgba(52, 152, 219, 0.6)",
            "line": {"width": 0},
            "hovertemplate": "%{y:,.2f} ktCO2<extra>%{fullData.name}</extra>",
        },
        {
            "type": "scatter",
            "mode": "lines",
            "name": "Voluntary Credits (ktCO2)",
            "x": years,
            "y": [round(-(abs(d) + abs(c)), 6) for d, c in zip(dac_kt, credits_kt)],
            "fill": "tonexty",
            "fillcolor": "rgba(39, 174, 96, 0.6)",
            "line": {"width": 0},
            "hovertemplate": "%{y:,.2f} ktCO2<extra>%{fullData.name}</extra>",
        },
        # Quotas Gratuits (0 vers Free_Quota)
        {
            "type": "scatter",
            "mode": "lines",
            "name": "Free Quotas (Direct)",
            "x": years,
            "y": free_quota_kt,
            "fill": "tozeroy",
            "fillcolor": "rgba(0, 128, 0, 0.3)",
            "line": {"width": 0},
            "hovertemplate": "%{y:,.2f} ktCO2<extra>%{fullData.name}</extra>",
        },
        # Émissions Taxées (Free_Quota vers Total) avec motifs
        {
            "type": "scatter",
            "mode": "lines",
            "name": "Taxed Emissions (Surface)",
            "x": years,
            "y": [round(f + t, 6) for f, t in zip(free_quota_kt, taxed_kt)],
            "fill": "tonexty",
            "fillcolor": "rgba(128, 128, 128, 0.4)",
            "fillpattern": {"shape": ".", "solidity": 0.3},
            "line": {"width": 0},
            "hovertemplate": "%{y:,.2f} ktCO2<extra>%{fullData.name}</extra>",
        },

        # --- 2. COURBES DE TRAJECTOIRE ---
        {
            "type": "scatter",
            "mode": "lines+markers",
            "name": "Direct Emissions",
            "x": years,
            "y": direct_kt,
            "line": {"width": 3, "color": "#111827"},
            "marker": {"size": 5, "color": "#111827"},
            "hovertemplate": "%{y:,.2f} ktCO2<extra>%{fullData.name}</extra>",
        },
        {
            "type": "scatter",
            "mode": "lines+markers",
            "name": "Indirect Emissions",
            "x": years,
            "y": indirect_kt,
            "line": {"width": 2, "color": "#111827", "dash": "dot"},
            "marker": {"size": 4, "color": "#111827"},
            "hovertemplate": "%{y:,.2f} ktCO2<extra>%{fullData.name}</extra>",
        },
        {
            "type": "scatter",
            "mode": "lines+markers",
            "name": "Total CO2 (Net)",
            "x": years,
            "y": [round(n + i, 6) for n, i in zip(net_direct_bal, indirect_kt)],
            "line": {"width": 3, "color": "darkred", "dash": "dash"},
            "marker": {"size": 5, "color": "darkred"},
            "hovertemplate": "%{y:,.2f} ktCO2<extra>%{fullData.name}</extra>",
        },
        {
            "type": "scatter",
            "mode": "lines+markers",
            "name": "Net Direct Emissions",
            "x": years,
            "y": net_direct_bal,
            "line": {"width": 3, "color": "#3498db", "dash": "dashdot"},
            "marker": {"size": 6, "color": "#3498db"},
            "hovertemplate": "%{y:,.2f} ktCO2<extra>%{fullData.name}</extra>",
        },
    ]
    layout = base_layout(title, "ktCO2 (Carbon Balance)", years, barmode="relative")
    layout["yaxis"]["zeroline"] = True
    layout["yaxis"]["zerolinewidth"] = 2
    layout["yaxis"]["zerolinecolor"] = "#333"
    description = (
        "The trajectory combines direct, indirect, and total emissions with the free-quota and taxed-emissions "
        "decomposition from CO2_Trajectory. Emission values are displayed in ktCO2 for readability."
    )
    return {"data": traces, "layout": layout}, description


def build_energy_mix_full_graph(df_energy: pd.DataFrame, df_metadata: pd.DataFrame = None) -> Tuple[Dict[str, Any], str]:
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

    # 1. Map columns to categories
    cat_map = {}
    metadata_map = {}
    if df_metadata is not None and not df_metadata.empty:
        for _, row in df_metadata.iterrows():
            metadata_map[str(row["ID"])] = {
                "Type": str(row.get("Type", "Unspecified")).upper(),
                "Category": str(row.get("Category", "Uncategorized")),
                "Unit": str(row.get("Unit", "Units")),
                "Name": str(row.get("Name", row["ID"]))
            }

    for c in cols:
        meta = metadata_map.get(c, {"Type": "Unspecified", "Category": "Uncategorized", "Unit": "Units", "Name": c})
        full_cat = f"[{meta['Type']}] {meta['Category']}"
        unit = meta["Unit"]
        name = meta["Name"]
        
        if full_cat not in cat_map:
            cat_map[full_cat] = {"unit": unit, "series": {}}
        
        cat_map[full_cat]["series"][name] = to_float_list(pd.to_numeric(df_energy[c], errors="coerce").fillna(0.0))

    # Sort map
    sorted_keys = sorted(cat_map.keys())
    cat_map = {k: cat_map[k] for k in sorted_keys}

    # 2. Build figure using the same shared component
    fig = build_energy_mix_figure(
        years=years,
        category_data=cat_map,
        title=title
    )
    
    # 3. Adapt to dashboard layout
    fig.update_layout(
        font=dict(family="Manrope, sans-serif", size=13, color="#1e293b"),
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="rgba(255,255,255,0)",
    )

    description = (
        "Energy_Mix is presented here as a categorical breakdown of all energy flows. "
        "Use the dropdown menu to select specific categories (defined in OverView)."
    )
    return fig_to_dict(fig), description


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


def build_investment_plan_graph(df_investments: pd.DataFrame, df_costs: pd.DataFrame = None, df_hf_invest: pd.DataFrame = None) -> Tuple[Dict[str, Any], str]:
    title = "INVESTMENT PLAN"
    
    # CASE 1: High-fidelity data available from Excel export
    if df_hf_invest is not None and not df_hf_invest.empty and "Year" in df_hf_invest.columns:
        years = year_axis(df_hf_invest["Year"])
        
        # Exact matching based on existing Plotly module logic
        fig = build_investment_plan_figure(
            df_projects=df_hf_invest,
            df_costs=df_costs,
            years=years,
            is_dark_bg=False,
            title="INVESTMENT PLAN: Implementation Costs (M€)"
        )
        description = "This graph uses high-fidelity CAPEX data exported directly from the optimizer for perfect consistency with the PDF report."
        return fig_to_dict(fig), description

    # CASE 2: Fallback to basic Investments sheet aggregation
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
        "#1D4ED8", "#2563EB", "#0EA5E9", "#10B981", "#22C55E", "#4F46E5", "#EA580C", "#475569",
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
        "Reconstructed from raw Investments sheet. aggregation by Year and Technology (sum of Capex_Euros). "
        "For best results, re-run the simulation to export high-fidelity investment data."
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
        res_price = to_nullable_float_list(pd.to_numeric(df_res["Price"], errors="coerce"))
        res_co2 = to_nullable_float_list(pd.to_numeric(df_res["CO2_Emissions"], errors="coerce"))

        is_first = (i == 0)

        # Trace 1: Price (left axis)
        traces.append({
            "type": "scatter",
            "mode": "lines+markers",
            "name": f"Price",
            "x": res_years,
            "y": res_price,
            "connectgaps": True,
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
            "connectgaps": True,
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
    }

    return {"data": traces, "layout": layout}, description


def _calculate_resource_costs(df_energy: pd.DataFrame, df_prices: pd.DataFrame) -> List[float]:
    """Calculates annual resource costs in M€ from consumption and price data."""
    if df_energy.empty or df_prices.empty or "Year" not in df_energy.columns:
        return []

    years = year_axis(df_energy["Year"])
    p_map = {}
    for _, row in df_prices.iterrows():
        try:
            r_id = str(row.get("Resource", "")).strip()
            y_id = int(float(row.get("Year", 0)))
            p_val = float(row.get("Price", 0.0))
            if r_id and y_id > 0:
                p_map[(r_id, y_id)] = p_val
        except (TypeError, ValueError):
            continue

    costs = []
    for i, yr in enumerate(years):
        s_cost = 0.0
        yr_int = int(yr) if yr is not None else 0
        for col in [c for c in df_energy.columns if c != "Year"]:
            con = clean_numeric(df_energy.iloc[i].get(col, 0.0))
            pri = p_map.get((col, yr_int), 0.0)
            if pri == 0:  # Fuzzy match
                short_id = col.replace("EN_", "")
                for (pr_res, pr_yr), pr_val in p_map.items():
                    if pr_yr == yr_int and (short_id in pr_res or pr_res in col):
                        pri = pr_val
                        break
            if con != 0 and pri != 0:
                s_cost += (con * pri) / 1_000_000.0
        costs.append(round(s_cost, 6))
    return costs


def build_transition_cost_graph(
    df_financing: pd.DataFrame, 
    df_costs: pd.DataFrame, 
    df_co2: pd.DataFrame, 
    df_transition_balance: pd.DataFrame, 
    df_energy: pd.DataFrame = None, 
    df_data_used: pd.DataFrame = None, 
    bau_res_costs: List[float] = None, 
    bau_tax_costs: Dict[Any, float] = None,
    df_hf_transition: pd.DataFrame = None
) -> Tuple[Dict[str, Any], str]:
    title = "TRANSITION COST"
    
    # CASE 1: High-Fidelity Data available in Excel
    if df_hf_transition is not None and not df_hf_transition.empty and "Year" in df_hf_transition.columns:
        years = year_axis(df_hf_transition["Year"])
        all_cols = [c for c in df_hf_transition.columns if c != "Year" and not str(c).startswith("Unnamed:")]
        
        # Exact matching based on prefixes Effort: and Saving:
        pos_cols = [c for c in all_cols if str(c).startswith("Effort:")]
        neg_cols = [c for c in all_cols if str(c).startswith("Saving:")]

        # Ensure all used columns are numeric (fix for TypeError: unsupported operand +: int and str)
        for c in (pos_cols + neg_cols):
             df_hf_transition[c] = pd.to_numeric(df_hf_transition[c], errors="coerce").fillna(0.0)
        
        # High-fidelity data is used exactly as exported to maintain full parity with the static PNG report.

        fig = build_transition_cost_figure(
            df_annual=df_hf_transition,
            years=years,
            pos_cols=pos_cols,
            neg_cols=neg_cols,
            title="ECOLOGICAL TRANSITION: ANNUAL EFFORTS & SAVINGS"
        )
        description = "This graph uses high-fidelity data exported directly from the optimizer for perfect consistency."
        return fig_to_dict(fig), description

    # CASE 2: Fallback manual calculation (if high-fidelity data missing)
    if df_financing.empty or "Year" not in df_financing.columns:
        return (placeholder_figure(title, "No data."), "No financing data.")

    years = year_axis(df_financing["Year"])
    n_yrs = len(years)

    # efforts
    col_oop = find_column(df_financing, ["out_of_pocket_capex", "out of pocket capex"])
    col_principal = find_column(df_financing, ["principal_repayment"])
    col_interest = find_column(df_financing, ["interest_paid"])
    
    capex_effort = pd.Series([0.0] * n_yrs)
    if col_oop: capex_effort += pd.to_numeric(df_financing[col_oop], errors="coerce").fillna(0.0)
    if col_principal: capex_effort += pd.to_numeric(df_financing[col_principal], errors="coerce").fillna(0.0)
    
    interest_effort = pd.Series([0.0] * n_yrs)
    if col_interest: interest_effort += pd.to_numeric(df_financing[col_interest], errors="coerce").fillna(0.0)

    opex_effort = pd.Series([0.0] * n_yrs)
    credits_effort = pd.Series([0.0] * n_yrs)
    if not df_costs.empty and "Year" in df_costs.columns:
        opex_col = find_column(df_costs, ["Total_OPEX", "opex_cost_meuros"])
        if opex_col:
            opex_map = _series_map_by_year(df_costs, opex_col, scale=1_000_000.0)
            opex_effort = pd.Series([opex_map.get(y, 0.0) for y in years])
        
        cred_col = find_column(df_costs, ["Credit_Cost", "carbon_credit_cost"])
        if cred_col:
            cred_map = _series_map_by_year(df_costs, cred_col, scale=1_000_000.0)
            credits_effort = pd.Series([cred_map.get(y, 0.0) for y in years])

    tax_actual = pd.Series([0.0] * n_yrs)
    # Use GROSS tax for efforts (refunds are handled in savings)
    tax_col = find_column(df_co2, ["tax_cost_meuros"]) 
    if tax_col:
        tax_map = _series_map_by_year(df_co2, tax_col)
        tax_actual = pd.Series([tax_map.get(y, 0.0) for y in years])

    # savings
    aids_saving = pd.Series([0.0] * n_yrs)
    # 1. Direct grants from Tech Costs
    aid_cols = [col for col in df_costs.columns if str(col).startswith("Aid_")]
    if aid_cols:
        aids_df = pd.DataFrame({c: pd.to_numeric(df_costs[c], errors="coerce").fillna(0.0) for c in aid_cols})
        aids_val = aids_df.sum(axis=1)
        aids_map = dict(zip(year_axis(df_costs["Year"]), to_float_list(aids_val, scale=1_000_000.0)))
        aids_saving += pd.Series([aids_map.get(y, 0.0) for y in years])
    
    # 2. CCfD refunds from CO2 Trajectory
    refund_col = find_column(df_co2, ["ccfd_refund_meuros"])
    if refund_col:
        refund_map = _series_map_by_year(df_co2, refund_col)
        aids_saving += pd.Series([refund_map.get(y, 0.0) for y in years])

    tax_offset = pd.Series([0.0] * n_yrs)
    if bau_tax_costs:
        tax_offset = pd.Series([-(bau_tax_costs.get(y, 0.0)) for y in years])

    res_add = pd.Series([0.0] * n_yrs)
    res_avoid = pd.Series([0.0] * n_yrs)
    if df_energy is not None and df_data_used is not None and bau_res_costs is not None:
        actual_res = _calculate_resource_costs(df_energy, df_data_used)
        for i, a_cost in enumerate(actual_res):
            if i < n_yrs:
                b_cost = bau_res_costs[i] if i < len(bau_res_costs) else bau_res_costs[-1]
                delta = a_cost - b_cost
                if delta > 1e-4: res_add[i] = delta
                elif delta < -1e-4: res_avoid[i] = delta

    # Build annual DF with Effort/Saving prefixes
    df_annual = pd.DataFrame({
        "Year": years,
        "Effort: Self-funded CAPEX": capex_effort.tolist(),
        "Effort: Bank Loan Service": interest_effort.tolist(),
        "Effort: Tech & DAC OPEX": opex_effort.tolist(),
        "Effort: Voluntary Credits": credits_effort.tolist(),
        "Effort: Additional Resource Cost": res_add.tolist(),
        "Effort: Carbon Tax (Actual)": tax_actual.tolist(),
        "Saving: Public Aids": (-aids_saving).tolist(),
        "Saving: Baseline Carbon Tax Offset": tax_offset.tolist(),
        "Saving: Avoided Resource Saving": res_avoid.tolist()
    })
    
    all_cols = [c for c in df_annual.columns if c != "Year"]
    pos_cols = [c for c in all_cols if c.startswith("Effort: ")]
    neg_cols = [c for c in all_cols if c.startswith("Saving: ")]
    
    fig = build_transition_cost_figure(
        df_annual=df_annual,
        years=years,
        pos_cols=pos_cols,
        neg_cols=neg_cols,
        title=title
    )

    description = (
        "Reconstructed Transition Cost chart from raw Excel sheets. "
        "For best results, re-run the simulation to export high-fidelity transition data."
    )
    return fig_to_dict(fig), description


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
            "y": [round(v if v > 0 else -abs(v), 6) for v in credits],
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

    layout = base_layout(title, "MEUR", years, barmode="relative")
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


def load_high_fidelity_table(workbook: Path, sheet_marker: str) -> pd.DataFrame:
    """Loads a specific high-fidelity table from the Charts sheet."""
    try:
        raw = pd.read_excel(workbook, sheet_name="Charts", header=None)
    except Exception:
        return pd.DataFrame()

    header_row = None
    for idx, row in raw.iterrows():
        values = [str(v).strip() for v in row.tolist() if pd.notna(v)]
        if any(sheet_marker in v for v in values):
            header_row = idx
            break

    if header_row is None:
        return pd.DataFrame()

    df = pd.read_excel(workbook, sheet_name="Charts", header=header_row + 1)
    
    # 1. Stop at the first entirely empty row to avoid reading subsequent tables
    # Find the index of the first row that is all NaN
    nan_rows = df.isnull().all(axis=1)
    if nan_rows.any():
        first_nan = nan_rows.idxmax()
        df = df.iloc[:first_nan]
    
    df = df.dropna(how="all")
    
    # Filter out other tables below if they exist (stop at next empty row or specific marker)
    # For now, just drop unnamed columns
    unnamed = [col for col in df.columns if str(col).startswith("Unnamed:")]
    if unnamed:
        df = df.drop(columns=unnamed)
    
    if "Year" not in df.columns:
        # Try to find Year column by fuzzy match
        for col in df.columns:
            if "year" in str(col).lower():
                df = df.rename(columns={col: "Year"})
                break
    
    return df.reset_index(drop=True)


def build_dashboard_data(workbooks: Dict[str, Path], discount_rate: float) -> Dict[str, Any]:
    latest_ts = max(path.stat().st_mtime for path in workbooks.values())
    generation_date = datetime.fromtimestamp(latest_ts).strftime("%Y-%m-%d %H:%M:%S")

    scenarios: Dict[str, Any] = {}
    
    # 1. First pass to find BAU costs for resource delta calculation
    bau_name = next((n for n in workbooks if any(b in n.upper() for b in ["BUSINESS AS USUAL", "BASELINE", "BAU"])), None)
    bau_res_costs = None
    bau_tax_costs = None
    if bau_name:
        wb = workbooks[bau_name]
        df_en_bau = load_sheet(wb, "Energy_Mix")
        df_du_bau = load_sheet(wb, "Data_Used")
        df_co2_bau = load_sheet(wb, "CO2_Trajectory")
        if not df_en_bau.empty and not df_du_bau.empty:
            bau_res_costs = _calculate_resource_costs(df_en_bau, df_du_bau)
        if not df_co2_bau.empty:
            tax_col = find_column(df_co2_bau, ["tax_cost_meuros", "net_tax_cost_meuros"])
            if tax_col:
                bau_tax_costs = _series_map_by_year(df_co2_bau, tax_col)

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
        df_hf_transition = load_high_fidelity_table(workbook, "TRANSITION_COST_HIGH_FIDELITY")
        df_hf_invest = load_high_fidelity_table(workbook, "INVESTMENT_PLAN_HIGH_FIDELITY")
        df_metadata = load_sheet(workbook, "Resource_Metadata")

        carbon_price_fig, carbon_price_desc = build_carbon_price_graph(df_co2)
        carbon_tax_fig, carbon_tax_desc = build_carbon_tax_graph(df_co2)
        co2_traj_fig, co2_traj_desc = build_co2_trajectory_full_graph(df_co2)
        energy_mix_fig, energy_mix_desc = build_energy_mix_full_graph(df_energy, df_metadata)
        ext_finance_fig, ext_finance_desc = build_external_financing_graph(df_financing, df_costs)
        indirect_fig, indirect_desc = build_indirect_emissions_graph(df_indir)
        invest_fig, invest_desc = build_investment_plan_graph(df_invest, df_costs, df_hf_invest)
        resources_opex_fig, resources_opex_desc = build_resources_opex_graph(df_costs)
        data_used_fig, data_used_desc = build_data_used_graph(df_data_used)
        transition_fig, transition_desc = build_transition_cost_graph(
            df_financing, df_costs, df_co2, df_transition_balance, 
            df_energy, df_data_used, bau_res_costs, bau_tax_costs,
            df_hf_transition=df_hf_transition
        )
        total_opex_fig, total_opex_desc = build_total_annual_opex_graph(df_costs, df_co2)
        co2_abatement_fig, co2_abatement_desc = build_co2_abatement_graph(df_mac)

        # --- 2. Chart Mapping Logic (Hybrid: JSON if available, else embedded) ---
        def get_graph_payload(key, label, title, description, figure, json_filename):
            """Returns a figure object, either from a modular JSON file if present, or fallback to the provided figure."""
            chart_json = workbook.parent / "charts" / f"{json_filename}.json"
            if chart_json.exists():
                try:
                    with open(chart_json, 'r', encoding='utf-8') as f:
                        modular_fig = json.load(f)
                    return {
                        "label": label,
                        "title": title,
                        "description": description,
                        "downloadName": sanitize_filename(f"{scenario_name}_{key}"),
                        "figure": modular_fig
                    }
                except Exception as e:
                    print(f"  [WARN] Failed to read modular JSON for {key}: {e}")
            
            # Fallback to the figure generated by the dashboard script
            return {
                "label": label,
                "title": title,
                "description": description,
                "downloadName": sanitize_filename(f"{scenario_name}_{key}"),
                "figure": figure
            }

        scenarios[scenario_name] = {
            "displayName": scenario_name,
            "sourceWorkbook": str(workbook),
            "graphs": {
                "carbon_price": get_graph_payload("carbon_price", "CARBON PRICE", "CARBON PRICE", carbon_price_desc, carbon_price_fig, "Carbon_Prices"),
                "carbon_tax": get_graph_payload("carbon_tax", "CARBON TAX", "CARBON TAX", carbon_tax_desc, carbon_tax_fig, "Carbon_Tax"),
                "co2_trajectory": get_graph_payload("co2_trajectory", "CO2 TRAJECTORY", "CO2 TRAJECTORY", co2_traj_desc, co2_traj_fig, "CO2_Trajectory"),
                "energy_mix": get_graph_payload("energy_mix", "ENERGY MIX", "ENERGY MIX", energy_mix_desc, energy_mix_fig, "Energy_Mix"),
                "external_financing": get_graph_payload("external_financing", "FINANCING", "FINANCING", ext_finance_desc, ext_finance_fig, "External_Financing"),
                "indirect_emissions": get_graph_payload("indirect_emissions", "INDIRECT EMISSIONS", "INDIRECT EMISSIONS", indirect_desc, indirect_fig, "Indirect_Emissions"),
                "investment_plan": get_graph_payload("investment_plan", "INVESTMENT PLAN", "INVESTMENT PLAN", invest_desc, invest_fig, "Investment_Plan"),
                "ressources_opex": get_graph_payload("ressources_opex", "RESSOURCES OPEX", "RESSOURCES OPEX", resources_opex_desc, resources_opex_fig, "Resources_Opex"),
                "data_used": get_graph_payload("data_used", "DATA USED", "DATA USED", data_used_desc, data_used_fig, "Data_Used"),
                "transition_cost": get_graph_payload("transition_cost", "TRANSITION COST", "TRANSITION COST", transition_desc, transition_fig, "Transition_Cost"),
                "total_annual_opex": get_graph_payload("total_annual_opex", "TOTAL ANNUAL OPEX", "TOTAL ANNUAL OPEX", total_opex_desc, total_opex_fig, "Total_Annual_Opex"),
                "co2_abatement": get_graph_payload("co2_abatement", "CO2 ABATEMENT", "CO2 ABATEMENT", co2_abatement_desc, co2_abatement_fig, "CO2_Abatement"),
            },
        }

    payload = {
        "projectTitle": "Plant-Optimization-PathWay - Results Dashboard",
        "generationDate": generation_date,
        "discountRate": discount_rate,
        "scenarios": scenarios,
    }
    return sanitize_payload(payload)


def load_sensitivity_data(json_path: Optional[Path] = None) -> Optional[List[Dict[str, Any]]]:
    """
    Charge les résultats JSON de l'analyse de sensibilité si le fichier existe.
    Retourne None si le fichier est absent ou illisible.
    """
    if json_path is None:
        json_path = get_repo_root() / "artifacts" / "sensitivity" / "sensitivity_results.json"
    if not json_path.exists():
        return None
    try:
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
        print(f"[Sensitivity] Données chargées : {len(data)} enregistrements depuis {json_path}")
        return data
    except Exception as exc:
        print(f"[Sensitivity] Impossible de lire {json_path}: {exc}")
        return None


def build_html(payload: Dict[str, Any], sensitivity_data: Optional[List[Dict[str, Any]]] = None) -> str:
    payload_json = json.dumps(payload, ensure_ascii=True)
    sensitivity_json = json.dumps(sensitivity_data if sensitivity_data else [], ensure_ascii=True)

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
            heading: ['"Montserrat"', 'sans-serif'],
            body: ['"Bookman Old Style"', '"Bookman"', 'serif']
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
  <link href=\"https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&display=swap\" rel=\"stylesheet\" />
  <script src=\"https://cdn.plot.ly/plotly-2.35.2.min.js\"></script>
  <link rel=\"stylesheet\" href=\"https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css\" integrity=\"sha512-SnH5WK+bZxgPHs44uWix+LLJAJ9/2PkPKZ5QiAj6Ta86w+fsb2TkR4j8f5Z5gDQL4x0XSLwWf2fQJKfG8d8gQw==\" crossorigin=\"anonymous\" referrerpolicy=\"no-referrer\" />
  <style>
    :root {
      --bg-1: #f8fafc;
      --bg-2: #f1f5f9;
      --bg-3: #e2e8f0;
      --glass: rgba(255, 255, 255, 0.82);
      --glass-strong: rgba(255, 255, 255, 0.95);
      --text-main: #1e293b;
      --text-muted: #64748b;
      --border-soft: rgba(15, 23, 42, 0.08);
      --accent: #0ea5e9;
    }

    html, body {
      min-height: 100%;
    }

    body {
      margin: 0;
      font-family: 'Bookman Old Style', 'Bookman', serif;
      color: var(--text-main);
      background:
        radial-gradient(circle at 10% 20%, rgba(14, 165, 233, 0.06) 0%, rgba(14, 165, 233, 0) 36%),
        radial-gradient(circle at 85% 12%, rgba(16, 185, 129, 0.08) 0%, rgba(16, 185, 129, 0) 34%),
        linear-gradient(135deg, #ffffff 0%, #f8fafc 50%, #f1f5f9 100%);
      background-attachment: fixed;
    }

    .glass-card {
      background: linear-gradient(155deg, var(--glass) 0%, var(--glass-strong) 100%);
      backdrop-filter: blur(20px);
      -webkit-backdrop-filter: blur(20px);
      border: 1px solid var(--border-soft);
      box-shadow: 0 10px 30px -5px rgba(15, 23, 42, 0.04), 0 4px 12px -4px rgba(15, 23, 42, 0.03);
      transition: transform 260ms ease, box-shadow 260ms ease;
    }

    .glass-card:hover {
      transform: translateY(-2px);
      box-shadow: 0 20px 40px -8px rgba(15, 23, 42, 0.08);
    }

    .subtle-pill {
      background: rgba(15, 23, 42, 0.05);
      color: #475569;
      border: 1px solid rgba(15, 23, 42, 0.08);
    }

    .selector {
      width: 100%;
      border: 1px solid rgba(15, 23, 42, 0.12);
      border-radius: 0.9rem;
      padding: 0.75rem 0.9rem;
      background: #ffffff;
      color: #1e293b;
      font-size: 0.95rem;
      outline: none;
      transition: all 200ms ease;
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

    /* Tab styles */
    .nav-pill-container {
      background: #f1f5f9;
      border-radius: 9999px;
      padding: 0.4rem;
      display: inline-flex;
      gap: 0.5rem;
      border: 1px solid rgba(15, 23, 42, 0.05);
    }

    .nav-tab {
      padding: 0.6rem 1.5rem;
      border-radius: 9999px;
      background: #ffffff;
      color: #1e293b;
      font-weight: 600;
      font-size: 0.95rem;
      transition: all 200ms ease;
      cursor: pointer;
      border: 1px solid transparent;
      white-space: nowrap;
    }

    .nav-tab:hover:not(.active) {
      background: #f8fafc;
      transform: translateY(-1px);
    }

    /* Active Tab States */
    .nav-tab.active.tab-home { background: #bfdbfe; color: #1e3a8a; }
    .nav-tab.active.tab-details { background: #e2e8f0; color: #334155; }
    .nav-tab.active.tab-results { background: #bef264; color: #365314; }
    .nav-tab.active.tab-sensitivity { background: #fed7aa; color: #7c2d12; }
    .nav-tab.active.tab-licence { background: #fecaca; color: #7f1d1d; }

    .tab-content {
      display: none;
    }

    .tab-content.active {
      display: block;
    }
  </style>
</head>
<body>
  <header class=\"max-w-7xl mx-auto px-4 md:px-8 pt-8 md:pt-12 text-center\">
    <div class=\"nav-pill-container shadow-sm\">
      <div class=\"nav-tab tab-home\" onclick=\"switchTab('home', this)\">Home</div>
      <div class=\"nav-tab tab-details\" onclick=\"switchTab('details', this)\">Simulation details</div>
      <div class=\"nav-tab tab-results active\" onclick=\"switchTab('results', this)\">Results</div>
      <div class=\"nav-tab tab-sensitivity\" onclick=\"switchTab('sensitivity', this)\">Sensitivity analysis</div>
      <div class=\"nav-tab tab-licence\" onclick=\"switchTab('licence', this)\">Licence</div>
    </div>
  </header>

  <!-- Tab Contents -->
  <div id=\"home-tab\" class=\"tab-content max-w-7xl mx-auto px-4 md:px-8 py-10\">
    <section class=\"glass-card rounded-3xl p-12 text-center\">
      <h2 class=\"text-3xl font-heading font-bold mb-4\">Home</h2>
      <p class=\"text-slate-500\">This is the home tab placeholder.</p>
    </section>
  </div>

  <div id=\"details-tab\" class=\"tab-content max-w-7xl mx-auto px-4 md:px-8 py-10\">
    <section class=\"glass-card rounded-3xl p-12 text-center\">
      <h2 class=\"text-3xl font-heading font-bold mb-4\">Simulation Details</h2>
      <p class=\"text-slate-500\">Detailed simulation information will appear here.</p>
    </section>
  </div>

  <div id=\"results-tab\" class=\"tab-content active\">
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
            <span class=\"text-sm font-bold text-slate-500 uppercase tracking-tight\">Scenario Selector</span>
            <select id=\"scenarioSelect\" class=\"selector mt-2\"></select>
          </label>

          <label class=\"block\">
            <span class=\"text-sm font-bold text-slate-500 uppercase tracking-tight\">Graph Selector</span>
            <select id=\"graphSelect\" class=\"selector mt-2\"></select>
          </label>
        </div>
      </section>

      <section class=\"glass-card rounded-3xl p-4 md:p-6 mb-6 fade-in\">
        <div class=\"flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 mb-3\">
          <h2 id=\"graphTitle\" class=\"font-heading text-lg md:text-2xl font-bold text-slate-800\"></h2>
          <button id=\"downloadBtn\" class=\"primary-btn inline-flex items-center justify-center gap-2\">
            <i class=\"fa-solid fa-download\"></i>
            Download Chart as Image
          </button>
        </div>
        <div id=\"chart\"></div>
      </section>

      <section class=\"glass-card rounded-3xl p-5 md:p-6 fade-in\">
        <h3 class=\"font-heading text-lg md:text-xl font-bold mb-3 text-slate-800\">How this graph is constructed</h3>
        <p id=\"graphMethod\" class=\"text-slate-600 leading-relaxed font-medium\"></p>
      </section>
    </main>
  </div>

  <!-- ═══════════════════ ONGLET ANALYSE DE SENSIBILITÉ ═══════════════════ -->
  <div id=\"sensitivity-tab\" class=\"tab-content max-w-full mx-auto px-4 md:px-12 py-10\">

    <!-- En-tête -->
    <section class=\"glass-card rounded-3xl p-6 md:p-8 mb-6 fade-in\">
      <div class=\"flex flex-col md:flex-row md:items-end md:justify-between gap-4\">
        <div>
          <p class=\"subtle-pill inline-flex items-center gap-2 rounded-full px-3 py-1 text-xs font-semibold uppercase tracking-wide\">
            <i class=\"fa-solid fa-chart-line\"></i>
            Analyse de Risque — One-At-a-Time (OAT)
          </p>
          <h2 class=\"font-heading text-2xl md:text-3xl font-bold mt-3\">Analyse de Sensibilité au Prix EUA</h2>
          <p class=\"text-slate-500 mt-2 text-sm\">
            Variation symétrique du prix du carbone (EUA) appliquée au scénario de référence (BS).
            Chaque point représente une simulation MILP indépendante.
          </p>
        </div>
        <div id=\"sens-status-badge\" class=\"rounded-2xl subtle-pill px-4 py-3 text-sm\"></div>
      </div>
    </section>

    <!-- Grille des 4 graphiques -->
    <div class=\"grid grid-cols-1 gap-6\">

      <!-- 1. Vue Globale des Risques (Packed Bubble) -->
      <section class=\"glass-card rounded-3xl p-5 md:p-6 fade-in\">
        <h3 class=\"font-heading text-lg font-bold text-slate-800 mb-1\">Vue Globale des Risques</h3>
        <p class=\"text-xs text-slate-500 mb-3\">
          Chaque bulle représente un paramètre perturbé. Le rayon est proportionnel à la variance maximale du coût de transition.
        </p>
        <div id=\"sens-bubble-chart\" style=\"height:500px;\"></div>
      </section>

      <!-- 2. Tornado Chart (Impact Financier) -->
      <section class=\"glass-card rounded-3xl p-5 md:p-6 fade-in\">
        <h3 class=\"font-heading text-lg font-bold text-slate-800 mb-1\">Impact Financier (Tornado)</h3>
        <p class=\"text-xs text-slate-500 mb-3\">
          Barres horizontales signant l'écart du coût de transition par rapport au scénario de base pour les variations extrêmes testées.
        </p>
        <div id=\"sens-tornado-chart\" style=\"height:500px;\"></div>
      </section>

      <!-- 3. Trajectoires de Décarbonation (Net CO2) -->
      <section class=\"glass-card rounded-3xl p-5 md:p-6 fade-in\">
        <div class=\"grid grid-cols-1 lg:grid-cols-2 gap-12\">
          <!-- Gauche: Trajectoire Temporelle -->
          <div>
            <h3 class=\"font-heading text-lg font-bold text-slate-800 mb-1\">Trajectoires de Décarbonation</h3>
            <p class=\"text-xs text-slate-500 mb-3\">
              Projection annuelle des émissions nettes (Scope 1 + Scope 2 - Captage DAC - Crédits) pour chaque scénario de variation.
            </p>
            <div id=\"sens-trajectory-chart\" style=\"height:500px;\"></div>
          </div>
          <!-- Droite: Sensibilité Totale -->
          <div>
            <h3 class=\"font-heading text-lg font-bold text-slate-800 mb-1\">Émissions Totales vs Variation</h3>
            <p class=\"text-xs text-slate-500 mb-3\">
              Impact cumulé sur les émissions totales sur l'horizon en fonction du pourcentage de variation du paramètre ciblé.
            </p>
            <div id=\"sens-total-co2-chart\" style=\"height:500px;\"></div>
          </div>
        </div>
      </section>

      <!-- 4. Scatter Coût vs CO₂ -->
      <section class=\"glass-card rounded-3xl p-5 md:p-6 fade-in\">
        <h3 class=\"font-heading text-lg font-bold text-slate-800 mb-1\">Coût vs Émissions CO₂</h3>
        <p class=\"text-xs text-slate-500 mb-3\">
          Chaque point est une simulation. L'axe X représente la variation du coût de transition (%) et l'axe Y la variation des émissions totales (%).
        </p>
        <div id=\"sens-scatter-chart\" style=\"height:500px;\"></div>
      </section>

    </div>
  </div>

  <div id=\"licence-tab\" class=\"tab-content max-w-7xl mx-auto px-4 md:px-8 py-10\">
    <section class=\"glass-card rounded-3xl p-12 text-center\">
      <h2 class=\"text-3xl font-heading font-bold mb-4\">Licence</h2>
      <p class=\"text-slate-500\">Licencing and attribution details.</p>
    </section>
  </div>

  <script>
    const dashboardData = __DASHBOARD_DATA__;
    // Données d'analyse de sensibilité (injectées par generate_results_dashboard.py)
    const sensitivityData = __SENSITIVITY_DATA__;

    function switchTab(tabId, el) {
      // Hide all tabs
      document.querySelectorAll('.tab-content').forEach(tab => {
        tab.classList.remove('active');
      });
      // Remove active class from buttons
      document.querySelectorAll('.nav-tab').forEach(btn => {
        btn.classList.remove('active');
      });
      // Show target tab
      document.getElementById(tabId + '-tab').classList.add('active');
      // Set button as active
      el.classList.add('active');

      // Trigger Plotly resize if we switched to results
      if (tabId === 'results') {
        window.dispatchEvent(new Event('resize'));
      }
    }

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
      const currentSelection = graphSelect.value;
      const keys = graphKeysForScenario(selectedScenario);
      graphSelect.innerHTML = '';
      keys.forEach((key) => {
        const graph = dashboardData.scenarios[selectedScenario].graphs[key];
        const option = document.createElement('option');
        option.value = key;
        option.textContent = graph.label || key;
        graphSelect.appendChild(option);
      });
      if (currentSelection && keys.includes(currentSelection)) {
        graphSelect.value = currentSelection;
      }
    }

    async function renderGraph() {
      const scenarioKey = scenarioSelect.value;
      const graphKey = graphSelect.value;
      const scenario = dashboardData.scenarios[scenarioKey] || { graphs: {} };
      const payload = scenario.graphs[graphKey] || null;

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
      fig.layout.font = { family: 'Bookman Old Style, Bookman, serif', size: 13, color: '#1e293b' };

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

    // ═══════════════════════════════════════════════════════════════════════
    // GRAPHIQUES DE SENSIBILITÉ
    // ═══════════════════════════════════════════════════════════════════════

    (function buildSensitivityCharts() {
      const data = sensitivityData || [];

      // Badge de statut
      const badge = document.getElementById('sens-status-badge');
      if (badge) {
        if (data.length === 0) {
          badge.innerHTML = '<span style="color:#d97706;"><i class="fa-solid fa-triangle-exclamation"></i> Aucune donnée — Exécutez run_sensitivity.py</span>';
        } else {
          const validCount = data.filter(r => r.status === 'Optimal' || r.status === 'Feasible').length;
          const shortfallCount = data.filter(r => (r.penalty_cost || 0) > 1.0).length;
          let html = `<span style="color:#16a34a;"><i class="fa-solid fa-circle-check"></i> ${validCount} / ${data.length} simulations valides</span>`;
          if (shortfallCount > 0) {
            html += ` <span style="color:#dc2626; margin-left:10px;"><i class="fa-solid fa-circle-exclamation"></i> ${shortfallCount} cibles non atteintes (Pénalités réduites dans les graphiques)</span>`;
          }
          badge.innerHTML = html;
        }
      }

      if (data.length === 0) return;

      // ── Données valides uniquement ──────────────────────────────────────
      const valid = data.filter(r => r.transition_cost != null);

      // Paramètre de base (variation = 0)
      const baseRecord = valid.find(r => Math.abs(r.variation_pct) < 0.001) || valid[0];
      
      // transition_cost est désormais déjà le NET BALANCE (M€ ou €)
      // On convertit en M€ pour l'affichage
      const baseCost = baseRecord ? baseRecord.transition_cost / 1_000_000.0 : 0;
      const baseEmis = baseRecord ? (baseRecord.total_emissions || 0) : 0;

      // Groupement par cible
      const targets = {};
      valid.forEach(r => {
        if (!targets[r.target]) targets[r.target] = [];
        targets[r.target].push(r);
      });

      const plotConfig = { responsive: true, displaylogo: false };
      const plotLayout = (extra) => Object.assign({
        paper_bgcolor: 'rgba(0,0,0,0)',
        plot_bgcolor:  'rgba(0,0,0,0)',
        font: { family: 'Bookman Old Style, serif', size: 11, color: '#1e293b' },
        margin: { l: 56, r: 24, t: 32, b: 50 },
        legend: { orientation: 'h', y: -0.15, font: { size: 10 } },
        hovermode: 'closest',
      }, extra || {});

      // ── 1. Packed Bubble (Vue Globale) ───────────────────────────────────
      (function buildBubble() {
        const bubbleTraces = Object.entries(targets).map(([target, records]) => {
          const costs = records.map(r => r.transition_cost / 1_000_000.0);
          const maxVariance = Math.max(...costs) - Math.min(...costs);
          
          return {
            type: 'scatter',
            mode: 'markers+text',
            name: target,
            x: [0],
            y: [0],
            text: [target],
            textposition: 'middle center',
            textfont: { size: 13, color: '#fff' },
            marker: {
              size: [Math.max(60, Math.min(180, maxVariance * 5))], // Scale factors
              sizemode: 'diameter',
              color: ['#0ea5e9'],
              opacity: 0.85,
              line: { width: 2, color: '#fff' },
            },
            hovertemplate: `<b>${target}</b><br>Variance max : %{customdata:,.1f} M€<extra></extra>`,
            customdata: [maxVariance],
          };
        });
        const layout = plotLayout({
          title: { text: 'Variance maximale du Bilan de Transition (M€)', font: { size: 13 } },
          xaxis: { visible: false, zeroline: false },
          yaxis: { visible: false, zeroline: false },
          showlegend: false,
        });
        Plotly.newPlot('sens-bubble-chart', bubbleTraces, layout, plotConfig);
      })();

      // ── 2. Tornado Chart (Impact Financier) ─────────────────────────────
      (function buildTornado() {
        const tornadoTraces = [];

        Object.entries(targets).forEach(([target, records], idx) => {
          const sorted = [...records].sort((a, b) => a.variation_pct - b.variation_pct);
          const minRec = sorted[0];
          const maxRec = sorted[sorted.length - 1];

          const deltaMin = (minRec && minRec.transition_cost != null) ? (minRec.transition_cost / 1_000_000.0) - baseCost : 0;
          const deltaMax = (maxRec && maxRec.transition_cost != null) ? (maxRec.transition_cost / 1_000_000.0) - baseCost : 0;

          tornadoTraces.push({
            type: 'bar',
            orientation: 'h',
            name: `${target} (variation basse ${minRec ? minRec.variation_pct.toFixed(0) : ''}%)`,
            y: [target],
            x: [deltaMin],
            marker: { color: deltaMin < 0 ? '#16a34a' : '#dc2626' },
            hovertemplate: `<b>${target}</b><br>Variation : ${minRec ? minRec.variation_pct.toFixed(1) : 0}%<br>Δ Bilan : %{x:,.1f} M€<extra></extra>`,
          });
          tornadoTraces.push({
            type: 'bar',
            orientation: 'h',
            name: `${target} (variation haute ${maxRec ? maxRec.variation_pct.toFixed(0) : ''}%)`,
            y: [target],
            x: [deltaMax],
            marker: { color: deltaMax < 0 ? '#16a34a' : '#dc2626' },
            hovertemplate: `<b>${target}</b><br>Variation : ${maxRec ? maxRec.variation_pct.toFixed(1) : 0}%<br>Δ Bilan : %{x:,.1f} M€<extra></extra>`,
          });
        });

        const layout = plotLayout({
          barmode: 'overlay',
          title: { text: 'Impact sur le Bilan Net / Scénario de Base', font: { size: 13 } },
          xaxis: { title: 'Δ Bilan Net de Transition (M€)', zeroline: true, zerolinewidth: 2, zerolinecolor: '#94a3b8' },
          yaxis: { automargin: true },
        });
        Plotly.newPlot('sens-tornado-chart', tornadoTraces, layout, plotConfig);
      })();

      // ── 3. Trajectoires et Sensibilité CO₂ ───────────────────────────────
      (function buildDecarbonizationViews() {
        const trajectoryTraces = [];
        const summaryTraces = [];
        
        // ── 3a. Trajectoires Temporelles (Gauche) ──────────────────────────
        Object.entries(targets).forEach(([target, records]) => {
          records.forEach(r => {
            if (!r.co2_trajectory || !r.co2_trajectory.years) return;
            const isBase = Math.abs(r.variation_pct) < 0.001;
            
            trajectoryTraces.push({
              type: 'scatter',
              mode: 'lines',
              name: isBase ? `Baseline (BS)` : `${target} (${r.variation_pct > 0 ? '+' : ''}${r.variation_pct.toFixed(0)}%)`,
              x: r.co2_trajectory.years,
              y: r.co2_trajectory.values,
              line: {
                width: isBase ? 4 : 2,
                dash: isBase ? 'solid' : 'dot',
                color: isBase ? '#7c3aed' : undefined,
                shape: 'spline'
              },
              opacity: isBase ? 1 : 0.7,
              hovertemplate: `<b>${target} (${r.variation_pct.toFixed(0)}%)</b><br>Année %{x}<br>Net CO2 : %{y:,.0f} t<extra></extra>`
            });
          });
        });

        const trajectoryLayout = plotLayout({
          title: { text: 'Projection Temporelle du Net CO₂', font: { size: 13 } },
          xaxis: { title: 'Année' },
          yaxis: { title: 'Net CO₂ (t)', zeroline: true },
          showlegend: true,
          legend: { orientation: 'h', y: -0.25 }
        });
        
        Plotly.newPlot('sens-trajectory-chart', trajectoryTraces, trajectoryLayout, plotConfig);

        // ── 3b. Émissions Totales vs Variation (Droite) ────────────────────
        Object.entries(targets).forEach(([target, records]) => {
          const sorted = [...records].sort((a, b) => a.variation_pct - b.variation_pct);
          
          summaryTraces.push({
            type: 'scatter',
            mode: 'markers',
            name: target,
            x: sorted.map(r => r.variation_pct),
            y: sorted.map(r => r.total_emissions),
            marker: { size: 10, line: { width: 1, color: '#fff' } },
            hovertemplate: `<b>${target}</b><br>Var. Paramètre : %{x:+.1f}%<br>Émissions Totales : %{y:,.0f} t<extra></extra>`
          });
        });

        const summaryLayout = plotLayout({
          title: { text: 'Sensibilité : Émissions Totales vs Variation Paramètre', font: { size: 13 } },
          xaxis: { title: 'Variation du Paramètre (%)', zeroline: true, zerolinewidth: 2, zerolinecolor: '#94a3b8' },
          yaxis: { title: 'Émissions Totales Horizon (tCO₂)', gridcolor: '#f1f5f9' },
          showlegend: true,
          legend: { orientation: 'h', y: -0.25 }
        });

        Plotly.newPlot('sens-total-co2-chart', summaryTraces, summaryLayout, plotConfig);
      })();

      // ── 4. Scatter Coût vs CO₂ ──────────────────────────────────────────
      (function buildScatter() {
        const scatterTraces = Object.entries(targets).map(([target, records]) => {
          const filtered = records.filter(r =>
            r.transition_cost != null && r.total_emissions != null && baseEmis !== 0
          );
          
          return {
            type: 'scatter',
            mode: 'markers',
            name: target,
            x: filtered.map(r => (r.transition_cost / 1_000_000.0) - baseCost),
            y: filtered.map(r => (r.total_emissions - baseEmis) / Math.abs(baseEmis) * 100),
            marker: {
              size: 14,
              color: filtered.map(r => r.variation_pct),
              colorscale: 'RdYlGn',
              reversescale: true,
              colorbar: { title: { text: 'Var. EUA (%)' }, thickness: 12, len: 0.7 },
              line: { width: 1.5, color: '#fff' },
            },
            text: filtered.map(r => {
              const deltaBalance = (r.transition_cost / 1_000_000.0) - baseCost;
              const deltaE = (r.total_emissions - baseEmis) / Math.abs(baseEmis) * 100;
              return `${r.timed_out ? '⏱ Temps limité<br>' : ''}Paramètre : ${target}<br>` +
                     `Variation : ${r.variation_pct.toFixed(1)}%<br>` +
                     `<b>Δ Bilan Net de Transition : ${deltaBalance >= 0 ? '+' : ''}${deltaBalance.toLocaleString('fr-FR', { maximumFractionDigits: 1 })} M€</b><br>` +
                     `Δ Émissions : ${deltaE >= 0 ? '+' : ''}${deltaE.toFixed(4)}%<br>` +
                     `Émissions : ${r.total_emissions.toLocaleString('fr-FR')} tCO₂` +
                     (r.penalty_cost > 1.0 ? `<br><span style="color:red">⚠️ Cible non atteinte (Gap: ${r.gap_from_final_target.toFixed(0)}t)</span>` : '');
            }),
            hovertemplate: '%{text}<extra></extra>',
          };
        });

        // Ligne de base (Origine)
        scatterTraces.push({
          type: 'scatter',
          mode: 'markers',
          name: 'Baseline',
          x: [0], y: [0],
          marker: { size: 18, color: '#7c3aed', symbol: 'star', line: { width: 2, color: '#fff' } },
          hovertemplate: 'Scénario de base (Origine)<br>Δ Bilan : 0 M€<br>Δ Émissions : 0%<extra></extra>',
        });

        const layout = plotLayout({
          title: { text: 'Variation du Bilan Net (M€) vs Variation des Émissions (%)', font: { size: 13 } },
          xaxis: { title: 'Δ Bilan Net de Transition (M€ / Base)', zeroline: true, zerolinewidth: 2, zerolinecolor: '#94a3b8', tickformat: ',.0f' },
          yaxis: { title: 'Δ Émissions Totales (%)',  zeroline: true, zerolinewidth: 2, zerolinecolor: '#94a3b8', tickformat: '.3f' },
          shapes: [
            { type: 'line', x0: 0, x1: 0, y0: -1, y1: 1, line: { color: '#cbd5e1', width: 1, dash: 'dot' } },
          ],
        });
        Plotly.newPlot('sens-scatter-chart', scatterTraces, layout, plotConfig);
      })();

    })();

  </script>
</body>
</html>
"""

    return (
        template
        .replace("__DASHBOARD_DATA__", payload_json)
        .replace("__SENSITIVITY_DATA__", sensitivity_json)
    )


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
    sensitivity_data = load_sensitivity_data()
    html = build_html(payload, sensitivity_data=sensitivity_data)

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