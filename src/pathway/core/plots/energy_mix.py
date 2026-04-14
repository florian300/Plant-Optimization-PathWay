import plotly.graph_objects as go
import pandas as pd
from typing import List, Dict, Any

def build_resources_mix_figure(
    years: List[int],
    type_data: Dict[str, Dict[str, Dict[str, Any]]],
    title: str = "RESOURCES MIX",
    theme: str = "report"
) -> go.Figure:
    """
    Builds a Plotly figure for Resources Mix with a dropdown to select Type (Emissions, Consumption, Production).
    
    type_data structure: {
        'CONSUMPTION': { 'Electricity': { 'unit': 'MWh', 'series': { 'Grid': [...] } } },
        'PRODUCTION': { ... },
        'EMISSIONS': { ... }
    }
    """
    fig = go.Figure()
    
    # Sort types: CONSUMPTION, PRODUCTION, EMISSIONS
    available_types = []
    for t in ['CONSUMPTION', 'PRODUCTION', 'EMISSIONS']:
        if t in type_data and type_data[t]:
            available_types.append(t)
    
    if not available_types:
        fig.add_annotation(text="No data available for Resources Mix", showarrow=False)
        return fig

    # --- Theme Configuration ---
    is_dashboard = (theme == "dashboard")
    bg_color = "rgba(0,0,0,0)" if is_dashboard else "white"
    font_family = "Bookman Old Style, Bookman, serif" if is_dashboard else "Arial"
    text_color = "#1e293b" if is_dashboard else "#0f172a"
    grid_color = "#f1f5f9"
    template = "plotly_white"

    # Create traces for all Types and Categories
    trace_indices_per_type = {}
    current_idx = 0
    
    # We want to identify the "initial" type to show
    initial_type = available_types[0]

    for t_id in available_types:
        trace_indices_per_type[t_id] = []
        categories = sorted(list(type_data[t_id].keys()))
        
        for cat in categories:
            data = type_data[t_id][cat]
            unit = data['unit']
            series_dict = data['series']
            
            # Sort series by average value
            sorted_series = sorted(series_dict.items(), key=lambda x: sum(abs(v) for v in x[1]), reverse=True)
            
            for name, values in sorted_series:
                fig.add_trace(go.Scatter(
                    x=years,
                    y=values,
                    name=name,
                    mode='lines',
                    line=dict(width=0.5),
                    stackgroup=f"{t_id}_{cat}", # Stack within Type-Category
                    visible=(t_id == initial_type),
                    hovertemplate=f"%{{y:,.1f}} {unit}<extra>{name} [{cat}]</extra>"
                ))
                trace_indices_per_type[t_id].append(current_idx)
                current_idx += 1

    # Create dropdown buttons for Types
    buttons = []
    for t_id in available_types:
        # Visibility list: True for traces in this Type, False otherwise
        visibility = [False] * current_idx
        for idx in trace_indices_per_type[t_id]:
            visibility[idx] = True
            
        # Get the first unit found in this Type
        first_cat = next(iter(type_data[t_id]))
        unit = type_data[t_id][first_cat]['unit']
        
        y_title = "Resources Volume"
        if t_id == 'EMISSIONS': y_title = "Indirect CO2 Emissions"
        elif t_id == 'PRODUCTION': y_title = "Local Production"
        else: y_title = "Gross Consumption"

        buttons.append(dict(
            label=t_id,
            method="update",
            args=[
                {"visible": visibility},
                {"yaxis": {
                    "title": f"{y_title} ({unit})", 
                    "gridcolor": grid_color,
                    "tickfont": dict(family=font_family, color=text_color)
                }}
            ]
        ))

    initial_y_title = "Resources Volume"
    if initial_type == 'EMISSIONS': initial_y_title = "Indirect CO2 Emissions"
    elif initial_type == 'PRODUCTION': initial_y_title = "Local Production"
    else: initial_y_title = "Gross Consumption"
    initial_unit = type_data[initial_type][next(iter(type_data[initial_type]))]['unit']

    fig.update_layout(
        updatemenus=[
            dict(
                buttons=buttons,
                direction="down",
                pad={"r": 10, "t": 10},
                showactive=True,
                x=0.0,
                xanchor="left",
                y=1.15,
                yanchor="top",
                bgcolor="white" if not is_dashboard else "rgba(255, 255, 255, 0.9)",
                bordercolor="#e2e8f0",
                font=dict(size=12, color="#1e293b", family=font_family)
            ),
        ],
        template=template,
        paper_bgcolor=bg_color,
        plot_bgcolor=bg_color,
        title=dict(
            text=title,
            font=dict(size=20, weight='bold', color=text_color, family=font_family),
            x=0.5,
            xanchor="center"
        ),
        xaxis=dict(
            title="Year",
            gridcolor=grid_color,
            tickmode='linear',
            dtick=5,
            tickfont=dict(family=font_family, color=text_color)
        ),
        yaxis=dict(
            title=f"{initial_y_title} ({initial_unit})",
            gridcolor=grid_color,
            tickfont=dict(family=font_family, color=text_color)
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.3,
            xanchor="center",
            x=0.5,
            bgcolor="rgba(255,255,255,0.5)" if is_dashboard else "rgba(255,255,255,0.8)",
            bordercolor="#e2e8f0",
            borderwidth=1,
            font=dict(family=font_family, color=text_color)
        ),
        margin=dict(l=60, r=40, t=100, b=100),
        hovermode="x unified",
        font=dict(family=font_family)
    )

    return fig
