import plotly.graph_objects as go
import pandas as pd
from typing import List, Dict, Any, Tuple

def build_resources_mix_figure(
    years: List[int],
    group_data: Dict[Tuple[str, str], Dict[str, Any]],
    title: str = "RESOURCES MIX: TYPE & CATEGORY",
    theme: str = "report"
) -> go.Figure:
    """
    Builds a Plotly figure for Resources Mix with a dropdown to select both Type and Category.
    
    group_data structure: {
        ('CONSUMPTION', 'ENERGY'): { 'unit': 'MWh', 'series': { 'Grid': [...] } },
        ('PRODUCTION', 'ENERGY'): { ... },
        ('EMISSIONS', 'POLLUTION'): { ... }
    }
    """
    fig = go.Figure()
    
    # Sort groups by Type then Category
    available_groups = sorted(group_data.keys(), key=lambda x: (x[0], x[1]))
    
    if not available_groups:
        fig.add_annotation(text="No data available for Resources Mix", showarrow=False)
        return fig

    # --- Theme Configuration ---
    is_dashboard = (theme == "dashboard")
    bg_color = "rgba(0,0,0,0)" if is_dashboard else "white"
    font_family = "Bookman Old Style, Bookman, serif" if is_dashboard else "Arial"
    text_color = "#1e293b" if is_dashboard else "#0f172a"
    grid_color = "#f1f5f9"
    template = "plotly_white"

    # Create traces for each group
    trace_indices_per_group = {}
    current_idx = 0
    
    # Identify the initial group to show (prefer CONSUMPTION ENERGY if available)
    initial_group = next((g for g in available_groups if g == ('CONSUMPTION', 'ENERGY')), available_groups[0])

    for group_key in available_groups:
        trace_indices_per_group[group_key] = []
        data = group_data[group_key]
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
                stackgroup=f"{group_key[0]}_{group_key[1]}", # Stack within the specific Type-Category group
                visible=(group_key == initial_group),
                hovertemplate=f"%{{y:,.1f}} {unit}<extra>{name} [{group_key[1]}]</extra>"
            ))
            trace_indices_per_group[group_key].append(current_idx)
            current_idx += 1

    # Create dropdown buttons for the combined Type-Category groups
    buttons = []
    for group_key in available_groups:
        t_id, cat_id = group_key
        # Visibility list: True for traces in this group, False otherwise
        visibility = [False] * current_idx
        for idx in trace_indices_per_group[group_key]:
            visibility[idx] = True
            
        unit = group_data[group_key]['unit']
        
        # Determine Y-axis title based on Type
        y_title = "Resources Volume"
        if t_id == 'EMISSIONS': y_title = "Indirect CO2 Emissions"
        elif t_id == 'PRODUCTION': y_title = "Local Production"
        else: y_title = "Gross Consumption"

        # Label: "TYPE CATEGORY"
        btn_label = f"{t_id} {cat_id}"

        buttons.append(dict(
            label=btn_label,
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

    # Initial Y-axis title and unit
    init_type, init_cat = initial_group
    initial_y_title = "Resources Volume"
    if init_type == 'EMISSIONS': initial_y_title = "Indirect CO2 Emissions"
    elif init_type == 'PRODUCTION': initial_y_title = "Local Production"
    else: initial_y_title = "Gross Consumption"
    initial_unit = group_data[initial_group]['unit']

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
