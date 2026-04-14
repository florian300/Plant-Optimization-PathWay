import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from typing import List, Dict, Any, Optional
import math

def build_simulation_prices_figure(
    price_series: Dict[str, Dict[str, Any]],
    years: List[int],
    title: str = "SIMULATION PRICE PARAMETERS",
    theme: str = "report"
) -> go.Figure:
    """
    Builds the high-fidelity Simulation Prices figure (Grid of subplots).
    Each subplot shows a resource price and its CO2 intensity if available.
    """
    n_plots = len(price_series)
    if n_plots == 0:
        return go.Figure()

    cols = math.ceil(math.sqrt(n_plots))
    rows = math.ceil(n_plots / cols)

    # --- Theme Configuration ---
    is_dashboard = (theme == "dashboard")
    is_dark = (theme == "dark")
    bg_color = "rgba(0,0,0,0)" if is_dashboard else ("#111827" if is_dark else "white")
    text_color = "#EEEEEE" if (is_dark or is_dashboard) else "#2c3e50"
    grid_color = "#2B2B36" if (is_dark or is_dashboard) else "#eeeeee"
    font_family = "Bookman Old Style, Bookman, serif" if is_dashboard else "Arial"
    template = "plotly_dark" if (is_dark or is_dashboard) else "plotly_white"

    fig = make_subplots(
        rows=rows, cols=cols,
        subplot_titles=[info['name'].upper() for info in price_series.values()],
        shared_xaxes=False,
        vertical_spacing=0.15,
        horizontal_spacing=0.1
    )

    for i, (key, info) in enumerate(price_series.items()):
        curr_row = (i // cols) + 1
        curr_col = (i % cols) + 1
        
        p_years = sorted(info['data'].keys())
        p_values = [info['data'][y] for y in p_years]
        
        main_color = info.get('color', '#3B82F6')

        # 1. Price Line
        fig.add_trace(
            go.Scatter(
                x=p_years, y=p_values,
                name=f"{info['name']} Price",
                mode='lines+markers',
                line=dict(color=main_color, width=3),
                marker=dict(size=6, color='white', line=dict(color=main_color, width=2)),
                hovertemplate=f"<b>{info['name']} Price</b><br>Year: %{{x}}<br>Price: %{{y:.2f}} {info['unit']}<extra></extra>"
            ),
            row=curr_row, col=curr_col
        )

        # 2. CO2 Intensity (Secondary Axis if we were doing it in Plotly, 
        # but for simplicity and parity with the grid, we might just put it as a second line 
        # or use secondary_y but make_subplots needs to be initialized with specs)
        
        # Updating axes style for this subplot
        fig.update_xaxes(
            title_text="Year", gridcolor=grid_color, tickmode='linear', dtick=5,
            row=curr_row, col=curr_col, color=text_color, tickfont=dict(family=font_family)
        )
        fig.update_yaxes(
            title_text=info['unit'], gridcolor=grid_color, color=text_color,
            row=curr_row, col=curr_col, tickfont=dict(family=font_family),
            zeroline=True, zerolinecolor=text_color
        )

    fig.update_layout(
        template=template,
        paper_bgcolor=bg_color, plot_bgcolor=bg_color,
        title=dict(text=title, font=dict(size=20, weight='bold', color=text_color, family=font_family), x=0.5),
        showlegend=False,
        margin=dict(l=60, r=60, t=100, b=100),
        font=dict(family=font_family)
    )

    # Global annotation / subtitle
    fig.add_annotation(
        text="PATHFINDER Price & CO2 Intensity Metrics",
        xref="paper", yref="paper", x=0.5, y=1.05,
        showarrow=False, font=dict(size=12, color="gray", family=font_family),
        opacity=0.7
    )

    return fig
