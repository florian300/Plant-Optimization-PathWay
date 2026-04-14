import pandas as pd
import plotly.graph_objects as go
from typing import List, Dict, Any, Optional

def build_opex_figure(
    df_opex: pd.DataFrame,
    years: List[int],
    title: str = "RESSOURCES OPEX: ANNUAL OPERATIONAL EXPENDITURE BREAKDOWN",
    is_dark_bg: bool = False
) -> go.Figure:
    """
    Builds a high-fidelity OPEX breakdown figure (Stacked Area + Total Line).
    This function is the SOLE Source of Truth for both PNG reports and the HTML dashboard.
    """
    fig = go.Figure()

    # Dynamic Color Palette (Mixing shades for categories)
    # Resources: Blues/Greens, Tech: Purples, Taxes/Credits: Oranges/Reds
    palette = [
        '#1a5276', '#2471a3', '#2980b9', '#5499c7', # Blues
        '#1e8449', '#27ae60', '#2ecc71', '#58d68d', # Greens
        '#8e44ad', '#a569bd', '#bb8fce', '#d2b4de', # Purples
        '#d35400', '#e67e22', '#f39c12', '#f4d03f', # Oranges/Yellows
        '#c0392b', '#e74c3c', '#ec7063', '#f1948a', # Reds
        '#2c3e50', '#34495e', '#7f8c8d', '#95a5a6'  # Grays
    ]

    # 1. Identify categories (columns except 'Year')
    categories = [col for col in df_opex.columns if col not in ['Year', 'TOTAL ANNUAL OPEX (M€)']]
    
    # 2. Add Stacked Areas
    for i, cat in enumerate(categories):
        color = palette[i % len(palette)]
        fig.add_trace(go.Scatter(
            x=years, y=df_opex[cat],
            name=cat,
            stackgroup='one',
            mode='lines',
            line=dict(width=0.5, color='white'),
            fill='tonexty',
            fillcolor=color,
            marker=dict(color=color),
            hovertemplate="<b>" + cat + "</b><br>Year: %{x}<br>Value: %{y:.2f} M€<extra></extra>"
        ))

    # 3. Add Total Line
    total_opex = df_opex.get('TOTAL ANNUAL OPEX (M€)', df_opex[categories].sum(axis=1))
    line_color = '#111827' if not is_dark_bg else '#EEEEEE'
    
    fig.add_trace(go.Scatter(
        x=years, y=total_opex,
        name='Total Annual OPEX',
        mode='lines+markers',
        line=dict(color=line_color, width=3, dash='dash'),
        marker=dict(size=8, color=line_color, symbol='circle', line=dict(color='white', width=1)),
        hovertemplate="<b>Total OPEX</b><br>Year: %{x}<br>Value: %{y:.2f} M€<extra></extra>"
    ))

    # 4. Annotations for start/end points
    for idx in [0, -1]:
        if idx < len(years):
            t = years[idx]
            val = total_opex.iloc[idx]
            fig.add_annotation(
                x=t, y=val,
                text=f"<b>{val:.1f} M€</b>",
                showarrow=True, arrowhead=2, arrowcolor=line_color,
                ax=0, ay=-40,
                font=dict(color='white', size=11),
                bgcolor=line_color, borderpad=4, opacity=0.9
            )

    # 5. Layout Styling
    bg_color = 'rgba(0,0,0,0)' if is_dark_bg else 'white'
    text_color = '#EEEEEE' if is_dark_bg else '#2c3e50'
    grid_color = '#2B2B36' if is_dark_bg else '#eeeeee'
    template = 'plotly_dark' if is_dark_bg else 'plotly_white'

    fig.update_layout(
        template=template,
        paper_bgcolor=bg_color, plot_bgcolor=bg_color,
        title=dict(
            text=title,
            font=dict(size=20, weight='bold', color=text_color),
            x=0.02, y=0.95
        ),
        xaxis=dict(
            title='Year', showgrid=True, gridcolor=grid_color, 
            tickmode='linear', color=text_color,
            tickfont=dict(size=12)
        ),
        yaxis=dict(
            title='Annual OPEX (M€)', showgrid=True, gridcolor=grid_color,
            zeroline=True, zerolinecolor=text_color, color=text_color,
            tickfont=dict(size=12)
        ),
        legend=dict(
            orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5,
            bgcolor='rgba(255,255,255,0.1)', bordercolor=grid_color, borderwidth=1,
            font=dict(size=11, color=text_color)
        ),
        margin=dict(l=60, r=40, t=100, b=100),
        hovermode='x unified'
    )

    # 6. Watermark / Subtitle
    fig.add_annotation(
        text="PATHFINDER Industrial Decarbonization Simulation",
        xref="paper", yref="paper", x=0.5, y=1.05,
        showarrow=False, font=dict(size=12, color="gray"),
        opacity=0.6
    )

    return fig
