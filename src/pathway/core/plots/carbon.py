import pandas as pd
import plotly.graph_objects as go
from typing import List, Dict, Any, Optional

def build_carbon_price_figure(
    years: List[int],
    market_prices: List[float],
    effective_prices: Optional[List[float]] = None,
    penalties: Optional[List[float]] = None,
    strike_prices: Optional[List[Dict[str, Any]]] = None,
    title: str = "CARBON PRICE & POLICY TRAJECTORY",
    is_dark_bg: bool = False
) -> go.Figure:
    """
    Builds a high-fidelity Carbon Price figure using Plotly.
    Includes Market Price, Effective Price (Penalty), and CCfD Strike Prices.
    """
    fig = go.Figure()

    # Colors (Clear Style)
    market_color = "#3B82F6"    # Modern Blue
    effective_color = "#F59E0B" # Amber for warnings/penalties
    gap_color = "rgba(245, 158, 11, 0.15)" # Light amber fill
    strike_color = "#8B5CF6"    # Purple for contracts
    
    # 1. Market Carbon Price
    fig.add_trace(go.Scatter(
        x=years, y=market_prices,
        name="Market Carbon Price",
        mode="lines+markers",
        line=dict(color=market_color, width=3),
        marker=dict(size=6, color=market_color, line=dict(color='white', width=1)),
        hovertemplate="<b>Market Price</b><br>Year: %{x}<br>Price: %{y:.2f} €/tCO2<extra></extra>"
    ))

    # 2. Effective Price (with Penalty)
    if effective_prices and any(ep > mp + 0.01 for ep, mp in zip(effective_prices, market_prices)):
        fig.add_trace(go.Scatter(
            x=years, y=effective_prices,
            name="Effective Price (incl. Penalty)",
            mode="lines+markers",
            line=dict(color=effective_color, width=3, dash='dash'),
            marker=dict(size=6, color=effective_color, symbol='diamond'),
            hovertemplate="<b>Effective Price</b><br>Year: %{x}<br>Price: %{y:.2f} €/tCO2<extra></extra>"
        ))

        # Fill the Penalty Gap
        fig.add_trace(go.Scatter(
            x=years + years[::-1],
            y=effective_prices + market_prices[::-1],
            fill='toself',
            fillcolor=gap_color,
            line=dict(color='rgba(255,255,255,0)'),
            hoverinfo='skip',
            showlegend=True,
            name="Penalty Impact Gap"
        ))

    # 3. CCfD Strike Prices
    if strike_prices:
        for i, strike in enumerate(strike_prices):
            s_years = strike.get('years', [])
            s_val = strike.get('val', 0.0)
            s_name = strike.get('name', 'CCfD')
            
            if s_years:
                fig.add_trace(go.Scatter(
                    x=s_years, y=[s_val] * len(s_years),
                    name=f"Strike: {s_name}",
                    mode="lines",
                    line=dict(color=strike_color, width=2.5, dash='dashdot'),
                    hovertemplate=f"<b>Strike Price: {s_name}</b><br>Value: %{{y:.2f}} €/tCO2<extra></extra>"
                ))
    
    # 4. Layout
    template = "plotly_white"
    font_color = "#1f2937"
    grid_color = "#f3f4f6"
    
    fig.update_layout(
        template=template,
        title=dict(
            text=title.upper(),
            font=dict(size=18, weight='bold', color=font_color),
            x=0.5, xanchor='center'
        ),
        xaxis=dict(
            title="Year",
            gridcolor=grid_color,
            tickmode='linear',
            range=[min(years)-0.5, max(years)+0.5]
        ),
        yaxis=dict(
            title="Euro (€) / tCO2",
            gridcolor=grid_color,
            zeroline=True,
            zerolinecolor="#d1d5db"
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.25,
            xanchor="center",
            x=0.5,
            bgcolor="rgba(255,255,255,0.8)",
            bordercolor="#e5e7eb",
            borderwidth=1
        ),
        margin=dict(l=60, r=60, t=80, b=100),
        hovermode="closest"
    )

    # 5. Watermark / Subtitle
    fig.add_annotation(
        text="PATHFINDER Carbon Policy Simulation",
        xref="paper", yref="paper",
        x=0.5, y=1.06,
        showarrow=False,
        font=dict(size=12, color="gray"),
        opacity=0.7
    )

    return fig
