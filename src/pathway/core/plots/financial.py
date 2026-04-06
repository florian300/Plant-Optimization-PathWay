import pandas as pd
import plotly.graph_objects as go
from typing import List, Dict, Any, Optional

def build_transition_cost_figure(
    df_annual: pd.DataFrame,
    years: List[int],
    pos_cols: List[str],
    neg_cols: List[str],
    investment_cap: float = 0.0,
    is_dark_bg: bool = False,
    title: str = "ECOLOGICAL TRANSITION: ANNUAL EFFORTS & SAVINGS"
) -> go.Figure:
    """
    Builds the high-fidelity Transition Cost figure (Area Charts + Cumulative Balance).
    This function is the SOLE Source of Truth for both PNG and HTML dashboard.
    """
    fig = go.Figure()

    # Colors (Standard Claire Palette)
    colors_efforts = ['#1a5276', '#5499c7', '#8e44ad', '#5dade2', '#aed6f1', '#111827']
    colors_savings = ['#1e8449', '#58d68d', '#f39c12', '#2ecc71']

    # 1. Calculate Net and Cumulative Balance for secondary axis
    # Ensure all columns are numeric to prevent TypeErrors during sum(axis=1)
    df_pos = df_annual[pos_cols].apply(pd.to_numeric, errors='coerce').fillna(0.0) if pos_cols else pd.DataFrame(index=df_annual.index)
    df_neg = df_annual[neg_cols].apply(pd.to_numeric, errors='coerce').fillna(0.0) if neg_cols else pd.DataFrame(index=df_annual.index)
    
    annual_pos = df_pos.sum(axis=1)
    annual_neg = df_neg.sum(axis=1)
    df_annual_net = annual_pos + annual_neg
    df_net_cumul = df_annual_net.cumsum()

    # 2. Efforts (Positive Area Stack)
    for i, col in enumerate(pos_cols):
        if col in df_pos.columns:
            display_name = col.replace("Effort: ", "")
            fig.add_trace(go.Scatter(
                x=years, y=df_pos[col],
                name=display_name, stackgroup='pos',
                mode='lines', 
                line=dict(width=0.5, color='white'),
                fill='tonexty', 
                fillcolor=colors_efforts[i % len(colors_efforts)],
                marker=dict(color=colors_efforts[i % len(colors_efforts)]),
                customdata=[display_name]*len(years),
                hovertemplate="<b>%{customdata}</b><br>Year: %{x}<br>Value: %{y:.2f} M€<extra></extra>"
            ))

    # 3. Savings (Negative Area Stack)
    for i, col in enumerate(neg_cols):
        if col in df_neg.columns:
            display_name = col.replace("Saving: ", "")
            fig.add_trace(go.Scatter(
                x=years, y=df_neg[col],
                name=display_name, stackgroup='neg',
                mode='lines', 
                line=dict(width=0.5, color='white'),
                fill='tonexty',
                fillcolor=colors_savings[i % len(colors_savings)],
                marker=dict(color=colors_savings[i % len(colors_savings)]),
                customdata=[display_name]*len(years),
                hovertemplate="<b>%{customdata}</b><br>Year: %{x}<br>Value: %{y:.2f} M€<extra></extra>"
            ))

    # 4. Cumulative Net Balance (Secondary Axis)
    glow_color = '#e74c3c'
    fig.add_trace(go.Scatter(
        x=years, y=df_net_cumul.tolist(),
        mode='lines+markers', name='Net Transition Balance (Cumulative)',
        line=dict(color=glow_color, width=4),
        marker=dict(color=glow_color, size=10, line=dict(color='white', width=2)),
        yaxis='y2', hovertemplate="<b>Cumulative Balance</b><br>Year: %{x}<br>Value: %{y:.2f} M€<extra></extra>"
    ))

    # 5. Layout (Clear Style)
    fig.update_layout(
        template='plotly_white',
        paper_bgcolor='white', plot_bgcolor='white',
        title=dict(text=title, font=dict(size=20, weight='bold', color='#2c3e50')),
        xaxis=dict(title='Year', gridcolor='#eeeeee', tickmode='linear'),
        yaxis=dict(title='Annual Impact (M€)', gridcolor='#eeeeee', zeroline=True, zerolinecolor='black', zerolinewidth=1.5),
        yaxis2=dict(title=dict(text='Cumulative Net Balance (M€)', font=dict(color=glow_color)), 
                    overlaying='y', side='right', showgrid=False, tickfont=dict(color=glow_color)),
        legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, bordercolor="#e0e0e0", borderwidth=1),
        margin=dict(l=60, r=60, t=100, b=120),
    )

    # 6. Investment Cap Line
    if investment_cap > 0:
        fig.add_shape(type="line", x0=years[0], x1=years[-1], y0=investment_cap, y1=investment_cap,
                      line=dict(color="#E74C3C", width=2, dash="dash"), layer='above')
        fig.add_annotation(x=years[len(years)//2], y=investment_cap, text=f"Annual Effort Cap ({investment_cap} M€)",
                           showarrow=False, font=dict(color="#E74C3C"), yshift=10)

    # 7. Labels for start/end points
    for idx in [0, -1]:
        t = years[idx]
        val = df_net_cumul.iloc[idx]
        fig.add_annotation(
            x=t, y=val, ay=-40 if val >= 0 else 40,
            text=f"<b>{val:.1f} M€</b>", yref='y2',
            showarrow=True, arrowhead=2, arrowcolor=glow_color,
            font=dict(color='white'), bgcolor=glow_color, borderpad=4
        )
    
    # 8. Watermark
    fig.add_annotation(
        text="PATHFINDER Industrial Decarbonization Simulation",
        xref="paper", yref="paper", x=0.5, y=1.07,
        showarrow=False, font=dict(size=12, color="gray"),
        opacity=0.6
    )

    return fig
