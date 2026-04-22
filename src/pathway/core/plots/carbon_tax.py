import plotly.graph_objects as go
from typing import List, Optional, Dict, Any

def build_carbon_tax_figure(
    years: List[int],
    standard_tax: List[float],
    penalties: List[float],
    avoided_reduced: List[float],
    avoided_captured: List[float],
    indirect_tax: Optional[List[float]] = None,
    ccfd_refunds: Optional[List[float]] = None,
    title: str = "CARBON TAX & AVOIDED COSTS BALANCE",
    theme: str = "report"
) -> go.Figure:
    """
    Creates a high-fidelity Plotly visualization of carbon costs (tax, penalties) 
    and avoided costs (savings from reduction or capture).
    """
    fig = go.Figure()

    # --- 1. COSTS (Positive Stacked Bars) ---
    fig.add_trace(go.Bar(
        x=years,
        y=standard_tax,
        name="Standard Carbon Tax (S1)",
        marker_color="#EF4444", # Red-500
        hovertemplate="Standard Tax: %{y:.2f} M€<extra></extra>"
    ))

    if indirect_tax and any(v > 1e-4 for v in indirect_tax):
        fig.add_trace(go.Bar(
            x=years,
            y=indirect_tax,
            name="Indirect Carbon Tax (S2/3)",
            marker_color="#F87171", # Red-400 (lighter)
            hovertemplate="Indirect Tax: %{y:.2f} M€<extra></extra>"
        ))

    fig.add_trace(go.Bar(
        x=years,
        y=penalties,
        name="Carbon Penalties",
        marker_color="#F97316", # Orange-500
        hovertemplate="Extra Penalties: %{y:.2f} M€<extra></extra>"
    ))

    # --- 2. SAVINGS (Negative Grouped/Stacked Bars) ---
    # We display them as negative bars (pointing down)
    fig.add_trace(go.Bar(
        x=years,
        y=[-v for v in avoided_reduced],
        name="Avoided (Reduced at Source)",
        marker_color="#10B981", # Emerald-500
        hovertemplate="Savings (Reduction): %{y:.2f} M€<extra></extra>"
    ))

    fig.add_trace(go.Bar(
        x=years,
        y=[-v for v in avoided_captured],
        name="Avoided (Captured via CCS)",
        marker_color="#06B6D4", # Cyan-500
        hovertemplate="Savings (Capture): %{y:.2f} M€<extra></extra>"
    ))

    # --- 3. REFUNDS (Markers) ---
    if ccfd_refunds and any(v > 1e-4 for v in ccfd_refunds):
        fig.add_trace(go.Scatter(
            x=years,
            y=ccfd_refunds,
            name="CCfD State Refund",
            mode="markers",
            marker=dict(
                symbol="diamond",
                size=10,
                color="#8B5CF6", # Violet-500
                line=dict(width=1, color="white")
            ),
            hovertemplate="CCfD Refund: %{y:.2f} M€<extra></extra>"
        ))

    # --- Theme Configuration ---
    is_dashboard = (theme == "dashboard")
    bg_color = "rgba(0,0,0,0)" if is_dashboard else "white"
    font_family = "Bookman Old Style, Bookman, serif" if is_dashboard else "Arial"
    text_color = "#1e293b" if is_dashboard else "#1F2937"
    grid_color = "#f1f5f9" if is_dashboard else "#E5E7EB"

    # --- Layout ---
    fig.update_layout(
        title=dict(
            text=title,
            x=0.5,
            font=dict(size=18, color=text_color, family=font_family)
        ),
        template="plotly_white",
        paper_bgcolor=bg_color,
        plot_bgcolor=bg_color,
        barmode="relative", # Stacked for positive, stacked for negative
        hovermode="x unified",
        xaxis=dict(
            title="Year",
            tickmode="linear",
            dtick=5,
            gridcolor=grid_color,
            tickfont=dict(family=font_family, color=text_color)
        ),
        yaxis=dict(
            title="Annual Financial Impact (M€)",
            gridcolor=grid_color,
            zeroline=True,
            zerolinecolor="#9CA3AF",
            zerolinewidth=2,
            tickfont=dict(family=font_family, color=text_color)
        ),
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.15,
            xanchor="center",
            x=0.5,
            bgcolor="rgba(255, 255, 255, 0.5)" if is_dashboard else "rgba(255, 255, 255, 0.8)",
            bordercolor=grid_color,
            borderwidth=1,
            font=dict(family=font_family, color=text_color, size=11)
        ),
        margin=dict(t=80, b=160, l=60, r=40),
        height=600,
        font=dict(family=font_family)
    )

    return fig
