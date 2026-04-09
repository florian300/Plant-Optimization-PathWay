import plotly.graph_objects as go
from typing import List, Optional, Dict, Any

def build_carbon_tax_figure(
    years: List[int],
    standard_tax: List[float],
    penalties: List[float],
    avoided_reduced: List[float],
    avoided_captured: List[float],
    ccfd_refunds: Optional[List[float]] = None,
    title: str = "CARBON TAX & AVOIDED COSTS BALANCE"
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
        name="Standard Carbon Tax",
        marker_color="#EF4444", # Red-500
        hovertemplate="Standard Tax: %{y:.2f} M€<extra></extra>"
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

    # --- Layout ---
    fig.update_layout(
        title=dict(
            text=title,
            x=0.5,
            font=dict(size=18, color="#1F2937", family="Arial")
        ),
        template="plotly_white",
        barmode="relative", # Stacked for positive, stacked for negative
        hovermode="x unified",
        xaxis=dict(
            title="Year",
            tickmode="linear",
            dtick=5,
            gridcolor="#E5E7EB"
        ),
        yaxis=dict(
            title="Annual Financial Impact (M€)",
            gridcolor="#E5E7EB",
            zeroline=True,
            zerolinecolor="#9CA3AF",
            zerolinewidth=2
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.3,
            xanchor="center",
            x=0.5,
            bgcolor="rgba(255, 255, 255, 0.8)",
            bordercolor="#E5E7EB",
            borderwidth=1
        ),
        margin=dict(t=80, b=100, l=60, r=40),
        height=600
    )

    return fig
