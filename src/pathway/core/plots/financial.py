import pandas as pd
import plotly.graph_objects as go
from typing import List, Dict, Any, Optional

def build_transition_cost_figure(
    df_annual: pd.DataFrame,
    years: List[int],
    pos_cols: List[str],
    neg_cols: List[str],
    investment_cap: float = 0.0,
    theme: str = "report",
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

    # --- Theme Configuration ---
    is_dashboard = (theme == "dashboard")
    is_dark = (theme == "dark")
    bg_color = "rgba(0,0,0,0)" if is_dashboard else ("#111827" if is_dark else "white")
    text_color = "#EEEEEE" if (is_dark or is_dashboard) else "#2c3e50"
    grid_color = "#2B2B36" if (is_dark or is_dashboard) else "#eeeeee"
    zeroline_color = "#555555" if (is_dark or is_dashboard) else "black"
    legend_border = "#333333" if (is_dark or is_dashboard) else "#e0e0e0"
    font_family = "Bookman Old Style, Bookman, serif" if is_dashboard else "Arial"
    template = "plotly_dark" if (is_dark or is_dashboard) else "plotly_white"

    # 5. Layout (Dynamic Style based on mode)
    fig.update_layout(
        template=template,
        paper_bgcolor=bg_color, plot_bgcolor=bg_color,
        title=dict(text=title, font=dict(size=20, weight='bold', color=text_color, family=font_family)),
        xaxis=dict(title='Year', gridcolor=grid_color, tickmode='linear', color=text_color, tickfont=dict(family=font_family)),
        yaxis=dict(title='Annual Impact (M€)', gridcolor=grid_color, zeroline=True, zerolinecolor=zeroline_color, zerolinewidth=1.5, color=text_color, tickfont=dict(family=font_family)),
        yaxis2=dict(title=dict(text='Cumulative Net Balance (M€)', font=dict(color=glow_color, family=font_family)), 
                    overlaying='y', side='right', showgrid=False, tickfont=dict(color=glow_color, family=font_family)),
        legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, bordercolor=legend_border, borderwidth=1, font=dict(color=text_color, family=font_family)),
        margin=dict(l=60, r=60, t=100, b=120),
        font=dict(family=font_family)
    )

    # 6. Investment Cap Line
    if investment_cap > 0:
        fig.add_shape(type="line", x0=years[0], x1=years[-1], y0=investment_cap, y1=investment_cap,
                      line=dict(color="#E74C3C", width=2, dash="dash"), layer='above')
        fig.add_annotation(x=years[len(years)//2], y=investment_cap, text=f"Annual Effort Cap ({investment_cap} M€)",
                           showarrow=False, font=dict(color="#E74C3C", family=font_family), yshift=10)

    # 7. Labels for start/end points
    for idx in [0, -1]:
        t = years[idx]
        val = df_net_cumul.iloc[idx]
        fig.add_annotation(
            x=t, y=val, ay=-40 if val >= 0 else 40,
            text=f"<b>{val:.1f} M€</b>", yref='y2',
            showarrow=True, arrowhead=2, arrowcolor=glow_color,
            font=dict(color='white', family=font_family), bgcolor=glow_color, borderpad=4
        )
    
    # 8. Watermark
    fig.add_annotation(
        text="PATHFINDER Industrial Decarbonization Simulation",
        xref="paper", yref="paper", x=0.5, y=1.07,
        showarrow=False, font=dict(size=12, color="gray", family=font_family),
        opacity=0.6
    )

    return fig

def build_external_financing_figure(
    df_plot: pd.DataFrame,
    years: List[int],
    title: str = "FINANCING STRATEGY",
    theme: str = "report"
) -> go.Figure:
    """
    Builds the high-fidelity Financing Strategy figure (Stacked Bar + Cumulative Line).
    Supports Grants, CCfD and Private Loans.
    """
    fig = go.Figure()

    # Colors
    colors = ['#00E5FF', '#FF007F', '#00FF7F', '#FFD700', '#FF8C00', '#9D00FF', '#FF00FF', '#CCFF00']
    glow_color = '#00ffcc'

    # 1. Stacked Bars
    cols = [c for c in df_plot.columns if c != 'Year']
    for i, col in enumerate(cols):
        fig.add_trace(go.Bar(
            x=years, y=df_plot[col],
            name=col,
            marker=dict(color=colors[i % len(colors)]),
            hovertemplate=f"<b>{col}</b><br>Year: %{{x}}<br>Value: %{{y:.2f}} M€<extra></extra>"
        ))

    # 2. Cumulative Line
    yearly_net = df_plot[cols].sum(axis=1)
    cumul_sum = yearly_net.cumsum()

    fig.add_trace(go.Scatter(
        x=years, y=cumul_sum,
        name='Cumulative Support (M€)',
        mode='lines+markers',
        line=dict(color=glow_color, width=4),
        marker=dict(color=glow_color, size=8, line=dict(color='white', width=1.5)),
        yaxis='y2',
        hovertemplate="<b>Cumulative Support</b><br>Year: %{x}<br>Value: %{y:.2f} M€<extra></extra>"
    ))

    # --- Theme Configuration ---
    is_dashboard = (theme == "dashboard")
    is_dark = (theme == "dark")
    bg_color = "rgba(0,0,0,0)" if is_dashboard else ("#111827" if is_dark else "white")
    text_color = "#EEEEEE" if (is_dark or is_dashboard) else "#2c3e50"
    grid_color = "#2B2B36" if (is_dark or is_dashboard) else "#eeeeee"
    font_family = "Bookman Old Style, Bookman, serif" if is_dashboard else "Arial"
    template = "plotly_dark" if (is_dark or is_dashboard) else "plotly_white"

    # 3. Layout
    fig.update_layout(
        template=template,
        barmode='relative',
        paper_bgcolor=bg_color, plot_bgcolor=bg_color,
        title=dict(text=title, font=dict(size=20, weight='bold', color=text_color, family=font_family)),
        xaxis=dict(title='Year', gridcolor=grid_color, tickmode='linear', color=text_color, tickfont=dict(family=font_family)),
        yaxis=dict(title='Annual Triggered Support (M€)', gridcolor=grid_color, color=text_color, tickfont=dict(family=font_family)),
        yaxis2=dict(title=dict(text='Cumulative Support (M€)', font=dict(color=glow_color, family=font_family)),
                    overlaying='y', side='right', showgrid=False, tickfont=dict(color=glow_color, family=font_family)),
        legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, bordercolor=grid_color, borderwidth=1, font=dict(color=text_color, family=font_family)),
        margin=dict(l=60, r=60, t=100, b=120),
        hovermode="x unified",
        font=dict(family=font_family)
    )

    return fig

def build_interest_paid_figure(
    df_plot: pd.DataFrame,
    years: List[int],
    title: str = "BANK LOANS: INTEREST & CONTRACTED AMOUNTS",
    theme: str = "report"
) -> go.Figure:
    """
    Builds the high-fidelity Interest Paid figure (Bars + Cumulative Line + Contracted Dots).
    """
    fig = go.Figure()

    interest = df_plot['Interest_Paid (M€)'].tolist()
    loans_taken = df_plot['Loan_Principal_Taken (M€)'].tolist()
    cumul_interest = df_plot['Interest_Paid (M€)'].cumsum().tolist()

    # 1. Annual Interest Bars
    fig.add_trace(go.Bar(
        x=years, y=interest,
        name='Annual Interest Paid (M€)',
        marker=dict(color='#E74C3C', opacity=0.8),
        hovertemplate="<b>Annual Interest</b><br>Year: %{x}<br>Value: %{y:.2f} M€<extra></extra>"
    ))

    # 2. Cumulative Interest Line
    fig.add_trace(go.Scatter(
        x=years, y=cumul_interest,
        name='Cumulative Interest Paid (M€)',
        mode='lines+markers',
        line=dict(color='#2C3E50', width=3),
        marker=dict(color='white', size=6, line=dict(color='#2C3E50', width=2)),
        yaxis='y2',
        hovertemplate="<b>Cumulative Interest</b><br>Year: %{x}<br>Value: %{y:.2f} M€<extra></extra>"
    ))

    # 3. Loan Amounts Contracted (Dots on a 3rd axis or shared axis)
    fig.add_trace(go.Scatter(
        x=years, y=loans_taken,
        name='Loan Amount Contracted (M€)',
        mode='markers',
        marker=dict(color='#27AE60', size=12, symbol='square', line=dict(color='white', width=1.5)),
        yaxis='y3',
        hovertemplate="<b>Loan Contracted</b><br>Year: %{x}<br>Value: %{y:.2f} M€<extra></extra>"
    ))

    # --- Theme Configuration ---
    is_dashboard = (theme == "dashboard")
    is_dark = (theme == "dark")
    bg_color = "rgba(0,0,0,0)" if is_dashboard else ("#111827" if is_dark else "white")
    text_color = "#EEEEEE" if (is_dark or is_dashboard) else "#2c3e50"
    grid_color = "#2B2B36" if (is_dark or is_dashboard) else "#eeeeee"
    font_family = "Bookman Old Style, Bookman, serif" if is_dashboard else "Arial"
    template = "plotly_dark" if (is_dark or is_dashboard) else "plotly_white"

    # 4. Layout
    fig.update_layout(
        template=template,
        paper_bgcolor=bg_color, plot_bgcolor=bg_color,
        title=dict(text=title, font=dict(size=18, weight='bold', color=text_color, family=font_family)),
        xaxis=dict(title='Year', gridcolor=grid_color, tickmode='linear', color=text_color, tickfont=dict(family=font_family)),
        yaxis=dict(title='Annual Interest Paid (M€)', gridcolor=grid_color, color='#E74C3C', tickfont=dict(family=font_family)),
        yaxis2=dict(title='Cumulative Interest (M€)', overlaying='y', side='right', showgrid=False, color='#2C3E50', tickfont=dict(family=font_family)),
        yaxis3=dict(title='Loan Principal Taken (M€)', overlaying='y', side='left', anchor='free', position=0.08, showgrid=False, color='#27AE60', tickfont=dict(family=font_family)),
        legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, bordercolor=grid_color, borderwidth=1, font=dict(color=text_color, family=font_family)),
        margin=dict(l=100, r=80, t=100, b=120),
        hovermode="x unified",
        font=dict(family=font_family)
    )

    return fig

def build_mac_figure(
    df_plot: pd.DataFrame,
    avg_carbon_price: float = 0.0,
    total_simulation_abatement: float = 0.0,
    title: str = "CO2 ABATEMENT COST BY TECHNOLOGY (MAC)",
    theme: str = "report"
) -> go.Figure:
    """
    Builds the high-fidelity Marginal Abatement Cost figure using Plotly.
    Includes CAPEX/OPEX components and Profitable Zone.
    """
    fig = go.Figure()

    # Colors
    is_invested = [s == 'Invested' for s in df_plot['Status']]
    
    # Total Abatement for text
    total_kt = total_simulation_abatement / 1000.0

    # 1. CAPEX Part
    fig.add_trace(go.Bar(
        x=df_plot['Project'], y=df_plot['MAC CAPEX (€/tCO2)'],
        name='CAPEX Component',
        marker=dict(color='#3498DB', opacity=[0.9 if inv else 0.5 for inv in is_invested]),
        hovertemplate="<b>%{x} (CAPEX)</b><br>Cost: %{y:.2f} €/tCO2<extra></extra>"
    ))

    # 2. OPEX Part
    fig.add_trace(go.Bar(
        x=df_plot['Project'], y=df_plot['MAC OPEX (€/tCO2)'],
        name='OPEX Component',
        marker=dict(color='#E67E22', opacity=[0.9 if inv else 0.5 for inv in is_invested]),
        hovertemplate="<b>%{x} (OPEX)</b><br>Cost: %{y:.2f} €/tCO2<extra></extra>"
    ))

    # 3. Profitable Zone Line
    if avg_carbon_price > 0:
        fig.add_shape(type="line", x0=-0.5, x1=len(df_plot)-0.5, y0=avg_carbon_price, y1=avg_carbon_price,
                      line=dict(color="#E74C3C", width=2, dash="dash"))
        fig.add_annotation(x=len(df_plot)-1, y=avg_carbon_price, text=f"Avg Carbon Price ({avg_carbon_price:,.0f} €/t)",
                           showarrow=False, bgcolor="white", opacity=0.8, xanchor="right", yshift=10)

    # --- Theme Configuration ---
    is_dashboard = (theme == "dashboard")
    is_dark = (theme == "dark")
    bg_color = "rgba(0,0,0,0)" if is_dashboard else ("#111827" if is_dark else "white")
    text_color = "#EEEEEE" if (is_dark or is_dashboard) else "#2c3e50"
    grid_color = "#2B2B36" if (is_dark or is_dashboard) else "#eeeeee"
    font_family = "Bookman Old Style, Bookman, serif" if is_dashboard else "Arial"
    template = "plotly_dark" if (is_dark or is_dashboard) else "plotly_white"

    # 4. Layout
    fig.update_layout(
        template=template,
        barmode='stack',
        paper_bgcolor=bg_color, plot_bgcolor=bg_color,
        title=dict(text=title, font=dict(size=18, weight='bold', color=text_color, family=font_family)),
        xaxis=dict(title='Implemented Technology per Process', gridcolor=grid_color, color=text_color, tickfont=dict(family=font_family)),
        yaxis=dict(title='Cost of Abatement (€ / tCO2 avoided)', gridcolor=grid_color, color=text_color, tickfont=dict(family=font_family), zeroline=True, zerolinecolor=text_color),
        legend=dict(orientation="h", yanchor="top", y=-0.3, xanchor="center", x=0.5, bordercolor=grid_color, borderwidth=1, font=dict(color=text_color, family=font_family)),
        margin=dict(l=60, r=60, t=100, b=150),
        font=dict(family=font_family)
    )

    # 5. Annotation for total
    fig.add_annotation(
        text=f"Total Simulation Abatement: {total_kt:,.0f} ktCO2",
        xref="paper", yref="paper", x=0.98, y=0.02,
        showarrow=False, font=dict(size=12, color=text_color, weight='bold', family=font_family),
        bgcolor="rgba(255,255,255,0.1)", bordercolor=text_color, borderwidth=1
    )

    return fig
