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
    theme: str = "report"
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
        paper_bgcolor=bg_color,
        plot_bgcolor=bg_color,
        title=dict(
            text=title.upper(),
            font=dict(size=18, weight='bold', color=text_color, family=font_family),
            x=0.5, xanchor='center'
        ),
        xaxis=dict(
            title="Year",
            gridcolor=grid_color,
            tickmode='linear',
            range=[min(years)-0.5, max(years)+0.5],
            tickfont=dict(family=font_family, color=text_color)
        ),
        yaxis=dict(
            title="Euro (€) / tCO2",
            gridcolor=grid_color,
            zeroline=True,
            zerolinecolor=text_color,
            tickfont=dict(family=font_family, color=text_color)
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.25,
            xanchor="center",
            x=0.5,
            bgcolor="rgba(0,0,0,0)",
            bordercolor=grid_color,
            borderwidth=1,
            font=dict(family=font_family, color=text_color, size=11)
        ),
        margin=dict(l=60, r=60, t=80, b=100),
        hovermode="closest",
        font=dict(family=font_family)
    )

    # 5. Watermark / Subtitle
    fig.add_annotation(
        text="PATHFINDER Carbon Policy Simulation",
        xref="paper", yref="paper",
        x=0.5, y=1.06,
        showarrow=False,
        font=dict(size=12, color="gray", family=font_family),
        opacity=0.7
    )

    return fig

def build_co2_trajectory_figure(
    df: pd.DataFrame,
    objectives: List[Any] = None,
    base_emissions: float = 0.0,
    title: str = "CO2 EMISSIONS TRAJECTORY & GOALS",
    theme: str = "report"
) -> go.Figure:
    """
    Builds the high-fidelity CO2 Trajectory figure (Aires + Curves).
    This function is the SOLE Source of Truth for both PNG and HTML dashboard.
    """
    KT = 1000.0
    df_plot = df.copy()
    
    # 1. Normalize and Calculate
    for col in ['Direct_CO2', 'Indirect_CO2', 'Total_CO2', 'Taxed_CO2', 'Free_Quota']:
        if col in df_plot.columns:
            df_plot[col] = df_plot[col] / KT
    
    dac_cap = df_plot.get('DAC_Captured_kt', pd.Series(0, index=df_plot.index))
    cred = df_plot.get('Credits_Purchased_kt', pd.Series(0, index=df_plot.index))
    has_dac_or_cred = (dac_cap + cred).max() > 1e-4

    net_direct = df_plot['Direct_CO2'] - dac_cap - cred
    net_total = net_direct + df_plot['Indirect_CO2']

    fig = go.Figure()

    # 2. Areas
    # Shade under Net Direct
    fig.add_trace(go.Scatter(
        x=df_plot['Year'], y=net_direct,
        fill='tozeroy', fillcolor='rgba(52, 152, 219, 0.1)',
        mode='none', showlegend=False, name='Net Direct Shade'
    ))

    # DAC & Credits (Negative areas)
    if has_dac_or_cred:
        fig.add_trace(go.Scatter(
            x=df_plot['Year'], y=[-x for x in dac_cap],
            name='DAC Captured (ktCO2)',
            fill='tozeroy', fillcolor='rgba(52, 152, 219, 0.6)',
            mode='lines', line=dict(width=0)
        ))
        fig.add_trace(go.Scatter(
            x=df_plot['Year'], y=[-(d + c) for d, c in zip(dac_cap, cred)],
            name='Voluntary Credits (ktCO2)',
            fill='tonexty', fillcolor='rgba(39, 174, 96, 0.6)',
            mode='lines', line=dict(width=0)
        ))

    # Free Quotas
    fig.add_trace(go.Scatter(
        x=df_plot['Year'], y=df_plot['Free_Quota'],
        name='Free Quotas (Direct)',
        fill='tozeroy', fillcolor='rgba(0, 128, 0, 0.3)',
        mode='lines', line=dict(width=0)
    ))

    # Taxed Emissions
    fig.add_trace(go.Scatter(
        x=df_plot['Year'], y=df_plot['Free_Quota'] + df_plot['Taxed_CO2'],
        name='Taxed Emissions (Surface)',
        fill='tonexty',
        fillcolor='rgba(128, 128, 128, 0.4)',
        fillpattern=dict(shape=".", solidity=0.3),
        mode='lines', line=dict(width=0)
    ))

    # 3. Trajectory Curves
    fig.add_trace(go.Scatter(
        x=df_plot['Year'], y=df_plot['Direct_CO2'],
        name='Direct Emissions', line=dict(color='black', width=3), mode='lines'
    ))
    fig.add_trace(go.Scatter(
        x=df_plot['Year'], y=df_plot['Indirect_CO2'],
        name='Indirect Emissions', line=dict(color='black', width=2, dash='dot'), mode='lines'
    ))
    fig.add_trace(go.Scatter(
        x=df_plot['Year'], y=net_total,
        name='Total Emissions (Net)', line=dict(color='darkred', width=3, dash='dash'), mode='lines'
    ))
    fig.add_trace(go.Scatter(
        x=df_plot['Year'], y=net_direct,
        name='Net Direct Emissions', line=dict(color='#3498db', width=3, dash='dashdot'), mode='lines'
    ))

    # 4. Objectives (Markers)
    if objectives:
        plotted_groups = set()
        available_colors = ['#e74c3c', '#9b59b6', '#f39c12', '#1abc9c', '#34495e', '#d35400', '#2ecc71']
        group_colors = {}

        for obj in objectives:
            if obj.resource == 'CO2_EM':
                if hasattr(obj, 'comparison_year') and obj.comparison_year and -1.0 <= obj.cap_value <= 1.0:
                    limit = base_emissions * (1 + obj.cap_value) / KT
                else:
                    limit = obj.cap_value / KT
                
                display_name = obj.name if obj.name else (obj.group if obj.group else 'Goal')
                show_legend = display_name not in plotted_groups
                plotted_groups.add(display_name)
                
                grp = obj.group if obj.group else 'Default'
                if grp not in group_colors:
                    group_colors[grp] = available_colors[len(group_colors) % len(available_colors)]
                
                fig.add_trace(go.Scatter(
                    x=[obj.target_year], y=[limit],
                    mode='markers', name=display_name, showlegend=show_legend,
                    marker=dict(symbol='x', size=12, line=dict(width=3), color=group_colors[grp])
                ))

    # --- Theme Configuration ---
    is_dashboard = (theme == "dashboard")
    is_dark = (theme == "dark")
    bg_color = "rgba(0,0,0,0)" if is_dashboard else ("#111827" if is_dark else "white")
    text_color = "#EEEEEE" if (is_dark or is_dashboard) else "#2c3e50"
    grid_color = "#2B2B36" if (is_dark or is_dashboard) else "#eeeeee"
    font_family = "Bookman Old Style, Bookman, serif" if is_dashboard else "Arial"
    template = "plotly_dark" if (is_dark or is_dashboard) else "plotly_white"

    # 5. Layout
    fig.update_layout(
        template=template,
        paper_bgcolor=bg_color,
        plot_bgcolor=bg_color,
        title=dict(
            text=title.upper(),
            font=dict(size=18, weight='bold', color=text_color, family=font_family),
            x=0.5, xanchor='center'
        ),
        xaxis=dict(
            title="Year",
            gridcolor=grid_color,
            tickmode='linear',
            tickfont=dict(family=font_family, color=text_color)
        ),
        yaxis=dict(
            title="ktCO2",
            gridcolor=grid_color,
            zeroline=True,
            zerolinecolor=text_color,
            tickfont=dict(family=font_family, color=text_color)
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.25,
            xanchor="center",
            x=0.5,
            bgcolor="rgba(0,0,0,0)",
            bordercolor=grid_color if is_dashboard else "#e5e7eb",
            borderwidth=1,
            font=dict(family=font_family, color=text_color, size=11)
        ),
        margin=dict(l=60, r=60, t=100, b=120),
        hovermode="x unified",
        font=dict(family=font_family)
    )

    return fig

def build_indirect_emissions_figure(
    df_cat: pd.DataFrame,
    years: List[int],
    title: str = "INDIRECT EMISSIONS BREAKDOWN (SCOPE 2 & 3)",
    theme: str = "report"
) -> go.Figure:
    """
    Builds the high-fidelity Indirect Emissions figure (Stacked Area).
    """
    fig = go.Figure()

    # Premium corporate palette
    premium_palette = ['#2C3E50', '#E67E22', '#2980B9', '#8E44AD', '#16A085', '#D35400']

    # 1. Stacked Area Traces
    cols = [c for c in df_cat.columns if c != 'Year']
    for i, col in enumerate(cols):
        fig.add_trace(go.Scatter(
            x=years, y=df_cat[col],
            name=col,
            stackgroup='one',
            mode='lines',
            line=dict(width=0.5, color='white'),
            fillcolor=premium_palette[i % len(premium_palette)],
            hovertemplate=f"<b>{col}</b><br>Year: %{{x}}<br>Emissions: %{{y:.2f}} ktCO2<extra></extra>"
        ))

    # --- Theme Configuration ---
    is_dashboard = (theme == "dashboard")
    is_dark = (theme == "dark")
    bg_color = "rgba(0,0,0,0)" if is_dashboard else ("#111827" if is_dark else "white")
    text_color = "#EEEEEE" if (is_dark or is_dashboard) else "#2c3e50"
    grid_color = "#2B2B36" if (is_dark or is_dashboard) else "#eeeeee"
    font_family = "Bookman Old Style, Bookman, serif" if is_dashboard else "Arial"
    template = "plotly_dark" if (is_dark or is_dashboard) else "plotly_white"

    # 2. Layout
    fig.update_layout(
        template=template,
        paper_bgcolor=bg_color, plot_bgcolor=bg_color,
        title=dict(text=title.upper(), font=dict(size=18, weight='bold', color=text_color, family=font_family), x=0.5),
        xaxis=dict(title='Year', gridcolor=grid_color, tickmode='linear', color=text_color, tickfont=dict(family=font_family)),
        yaxis=dict(title='ktCO2', gridcolor=grid_color, color=text_color, tickfont=dict(family=font_family)),
        legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, bordercolor=grid_color, borderwidth=1, font=dict(color=text_color, family=font_family)),
        margin=dict(l=60, r=60, t=100, b=120),
        hovermode="x unified",
        font=dict(family=font_family)
    )

    return fig
