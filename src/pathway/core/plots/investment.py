import pandas as pd
import plotly.graph_objects as go
from typing import List, Dict, Any, Optional

def build_investment_plan_figure(
    df_projects: pd.DataFrame,
    df_costs: pd.DataFrame,
    years: List[int],
    theme: str = "report",
    title: str = "INVESTMENT PLAN: IMPLEMENTATION COSTS"
) -> go.Figure:
    """
    Builds the high-fidelity Investment Plan figure (Stacked Bars for CAPEX + Budget Limits).
    This function is the SOLE Source of Truth for both PNG and HTML dashboard.
    """
    fig = go.Figure()

    # 1. Colors & Palettes (Corporate standard)
    proc_palette = [
        '#004B87', '#007AA5', '#0095A8', '#00A69C', '#24A148',
        '#63BA3C', '#BFD02C', '#FADA00', '#F8B400', '#F08000'
    ]
    tech_palette = [
        '#D0021B', '#F5A623', '#F8E71C', '#8B572A', '#BD10E0',
        '#9013FE', '#4A90E2', '#E24A8D', '#00FF00', '#FF00FF'
    ]

    # 2. Identify and Filter Data
    excluded_suffixes = ('##tCO2', '_labels', '_is_new', 'Financing Interests', 'Year', 'Yearly_Total')
    capex_cols = [c for c in df_projects.columns if not any(c.endswith(s) for s in excluded_suffixes) and c != 'Year']
    
    # Filter only columns with non-zero values
    active_cols = [col for col in capex_cols if df_projects[col].sum() > 1.0]

    # Map processes and technologies to colors
    proc_ids = sorted(list({col.split('##')[0] for col in active_cols if '##' in col}))
    proc_to_color = {pid: proc_palette[i % len(proc_palette)] for i, pid in enumerate(proc_ids)}
    
    tech_ids = sorted(list({col.split('##')[1] if '##' in col else col.split('##')[0] for col in active_cols if '##' in col}))
    tech_to_edge_color = {tid: tech_palette[i % len(tech_palette)] for i, tid in enumerate(tech_ids)}

    # 3. Add categorical legend references (Dummy traces)
    # This recreates the "separated" legend look the user requested
    
    # PROCESSES Section
    for pid in proc_ids:
        proc_name = pid if pid != 'INDIRECT' else "Indirect Tech (DAC/Credits)"
        fig.add_trace(go.Bar(
            x=[None], y=[None],
            name=f"Process: {proc_name}",
            marker=dict(color=proc_to_color[pid]),
            legendgroup="PROCESSES",
            legendgrouptitle_text="<b>PROCESSES (Fill)</b>",
            showlegend=True
        ))

    # TECHNOLOGIES Section
    for tid in tech_ids:
        fig.add_trace(go.Bar(
            x=[None], y=[None],
            name=f"Tech: {tid}",
            marker=dict(color='rgba(0,0,0,0)', line=dict(color=tech_to_edge_color[tid], width=2)),
            legendgroup="TECHNOLOGIES",
            legendgrouptitle_text="<b>TECHNOLOGIES (Borders)</b>",
            showlegend=True
        ))

    # 4. Add Actual Data Bars (Hidden from legend to keep it clean, but toggleable as a group if needed)
    for col in active_cols:
        parts = col.split('##')
        pid = parts[0]
        tid = parts[1] if len(parts) > 1 else pid
        
        display_name = col.replace("##", " - ")
        label_col = f"{col}_labels"
        labels = df_projects[label_col].tolist() if label_col in df_projects.columns else [""] * len(years)
        values_meur = (df_projects[col] / 1_000_000.0).tolist()
        
        fig.add_trace(go.Bar(
            x=years,
            y=values_meur,
            name=display_name,
            marker=dict(
                color=proc_to_color.get(pid, '#333333'),
                line=dict(color=tech_to_edge_color.get(tid, 'white'), width=1.5)
            ),
            showlegend=False, # Hide individual combinations to prioritize the categorical legend
            customdata=list(zip([pid]*len(years), [tid]*len(years), labels)),
            hovertemplate=(
                "<b>%{name}</b><br>" +
                "Implementation: %{customdata[2]}<br>" +
                "Cost: %{y:.2f} M€<extra></extra>"
            )
        ))

    # 5. Add Investment Limits (Step Lines)
    if not df_costs.empty and 'Budget_Limit' in df_costs.columns:
        df_costs_sorted = df_costs.sort_values('Year')
        budget_lim = (df_costs_sorted['Budget_Limit'] / 1_000_000.0).tolist()
        total_lim = (df_costs_sorted['Total_Limit'] / 1_000_000.0).tolist()
        lim_years = df_costs_sorted['Year'].tolist()

        fig.add_trace(go.Scatter(
            x=lim_years, y=budget_lim,
            name='Self-funded Limit (Own Cash)',
            mode='lines',
            line=dict(color='#2ECC71', width=3, dash='dash', shape='hv'),
            legendgroup="LIMITS",
            legendgrouptitle_text="<b>INVESTMENT LIMITS</b>",
            hovertemplate="Self-funded Limit: %{y:.1f} M€<extra></extra>"
        ))

        fig.add_trace(go.Scatter(
            x=lim_years, y=total_lim,
            name='Total Investment Limit (Incl. Loans)',
            mode='lines',
            line=dict(color='#E74C3C', width=3, dash='dot', shape='hv'),
            legendgroup="LIMITS",
            hovertemplate="Total Limit: %{y:.1f} M€<extra></extra>"
        ))

    # --- Theme Configuration ---
    is_dashboard = (theme == "dashboard")
    is_dark = (theme == "dark")
    bg_color = "rgba(0,0,0,0)" if is_dashboard else ("#111827" if is_dark else "white")
    text_color = "#EEEEEE" if (is_dark or is_dashboard) else "#2c3e50"
    grid_color = "#2B2B36" if (is_dark or is_dashboard) else "#eeeeee"
    font_family = "Bookman Old Style, Bookman, serif" if is_dashboard else "Arial"
    template = "plotly_dark" if (is_dark or is_dashboard) else "plotly_white"

    # 6. Layout & styling
    fig.update_layout(
        template=template,
        barmode='stack',
        paper_bgcolor=bg_color, plot_bgcolor=bg_color,
        title=dict(text=title, font=dict(size=20, weight='bold', color=text_color, family=font_family)),
        xaxis=dict(
            title='Year', gridcolor=grid_color, 
            tickmode='linear', dtick=5, color=text_color,
            tickfont=dict(family=font_family)
        ),
        yaxis=dict(
            title='Annual Implementation Cost (M€)', 
            gridcolor=grid_color, color=text_color,
            zeroline=True, zerolinecolor=text_color,
            tickfont=dict(family=font_family)
        ),
        legend=dict(
            orientation="h", 
            yanchor="top", y=-0.15, 
            xanchor="center", x=0.5,
            groupclick="toggleitem",
            font=dict(color=text_color, size=11, family=font_family),
            bgcolor='rgba(0,0,0,0)',
            bordercolor=grid_color,
            borderwidth=1
        ),
        margin=dict(l=60, r=60, t=100, b=200),
        hovermode="x unified",
        font=dict(family=font_family)
    )

    # 7. Watermark
    fig.add_annotation(
        text="PATHFINDER Industrial Decarbonization Simulation",
        xref="paper", yref="paper", x=0.5, y=1.07,
        showarrow=False, font=dict(size=12, color="gray", family=font_family),
        opacity=0.6
    )

    return fig
