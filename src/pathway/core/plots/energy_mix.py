import plotly.graph_objects as go
import pandas as pd
from typing import List, Dict, Any

def build_energy_mix_figure(
    years: List[int],
    category_data: Dict[str, Dict[str, Any]],
    title: str = "ENERGY MIX",
    theme: str = "report"
) -> go.Figure:
    """
    Builds a Plotly figure for Energy Mix with a dropdown to select categories.
    """
    fig = go.Figure()
    
    categories = sorted(list(category_data.keys()))
    if not categories:
        fig.add_annotation(text="No data available for Energy Mix", showarrow=False)
        return fig

    # --- Theme Configuration ---
    is_dashboard = (theme == "dashboard")
    bg_color = "rgba(0,0,0,0)" if is_dashboard else "white"
    font_family = "Bookman Old Style, Bookman, serif" if is_dashboard else "Arial"
    text_color = "#1e293b" if is_dashboard else "#0f172a"
    grid_color = "#f1f5f9"
    template = "plotly_white"

    # Create traces for all categories, but only show the first one initially
    trace_indices_per_cat = {}
    current_idx = 0
    
    for cat in categories:
        data = category_data[cat]
        unit = data['unit']
        series_dict = data['series']
        
        trace_indices_per_cat[cat] = []
        
        # Sort series by average value to have a consistent look
        sorted_series = sorted(series_dict.items(), key=lambda x: sum(abs(v) for v in x[1]), reverse=True)
        
        for name, values in sorted_series:
            fig.add_trace(go.Scatter(
                x=years,
                y=values,
                name=name,
                mode='lines',
                line=dict(width=0.5),
                stackgroup=cat, # Stack by category
                visible=(cat == categories[0]),
                hovertemplate=f"%{{y:,.1f}} {unit}<extra>{name}</extra>"
            ))
            trace_indices_per_cat[cat].append(current_idx)
            current_idx += 1

    # Create dropdown buttons
    buttons = []
    for cat in categories:
        # Visibility list: True for traces in this category, False otherwise
        visibility = [False] * current_idx
        for idx in trace_indices_per_cat[cat]:
            visibility[idx] = True
            
        unit = category_data[cat]['unit']
        
        buttons.append(dict(
            label=cat,
            method="update",
            args=[
                {"visible": visibility},
                {"yaxis": {
                    "title": f"Energy Flow ({unit})", 
                    "gridcolor": grid_color,
                    "tickfont": dict(family=font_family, color=text_color)
                }}
            ]
        ))

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
            title=f"Energy Flow ({category_data[categories[0]]['unit']})",
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
