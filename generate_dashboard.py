import os
import glob
import json
import pandas as pd
from datetime import datetime
import sys
from pathlib import Path

# Ensure we can import from src/
repo_root = str(Path(__file__).resolve().parent)
src_path = os.path.join(repo_root, "src")
if src_path not in sys.path:
    sys.path.insert(0, src_path)

from pathway.core.plots.financial import build_transition_cost_figure

# --- HTML TEMPLATE ---
HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PathWay - Premium Dashboard</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-dark: #0D0D14;
            --card-bg: rgba(22, 22, 30, 0.7);
            --accent-green: #00FF7F;
            --accent-blue: #3498DB;
            --accent-red: #E74C3C;
            --text-main: #EEEEEE;
            --text-dim: #94A3B8;
            --grid-color: #2B2B36;
        }
        body { 
            background-color: var(--bg-dark); 
            color: var(--text-main); 
            font-family: 'Inter', sans-serif;
            background-image: radial-gradient(circle at 50% 50%, #161622 0%, #0D0D14 100%);
        }
        .glass-card {
            background: var(--card-bg);
            backdrop-filter: blur(12px);
            border: 1px solid rgba(255, 255, 255, 0.05);
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.5);
        }
        select, option {
            background-color: #1E293B !important;
            border-color: #334155 !important;
        }
        .custom-scrollbar::-webkit-scrollbar { width: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: transparent; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #334155; border-radius: 10px; }
        
        /* Premium Glow Effects */
        .glow-text { text-shadow: 0 0 10px rgba(0, 255, 127, 0.3); }
        .btn-premium {
            background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
            border: 1px solid rgba(255,255,255,0.1);
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }
        .btn-premium:hover {
            border-color: var(--accent-green);
            box-shadow: 0 0 15px rgba(0, 255, 127, 0.2);
            transform: translateY(-1px);
        }
    </style>
</head>
<body class="min-h-screen p-8 custom-scrollbar">

    <!-- Header Section -->
    <header class="max-w-7xl mx-auto mb-10 flex flex-col md:flex-row md:items-end justify-between gap-4">
        <div>
            <div class="flex items-center gap-3 mb-2">
                <div class="w-10 h-10 bg-emerald-500 rounded-lg flex items-center justify-center shadow-[0_0_20px_rgba(16,185,129,0.4)]">
                    <span class="text-white font-bold text-xl">P</span>
                </div>
                <h1 class="text-4xl font-bold tracking-tight text-white glow-text">PathWay <span class="text-emerald-400">Intelligence</span></h1>
            </div>
            <p class="text-slate-400 font-medium" id="generationDate">Report generated on: {{GENERATION_DATE}}</p>
        </div>
        <div class="flex gap-3">
             <button onclick="downloadGraph()" class="btn-premium px-6 py-2.5 rounded-xl text-sm font-semibold flex items-center gap-2">
                <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a2 2 0 002 2h12a2 2 0 002-2v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/></svg>
                Export PNG
            </button>
        </div>
    </header>

    <main class="max-w-7xl mx-auto grid grid-cols-12 gap-8">
        
        <!-- Sidebar Controls -->
        <aside class="col-span-12 lg:col-span-3 space-y-6">
            <div class="glass-card rounded-2xl p-6">
                <h3 class="text-xs font-bold uppercase tracking-widest text-slate-500 mb-4">Configuration</h3>
                
                <div class="space-y-5">
                    <div>
                        <label class="block text-sm font-semibold text-slate-300 mb-2">Scenario Analysis</label>
                        <select id="scenarioSelect" class="w-full bg-slate-900 text-white rounded-xl p-3 border border-slate-700 outline-none focus:ring-2 focus:ring-emerald-500/50 transition duration-300">
                        </select>
                    </div>

                    <div>
                        <label class="block text-sm font-semibold text-slate-300 mb-2">Insight Category</label>
                        <select id="graphSelect" class="w-full bg-slate-900 text-white rounded-xl p-3 border border-slate-700 outline-none focus:ring-2 focus:ring-blue-500/50 transition duration-300">
                            <optgroup label="Core Trajectories">
                                <option value="co2">CO2 Emissions & Targets</option>
                                <option value="mac">Marginal Abatement Cost (MAC)</option>
                                <option value="carbon_policy">Carbon Price & CCfD</option>
                            </optgroup>
                            <optgroup label="Economic Balance">
                                <option value="financial">Financial Balance (NPV)</option>
                                <option value="transition">Ecological Transition Costs</option>
                                <option value="investment">Investment Plan (CAPEX)</option>
                            </optgroup>
                            <optgroup label="Resource Flows">
                                <option value="hydrogen">Hydrogen Mass Balance</option>
                                <option value="power">Power & Market Pricing</option>
                            </optgroup>
                        </select>
                    </div>
                </div>
            </div>

            <!-- Context Card -->
            <div class="glass-card rounded-2xl p-6 border-l-4 border-l-emerald-500">
                <h3 class="text-sm font-bold text-emerald-400 mb-2 flex items-center gap-2">
                    <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
                    Methodology
                </h3>
                <p id="methodologyText" class="text-sm text-slate-400 leading-relaxed italic">
                    Loading methodology parameters...
                </p>
            </div>
        </aside>

        <!-- Main Chart Display -->
        <section class="col-span-12 lg:col-span-9 space-y-6">
            <div class="glass-card rounded-[2rem] p-8 min-h-[650px] flex flex-col">
                <div class="flex items-center justify-between mb-8">
                    <h2 id="graphTitle" class="text-2xl font-bold text-white tracking-tight">Strategy Overview</h2>
                    <div class="flex items-center gap-2 px-3 py-1 bg-slate-800/50 rounded-full border border-white/5">
                        <span class="w-2 h-2 bg-emerald-500 rounded-full animate-pulse"></span>
                        <span class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Interactive Result</span>
                    </div>
                </div>
                
                <div id="plotlyChart" class="flex-1 w-full"></div>
            </div>
        </section>

    </main>

    <script>
        const dashboardData = {{INJECTED_JSON_HERE}};

        const methodologies = {
            'co2': "Comprehensive CO2 trajectory. Areas represent the gross emissions and the taxed/unquoted portions, while lines track the Net emissions after DAC/Credits sequestration.",
            'mac': "Marginal Abatement Cost curve identifying the most cost-effective technologies. Sorted from lowest to highest cost per abated ton of CO2.",
            'carbon_policy': "Trajectory of market carbon prices and the effective strike prices for any active Carbon Contract for Difference (CCfD).",
            'financial': "Macro-financial waterfall showing the balance between investments (CAPEX), operational costs (OPEX), and returns (Aids/Revenues).",
            'transition': "Cumulative cost of the ecological transition. Areas break down annual deltas (Sunk CAPEX, interests, credits) against the cumulative net total line.",
            'investment': "Detailed annual CAPEX allocation. Highlights technological application to specific industrial processes, monitored against self-funding and loan limits.",
            'hydrogen': "Mass balance for the H2 ecosystem, integrating production units, user processes, and storage buffering capacity.",
            'power': "Grid interaction profile. Compares power consumption units against the volatile spot market prices."
        };

        // Layout Constants from Reporting.py
        const darkTheme = {
            bg: '#0D0D14',
            grid: '#2B2B36',
            text: '#FFFFFF',
            spine: '#EEEEEE',
            palettes: {
                standard: ['#1A5276', '#E67E22', '#2980B9', '#8E44AD', '#16A085', '#D35400'],
                neon: ['#00E5FF', '#FF007F', '#00FF7F', '#FFD700', '#9D00FF', '#FF00FF'],
                co2: {
                    direct: '#1B4965',
                    indirect: '#A9A9A9',
                    total: '#BEBEBE',
                    net: '#3498DB',
                    free: 'rgba(0, 255, 127, 0.2)',
                    taxed: 'rgba(148, 163, 184, 0.3)',
                    dac: '#3498DB',
                    credits: '#27AE60'
                }
            }
        };

        function updateDashboard() {
            const scenario = document.getElementById('scenarioSelect').value;
            const type = document.getElementById('graphSelect').value;
            if (!dashboardData.data[scenario]) return;
            const data = dashboardData.data[scenario];

            document.getElementById('methodologyText').textContent = methodologies[type] || "No methodology defined.";
            document.getElementById('graphTitle').textContent = document.querySelector(`#graphSelect option[value="${type}"]`)?.textContent || "Chart";

            const layout = {
                paper_bgcolor: 'rgba(0,0,0,0)',
                plot_bgcolor: 'rgba(0,0,0,0)',
                font: { color: darkTheme.text, family: 'Inter, Arial, sans-serif' },
                margin: { t: 20, b: 80, l: 60, r: 60 },
                xaxis: { 
                    gridcolor: darkTheme.grid, 
                    showline: true, linecolor: darkTheme.spine,
                    tickfont: { size: 10 }
                },
                yaxis: { 
                    gridcolor: darkTheme.grid, 
                    showline: true, linecolor: darkTheme.spine,
                    tickfont: { size: 10 }
                },
                legend: { orientation: 'h', y: -0.2, x: 0.5, xanchor: 'center', font: {size: 11} },
                hovermode: 'x unified'
            };

            if (type === 'co2') {
                const c = darkTheme.palettes.co2;
                const traces = [
                    // Negative Sinks (DAC/Credits) - Plotted first to be at physical bottom
                    { x: data.co2.time, y: data.co2.dac.map(v => -v), name: 'DAC Captured', type: 'scatter', fill: 'tozeroy', fillcolor: 'rgba(52, 152, 219, 0.4)', line: {width:0}, stackgroup: 'sinks' },
                    { x: data.co2.time, y: data.co2.credits.map(v => -v), name: 'Voluntary Credits', type: 'scatter', fill: 'tonexty', fillcolor: 'rgba(39, 174, 96, 0.4)', line: {width:0}, stackgroup: 'sinks' },

                    // Positive Emissions (Areas)
                    { x: data.co2.time, y: data.co2.free_quotas_direct, name: 'Free Quotas', type: 'scatter', fill: 'tozeroy', fillcolor: c.free, line: {width:0} },
                    { x: data.co2.time, y: data.co2.taxed_emissions, name: 'Taxed Emissions', type: 'scatter', fill: 'tonexty', fillcolor: c.taxed, line: {width:0} },
                    
                    // Main lines
                    { x: data.co2.time, y: data.co2.direct_emissions, name: 'Direct Emissions', line: {color: '#FFFFFF', width: 3} },
                    { x: data.co2.time, y: data.co2.indirect_emissions, name: 'Indirect Emissions', line: {color: '#888888', width: 2, dash: 'dot'} },
                    { x: data.co2.time, y: data.co2.net_direct_emissions, name: 'Net Direct Emissions', line: {color: c.net, width: 3, dash: 'dashdot'} }
                ];

                // Targets
                if (data.co2.target_time && data.co2.target_time.length > 0) {
                    traces.push({
                        x: data.co2.target_time, y: data.co2.target_values,
                        mode: 'markers', name: 'Goals',
                        marker: { symbol: 'x', size: 14, color: '#FF0000', line: {width: 3} },
                        text: data.co2.target_names
                    });
                }

                const yMax = Math.max(...data.co2.direct_emissions, ...data.co2.taxed_emissions.map((v,i) => v + (data.co2.free_quotas_direct[i]||0))) * 1.25;
                const yMin = Math.min(...data.co2.dac.map(v => -v), ...data.co2.credits.map(v => -v)) * 1.3;

                Plotly.newPlot('plotlyChart', traces, {
                    ...layout,
                    yaxis: { ...layout.yaxis, title: 'ktCO2 (Carbon Balance)', range: [yMin < -5 ? yMin : -5, yMax] },
                    shapes: [{ type: 'line', x0: data.co2.time[0], x1: data.co2.time[data.co2.time.length-1], y0: 0, y1: 0, line: {color: 'rgba(255,255,255,0.5)', width: 1.5, dash: 'solid'} }]
                }, {responsive: true});

            } else if (type === 'mac') {
                if (!data.mac) return showNoData();
                const trace = {
                    type: 'bar', orientation: 'h',
                    y: data.mac.projects, x: data.mac.mac_values,
                    marker: {
                        color: data.mac.mac_values.map(v => v < 50 ? '#27AE60' : v < 150 ? '#F1C40F' : '#E74C3C'),
                        line: { color: 'white', width: 0.5 }
                    },
                    text: data.mac.mac_values.map(v => v.toFixed(0) + ' €/t'),
                    textposition: 'outside'
                };
                Plotly.newPlot('plotlyChart', [trace], {
                    ...layout,
                    xaxis: { ...layout.xaxis, title: 'MAC (€/tCO2)' },
                    margin: { ...layout.margin, l: 200 }
                }, {responsive: true});

            } else if (type === 'investment') {
                const traces = [];
                const p = darkTheme.palettes.standard;
                let i = 0;
                for (const [name, vals] of Object.entries(data.investment.capex_by_process_tech)) {
                    if (vals.some(v => v > 0)) {
                        traces.push({
                            x: data.investment.time, y: vals.map(v => v/1e6),
                            name: name.split('##')[0], type: 'bar',
                            marker: { color: p[i % p.length] }
                        });
                        i++;
                    }
                }
                // Limits
                if (data.investment.budget_limit && data.investment.budget_limit.length > 0) {
                    traces.push({
                        x: data.investment.time, y: data.investment.budget_limit.map(v => v/1e6),
                        name: 'Self-Funded Limit', mode: 'lines', line: {color: '#2ECC71', dash: 'dash', shape: 'hv', width: 3}
                    });
                    traces.push({
                        x: data.investment.time, y: data.investment.total_limit.map(v => v/1e6),
                        name: 'Total Limit', mode: 'lines', line: {color: '#E74C3C', dash: 'dot', shape: 'hv', width: 3}
                    });
                }
                Plotly.newPlot('plotlyChart', traces, {
                    ...layout,
                    barmode: 'stack',
                    yaxis: { ...layout.yaxis, title: 'CAPEX (M€)' }
                }, {responsive: true});

            } else if (type === 'carbon_policy') {
                if (!data.carbon_policy) return showNoData();
                const traces = [
                    { x: data.carbon_policy.time, y: data.carbon_policy.tax_price, name: 'Market Carbon Price', line: {color: '#00E5FF', width: 3} }
                ];
                Plotly.newPlot('plotlyChart', traces, {
                    ...layout,
                    yaxis: { ...layout.yaxis, title: '€/tCO2' }
                }, {responsive: true});

            } else if (type === 'transition') {
                if (!data.transition_fig) return showNoData();
                const fig = data.transition_fig;
                Plotly.newPlot('plotlyChart', fig.data, {
                    ...layout,
                    ...fig.layout,
                    margin: { t: 60, b: 80, l: 60, r: 60 },
                    paper_bgcolor: 'rgba(0,0,0,0)',
                    plot_bgcolor: 'rgba(0,0,0,0)',
                    xaxis: { ...layout.xaxis, ...fig.layout.xaxis },
                    yaxis: { ...layout.yaxis, ...fig.layout.yaxis },
                    yaxis2: { ...fig.layout.yaxis2, overlaying: 'y', side: 'right', showgrid: false },
                    legend: { ...fig.layout.legend, orientation: "h", yanchor: "top", y: -0.2, xanchor: "center", x: 0.5 }
                }, {responsive: true});
                
            } else if (type === 'financial') {
                const trace = {
                    type: "waterfall",
                    orientation: "v",
                    measure: ["relative", "relative", "relative", "relative", "total"],
                    x: ["CAPEX", "OPEX", "Aids", "Revenues", "NPV"],
                    y: [-data.financial.capex, -data.financial.opex, data.financial.aids, data.financial.revenues, data.financial.npv],
                    connector: { line: { color: "rgb(63, 63, 63)" } },
                    increasing: { marker: { color: "#27AE60" } },
                    decreasing: { marker: { color: "#E74C3C" } },
                    totals: { marker: { color: "#2980B9" } }
                };
                Plotly.newPlot('plotlyChart', [trace], { ...layout, yaxis: { ...layout.yaxis, title: 'Amount (M€)' } }, {responsive: true});

            } else if (type === 'hydrogen') {
                const h = data.hydrogen;
                Plotly.newPlot('plotlyChart', [
                    { x: h.time, y: h.production, name: 'H2 Produced', line: {color: '#00FF7F', width: 3}, type: 'scatter' },
                    { x: h.time, y: h.consumption, name: 'H2 Consumed', line: {color: '#FF007F', width: 2, dash: 'dash'}, type: 'scatter' },
                    { x: h.time, y: h.storage_level, name: 'Storage', fill: 'tozeroy', fillcolor: 'rgba(0, 229, 255, 0.15)', yaxis: 'y2', line: {width:0}, type: 'scatter' }
                ], {
                    ...layout,
                    yaxis: { title: 'kg/h' },
                    yaxis2: { overlaying: 'y', side: 'right', title: 'Storage (kg)', showgrid: false }
                }, {responsive: true});

            } else if (type === 'power') {
                Plotly.newPlot('plotlyChart', [
                    { x: data.power.time, y: data.power.consumption, name: 'Power Mix', type: 'bar', marker: {color: '#9D00FF'} },
                    { x: data.power.time, y: data.power.spot_price, name: 'Spot Price', yaxis: 'y2', line: {color: '#FFD700', width: 3}, type: 'scatter' }
                ], {
                    ...layout,
                    yaxis: { title: 'MW' },
                    yaxis2: { overlaying: 'y', side: 'right', title: 'Price (€/MWh)', showgrid: false }
                }, {responsive: true});
            }
        }

        function showNoData() {
            Plotly.newPlot('plotlyChart', [], {
                annotations: [{ text: "No data available for this visualization.", xref: "paper", yref: "paper", showarrow: false, font: {size: 20} }]
            });
        }

        function init() {
            const select = document.getElementById('scenarioSelect');
            Object.keys(dashboardData.data).forEach(s => {
                const opt = document.createElement('option');
                opt.value = s; opt.textContent = s;
                select.appendChild(opt);
            });
            select.addEventListener('change', updateDashboard);
            document.getElementById('graphSelect').addEventListener('change', updateDashboard);
            updateDashboard();
        }

        function downloadGraph() {
            Plotly.downloadImage('plotlyChart', {format: 'png', width: 1200, height: 800, filename: 'PathWay_Insight'});
        }

        window.onload = init;
    </script>
</body>
</html>"""

# --- PYTHON EXTRACTOR ---

def parse_excel_scenarios(data_dir: str) -> dict:
    """
    Parses Excel files in the given directory to build the master JSON dictionary.
    Exhaustively extracts data from Master_Plan.xlsx format generated by reporting.py.
    """
def parse_excel_scenarios(data_dir: str) -> dict:
    """
    Search results directories for Master_Plan.xlsx and build a structured dictionary.
    """
    master_data = {"data": {}}
    excel_files = glob.glob(os.path.join(data_dir, "**", "Master_Plan.xlsx"), recursive=True)

    if not excel_files:
        print(f"Warning: No Master_Plan.xlsx found in {data_dir}. Generating mock data.")
        return generate_mock_data()

    for file_path in excel_files:
        scenario_name = os.path.basename(os.path.dirname(file_path))
        print(f"Processing Scenario: {scenario_name}")
        try:
            xl = pd.ExcelFile(file_path)
            dfs = {name: pd.read_excel(xl, name) for name in xl.sheet_names}
            s_data = {}

            # --- CO2 Trajectory ---
            if 'CO2_Trajectory' in dfs:
                df_co2 = dfs['CO2_Trajectory']
                n = len(df_co2)
                def get_col(col):
                    if col in df_co2.columns: return pd.to_numeric(df_co2[col], errors='coerce').fillna(0.0).tolist()
                    return [0.0] * n
                s_data["co2"] = {
                    "time": [int(y) for y in df_co2['Year']],
                    "direct_emissions": get_col('Direct_CO2'),
                    "net_direct_emissions": get_col('Net_Direct_Emissions'),
                    "taxed_emissions": get_col('Taxed_CO2'),
                    "tax_cost": get_col('Tax_Cost_MEuros'), # Vital for baseline delta
                    "dac": get_col('DAC_Captured_kt'),
                    "credits": get_col('Credits_Purchased_kt')
                }

            # --- Financial & Opex proxies ---
            if 'Technology_Costs' in dfs:
                df_costs = dfs['Technology_Costs']
                n_costs = len(df_costs)
                def get_fin(col):
                    if col in df_costs.columns: return pd.to_numeric(df_costs[col], errors='coerce').fillna(0.0).tolist()
                    return [0.0] * n_costs
                s_data["financial"] = {
                    "self_funded_capex": get_fin('Self-funded_Capex'),
                    "loan_service": get_fin('Financing Interests'),
                    "tech_dac_opex": get_fin('DAC_Opex'),
                    "credit_cost": get_fin('Credit_Cost')
                }

            # --- Resource & Energy Mix ---
            if 'Energy_Mix' in dfs:
                df_mix = dfs['Energy_Mix']
                s_data["res_cons"] = {col: df_mix[col].tolist() for col in df_mix.columns if col != 'Year'}
            if 'Data_Used' in dfs:
                df_data = dfs['Data_Used']
                p_map = {}
                for idx, row in df_data.iterrows():
                    try:
                        r_id = str(row.get('Resource', '')).strip()
                        y_id = int(float(row.get('Year', 0)))
                        p_val = float(row.get('Price', 0.0))
                        if r_id and y_id > 0:
                            p_map[(r_id, y_id)] = p_val
                    except: continue
                s_data["res_prices"] = p_map
                # print(f"  [DEBUG] Found {len(p_map)} price entries in {scenario_name}")

            # --- Charts extraction (Try to find the table if it exists) ---
            if 'Charts' in dfs:
                df_charts = dfs['Charts']
                try:
                    if "TRANSITION_COST_HIGH_FIDELITY" in str(df_charts.columns[0]):
                        idx = -1
                    else:
                        trans_idx = df_charts[df_charts.iloc[:, 0].astype(str).str.contains("TRANSITION_COST_HIGH_FIDELITY", na=False)].index
                        idx = trans_idx[0] if not trans_idx.empty else None
                        
                    if idx is not None:
                        header = df_charts.iloc[idx+1].tolist()
                        df_trans_raw = df_charts.iloc[idx+2:].copy()
                        end_idx = df_trans_raw[df_trans_raw.iloc[:, 0].isna()].index
                        if not end_idx.empty: df_trans_raw = df_trans_raw.loc[:end_idx[0]-1]
                        df_trans_raw.columns = header
                        
                        year_col = next((c for c in df_trans_raw.columns if "year" in str(c).lower() or "index" in str(c).lower()), df_trans_raw.columns[1] if len(df_trans_raw.columns)>1 else df_trans_raw.columns[0])
                        years = [int(float(y)) for y in df_trans_raw[year_col].dropna()]
                        
                        all_cols = [str(c) for c in df_trans_raw.columns if c != year_col and not str(c).startswith("Unnamed:")]
                        pos_cols = [c for c in all_cols if c.startswith("Effort:")]
                        neg_cols = [c for c in all_cols if c.startswith("Saving:")]
                        
                        for c in (pos_cols + neg_cols):
                            df_trans_raw[c] = pd.to_numeric(df_trans_raw[c], errors="coerce").fillna(0.0)
                        
                        fig = build_transition_cost_figure(
                            df_annual=df_trans_raw,
                            years=years,
                            pos_cols=pos_cols,
                            neg_cols=neg_cols,
                            is_dark_bg=True,
                            title="ECOLOGICAL TRANSITION COSTS"
                        )
                        s_data["transition_fig"] = json.loads(fig.to_json())
                except Exception as e: print(f"  [DEBUG] Failed to parse transition_fig for {scenario_name}: {e}")

            master_data["data"][scenario_name] = s_data
        except Exception as e: print(f"Error reading {file_path}: {e}")

    # --- Robust Autonomous Delta Calculation ---
    baseline_scenario_name = next((n for n in master_data["data"] if any(b in n.upper() for b in ["BUSINESS AS USUAL", "BASELINE", "BAU"])), None)
    baseline_scenario = master_data["data"].get(baseline_scenario_name)
    if not baseline_scenario: print(f"  [WARNING] No baseline scenario found for delta calculation! (Checked: {list(master_data['data'].keys())})")
    else: print(f"  [DEBUG] Using '{baseline_scenario_name}' as baseline for resource deltas.")

    def calculate_scenario_resource_costs(data, scen_name=""):
        if "res_cons" not in data or "co2" not in data: return []
        years = [int(float(y)) for y in data["co2"]["time"]]
        cons_data = data["res_cons"]
        prices = data.get("res_prices", {}) 
        if not prices and baseline_scenario: prices = baseline_scenario.get("res_prices", {})

        total_costs = []
        for i, yr in enumerate(years):
            yr_cost = 0.0
            for res_id, cons_list in cons_data.items():
                if i < len(cons_list):
                    val = float(cons_list[i])
                    p = prices.get((res_id, yr), 0.0)
                    if p == 0:
                        short_id = res_id.replace('EN_', '')
                        for (pr_res, pr_yr), pr_val in prices.items():
                            if pr_yr == yr and (short_id in pr_res or pr_res in res_id):
                                p = pr_val
                                break
                    
                    if val != 0 and p != 0:
                        yr_cost += (val * p) / 1_000_000.0
            total_costs.append(round(yr_cost, 6))
        
        # Log a sample for Electricity if found
        if scen_name:
            e_cons = cons_data.get('EN_ELEC', [0])[0]
            e_price = prices.get(('EN_ELEC', 2025), 0)
            if e_price == 0: # Try fuzzy for log
                for (pr_res, pr_yr), pr_val in prices.items():
                    if pr_yr == 2025 and ('ELEC' in pr_res): e_price = pr_val; break
            # print(f"    - {scen_name} (2025): Elec Cons {e_cons:.1f}, Price {e_price:.2f}")

        return total_costs

    if baseline_scenario:
        bau_costs = calculate_scenario_resource_costs(baseline_scenario, baseline_scenario_name)
        bau_tax_cost = baseline_scenario.get("co2", {}).get("tax_cost", [0]*26)
        
        for name, data in master_data["data"].items():
            is_bau = (name == baseline_scenario_name)
            if is_bau:
                if "transition" in data:
                    z = [0.0] * len(data["transition"]["time"])
                    data["transition"]["self_funded_capex"] = z
                    data["transition"]["bank_loan_service"] = z
                    data["transition"]["tech_dac_opex"] = z
                    data["transition"]["voluntary_carbon_credits"] = z
                    data["transition"]["annual_avoided_tax"] = z
                    data["transition"]["avoided_resource_saving"] = z
                    data["transition"]["additional_resource_cost"] = z
                    data["transition"]["cumulative_net_cost"] = z
                continue
            
            years = [int(float(y)) for y in data["co2"]["time"]]
            n_yrs = len(years)
            
            s_tax = data["co2"].get("tax_cost", [0.0]*n_yrs)
            b_tax = bau_tax_cost[:n_yrs] if len(bau_tax_cost) >= n_yrs else bau_tax_cost + [0.0]*(n_yrs - len(bau_tax_cost))
            tax_savings_annual = [round(float(s) - float(b), 6) for s, b in zip(s_tax, b_tax)]
            
            s_res_costs = calculate_scenario_resource_costs(data)
            b_res_costs = bau_costs[:n_yrs] if len(bau_costs) >= n_yrs else bau_costs + [0.0]*(n_yrs - len(bau_costs))
            res_delta_annual = [round(float(s) - float(b), 6) for s, b in zip(s_res_costs, b_res_costs)]
            
            add_res_annual = [max(0, x) for x in res_delta_annual]
            avoid_res_annual = [min(0, x) for x in res_delta_annual]

            def to_cumul(l):
                c = 0
                res = []
                for x in l:
                    c += float(x)
                    res.append(round(c, 6))
                return res

            if "transition" not in data:
                self_funded = data.get("financial", {}).get("self_funded_capex", [0.0]*n_yrs)
                loan_service = data.get("financial", {}).get("loan_service", [0.0]*n_yrs)
                tech_opex = data.get("financial", {}).get("tech_dac_opex", [0.0]*n_yrs)
                credits = data.get("financial", {}).get("credit_cost", [0.0]*n_yrs)
                
                data["transition"] = {
                    "time": years,
                    "self_funded_capex": self_funded,
                    "bank_loan_service": loan_service,
                    "tech_dac_opex": tech_opex,
                    "voluntary_carbon_credits": credits,
                    "annual_avoided_tax": tax_savings_annual,
                }

            data["transition"]["additional_resource_cost"] = add_res_annual
            data["transition"]["avoided_resource_saving"] = avoid_res_annual
            
            sf = [float(x) for x in data.get("financial", {}).get("self_funded_capex", [0.0]*n_yrs)]
            ls = [float(x) for x in data.get("financial", {}).get("loan_service", [0.0]*n_yrs)]
            to = [float(x) for x in data.get("financial", {}).get("tech_dac_opex", [0.0]*n_yrs)]
            cr = [float(x) for x in data.get("financial", {}).get("credit_cost", [0.0]*n_yrs)]
            
            net_annual = [(sf[i] + ls[i] + to[i] + cr[i] + res_delta_annual[i] + tax_savings_annual[i]) for i in range(n_yrs)]
            data["transition"]["cumulative_net_cost"] = to_cumul(net_annual)

    for data in master_data["data"].values():
        data.pop("res_prices", None)
        data.pop("res_cons", None)

    return master_data

def generate_mock_data():
    return {"data": {"Baseline": {"co2": {"time": [2025, 2030, 2050], "direct_emissions": [3000, 2500, 500]}}}}

def main():
    print("--- Starting Dashboard Generation ---")
    RESULTS_DIR = "./artifacts/reports/"
    dashboard_dict = parse_excel_scenarios(RESULTS_DIR)
    json_string = json.dumps(dashboard_dict)
    
    html = HTML_TEMPLATE.replace('{{INJECTED_JSON_HERE}}', json_string)
    with open("dashboard.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("[OK] Dashboard generated successfully: dashboard.html")

if __name__ == "__main__":
    main()
