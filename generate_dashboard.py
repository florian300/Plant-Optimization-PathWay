import os
import glob
import json
import pandas as pd
from datetime import datetime

# --- HTML TEMPLATE ---
HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Plant-Optimization-PathWay Dashboard</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
    <style>
        body { background-color: #0D0D14; color: #FFFFFF; } /* Premium Dark Theme bg */
        .glass-card {
            background: rgba(43, 43, 54, 0.4); /* #2B2B36 overlay */
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255, 255, 255, 0.1);
        }
    </style>
</head>
<body class="min-h-screen p-6 font-sans">

    <!-- Header -->
    <header class="mb-8">
        <h1 class="text-3xl font-bold text-emerald-400">Plant-Optimization-PathWay Results</h1>
        <p class="text-slate-400 mt-1" id="generationDate">Generated on: {{GENERATION_DATE}}</p>
    </header>

    <!-- Control Panel -->
    <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
        <div class="glass-card rounded-xl shadow-lg p-5">
            <label class="block text-sm font-medium text-slate-300 mb-2">Select Scenario</label>
            <select id="scenarioSelect" class="w-full bg-slate-700 text-white rounded-lg p-2.5 border border-slate-600 focus:ring-emerald-500 focus:border-emerald-500 outline-none transition">
                <!-- Populated dynamically -->
            </select>
        </div>

        <div class="glass-card rounded-xl shadow-lg p-5">
            <label class="block text-sm font-medium text-slate-300 mb-2">Select Visualization</label>
            <select id="graphSelect" class="w-full bg-slate-700 text-white rounded-lg p-2.5 border border-slate-600 focus:ring-blue-500 focus:border-blue-500 outline-none transition">
                <option value="financial">Financial Summary (NPV)</option>
                <option value="hydrogen">H2 Mass Balance</option>
                <option value="power">Power Market & Consumption</option>
                <option value="energy_mix">Energy Mix (Detailed)</option>
                <option value="co2">CO2 Trajectory & Targets</option>
                <option value="investment">Investment Plan (Process/Tech Mapping)</option>
                <option value="transition">Ecological Transition Costs</option>
            </select>
        </div>

        <!-- Energy Mix Sub-controls (Hidden initially) -->
        <div id="energyMixControls" class="glass-card rounded-xl shadow-lg p-5 hidden col-span-1 md:col-span-3 grid-cols-1 md:grid-cols-2 gap-4">
            <div>
                <label class="block text-sm font-medium text-slate-300 mb-2">Resource Category</label>
                <select id="mixCategorySelect" class="w-full bg-slate-700 text-white rounded-lg p-2.5 border border-slate-600 focus:ring-purple-500 focus:border-purple-500 outline-none transition">
                    <option value="Electricity">Electricity</option>
                    <option value="Hydrogen">Hydrogen</option>
                </select>
            </div>
            <div>
                <label class="block text-sm font-medium text-slate-300 mb-2">Flow Direction</label>
                <select id="mixDirectionSelect" class="w-full bg-slate-700 text-white rounded-lg p-2.5 border border-slate-600 focus:ring-purple-500 focus:border-purple-500 outline-none transition">
                    <option value="production">Production</option>
                    <option value="consumption">Consumption</option>
                </select>
            </div>
        </div>

        <!-- Methodology Box -->
        <div class="glass-card rounded-xl shadow-lg p-5 flex flex-col justify-center col-span-1 md:col-span-3">
            <h3 class="text-sm font-bold text-emerald-400 mb-2 flex items-center">
                <span class="mr-2">🧠</span> How this graph is constructed
            </h3>
            <p id="methodologyText" class="text-sm text-slate-300">
                Methodology description will appear here.
            </p>
        </div>
    </div>

    <!-- Graph Container -->
    <div class="glass-card rounded-xl shadow-xl p-6 relative">
        <div class="flex justify-between items-center mb-4">
            <h2 id="graphTitle" class="text-xl font-semibold text-white">Interactive Graph</h2>
            <button onclick="downloadGraph()" class="bg-slate-700 hover:bg-slate-600 text-white py-1.5 px-4 rounded-lg text-sm border border-slate-500 transition shadow hover:shadow-md flex items-center gap-2">
                📸 Export PNG
            </button>
        </div>
        
        <div id="plotlyChart" class="w-full h-[500px]"></div>
    </div>

    <!-- Data Injection -->
    <script>
        // Data injected by Python script
        const dashboardData = {{INJECTED_JSON_HERE}};

        // Methodology mapping
        const methodologies = {
            'financial': "Global financial balance. CAPEX/OPEX (-), Aids/Revenues (+). Highlights the Net Present Value Waterfall.",
            'hydrogen': "Synchronization between H2 production, consumption, and storage buffering across the simulation timeframe.",
            'power': "Highlights grid power consumption behavior relative to spot market electricity prices.",
            'energy_mix': "Detailed Energy Mix showing individual technological processes producing or consuming the selected category of resource.",
            'co2': "CO2 trajectory vs Objectives (indicated by crosses), including net emission reductions through DAC (Direct Air Capture) and Voluntary Credits (plotted as negative values).",
            'investment': "Interactive visual mapping of CAPEX allocations, detailing exactly which existing process has received which decarbonization technology across the timeframe.",
            'transition': "Plots the cumulative net cost of the ecological transition alongside the volume of avoided emissions (relative to baseline) for each specific year."
        };

        const scenarioSelect = document.getElementById('scenarioSelect');
        const graphSelect = document.getElementById('graphSelect');
        const mixCategorySelect = document.getElementById('mixCategorySelect');
        const mixDirectionSelect = document.getElementById('mixDirectionSelect');
        const energyMixControls = document.getElementById('energyMixControls');
        const methodologyText = document.getElementById('methodologyText');
        const plotlyChart = document.getElementById('plotlyChart');
        const graphTitle = document.getElementById('graphTitle');

        // Initialize Options
        function init() {
            const scenarios = Object.keys(dashboardData.data || {});
            if (scenarios.length === 0) {
                scenarioSelect.innerHTML = '<option value="">No data available</option>';
                return;
            }
            
            scenarios.forEach(sc => {
                const opt = document.createElement('option');
                opt.value = sc;
                opt.textContent = sc;
                scenarioSelect.appendChild(opt);
            });

            // Listeners
            scenarioSelect.addEventListener('change', updateDashboard);
            graphSelect.addEventListener('change', updateDashboard);
            mixCategorySelect.addEventListener('change', updateDashboard);
            mixDirectionSelect.addEventListener('change', updateDashboard);

            // First render
            updateDashboard();
        }

        // Main Render Logic
        function updateDashboard() {
            const scenario = scenarioSelect.value;
            const graphType = graphSelect.value;
            
            if (!scenario || !dashboardData.data[scenario]) return;
            
            const sData = dashboardData.data[scenario];
            methodologyText.textContent = methodologies[graphType];

            // Toggle visibility of Energy Mix sub-controls
            if (graphType === 'energy_mix') {
                energyMixControls.classList.remove('hidden');
                energyMixControls.classList.add('grid');
            } else {
                energyMixControls.classList.add('hidden');
                energyMixControls.classList.remove('grid');
            }

            // Common Layout settings aligning with the premium dark theme of reporting.py
            const darkBg = '#0D0D14';
            const gridColor = '#2B2B36';
            const textColor = '#FFFFFF';
            const spineColor = '#EEEEEE';

            const commonLayout = {
                paper_bgcolor: darkBg,
                plot_bgcolor: darkBg,
                font: { color: textColor, family: 'Arial, sans-serif' },
                margin: { t: 50, b: 60, l: 80, r: 80 },
                xaxis: { 
                    gridcolor: gridColor, 
                    gridwidth: 1, 
                    zerolinecolor: gridColor,
                    showline: true,
                    linecolor: spineColor,
                    linewidth: 1,
                    tickfont: { color: textColor },
                    tickangle: 0,
                    dtick: 5,
                    tick0: 2025, // Make sure sequence starts cleanly
                    range: sData.hydrogen && sData.hydrogen.time ? 
                           [sData.hydrogen.time[0], sData.hydrogen.time[sData.hydrogen.time.length - 1]] : undefined
                },
                yaxis: { 
                    gridcolor: gridColor, 
                    gridwidth: 1, 
                    zerolinecolor: gridColor,
                    showline: true,
                    linecolor: spineColor,
                    linewidth: 1,
                    tickfont: { color: textColor }
                },
                legend: { 
                    orientation: 'h', 
                    y: -0.2, 
                    bgcolor: darkBg, // using exact same dark background to avoid box effect
                    bordercolor: 'rgba(0,0,0,0)', // Remove border
                    borderwidth: 0,
                    font: { color: textColor }
                },
                hovermode: 'x unified'
            };

            if (graphType === 'financial') {
                graphTitle.textContent = "Financial Summary & NPV";
                
                const trace = {
                    type: "waterfall",
                    orientation: "v",
                    measure: ["relative", "relative", "relative", "relative", "total"],
                    x: ["CAPEX", "OPEX", "Public Aids", "Revenues", "NPV"],
                    textposition: "outside",
                    textfont: { color: textColor, size: 13, weight: "bold" },
                    y: [
                        -sData.financial.capex, 
                        -sData.financial.opex, 
                        sData.financial.aids, 
                        sData.financial.revenues, 
                        0 // Total computed automatically
                    ],
                    text: [
                        -sData.financial.capex, 
                        -sData.financial.opex, 
                        sData.financial.aids, 
                        sData.financial.revenues, 
                        sData.financial.npv
                    ].map(v => v.toFixed(2) + ' M€'),
                    connector: { line: { color: gridColor, width: 2 } },
                    decreasing: { marker: { color: "#E74C3C" } }, // red from reporting.py
                    increasing: { marker: { color: "#27AE60" } }, // green from reporting.py
                    totals: { marker: { color: "#3498DB" } } // blue from reporting.py
                };

                Plotly.newPlot('plotlyChart', [trace], {
                    ...commonLayout,
                    yaxis: { ...commonLayout.yaxis, title: 'Amount (M€)' }
                }, { responsive: true });

            } else if (graphType === 'hydrogen') {
                graphTitle.textContent = "Hydrogen Mass Balance";
                const timeStr = sData.hydrogen.time;
                
                const traceProd = {
                    x: timeStr, y: sData.hydrogen.production,
                    type: 'scatter', mode: 'lines',
                    name: 'H2 Produced', line: { color: '#00FF7F', width: 3 } // neon green from reporting.py
                };
                
                const traceCons = {
                    x: timeStr, y: sData.hydrogen.consumption,
                    type: 'scatter', mode: 'lines',
                    name: 'H2 Consumed', line: { color: '#FF007F', dash: 'dash', width: 3 } // neon pink/red
                };

                const traceStorage = {
                    x: timeStr, y: sData.hydrogen.storage_level,
                    type: 'scatter', mode: 'none', fill: 'tozeroy',
                    name: 'Storage Level', fillcolor: 'rgba(0, 229, 255, 0.25)', // neon cyan with opacity
                    yaxis: 'y2'
                };

                Plotly.newPlot('plotlyChart', [traceStorage, traceProd, traceCons], {
                    ...commonLayout,
                    yaxis: { ...commonLayout.yaxis, title: 'H2 Flow (kg/h)' },
                    yaxis2: {
                        title: 'Storage Level (kg)',
                        titlefont: { color: '#00E5FF', weight: 'bold' },
                        tickfont: { color: '#00E5FF' },
                        overlaying: 'y', side: 'right',
                        gridcolor: 'rgba(0,0,0,0)',
                        showline: true, linecolor: spineColor
                    }
                }, { responsive: true });

            } else if (graphType === 'power') {
                graphTitle.textContent = "Power Market & Consumption";
                const timeStr = sData.power.time;

                const tracePower = {
                    x: timeStr, y: sData.power.consumption,
                    type: 'bar', name: 'Power Consumed',
                    marker: { color: '#9D00FF', opacity: 0.8 }, // neon purple from reporting.py
                    marker_line_width: 0 
                };

                const tracePrice = {
                    x: timeStr, y: sData.power.spot_price,
                    type: 'scatter', mode: 'lines', name: 'Spot Price',
                    line: { color: '#FFD700', width: 3 }, // neon yellow
                    yaxis: 'y2'
                };

                Plotly.newPlot('plotlyChart', [tracePower, tracePrice], {
                    ...commonLayout,
                    yaxis: { ...commonLayout.yaxis, title: 'Power Consumed (MW)' },
                    yaxis2: {
                        title: 'Spot Price (€/MWh)',
                        titlefont: { color: '#FFD700', weight: 'bold' },
                        tickfont: { color: '#FFD700' },
                        overlaying: 'y', side: 'right',
                        gridcolor: 'rgba(0,0,0,0)',
                        showline: true, linecolor: spineColor
                    },
                    barmode: 'stack'
                }, { responsive: true });

            } else if (graphType === 'energy_mix') {
                const category = mixCategorySelect.value;
                const direction = mixDirectionSelect.value; // 'production' or 'consumption'
                
                const properDirectionLabel = direction.charAt(0).toUpperCase() + direction.slice(1);
                graphTitle.textContent = `Energy Mix: ${properDirectionLabel} of ${category}`;

                const mixData = sData.energy_mix?.[category]?.[direction];
                const timeStr = sData.energy_mix?.time || sData.hydrogen.time;
                
                if (!mixData || Object.keys(mixData).length === 0) {
                    Plotly.newPlot('plotlyChart', [], {
                        ...commonLayout,
                        annotations: [{
                            text: `No ${properDirectionLabel} data available for ${category}`,
                            xref: 'paper', yref: 'paper',
                            x: 0.5, y: 0.5, showarrow: false,
                            font: { size: 16, color: '#f87171' }
                        }]
                    }, { responsive: true });
                    return;
                }

                // Premium stacked area chart for Energy Mix
                const traces = [];
                const colorPalette = ['#e74c3c', '#9b59b6', '#f39c12', '#1abc9c', '#34495e', '#d35400', '#2ecc71', '#3498DB'];
                let idx = 0;
                
                for (const [techName, values]  of Object.entries(mixData)) {
                    traces.push({
                        x: timeStr,
                        y: values,
                        name: techName,
                        type: 'scatter',
                        mode: 'none', // fills don't need lines by default in stacked layout, but we'll adapt
                        fill: 'tonexty',
                        stackgroup: 'one',
                        fillcolor: colorPalette[idx % colorPalette.length],
                        line: { color: colorPalette[idx % colorPalette.length], width: 1 }
                    });
                    idx++;
                }

                Plotly.newPlot('plotlyChart', Object.values(traces), {
                    ...commonLayout,
                    yaxis: { ...commonLayout.yaxis, title: `Quantity (MWh)` },
                    hovermode: 'x unified'
                }, { responsive: true });
            } else if (graphType === 'co2') {
                graphTitle.textContent = "CO2 Trajectory & Emission Targets";
                const timeStr = sData.co2.time;

                // Color palette matching the design
                const colorPaletteC02 = {
                    'Direct Emissions': '#1B4965',        // Dark Blue
                    'Indirect Emissions': '#A9A9A9',      // Grey (dotted line)
                    'Total Emissions (Net Direct + Indirect)': '#BEBEBE', // Light Grey
                    'Net Direct Emissions': '#1E3A8A',    // Deep Blue
                    'Free Quotas (Direct)': '#81C784',    // Light Green
                    'Taxed Emissions (Surface)': '#90A4AE', // Blue-Grey
                    'DAC Captured (ktCO2)': '#42A5F5',    // Bright Blue
                    'Voluntary Credits (ktCO2)': '#66BB6A'  // Green
                };

                // Stacked Areas - building from bottom to top
                const traceDirectEmissions = {
                    x: timeStr, y: sData.co2.direct_emissions,
                    type: 'scatter', mode: 'none',
                    name: 'Direct Emissions',
                    fill: 'tozeroy', fillcolor: colorPaletteC02['Direct Emissions'],
                    stackgroup: 'positive', legendgroup: 'positive', showlegend: true
                };

                const traceIndirectEmissions = {
                    x: timeStr, y: sData.co2.indirect_emissions,
                    type: 'scatter', mode: 'none',
                    name: 'Indirect Emissions',
                    fill: 'tonexty', fillcolor: colorPaletteC02['Indirect Emissions'],
                    stackgroup: 'positive', legendgroup: 'positive', showlegend: true
                };

                const traceTotalEmissions = {
                    x: timeStr, y: sData.co2.total_emissions,
                    type: 'scatter', mode: 'none',
                    name: 'Total Emissions (Net Direct + Indirect)',
                    fill: 'tonexty', fillcolor: colorPaletteC02['Total Emissions (Net Direct + Indirect)'],
                    stackgroup: 'positive', legendgroup: 'positive', showlegend: true
                };

                const traceNetDirectEmissions = {
                    x: timeStr, y: sData.co2.net_direct_emissions,
                    type: 'scatter', mode: 'none',
                    name: 'Net Direct Emissions',
                    fill: 'tonexty', fillcolor: colorPaletteC02['Net Direct Emissions'],
                    stackgroup: 'positive', legendgroup: 'positive', showlegend: true
                };

                const traceFreeQuotasDirect = {
                    x: timeStr, y: sData.co2.free_quotas_direct,
                    type: 'scatter', mode: 'none',
                    name: 'Free Quotas (Direct)',
                    fill: 'tonexty', fillcolor: colorPaletteC02['Free Quotas (Direct)'],
                    stackgroup: 'positive', legendgroup: 'positive', showlegend: true
                };

                const traceTaxedEmissions = {
                    x: timeStr, y: sData.co2.taxed_emissions,
                    type: 'scatter', mode: 'none',
                    name: 'Taxed Emissions (Surface)',
                    fill: 'tonexty', fillcolor: colorPaletteC02['Taxed Emissions (Surface)'],
                    stackgroup: 'positive', legendgroup: 'positive', showlegend: true
                };

                // Negative areas for reductions (DAC and Credits)
                const traceDAC = {
                    x: timeStr, y: sData.co2.dac.map(v => -v),
                    type: 'scatter', mode: 'none',
                    name: 'DAC Captured (ktCO2)',
                    fill: 'tonexty', fillcolor: colorPaletteC02['DAC Captured (ktCO2)'],
                    stackgroup: 'negative', legendgroup: 'negative', showlegend: true
                };

                const traceCredits = {
                    x: timeStr, y: sData.co2.credits.map(v => -v),
                    type: 'scatter', mode: 'none',
                    name: 'Voluntary Credits (ktCO2)',
                    fill: 'tonexty', fillcolor: colorPaletteC02['Voluntary Credits (ktCO2)'],
                    stackgroup: 'negative', legendgroup: 'negative', showlegend: true
                };

                // Target markers with distinct colors per target
                const traceTargets = sData.co2.target_names.map((name, i) => {
                    const color = name.includes('French') ? '#E74C3C' : '#9b59b6'; // Red for French, Purple for EU/Others
                    return {
                        x: [sData.co2.target_time[i]], 
                        y: [sData.co2.target_values[i]],
                        text: [name],
                        type: 'scatter', 
                        mode: 'markers',
                        name: name,
                        marker: { symbol: 'x', size: 16, color: color, line: {width: 4, color: color} },
                        cliponaxis: false,
                        hovertemplate: '<b>%{text}</b><br>Target: %{y} ktCO2<extra></extra>',
                        showlegend: true
                    };
                });

                Plotly.newPlot('plotlyChart', 
                    [traceDirectEmissions, traceIndirectEmissions, traceTotalEmissions, 
                     traceNetDirectEmissions, traceFreeQuotasDirect, traceTaxedEmissions,
                     traceDAC, traceCredits, ...traceTargets], {
                    ...commonLayout,
                    yaxis: { 
                        ...commonLayout.yaxis, 
                        title: 'Emissions / Capture (ktCO2)',
                        zeroline: true, zerolinecolor: spineColor 
                    },
                    legend: { 
                        ...commonLayout.legend, 
                        y: -0.25, 
                        x: 0,
                        orientation: 'h'
                    }
                }, { responsive: true, autosize: true });
                
            } else if (graphType === 'investment') {
                graphTitle.textContent = "Investment Plan: Tech applied to Process";
                const invData = sData.investment;
                const timeStr = invData.time || sData.hydrogen.time;
                
                if (!invData || Object.keys(invData.capex_by_process_tech || {}).length === 0) {
                    Plotly.newPlot('plotlyChart', [], {
                        ...commonLayout,
                        annotations: [{
                            text: 'No investment data found for this scenario.',
                            xref: 'paper', yref: 'paper', x: 0.5, y: 0.5, showarrow: false,
                            font: { size: 16, color: '#f87171' }
                        }]
                    }, { responsive: true });
                    return;
                }

                const traces = [];
                const colorPalette = ['#e74c3c', '#3498DB', '#f39c12', '#1abc9c', '#9b59b6', '#d35400', '#27AE60', '#F1C40F', '#34495e'];
                let idx = 0;

                // Create a stacked bar trace for each "Process - Technology" pairing
                for (const [processTechCombo, values] of Object.entries(invData.capex_by_process_tech)) {
                    traces.push({
                        x: timeStr,
                        y: values,
                        name: processTechCombo,
                        type: 'bar',
                        marker: { color: colorPalette[idx % colorPalette.length], opacity: 0.9 },
                        text: values.map(v => v > 0 ? `${v.toFixed(1)} M€` : ''),
                        textposition: 'inside',
                        hovertemplate: `<b>${processTechCombo}</b><br>CAPEX: %{y} M€<extra></extra>`
                    });
                    idx++;
                }

                Plotly.newPlot('plotlyChart', traces, {
                    ...commonLayout,
                    barmode: 'stack',
                    yaxis: { 
                        ...commonLayout.yaxis, 
                        title: 'CAPEX (M€)' 
                    },
                    hovermode: 'x'
                }, { responsive: true });

            } else if (graphType === 'transition') {
                graphTitle.textContent = "Ecological Transition Costs vs Avoided CO2";
                const transData = sData.transition;
                const timeStr = transData.time || sData.hydrogen.time;

                if (!transData || !transData.cumulative_net_cost) {
                    Plotly.newPlot('plotlyChart', [], {
                        ...commonLayout,
                        annotations: [{
                            text: 'No transition cost data available.',
                            xref: 'paper', yref: 'paper', x: 0.5, y: 0.5, showarrow: false,
                            font: { size: 16, color: '#f87171' }
                        }]
                    }, { responsive: true });
                    return;
                }

                // Color palette for cost breakdown
                const colorPaletteTransition = {
                    'Self-funded CAPEX': '#0D47A1',          // Deep Blue
                    'Bank Loan Service': '#1565C0',           // Blue
                    'Tech & DAC OPEX': '#1976D2',             // Medium Blue
                    'Resource Mix Change': '#1E88E5',         // Light Blue
                    'Voluntary Carbon Credits': '#64B5F6'     // Very Light Blue
                };

                // Stacked negative areas for costs
                const traceSelfFundedCapex = {
                    x: timeStr, y: transData.self_funded_capex.map(v => -v),
                    type: 'scatter', mode: 'none',
                    name: 'Self-funded CAPEX',
                    fill: 'tozeroy', fillcolor: colorPaletteTransition['Self-funded CAPEX'],
                    stackgroup: 'costs', legendgroup: 'costs', showlegend: true
                };

                const traceBankLoanService = {
                    x: timeStr, y: transData.bank_loan_service.map(v => -v),
                    type: 'scatter', mode: 'none',
                    name: 'Bank Loan Service',
                    fill: 'tonexty', fillcolor: colorPaletteTransition['Bank Loan Service'],
                    stackgroup: 'costs', legendgroup: 'costs', showlegend: true
                };

                const traceTechDacOpex = {
                    x: timeStr, y: transData.tech_dac_opex.map(v => -v),
                    type: 'scatter', mode: 'none',
                    name: 'Tech & DAC OPEX',
                    fill: 'tonexty', fillcolor: colorPaletteTransition['Tech & DAC OPEX'],
                    stackgroup: 'costs', legendgroup: 'costs', showlegend: true
                };

                const traceResourceMixChange = {
                    x: timeStr, y: transData.resource_mix_change.map(v => -v),
                    type: 'scatter', mode: 'none',
                    name: 'Resource Mix Change',
                    fill: 'tonexty', fillcolor: colorPaletteTransition['Resource Mix Change'],
                    stackgroup: 'costs', legendgroup: 'costs', showlegend: true
                };

                const traceVoluntaryCarbonCredits = {
                    x: timeStr, y: transData.voluntary_carbon_credits.map(v => -v),
                    type: 'scatter', mode: 'none',
                    name: 'Voluntary Carbon Credits',
                    fill: 'tonexty', fillcolor: colorPaletteTransition['Voluntary Carbon Credits'],
                    stackgroup: 'costs', legendgroup: 'costs', showlegend: true
                };

                // Line for Cumulative Transition Cost (top, in red)
                const traceCumCost = {
                    x: timeStr, y: transData.cumulative_net_cost,
                    type: 'scatter', mode: 'lines+markers', name: 'Net Transition Cost (Cumulative) (M€)',
                    line: { color: '#E74C3C', width: 4 }, // Red
                    marker: { size: 8, color: '#E74C3C', line: { color: darkBg, width: 2 } },
                    yaxis: 'y',
                    hovertemplate: '<b>%{x}</b><br>Cumul Cost: %{y} M€<extra></extra>'
                };

                Plotly.newPlot('plotlyChart', 
                    [traceSelfFundedCapex, traceBankLoanService, traceTechDacOpex, 
                     traceResourceMixChange, traceVoluntaryCarbonCredits, traceCumCost], {
                    ...commonLayout,
                    yaxis: { 
                        ...commonLayout.yaxis, 
                        title: 'Annual delta (Area) (M€)',
                        zeroline: true, zerolinecolor: spineColor 
                    },
                    yaxis2: {
                        title: 'Cumulative Net Cost (Line) (M€)',
                        titlefont: { color: '#E74C3C', weight: 'bold' },
                        tickfont: { color: '#333333' },
                        overlaying: 'y', side: 'right',
                        gridcolor: 'rgba(0,0,0,0)',
                        showline: true, linecolor: spineColor
                    },
                    legend: { 
                        ...commonLayout.legend, 
                        y: -0.25, 
                        x: 0,
                        orientation: 'h'
                    },
                    hovermode: 'x unified'
                }, { responsive: true, autosize: true });
            }
        }

        // Export image
        function downloadGraph() {
            Plotly.downloadImage('plotlyChart', {format: 'png', width: 1200, height: 800, filename: 'PathWay_Graph'});
        }

        // Boot
        init();
    </script>
</body>
</html>
"""

# --- PYTHON EXTRACTOR ---

def parse_excel_scenarios(data_dir: str) -> dict:
    """
    Parses Excel files in the given directory to build the master JSON dictionary.
    """
    master_data = {"data": {}}
    excel_files = glob.glob(os.path.join(data_dir, "*.xlsx"))

    if not excel_files:
        print(f"Warning: No Excel files found in {data_dir}. Generating mock data for demonstration.")
        return generate_mock_data()

    for file_path in excel_files:
        filename = os.path.basename(file_path)
        scenario_name = filename.replace(".xlsx", "").replace("Results_", "")
        print(f"Processing Scenario: {scenario_name}")
        
        try:
            # We attempt to load specific sheets. 
            # Replace these sheet names / columns with your actual mappings.
            # 1. Financial Data
            df_invest = pd.read_excel(file_path, sheet_name='Bilan_Investissement')
            df_opex = pd.read_excel(file_path, sheet_name='Bilan_Opex')
            df_synth = pd.read_excel(file_path, sheet_name='Synthèse')

            # Naive extraction - adjust logic based on exact Excel structure
            capex = abs(df_invest['Value'].sum()) if 'Value' in df_invest.columns else 100.0
            opex = abs(df_opex['Value'].sum()) if 'Value' in df_opex.columns else 20.0
            aids = df_synth.loc[df_synth['Metric'] == 'Aids', 'Value'].values[0] if 'Metric' in df_synth.columns else 15.0
            revenues = df_synth.loc[df_synth['Metric'] == 'Revenues', 'Value'].values[0] if 'Metric' in df_synth.columns else 150.0
            npv = revenues + aids - capex - opex

            # 2. Hydrogen Flow
            df_h2 = pd.read_excel(file_path, sheet_name='Bilan_H2')
            time_h2 = df_h2['Time'].tolist() if 'Time' in df_h2.columns else list(range(len(df_h2)))
            prod_h2 = df_h2['Production'].tolist() if 'Production' in df_h2.columns else [0] * len(df_h2)
            cons_h2 = df_h2['Consumption'].tolist() if 'Consumption' in df_h2.columns else [0] * len(df_h2)
            stor_h2 = df_h2['Storage_Level'].tolist() if 'Storage_Level' in df_h2.columns else [0] * len(df_h2)

            # 3. Power Consumption
            # Assuming power data might be in the same or separate time-series sheet
            df_power = pd.read_excel(file_path, sheet_name='Bilan_Elec')
            time_pwr = df_power['Time'].tolist() if 'Time' in df_power.columns else list(range(len(df_power)))
            cons_pwr = df_power['Grid_Consumption'].tolist() if 'Grid_Consumption' in df_power.columns else [0] * len(df_power)
            spot_pwr = df_power['Spot_Price'].tolist() if 'Spot_Price' in df_power.columns else [0] * len(df_power)

            master_data["data"][scenario_name] = {
                "financial": {
                    "capex": float(capex),
                    "opex": float(opex),
                    "aids": float(aids),
                    "revenues": float(revenues),
                    "npv": float(npv)
                },
                "hydrogen": {
                    "time": time_h2,
                    "production": prod_h2,
                    "consumption": cons_h2,
                    "storage_level": stor_h2
                },
                "power": {
                    "time": time_pwr,
                    "consumption": cons_pwr,
                    "spot_price": spot_pwr
                }
            }

        except Exception as e:
            print(f"Error reading {filename}: {e}. (Falling back to empty/mock data for this scenario)")
            # You can inject a fallback mock here if needed.

    return master_data

def generate_mock_data() -> dict:
    """Generate mock data if no Excel files are present for testing out of the box."""
    import random
    times = [year for year in range(2025, 2051)] # 26 years from 2025 to 2050
    return {
        "data": {
            "Baseline": {
                "financial": { "capex": 120.5, "opex": 45.2, "aids": 10.0, "revenues": 200.0, "npv": 44.3 },
                "hydrogen": {
                    "time": times,
                    "production": [random.randint(50, 100) for _ in times],
                    "consumption": [random.randint(40, 90) for _ in times],
                    "storage_level": [random.randint(200, 500) for _ in times]
                },
                "power": {
                    "time": times,
                    "consumption": [random.randint(10, 50) for _ in times],
                    "spot_price": [random.randint(20, 150) for _ in times]
                },
                "energy_mix": {
                    "time": times,
                    "Electricity": {
                        "production": {
                            "Solar PV": [random.randint(5, 20) for _ in times],
                            "Wind Farm": [random.randint(10, 30) for _ in times]
                        },
                        "consumption": {
                            "Electrolyzer": [random.randint(10, 40) for _ in times],
                            "Base Process": [5 for _ in times]
                        }
                    },
                    "Hydrogen": {
                        "production": {
                            "Electrolyzer ALK": [random.randint(20, 50) for _ in times],
                            "SMR Plant": [random.randint(10, 20) for _ in times]
                        },
                        "consumption": {
                            "Steel Process Unit": [random.randint(25, 60) for _ in times]
                        }
                    }
                },
                "co2": {
                    "time": times,
                    "direct_emissions": [3200, 3100, 2900, 2800, 2600, 2500, 2300, 2200, 2000, 1900, 1850] + [1850]*15,
                    "indirect_emissions": [100, 95, 90, 85, 80, 75, 70, 65, 60, 50, 40] + [40]*15,
                    "total_emissions": [3300, 3195, 2990, 2885, 2680, 2575, 2370, 2265, 2060, 1950, 1890] + [1890]*15,
                    "net_direct_emissions": [2800, 2700, 2400, 2200, 1800, 1600, 1200, 1000, 600, 400, 300] + [200]*15,
                    "free_quotas_direct": [400, 400, 400, 500, 600, 700, 800, 900, 1000, 1100, 1100] + [1100]*15,
                    "taxed_emissions": [0, 50, 200, 250, 400, 450, 600, 650, 800, 850, 900] + [900]*15,
                    "dac": [0, 20, 150, 200, 300, 350, 500, 550, 700, 750, 800] + [900]*15,
                    "credits": [0, 0, 50, 100, 250, 300, 450, 500, 650, 700, 750] + [750]*15,
                    "target_time": [2030, 2040, 2050],
                    "target_values": [2500, 1500, 500],
                    "target_names": ["Milestone 2030", "Milestone 2040", "Net Zero Objective"]
                },
                "investment": {
                    "time": times,
                    "capex_by_process_tech": {
                        "Main Furnace (Process) - Electrification (Tech)": [10, 0, 0, 50, 0, 0] + [0]*20,
                        "Logistics (Process) - H2 Trucks (Tech)": [0, 0, 15, 0, 0, 30] + [0]*20
                    }
                },
                "transition": {
                    "time": times,
                    "cumulative_net_cost": [0, 10, 35, 70, 130, 185, 250, 330, 420, 520, 600] + [650 + 20*i for i in range(15)],
                    "self_funded_capex": [0, 5, 10, 15, 25, 30, 40, 50, 60, 70, 80] + [85]*15,
                    "bank_loan_service": [0, 3, 8, 15, 25, 35, 45, 55, 65, 75, 80] + [80]*15,
                    "tech_dac_opex": [0, 2, 5, 10, 15, 20, 30, 40, 50, 60, 70] + [70]*15,
                    "resource_mix_change": [0, 0, 5, 15, 30, 40, 50, 60, 70, 80, 90] + [90]*15,
                    "voluntary_carbon_credits": [0, 0, 2, 8, 15, 20, 25, 30, 35, 40, 45] + [45]*15,
                    "annual_avoided_co2": [0, 5, 20, 35, 60, 65, 75, 80, 85, 95, 100] + [120]*15
                }
            },
            "Green_Scenario": {
                "financial": { "capex": 200.0, "opex": 30.0, "aids": 50.0, "revenues": 250.0, "npv": 70.0 },
                "hydrogen": {
                    "time": times,
                    "production": [random.randint(100, 150) for _ in times],
                    "consumption": [random.randint(80, 120) for _ in times],
                    "storage_level": [random.randint(300, 700) for _ in times]
                },
                "power": {
                    "time": times,
                    "consumption": [random.randint(5, 30) for _ in times],
                    "spot_price": [random.randint(20, 150) for _ in times]
                },
                "energy_mix": {
                    "time": times,
                    "Electricity": {
                        "production": {
                            "Wind Farm": [random.randint(40, 80) for _ in times],
                            "Solar Farm": [random.randint(20, 60) for _ in times]
                        },
                        "consumption": {
                            "Electrolyzer PEM": [random.randint(50, 100) for _ in times]
                        }
                    },
                    "Hydrogen": {
                        "production": {
                            "Electrolyzer PEM": [random.randint(80, 140) for _ in times]
                        },
                        "consumption": {
                            "Chemical Plant": [100 for _ in times]
                        }
                    }
                },
                "co2": {
                    "time": times,
                    "direct_emissions": [3200, 2900, 2600, 2300, 2000, 1700, 1400, 1100, 800, 500, 200] + [0]*15,
                    "indirect_emissions": [100, 95, 85, 75, 65, 55, 45, 35, 25, 15, 5] + [0]*15,
                    "total_emissions": [3300, 2995, 2685, 2375, 2065, 1755, 1445, 1135, 825, 515, 205] + [0]*15,
                    "net_direct_emissions": [2200, 1900, 1600, 1300, 1000, 700, 400, 150, 0, 0, 0] + [0]*15,
                    "free_quotas_direct": [800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800] + [800]*15,
                    "taxed_emissions": [200, 300, 400, 500, 600, 700, 800, 900, 1000, 1000, 1000] + [1000]*15,
                    "dac": [0, 100, 250, 400, 500, 600, 700, 800, 900, 1000, 1200] + [1500]*15,
                    "credits": [0, 50, 150, 350, 500, 650, 800, 950, 1100, 1200, 1300] + [1500]*15,
                    "target_time": [2030, 2040, 2050],
                    "target_values": [2000, 500, 0],
                    "target_names": ["French Target", "EU Target", "Net Zero"]
                },
                "investment": {
                    "time": times,
                    "capex_by_process_tech": {
                        "Reactor 1 (Process) - Heat Pump (Tech)": [0, 45, 0, 0, 0, 10] + [0]*20,
                        "Power Generation (Process) - Solar Farm (Tech)": [120, 0, 0, 0, 0, 0] + [0]*20,
                        "Boiler System (Process) - BioGas Conversion (Tech)": [0, 0, 0, 0, 60, 0] + [0]*20
                    }
                },
                "transition": {
                    "time": times,
                    "cumulative_net_cost": [0, 120, 130, 130, 190, 195, 195, 195, 195, 195, 195] + [195]*15,
                    "self_funded_capex": [0, 60, 5, 0, 30, 5, 0, 0, 0, 0, 0] + [0]*15,
                    "bank_loan_service": [0, 30, 10, 5, 20, 10, 5, 0, 0, 0, 0] + [0]*15,
                    "tech_dac_opex": [0, 20, 8, 10, 15, 10, 8, 5, 3, 2, 1] + [0]*15,
                    "resource_mix_change": [0, 8, 3, 5, 15, 10, 5, 3, 2, 1, 0] + [0]*15,
                    "voluntary_carbon_credits": [0, 2, 4, 5, 10, 10, 8, 5, 3, 2, 1] + [0]*15,
                    "annual_avoided_co2": [0, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500] + [500]*15
                }
            }
        }
    }

def main():
    print("🚀 Starting Dashboard Generation...")
    
    # Define folder containing Excel Results
    RESULTS_DIR = "./artifacts/reports/"
    
    # 1. Parse Excel into Dictionary
    dashboard_dict = parse_excel_scenarios(RESULTS_DIR)
    
    # 2. Serialize to JSON
    json_string = json.dumps(dashboard_dict)
    
    # 3. Inject into HTML Template
    current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    html_output = HTML_TEMPLATE.replace('{{INJECTED_JSON_HERE}}', json_string)
    html_output = html_output.replace('{{GENERATION_DATE}}', current_date)
    
    # 4. Save Final HTML
    output_filename = "dashboard.html"
    with open(output_filename, "w", encoding="utf-8") as f:
        f.write(html_output)
        
    print(f"✅ Dashboard generated successfully: {output_filename}")

if __name__ == "__main__":
    main()
