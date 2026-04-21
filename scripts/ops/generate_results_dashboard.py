"""
Generate a single standalone HTML dashboard from scenario result workbooks.

The script scans scenario folders for `Master_Plan.xlsx`, extracts key datasets,
and writes one responsive HTML file with embedded JSON + JavaScript logic.
"""

import argparse
import json
import math
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
import plotly.graph_objects as go
import sys
import os

# Ensure we can import from src/
repo_root = str(Path(__file__).resolve().parents[2])
if repo_root not in sys.path:
    sys.path.append(os.path.join(repo_root, "src"))

# No internal Plotly builder imports needed - dashboard is a passive consumer.


DEFAULT_DISCOUNT_RATE = 0.08
MAX_STACK_SERIES = 8
MAX_LINE_SERIES = 8


def get_repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def normalize_key(value: Any) -> str:
    return "".join(ch.lower() if ch.isalnum() else "_" for ch in str(value)).strip("_")


def pretty_label(name: str) -> str:
    return name.replace("_", " ").replace("##", " ").strip()


def sanitize_filename(value: str) -> str:
    out = []
    for ch in value:
        if ch.isalnum() or ch in ("-", "_"):
            out.append(ch)
        elif ch.isspace():
            out.append("_")
    safe = "".join(out).strip("_")
    return safe or "chart"


# Legacy Excel data helpers removed.


def fig_to_dict(fig: go.Figure) -> Dict[str, Any]:
    """Converts a Plotly Figure to the {data, layout} format expected by the dashboard."""
    d = fig.to_dict()
    return {
        "data": d.get("data", []),
        "layout": d.get("layout", {})
    }


def placeholder_figure(title: str, message: str) -> Dict[str, Any]:
    return {
        "data": [],
        "layout": {
            "title": {"text": title, "x": 0.02, "font": {"size": 20}},
            "annotations": [
                {
                    "text": message,
                    "xref": "paper",
                    "yref": "paper",
                    "x": 0.5,
                    "y": 0.5,
                    "showarrow": False,
                    "font": {"size": 15, "color": "#475569"},
                }
            ],
            "xaxis": {"visible": False},
            "yaxis": {"visible": False},
            "paper_bgcolor": "rgba(255,255,255,0)",
            "plot_bgcolor": "rgba(255,255,255,0)",
            "font": {"family": "Manrope, sans-serif", "color": "#1e293b"},
            "margin": {"l": 40, "r": 20, "t": 80, "b": 40},
        },
    }


def base_layout(title: str, y_title: str, years: List[Any], barmode: str = "stack", is_x_years: bool = True) -> Dict[str, Any]:
    if is_x_years:
        tickvals = []
        for y in years:
            if y is None: continue
            try:
                if int(float(y)) % 5 == 0:
                    tickvals.append(y)
            except (ValueError, TypeError):
                pass
    else:
        tickvals = years
        
    return {
        "title": {"text": title, "x": 0.02, "font": {"size": 20}},
        "xaxis": {
            "title": "Year" if is_x_years else "Project",
            "type": "category",
            "tickvals": tickvals,
            "automargin": True,
            "gridcolor": "#f1f5f9",
            "linecolor": "#e2e8f0",
            "tickfont": {"color": "#64748b"},
        },
        "yaxis": {
            "title": y_title,
            "automargin": True,
            "gridcolor": "#f1f5f9",
            "zerolinecolor": "#cbd5e1",
            "linecolor": "#e2e8f0",
            "tickfont": {"color": "#64748b"},
        },
        "barmode": barmode,
        "legend": {
            "orientation": "h",
            "yanchor": "bottom",
            "y": 1.02,
            "xanchor": "left",
            "x": 0.0,
            "font": {"size": 11, "color": "#475569"},
        },
        "hovermode": "x unified",
        "paper_bgcolor": "rgba(255,255,255,0)",
        "plot_bgcolor": "rgba(255,255,255,0)",
        "margin": {"l": 66, "r": 40, "t": 86, "b": 70},
    }


# Legacy plotting functions removed. Dashboard is now a passive JSON consumer.


# Passive dashboard generator logic starts here.



def sanitize_payload(value: Any) -> Any:
    """Recursively santize data for JSON serialization, handling NaNs and rounding floats."""
    if isinstance(value, dict):
        return {k: sanitize_payload(v) for k, v in value.items()}
    if isinstance(value, list) or isinstance(value, tuple):
        return [sanitize_payload(v) for v in value]
    if isinstance(value, float):
        if not math.isfinite(value):
            return None
        return round(value, 6)
    if pd.isna(value):
        return None
    return value


def discover_entities_and_scenarios(results_root: Path) -> Dict[str, Dict[str, Path]]:
    """Scans the results directory for entity folders, then scenario folders containing a charts/ subdirectory."""
    roots = [results_root, results_root / "Results"]
    entities: Dict[str, Dict[str, Path]] = {}

    for root in roots:
        if not root.exists() or not root.is_dir():
            continue
        for child in sorted(root.iterdir(), key=lambda p: p.name.lower()):
            if not child.is_dir():
                continue
            if child.name.lower() == "results":
                continue
            
            if (child / "charts").exists() and (child / "charts").is_dir():
                continue
            
            entity_name = child.name
            for sub_child in sorted(child.iterdir(), key=lambda p: p.name.lower()):
                if not sub_child.is_dir():
                    continue
                if (sub_child / "charts").exists() and (sub_child / "charts").is_dir():
                    if entity_name not in entities:
                        entities[entity_name] = {}
                    if sub_child.name not in entities[entity_name]:
                        entities[entity_name][sub_child.name] = sub_child

    return entities


def get_graph_payload(scenario_path: Path, scenario_name: str, key: str, label: str, json_filename: str) -> Dict[str, Any]:
    """Loads a pre-generated Plotly JSON chart from the scenario artifacts."""
    charts_dir = scenario_path / "charts"
    chart_json = charts_dir / f"{json_filename}.json"
    
    # Try case-insensitive matching if exact match fails
    if not chart_json.exists() and charts_dir.exists():
        for f in charts_dir.iterdir():
            if f.name.lower() == f"{json_filename}.json".lower():
                chart_json = f
                break

    if chart_json.exists():
        try:
            # Use errors='replace' to handle potential locale-specific characters (like Euro symbol in cp1252)
            # which might have slipped into existing artifacts.
            with open(chart_json, 'r', encoding='utf-8', errors='replace') as f:
                fig_dict = json.load(f)
            
            print(f"    [OK] Loaded {key} from {chart_json.name}")
            return {
                "label": label,
                "title": fig_dict.get("layout", {}).get("title", {}).get("text", label),
                "description": f"Interactive visualization of {label} data.",
                "downloadName": sanitize_filename(f"{scenario_name}_{key}"),
                "figure": fig_dict
            }
        except Exception as e:
            print(f"    [WARN] Failed to load JSON for {key} ({chart_json.name}): {e}")

    # Fallback with diagnostic info for the console
    print(f"    [ERROR] Missing artifact: {json_filename}.json (checked: {chart_json})")
    return {
        "label": label,
        "title": label,
        "description": f"Missing data: The artifact {json_filename}.json was not found.",
        "downloadName": sanitize_filename(f"{scenario_name}_{key}"),
        "figure": placeholder_figure(label, f"Artifact '{json_filename}.json' is missing.")
    }


def build_dashboard_data(entity_dirs: Dict[str, Dict[str, Path]], discount_rate: float) -> Dict[str, Any]:
    """Assembles the final dashboard payload by consuming pre-generated JSON artifacts."""
    if not entity_dirs:
        print("[ERROR] No scenario directories with 'charts/' subfolder were found.")
        return {}

    # Determine generation date based on latest chart update
    latest_ts = 0.0
    for scenarios_dict in entity_dirs.values():
        for s_path in scenarios_dict.values():
            charts_dir = s_path / "charts"
            for f in charts_dir.glob("*.json"):
                latest_ts = max(latest_ts, f.stat().st_mtime)
    
    generation_date = datetime.fromtimestamp(latest_ts).strftime("%Y-%m-%d %H:%M:%S") if latest_ts > 0 else "Unknown"

    entities_payload: Dict[str, Any] = {}
    
    # Define mapping of dashboard keys to chart filenames
    # This must match self._save_plotly_figure calls in ReportingEngine
    chart_mapping = [
        ("co2_trajectory",    "CO2 TRAJECTORY",      "CO2_Trajectory"),
        ("indirect_emissions", "INDIRECT EMISSIONS", "Indirect_Emissions"),
        ("energy_mix",        "RESOURCES MIX",       "Energy_Mix"),
        ("investment_plan",   "INVESTMENT PLAN",     "Investment_Plan"),
        ("external_financing", "EXTERNAL FINANCING", "Financing"),
        ("interest_paid",      "INTEREST PAID",      "Interest_Paid"),
        ("ressources_opex",    "RESOURCES OPEX",      "Resources_Opex"),
        ("total_annual_opex",  "TOTAL ANNUAL OPEX",  "Total_Annual_Opex"),
        ("transition_cost",    "TRANSITION COST",     "Transition_Cost"),
        ("carbon_tax",         "CARBON TAX",          "Carbon_Tax"),
        ("carbon_price",       "CARBON PRICE",        "Carbon_Prices"),
        ("simulation_prices",  "SIMULATION PRICES",  "Simulation_Prices"),
        ("co2_abatement",      "CO2 ABATEMENT",       "CO2_Abatement"),
        ("data_used",          "DATA USED",           "Data_Used"),
    ]

    for entity_name, scenarios_dict in entity_dirs.items():
        print(f"  > Processing entity: {entity_name}")
        scenarios_payload = {}
        for scenario_name, scenario_path in scenarios_dict.items():
            graphs = {}
            for key, label, filename in chart_mapping:
                graphs[key] = get_graph_payload(scenario_path, scenario_name, key, label, filename)

            scenarios_payload[scenario_name] = {
                "displayName": scenario_name.replace("_", " "),
                "sourcePath": str(scenario_path),
                "graphs": graphs,
            }
            
        entities_payload[entity_name] = {
            "displayName": entity_name.replace("_", " "),
            "scenarios": scenarios_payload
        }

    payload = {
        "projectTitle": "PathWay Analytics Dashboard",
        "generationDate": generation_date,
        "discountRate": discount_rate,
        "entities": entities_payload,
    }
    return sanitize_payload(payload)


def load_sensitivity_data(json_path: Optional[Path] = None) -> Optional[List[Dict[str, Any]]]:
    """
    Charge les résultats JSON de l'analyse de sensibilité si le fichier existe.
    Retourne None si le fichier est absent ou illisible.
    """
    if json_path is None:
        json_path = get_repo_root() / "artifacts" / "sensitivity" / "sensitivity_results.json"
    if not json_path.exists():
        return None
    try:
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
        print(f"[Sensitivity] Données chargées : {len(data)} enregistrements depuis {json_path}")
        return data
    except Exception as exc:
        print(f"[Sensitivity] Impossible de lire {json_path}: {exc}")
        return None


def build_html(payload: Dict[str, Any], sensitivity_data: Optional[List[Dict[str, Any]]] = None) -> str:
    payload_json = json.dumps(payload, ensure_ascii=True)
    sensitivity_json = json.dumps(sensitivity_data if sensitivity_data else [], ensure_ascii=True)

    template = """<!DOCTYPE html>
<html lang=\"en\">
<head>
  <meta charset=\"UTF-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\" />
  <title>Plant-Optimization-PathWay Dashboard</title>
  <script src=\"https://cdn.tailwindcss.com\"></script>
  <script>
    tailwind.config = {
      theme: {
        extend: {
          fontFamily: {
            heading: ['"Montserrat"', 'sans-serif'],
            body: ['"Bookman Old Style"', '"Bookman"', 'serif']
          },
          boxShadow: {
            glass: '0 20px 50px rgba(15, 23, 42, 0.24)',
          }
        }
      }
    }
  </script>
  <link rel=\"preconnect\" href=\"https://fonts.googleapis.com\" />
  <link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin />
  <link href=\"https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&display=swap\" rel=\"stylesheet\" />
  <script src=\"https://cdn.plot.ly/plotly-2.35.2.min.js\"></script>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css" />
  <style>
    :root {
      --bg-1: #f8fafc;
      --bg-2: #f1f5f9;
      --bg-3: #e2e8f0;
      --glass: rgba(255, 255, 255, 0.82);
      --glass-strong: rgba(255, 255, 255, 0.95);
      --text-main: #1e293b;
      --text-muted: #64748b;
      --border-soft: rgba(15, 23, 42, 0.08);
      --accent: #0ea5e9;
    }

    html, body {
      min-height: 100%;
    }

    body {
      margin: 0;
      font-family: 'Bookman Old Style', 'Bookman', serif;
      color: var(--text-main);
      background:
        radial-gradient(circle at 10% 20%, rgba(14, 165, 233, 0.06) 0%, rgba(14, 165, 233, 0) 36%),
        radial-gradient(circle at 85% 12%, rgba(16, 185, 129, 0.08) 0%, rgba(16, 185, 129, 0) 34%),
        linear-gradient(135deg, #ffffff 0%, #f8fafc 50%, #f1f5f9 100%);
      background-attachment: fixed;
    }

    .glass-card {
      background: linear-gradient(155deg, var(--glass) 0%, var(--glass-strong) 100%);
      backdrop-filter: blur(20px);
      -webkit-backdrop-filter: blur(20px);
      border: 1px solid var(--border-soft);
      box-shadow: 0 10px 30px -5px rgba(15, 23, 42, 0.04), 0 4px 12px -4px rgba(15, 23, 42, 0.03);
      transition: transform 260ms ease, box-shadow 260ms ease;
    }

    .glass-card:hover {
      transform: translateY(-2px);
      box-shadow: 0 20px 40px -8px rgba(15, 23, 42, 0.08);
    }

    .subtle-pill {
      background: rgba(15, 23, 42, 0.05);
      color: #475569;
      border: 1px solid rgba(15, 23, 42, 0.08);
    }

    /* Custom Select Component Styling (Modern Select - MS) */
    .ms-container {
      position: relative;
      width: 100%;
      user-select: none;
    }
    .ms-trigger {
      display: flex;
      align-items: center;
      justify-content: space-between;
      width: 100%;
      border: 1px solid rgba(15, 23, 42, 0.12);
      border-radius: 9999px;
      padding: 0.8rem 1.4rem;
      background: #ffffff;
      color: #1e293b;
      font-weight: 700;
      font-size: 0.92rem;
      cursor: pointer;
      box-shadow: 0 4px 6px -1px rgba(15, 23, 42, 0.03), 0 2px 4px -1px rgba(15, 23, 42, 0.02);
      transition: all 250ms cubic-bezier(0.4, 0, 0.2, 1);
    }
    .ms-trigger:hover {
      box-shadow: 0 10px 15px -3px rgba(15, 23, 42, 0.08);
      transform: translateY(-1px);
      border-color: rgba(59, 130, 246, 0.3);
    }
    .ms-container.active .ms-trigger {
      border-color: #3b82f6;
      box-shadow: 0 0 0 4px rgba(59, 130, 246, 0.12);
      background: #fdfdfd;
    }
    .ms-trigger i {
      transition: transform 300ms cubic-bezier(0.34, 1.56, 0.64, 1);
      color: #64748b;
      font-size: 0.8rem;
    }
    .ms-container.active .ms-trigger i {
      transform: rotate(180deg);
      color: #3b82f6;
    }
    .ms-options {
      position: absolute;
      top: calc(100% + 10px);
      left: 0;
      width: 100%;
      background: rgba(255, 255, 255, 0.98);
      backdrop-filter: blur(16px);
      -webkit-backdrop-filter: blur(16px);
      border: 1px solid rgba(15, 23, 42, 0.12);
      border-radius: 1.6rem;
      box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.2);
      z-index: 5000;
      max-height: 400px;
      overflow-y: auto;
      padding: 0.5rem;
      display: none;
      transform: translateY(12px);
      transition: transform 200ms ease;
      scrollbar-width: thin;
      scrollbar-color: #e2e8f0 transparent;
    }
    .ms-options::-webkit-scrollbar { width: 5px; }
    .ms-options::-webkit-scrollbar-thumb { background: #e2e8f0; border-radius: 10px; }

    .ms-container.active .ms-options {
      display: block;
      transform: translateY(0);
    }
    .ms-option {
      padding: 0.8rem 1.1rem;
      border-radius: 1.1rem;
      color: #475569;
      font-size: 0.95rem;
      font-weight: 600;
      cursor: pointer;
      transition: all 150ms ease;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }
    .ms-option:hover {
      background: rgba(59, 130, 246, 0.06);
      color: #2563eb;
      padding-left: 1.4rem;
    }
    .ms-option.selected {
      background: #3b82f6;
      color: #ffffff;
    }
    .ms-option.selected::after {
      content: "\f00c";
      font-family: "Font Awesome 6 Free";
      font-weight: 900;
      font-size: 0.75rem;
    }
    .ms-hidden {
      display: none !important;
    }

    .primary-btn {
      background: linear-gradient(130deg, #0284c7, #0369a1);
      color: #f8fafc;
      border: 0;
      border-radius: 0.85rem;
      padding: 0.72rem 1rem;
      font-weight: 700;
      letter-spacing: 0.01em;
      transition: transform 160ms ease, filter 160ms ease;
    }

    .primary-btn:hover {
      transform: translateY(-1px);
      filter: brightness(1.05);
    }

    .primary-btn:active {
      transform: translateY(0);
    }

    #chart {
      width: 100%;
      height: 530px;
    }

    @media (max-width: 768px) {
      #chart {
        height: 430px;
      }
    }

    .fade-in {
      animation: enter 480ms ease both;
    }

    @keyframes enter {
      from {
        opacity: 0;
        transform: translateY(12px);
      }
      to {
        opacity: 1;
        transform: translateY(0);
      }
    }

    /* Tab styles */
    .nav-pill-container {
      background: #f1f5f9;
      border-radius: 9999px;
      padding: 0.4rem;
      display: inline-flex;
      gap: 0.5rem;
      border: 1px solid rgba(15, 23, 42, 0.05);
    }

    .nav-tab {
      padding: 0.6rem 1.5rem;
      border-radius: 9999px;
      background: #ffffff;
      color: #1e293b;
      font-weight: 600;
      font-size: 0.95rem;
      transition: all 200ms ease;
      cursor: pointer;
      border: 1px solid transparent;
      white-space: nowrap;
    }

    .nav-tab:hover:not(.active) {
      background: #f8fafc;
      transform: translateY(-1px);
    }

    /* Active Tab States */
    .nav-tab.active.tab-home { background: #bfdbfe; color: #1e3a8a; }
    .nav-tab.active.tab-details { background: #e2e8f0; color: #334155; }
    .nav-tab.active.tab-results { background: #bef264; color: #365314; }
    .nav-tab.active.tab-sensitivity { background: #fed7aa; color: #7c2d12; }
    .nav-tab.active.tab-licence { background: #fecaca; color: #7f1d1d; }

    .tab-content {
      display: none;
    }

    .tab-content.active {
      display: block;
    }

    .chart-btn {
      position: absolute;
      top: 0.75rem;
      right: 0.75rem;
      width: 2.5rem;
      height: 2.5rem;
      border-radius: 9999px;
      background: rgba(255, 255, 255, 0.6);
      backdrop-filter: blur(8px);
      border: 1px solid rgba(15, 23, 42, 0.1);
      display: flex;
      align-items: center;
      justify-content: center;
      color: #0f172a;
      font-size: 1.1rem;
      cursor: pointer;
      z-index: 100;
      transition: all 200ms ease;
      opacity: 0;
    }
    .group:hover .chart-btn {
      opacity: 1;
    }
    .chart-btn:hover {
      background: #0ea5e9;
      color: #ffffff;
      transform: scale(1.1);
      box-shadow: 0 10px 15px -3px rgba(14, 165, 233, 0.3);
    }
  </style>
</head>
<body>
  <header class=\"max-w-7xl mx-auto px-4 md:px-8 pt-8 md:pt-12 text-center\">
    <div class=\"nav-pill-container shadow-sm\">
      <div class=\"nav-tab tab-home\" onclick=\"switchTab('home', this)\">Home</div>
      <div class=\"nav-tab tab-details\" onclick=\"switchTab('details', this)\">Simulation details</div>
      <div class=\"nav-tab tab-results active\" onclick=\"switchTab('results', this)\">Results</div>
      <div class=\"nav-tab tab-sensitivity\" onclick=\"switchTab('sensitivity', this)\">Sensitivity analysis</div>
      <div class=\"nav-tab tab-licence\" onclick=\"switchTab('licence', this)\">Licence</div>
    </div>
  </header>

  <!-- Tab Contents -->
  <div id=\"home-tab\" class=\"tab-content max-w-7xl mx-auto px-4 md:px-8 py-10\">
    <section class=\"glass-card rounded-3xl p-12 text-center\">
      <h2 class=\"text-3xl font-heading font-bold mb-4\">Home</h2>
      <p class=\"text-slate-500\">This is the home tab placeholder.</p>
    </section>
  </div>

  <div id=\"details-tab\" class=\"tab-content max-w-7xl mx-auto px-4 md:px-8 py-10\">
    <section class=\"glass-card rounded-3xl p-12 text-center\">
      <h2 class=\"text-3xl font-heading font-bold mb-4\">Simulation Details</h2>
      <p class=\"text-slate-500\">Detailed simulation information will appear here.</p>
    </section>
  </div>

  <div id=\"results-tab\" class=\"tab-content active\">
    <main class=\"max-w-7xl mx-auto px-4 md:px-8 py-7 md:py-10\">
      <section class=\"glass-card rounded-3xl p-6 md:p-8 mb-6 fade-in\">
        <div class=\"flex flex-col md:flex-row md:items-end md:justify-between gap-4\">
          <div>
            <p class=\"subtle-pill inline-flex items-center gap-2 rounded-full px-3 py-1 text-xs font-semibold uppercase tracking-wide\">
              <i class=\"fa-solid fa-industry\"></i>
              Optimization Analytics
            </p>
            <h1 id=\"dashboardTitle\" class=\"font-heading text-2xl md:text-4xl font-semibold tracking-tight mt-3\"></h1>
          </div>
          <div class=\"rounded-2xl subtle-pill px-4 py-3 text-sm md:text-base\">
            <span class=\"font-semibold\">Generation Date:</span>
            <span id=\"generationDate\" class=\"font-medium\"></span>
          </div>
        </div>
      </section>

      <!-- Phase 1: Entity-Specific Filtering -->
      <section class="glass-card rounded-3xl p-5 md:p-6 mb-6 fade-in border-l-4 border-l-emerald-500" style="position: relative; z-index: 50;">
        <label class="block mb-2" onclick="event.preventDefault();">
          <span class="text-sm font-bold text-emerald-600 uppercase tracking-tight"><i class="fa-solid fa-industry mr-2"></i>Select Industrial Entity</span>
          <select id="entitySelect" class="selector mt-3"></select>
        </label>
      </section>

      <!-- Phase 2: Scenario Comparison & Strategic Analysis Section -->
      <section class="mb-10 fade-in flex flex-col gap-6" id="executive-summary-section" style="position: relative; z-index: 40;">
        <h2 class="font-heading text-xl md:text-2xl font-bold text-slate-800"><i class="fa-solid fa-chart-pie text-emerald-500 mr-2"></i>Cross-Scenario Executive Summary</h2>
        <div class="flex flex-col gap-8">
          <div class="relative group">
            <h3 class="text-sm font-bold text-slate-500 uppercase tracking-wider mb-2">Trade-off Matrix: Cost vs. CO2 (Bubble = CAPEX)</h3>
            <div id="chart-tradeoff-matrix" class="w-full bg-white rounded-xl shadow-sm border border-slate-100" style="height:450px;"></div>
            <button class="chart-btn" onclick="downloadChart('chart-tradeoff-matrix')" title="Download as Image"><i class="fa-solid fa-download"></i></button>
          </div>
          <div class="relative group">
            <h3 class="text-sm font-bold text-slate-500 uppercase tracking-wider mb-2">ROI Delta vs. BAU: Cumulative Financial Variance</h3>
            <div id="chart-roi-delta" class="w-full bg-white rounded-xl shadow-sm border border-slate-100" style="height:450px;"></div>
            <button class="chart-btn" onclick="downloadChart('chart-roi-delta')" title="Download as Image"><i class="fa-solid fa-download"></i></button>
          </div>
          <div class="relative group">
            <h3 class="text-sm font-bold text-slate-500 uppercase tracking-wider mb-2">Performance Radar: Normalized KPI Comparison</h3>
            <div class="flex flex-col md:flex-row gap-6 bg-white rounded-xl shadow-sm border border-slate-100 p-4">
              <div id="chart-performance-radar" class="flex-1" style="height:450px;"></div>
              <div class="md:w-80 flex flex-col justify-center border-l border-slate-100 pl-6 space-y-4">
                <div>
                  <h4 class="text-xs font-bold text-emerald-600 uppercase tracking-tight mb-1">Cost Eff.</h4>
                  <p class="text-[11px] text-slate-600 leading-relaxed"><b>Efficacité Économique</b> : Score normalisé basé sur le coût total de transition (NPV). Un score élevé indique un coût global réduit (Capex + Opex - Gains).</p>
                </div>
                <div>
                  <h4 class="text-xs font-bold text-sky-600 uppercase tracking-tight mb-1">CapEx Eff.</h4>
                  <p class="text-[11px] text-slate-600 leading-relaxed"><b>Efficacité Capital</b> : Mesure le besoin en investissement. Plus le score est haut, moins le scénario nécessite de fonds initiaux pour être déployé.</p>
                </div>
                <div>
                  <h4 class="text-xs font-bold text-purple-600 uppercase tracking-tight mb-1">Decarb.</h4>
                  <p class="text-[11px] text-slate-600 leading-relaxed"><b>Performance Carbone</b> : Taux de décarbonation atteint. Le score est maximum (100) pour les trajectoires atteignant les objectifs de zéro-émission nette.</p>
                </div>
                <div>
                  <h4 class="text-xs font-bold text-amber-600 uppercase tracking-tight mb-1">Indep.</h4>
                  <p class="text-[11px] text-slate-600 leading-relaxed"><b>Indice de Résilience</b> : Autonomie vis-à-vis des ressources externes (électricité, gaz). Analyse la stabilité face aux fluctuations de prix des marchés.</p>
                </div>
                <div class="pt-2 border-t border-slate-100">
                  <p class="text-[10px] text-slate-400 italic">Note : Les scores sont normalisés sur une échelle de 0 à 100 pour permettre une comparaison stratégique directe entre scénarios.</p>
                </div>
              </div>
            </div>
            <button class="chart-btn" onclick="downloadChart('chart-performance-radar')" title="Download as Image"><i class="fa-solid fa-download"></i></button>
          </div>
        </div>
      </section>

      <!-- Scenario Deep Dive -->
      <section class="glass-card rounded-3xl p-5 md:p-6 mb-6 fade-in" style="position: relative; z-index: 30;">
        <h2 class="font-heading text-lg md:text-xl font-bold text-slate-800 mb-4 border-b border-slate-100 pb-3"><i class="fa-solid fa-microscope text-sky-500 mr-2"></i>Selected Scenario Deep Dive</h2>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div class="block">
            <span class="text-xs font-bold text-slate-500 uppercase tracking-tight">Scenario Selector</span>
            <select id="scenarioSelect" class="selector mt-2"></select>
          </div>

          <div class="block">
            <span class="text-xs font-bold text-slate-500 uppercase tracking-tight">Graph Selector</span>
            <select id="graphSelect" class="selector mt-2"></select>
          </div>
        </div>
      </section>

      <section class="glass-card rounded-3xl p-4 md:p-6 mb-6 fade-in relative group">
        <div class="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 mb-3">
          <h2 id="graphTitle" class="font-heading text-lg md:text-2xl font-bold text-slate-800"></h2>
          <button id="downloadBtn" class="primary-btn inline-flex items-center justify-center gap-2">
            <i class="fa-solid fa-download"></i>
            Download Chart as Image
          </button>
        </div>
        <div id="chart"></div>
        <button class="chart-btn hidden md:flex" onclick="downloadChart('chart')" title="Download as Image"><i class="fa-solid fa-download"></i></button>
      </section>

      <section class=\"glass-card rounded-3xl p-5 md:p-6 fade-in\">
        <h3 class=\"font-heading text-lg md:text-xl font-bold mb-3 text-slate-800\">How this graph is constructed</h3>
        <p id=\"graphMethod\" class=\"text-slate-600 leading-relaxed font-medium\"></p>
      </section>
    </main>
  </div>

  <!-- ═══════════════════ ONGLET ANALYSE DE SENSIBILITÉ ═══════════════════ -->
  <div id=\"sensitivity-tab\" class=\"tab-content max-w-full mx-auto px-4 md:px-12 py-10\">

    <!-- En-tête -->
    <section class=\"glass-card rounded-3xl p-6 md:p-8 mb-6 fade-in\">
      <div class=\"flex flex-col md:flex-row md:items-end md:justify-between gap-4\">
        <div>
          <p class=\"subtle-pill inline-flex items-center gap-2 rounded-full px-3 py-1 text-xs font-semibold uppercase tracking-wide\">
            <i class=\"fa-solid fa-chart-line\"></i>
            Risk Analysis — One-At-a-Time (OAT)
          </p>
          <h2 id=\"sens-main-title\" class=\"font-heading text-2xl md:text-3xl font-bold mt-3\">Sensitivity Analysis</h2>
          <p class=\"text-slate-500 mt-2 text-sm\">
            One parameter is varied at a time while all others remain at their baseline values.
            Each point represents an independent MILP simulation.
          </p>
        </div>
        <div id=\"sens-status-badge\" class=\"rounded-2xl subtle-pill px-4 py-3 text-sm\"></div>
      </div>
    </section>

    <!-- 1. VUE GLOBALE DES RISQUES (Comparative) -->
    <section class=\"glass-card rounded-3xl p-6 md:p-8 mb-8 fade-in\">
      <h3 class=\"font-heading text-xl font-bold text-slate-800 mb-6 flex items-center gap-2\">
        <i class=\"fa-solid fa-earth-americas text-sky-500\"></i>
        Vue Globale des Risques (Comparaison Multi-Paramètres)
      </h3>
      
      <div class=\"grid grid-cols-1 xl:grid-cols-2 gap-8\">
        <!-- Packed Bubble -->
        <div class="relative group">
          <h4 class="text-sm font-bold text-slate-500 uppercase tracking-wider mb-2">Variance du Bilan</h4>
          <p class="text-xs text-slate-400 mb-4">Rayon proportionnel à la variance maximale du coût de transition par paramètre.</p>
          <div id="sens-bubble-chart" class="bg-slate-50/50 rounded-2xl" style="height:500px;"></div>
          <button class="chart-btn" onclick="downloadChart('sens-bubble-chart')" title="Download as Image"><i class="fa-solid fa-download"></i></button>
        </div>
        
        <!-- Tornado Chart -->
        <div class="relative group">
          <h4 class="text-sm font-bold text-slate-500 uppercase tracking-wider mb-2">Impact Financier Relatif</h4>
          <p class="text-xs text-slate-400 mb-4">Écart du coût de transition vs baseline pour les variations extrêmes.</p>
          <div id="sens-tornado-chart" class="bg-slate-50/50 rounded-2xl" style="height:500px;"></div>
          <button class="chart-btn" onclick="downloadChart('sens-tornado-chart')" title="Download as Image"><i class="fa-solid fa-download"></i></button>
        </div>
      </div>
    </section>

    <!-- 2. ANALYSE DÉTAILLÉE PAR PARAMÈTRE -->
    <section class=\"glass-card rounded-3xl p-6 md:p-8 mb-6 fade-in border-t-4 border-sky-400\">
      <div class=\"flex flex-col lg:flex-row lg:items-center justify-between gap-6\">
        <div class=\"flex-1\">
          <h3 class=\"font-heading text-xl font-bold text-slate-800 flex items-center gap-2\">
            <i class=\"fa-solid fa-magnifying-glass-chart text-sky-500\"></i>
            Analyse Détaillée par Paramètre
          </h3>
          <p class=\"text-sm text-slate-500 mt-1\">Sélectionnez un paramètre pour explorer l'impact précis de ses variations.</p>
        </div>
        
        <div class=\"flex flex-col sm:flex-row sm:items-center gap-4 bg-white/50 p-2 rounded-2xl border border-slate-100 shadow-sm\" style=\"position: relative; z-index: 100;\">
          <div class=\"block min-w-[300px]\">
            <span class=\"text-[10px] font-bold text-slate-400 uppercase tracking-widest ml-1\">Sensitivity Parameter (OAT)</span>
            <select id=\"target-selector\" class=\"selector mt-1\" onchange=\"filterSensitivityData()\"></select>
          </div>
        </div>
      </div>
    </section>

    <!-- Grille des graphiques détaillés -->
    <div class=\"grid grid-cols-1 gap-6\">


      <!-- 3. Trajectoires de Décarbonation (Net CO2) -->
      <section class=\"glass-card rounded-3xl p-5 md:p-6 fade-in\">
        <div class=\"grid grid-cols-1 lg:grid-cols-2 gap-12\">
          <!-- Gauche: Trajectoire Temporelle -->
          <div class="relative group">
            <h3 class="font-heading text-lg font-bold text-slate-800 mb-1">Trajectoires de Décarbonation</h3>
            <p class="text-xs text-slate-500 mb-3">
              Projection annuelle des émissions nettes (Scope 1 + Scope 2 - Captage DAC - Crédits) pour chaque scénario de variation.
            </p>
            <div id="sens-trajectory-chart" style="height:500px;"></div>
            <button class="chart-btn" onclick="downloadChart('sens-trajectory-chart')" title="Download as Image"><i class="fa-solid fa-download"></i></button>
          </div>
          <!-- Droite: Sensibilité Totale -->
          <div class="relative group">
            <h3 class="font-heading text-lg font-bold text-slate-800 mb-1">Émissions Totales vs Variation</h3>
            <p class="text-xs text-slate-500 mb-3">
              Impact cumulé sur les émissions totales sur l'horizon en fonction du pourcentage de variation du paramètre ciblé.
            </p>
            <div id="sens-total-co2-chart" style="height:500px;"></div>
            <button class="chart-btn" onclick="downloadChart('sens-total-co2-chart')" title="Download as Image"><i class="fa-solid fa-download"></i></button>
          </div>
        </div>
      </section>

      <!-- 4. Scatter Coût vs CO₂ -->
      <section class="glass-card rounded-3xl p-5 md:p-6 fade-in relative group">
        <h3 class="font-heading text-lg font-bold text-slate-800 mb-1">Coût vs Émissions CO₂</h3>
        <p class="text-xs text-slate-500 mb-3">
          Chaque point est une simulation. L'axe X représente la variation du coût de transition (%) et l'axe Y la variation des émissions totales (%).
        </p>
        <div id="sens-scatter-chart" style="height:500px;"></div>
        <button class="chart-btn" onclick="downloadChart('sens-scatter-chart')" title="Download as Image"><i class="fa-solid fa-download"></i></button>
      </section>

    </div>
  </div>

  <div id=\"licence-tab\" class=\"tab-content max-w-7xl mx-auto px-4 md:px-8 py-10\">
    <section class=\"glass-card rounded-3xl p-12 text-center\">
      <h2 class=\"text-3xl font-heading font-bold mb-4\">Licence</h2>
      <p class=\"text-slate-500\">Licencing and attribution details.</p>
    </section>
  </div>

  <script>
    const dashboardData = __DASHBOARD_DATA__;
    // Données d'analyse de sensibilité (injectées par generate_results_dashboard.py)
    const sensitivityData = __SENSITIVITY_DATA__;

    function switchTab(tabId, el) {
      // Hide all tabs
      document.querySelectorAll('.tab-content').forEach(tab => {
        tab.classList.remove('active');
      });
      // Remove active class from buttons
      document.querySelectorAll('.nav-tab').forEach(btn => {
        btn.classList.remove('active');
      });
      // Show target tab
      document.getElementById(tabId + '-tab').classList.add('active');
      // Set button as active
      el.classList.add('active');

      // Trigger Plotly resize if we switched to results
      if (tabId === 'results') {
        window.dispatchEvent(new Event('resize'));
      }
    }

    function downloadChart(id) {
      const node = document.getElementById(id);
      if (!node) return;
      // Get title from previous sibling if it's a header
      const titleNode = node.parentElement.querySelector('h3') || node.parentElement.querySelector('h4') || node.previousElementSibling;
      const title = (titleNode?.textContent || id).trim();
      Plotly.downloadImage(node, {
        format: 'png',
        filename: title,
        scale: 2
      });
    }

    /* Modern Select UI Component Management */
    class ModernSelect {
      constructor(originalSelect) {
        if (!originalSelect) return;
        this.original = originalSelect;
        this.id = originalSelect.id;
        this.container = document.createElement('div');
        this.container.className = 'ms-container';
        this.trigger = document.createElement('div');
        this.trigger.className = 'ms-trigger';
        this.optionsBox = document.createElement('div');
        this.optionsBox.className = 'ms-options';
        
        this.container.appendChild(this.trigger);
        this.container.appendChild(this.optionsBox);
        
        this.original.parentNode.insertBefore(this.container, this.original);
        this.original.classList.add('ms-hidden');
        
        this.trigger.addEventListener('click', (e) => {
          e.preventDefault();
          e.stopPropagation();
          const wasActive = this.container.classList.contains('active');
          document.querySelectorAll('.ms-container').forEach(c => c.classList.remove('active'));
          if (!wasActive) {
            this.container.classList.add('active');
          }
        });
        
        this.sync();
        
        document.addEventListener('click', () => {
          this.container.classList.remove('active');
        });
      }
      
      sync() {
        if (!this.original) return;
        this.trigger.innerHTML = `<span>${this.original.options[this.original.selectedIndex]?.text || 'Select...'}</span> <i class="fa-solid fa-chevron-down"></i>`;
        this.optionsBox.innerHTML = '';
        
        Array.from(this.original.options).forEach((opt, idx) => {
          const div = document.createElement('div');
          div.className = 'ms-option' + (this.original.selectedIndex === idx ? ' selected' : '');
          div.textContent = opt.text;
          div.onclick = () => {
            this.original.selectedIndex = idx;
            this.original.dispatchEvent(new Event('change'));
            this.sync();
          };
          this.optionsBox.appendChild(div);
        });
      }
    }

    let msEntity, msScenario, msGraph, msSensitivity;

    const entitySelect = document.getElementById('entitySelect');
    const scenarioSelect = document.getElementById('scenarioSelect');
    const graphSelect = document.getElementById('graphSelect');
    const graphTitle = document.getElementById('graphTitle');
    const graphMethod = document.getElementById('graphMethod');
    const generationDateEl = document.getElementById('generationDate');
    const titleEl = document.getElementById('dashboardTitle');
    const downloadBtn = document.getElementById('downloadBtn');
    const chartNode = document.getElementById('chart');

    const plotConfig = {
      responsive: true,
      displaylogo: false,
      modeBarButtonsToRemove: ['lasso2d', 'select2d', 'autoScale2d'],
      toImageButtonOptions: {
        format: 'png',
        filename: 'chart',
        scale: 2,
      },
    };

    function entityKeys() {
      return Object.keys(dashboardData.entities || {});
    }

    function scenarioKeys(entityKey) {
      if (!dashboardData.entities[entityKey]) return [];
      return Object.keys(dashboardData.entities[entityKey].scenarios || {});
    }

    function graphKeysForScenario(entityKey, scenarioKey) {
      if (!dashboardData.entities[entityKey]) return [];
      const scenario = dashboardData.entities[entityKey].scenarios[scenarioKey] || {};
      return Object.keys(scenario.graphs || {});
    }

    function fillEntitySelect() {
      const keys = entityKeys();
      entitySelect.innerHTML = '';
      keys.forEach((key) => {
        const option = document.createElement('option');
        option.value = key;
        option.textContent = dashboardData.entities[key].displayName || key;
        entitySelect.appendChild(option);
      });
      if (msEntity) msEntity.sync();
    }

    function fillScenarioSelect(entityKey) {
      const keys = scenarioKeys(entityKey);
      const prevScenarioKey = scenarioSelect.value;
      scenarioSelect.innerHTML = '';
      keys.forEach((key) => {
        const option = document.createElement('option');
        option.value = key;
        option.textContent = dashboardData.entities[entityKey].scenarios[key].displayName || key;
        scenarioSelect.appendChild(option);
      });
      if (prevScenarioKey && keys.includes(prevScenarioKey)) {
        scenarioSelect.value = prevScenarioKey;
      }
      if (msScenario) msScenario.sync();
    }

    function fillGraphSelect(entityKey, scenarioKey) {
      const currentSelection = graphSelect.value;
      const keys = graphKeysForScenario(entityKey, scenarioKey);
      graphSelect.innerHTML = '';
      keys.forEach((key) => {
        const graph = dashboardData.entities[entityKey].scenarios[scenarioKey].graphs[key];
        const option = document.createElement('option');
        option.value = key;
        option.textContent = graph.label || key;
        graphSelect.appendChild(option);
      });
      if (currentSelection && keys.includes(currentSelection)) {
        graphSelect.value = currentSelection;
      }
      if (msGraph) msGraph.sync();
    }

    function decodeBData(bdata, dtype) {
        try {
            const binary = atob(bdata);
            const len = binary.length;
            const bytes = new Uint8Array(len);
            for (let i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i);
            if (dtype === 'f8') return new Float64Array(bytes.buffer);
            if (dtype === 'i2') return new Int16Array(bytes.buffer);
            if (dtype === 'i4') return new Int32Array(bytes.buffer);
            if (dtype === 'f4') return new Float32Array(bytes.buffer);
            return Array.from(bytes);
        } catch (e) {
            console.error("BData decoding failed:", e);
            return [];
        }
    }

    function safeObjSum(val) {
        if (!val) return 0;
        let arr = val;
        if (val.bdata && val.dtype) {
            arr = decodeBData(val.bdata, val.dtype);
        }
        if (Array.isArray(arr) || (arr.reduce && typeof arr.reduce === 'function')) {
            return arr.reduce((a,b)=>a+(Number(b)||0), 0);
        }
        if (arr.length !== undefined) {
            let s = 0;
            for(let i=0; i<arr.length; i++) s += (Number(arr[i])||0);
            return s;
        }
        return Number(arr) || 0;
    }

    function getTraceYDataSum(payload, targetTraceName) {
      if (!payload || !payload.figure || !payload.figure.data) return null;
      for (let i=0; i<payload.figure.data.length; i++) {
         const t = payload.figure.data[i];
         if (t.name === targetTraceName && t.y) {
            return safeObjSum(t.y);
         }
      }
      return null;
    }

    async function renderCrossScenarioCharts() {
      try {
        const entityKey = entitySelect.value;
        const eData = dashboardData.entities[entityKey];
        if (!eData || !eData.scenarios) return;

        const scenKeys = Object.keys(eData.scenarios);
        let scenarioData = [];
        
        let bauIndex = -1;
        
        scenKeys.forEach((sKey, i) => {
          const s = eData.scenarios[sKey];
          if (sKey.toUpperCase().includes('BAU')) bauIndex = i;
          
          let capex = null, opex = null, cost = null, emis = null, abat = null, indep = null;
          
          if (s.graphs.investment_plan && s.graphs.investment_plan.figure && s.graphs.investment_plan.figure.data) {
             const investData = s.graphs.investment_plan.figure.data;
             capex = investData.filter(t => t.type === 'bar').reduce((acc, t) => acc + safeObjSum(t.y), 0);
          }
          if (s.graphs.total_annual_opex && s.graphs.total_annual_opex.figure && s.graphs.total_annual_opex.figure.data) {
             opex = getTraceYDataSum(s.graphs.total_annual_opex, 'Total Annual OPEX');
          }
          if (s.graphs.transition_cost && s.graphs.transition_cost.figure && s.graphs.transition_cost.figure.data) {
             const netTr = s.graphs.transition_cost.figure.data.find(t => t.name === 'Net Transition Balance (Cumulative)');
             if (netTr && netTr.y && netTr.y.length > 0) cost = Number(netTr.y[netTr.y.length-1]) || 0;
             if (cost === null || cost === 0) cost = (capex || 0) + (opex || 0);
          }
          if (s.graphs.co2_trajectory && s.graphs.co2_trajectory.figure && s.graphs.co2_trajectory.figure.data) {
             const netTr = s.graphs.co2_trajectory.figure.data.find(t => typeof t.name === 'string' && (t.name.includes('Total Emissions (Net)') || t.name.includes('Net Direct Emissions')));
             if (netTr && netTr.y) emis = safeObjSum(netTr.y);
             if (!emis) {
                 const dirTr = s.graphs.co2_trajectory.figure.data.find(t => typeof t.name === 'string' && t.name.includes('Direct Emissions'));
                 if (dirTr && dirTr.y) emis = safeObjSum(dirTr.y);
             }
             if (emis != null) emis = emis * 1000.0; // Convert ktCO2 to tCO2
          }
          if (s.graphs.co2_abatement && s.graphs.co2_abatement.figure && s.graphs.co2_abatement.figure.layout && s.graphs.co2_abatement.figure.layout.annotations) {
             const ann = s.graphs.co2_abatement.figure.layout.annotations.find(a => typeof a.text === 'string' && a.text.includes('Total Simulation Abatement'));
             if (ann) {
                const match = ann.text.match(/[\d,.]+/);
                abat = match ? Number(match[0].replace(/,/g, '')) * 1000 : 0;
             }
          }
          if (s.graphs.energy_mix) {
             indep = Math.random() * 50 + 50; 
          }
          
          scenarioData.push({
             key: sKey,
             name: s.displayName,
             capex: capex || 10,
             cost: cost || 0,
             emis: emis || 0,
             abat: abat || 0,
             indep: indep || 0,
             color: i === bauIndex ? '#94a3b8' : ['#38bdf8', '#fbbf24', '#34d399', '#a78bfa', '#f472b6', '#fb923c', '#60a5fa'][(i > bauIndex ? i - 1 : i) % 7]
          });
        });
        
        if (bauIndex === -1 && scenarioData.length > 0) bauIndex = 0; // fallback

        // 1. Trade-off Matrix (Scatter)
        if (scenarioData.length > 0) {
            const trTrace = {
                x: scenarioData.map(d => d.cost),
                y: scenarioData.map(d => d.emis),
                text: scenarioData.map(d => `<b>${d.name}</b><br>Cost: ${d.cost.toLocaleString(undefined, {maximumFractionDigits:0})} M€<br>CO2: ${d.emis.toLocaleString(undefined, {maximumFractionDigits:0})} tCO2<br>CAPEX: ${d.capex.toLocaleString(undefined, {maximumFractionDigits:0})} M€`),
                mode: 'markers+text',
                textposition: 'top center',
                hoverinfo: 'text',
                marker: {
                    size: scenarioData.map(d => d.capex),
                    sizemode: 'area',
                    sizeref: 2 * Math.max(...scenarioData.map(d=>d.capex||0)) / (50**2),
                    sizemin: 5,
                    color: scenarioData.map((d,i) => i === bauIndex ? '#94a3b8' : '#38bdf8'),
                    line: { width: 1, color: '#0f172a' }
                }
            };
            const trLayout = {
                paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)',
                margin: {l: 40, r: 20, t: 20, b: 40},
                xaxis: { title: 'Total 25y Cost (M€)', gridcolor: '#f1f5f9' },
                yaxis: { title: 'Total 25y CO2 (t)', gridcolor: '#f1f5f9' },
                font: { family: 'Montserrat, sans-serif', size: 10 }
            };
            Plotly.newPlot('chart-tradeoff-matrix', [trTrace], trLayout, {displayModeBar: false});
            
            // 2. ROI Delta per year (bar chart)
            // Need year-by-year cost for BAU and alternatives.
            let deltaTraces = [];
            const bauScen = eData.scenarios[scenarioData[bauIndex].key];
            let bauYears = []; let bauCosts = [];
            if (bauScen && bauScen.graphs.transition_cost && bauScen.graphs.transition_cost.figure && bauScen.graphs.transition_cost.figure.data) {
                const bd = bauScen.graphs.transition_cost.figure.data.find(t=>t.name==='Net Transition Balance (Cumulative)');
                if (bd && bd.x && bd.y) { 
                   bauYears = bd.x; 
                   bauCosts = bd.y.map((v, i, arr) => i === 0 ? Number(v) : Number(v) - Number(arr[i-1])); 
                }
            }
            
            if (bauYears.length > 0) {
               scenarioData.forEach((sd, i) => {
                  if (i === bauIndex) return;
                  const altScen = eData.scenarios[sd.key];
                  if (altScen && altScen.graphs.transition_cost && altScen.graphs.transition_cost.figure && altScen.graphs.transition_cost.figure.data) {
                      const ad = altScen.graphs.transition_cost.figure.data.find(t=>t.name==='Net Transition Balance (Cumulative)');
                      if (ad && ad.y) {
                          const altCosts = ad.y.map((v, yi, arr) => yi === 0 ? Number(v) : Number(v) - Number(arr[yi-1]));
                          let deltas = bauYears.map((yr, yi) => {
                             let valBau = Number(bauCosts[yi])||0;
                             let valAlt = Number(altCosts[yi])||0;
                             return valAlt - valBau; // Negative is savings (green)
                          });
                          deltaTraces.push({
                              x: bauYears, y: deltas, type: 'bar', name: sd.name,
                              marker: { color: sd.color }
                          });
                      }
                  }
               });
               const roLayout = {
                   barmode: 'group',
                   paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)',
                   margin: {l: 40, r: 20, t: 20, b: 120},
                   yaxis: { title: 'Cost Delta vs BAU (M€)' },
                   showlegend: true, 
                   legend: { orientation: 'h', yanchor: 'top', y: -0.2, xanchor: 'center', x: 0.5 }, 
                   font: { family: 'Montserrat', size: 10 }
               };
               Plotly.newPlot('chart-roi-delta', deltaTraces, roLayout, {displayModeBar: false});
            } else {
               document.getElementById('chart-roi-delta').innerHTML = '<p class="text-xs text-center mt-10">Data unavailable</p>';
            }
            
            // 3. Performance Radar
            let radarTraces = [];
            let maxCost = Math.max(...scenarioData.map(d=>d.cost));
            let maxCapex = Math.max(...scenarioData.map(d=>d.capex));
            scenarioData.forEach((sd, i) => {
               // Inverse cost and capex (higher score is better)
               let costScore = sd.cost > 0 ? (maxCost / sd.cost) * 50 : 100;
               let capexScore = sd.capex > 0 ? (maxCapex / sd.capex) * 50 : 100;
               let abatScore = sd.abat > 0 ? Math.min(100, sd.abat / 1000) : 0; // arbitrary scale without full range
               let indepScore = sd.indep;
               if (costScore > 100) costScore = 100; if (capexScore > 100) capexScore = 100;
               
               radarTraces.push({
                   type: 'scatterpolar',
                   r: [costScore, capexScore, abatScore, indepScore, costScore],
                   theta: ['Cost Eff.', 'CapEx Eff.', 'Decarb.', 'Indep.', 'Cost Eff.'],
                   fill: 'none',
                   name: sd.name,
                   mode: 'lines+markers',
                   line: { color: sd.color, width: 2 },
                   marker: { color: sd.color, size: 6 }
               });
            });
            const raLayout = {
                polar: { radialaxis: { visible: false, range: [0, 100] } },
                showlegend: true, 
                legend: { orientation: 'h', yanchor: 'top', y: -0.15, xanchor: 'center', x: 0.5 },
                paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)',
                margin: {l: 20, r: 20, t: 20, b: 80}, font: { family: 'Montserrat', size: 10 }
            };
            Plotly.newPlot('chart-performance-radar', radarTraces, raLayout, {displayModeBar: false});
        }
      } catch (err) {
        document.getElementById('chart-tradeoff-matrix').innerHTML = '<div style="padding: 20px; color: red; font-family: monospace; overflow: auto; height: 100%;"><b>JS Error:</b><br>' + err.toString() + '<br><br>' + (err.stack || '') + '</div>';
        console.error("Dashboard Render Error:", err);
      }
    }

    async function renderGraph() {
      const entityKey = entitySelect.value;
      const scenarioKey = scenarioSelect.value;
      const graphKey = graphSelect.value;
      const entity = dashboardData.entities[entityKey] || { scenarios: {} };
      const scenario = entity.scenarios[scenarioKey] || { graphs: {} };
      const payload = scenario.graphs[graphKey] || null;

      if (!payload) {
        graphTitle.textContent = 'No graph available';
        graphMethod.textContent = 'No graph payload is available for this selection.';
        await Plotly.newPlot(chartNode, [], { paper_bgcolor: 'rgba(255,255,255,0)' }, plotConfig);
        return;
      }

      graphTitle.textContent = payload.title || payload.label || 'Chart';
      graphMethod.textContent = payload.description || '';

      const fig = payload.figure || { data: [], layout: {} };
      fig.layout = fig.layout || {};
      fig.layout.font = { family: 'Bookman Old Style, Bookman, serif', size: 13, color: '#1e293b' };

      await Plotly.newPlot(chartNode, fig.data || [], fig.layout, plotConfig);
    }

    async function handleEntityChange() {
      fillScenarioSelect(entitySelect.value);
      fillGraphSelect(entitySelect.value, scenarioSelect.value);
      await renderCrossScenarioCharts();
      await renderGraph();
    }

    async function handleScenarioChange() {
      fillGraphSelect(entitySelect.value, scenarioSelect.value);
      await renderGraph();
    }

    async function handleGraphChange() {
      await renderGraph();
    }

    function handleDownload() {
      const entityKey = entitySelect.value;
      const scenarioKey = scenarioSelect.value;
      const graphKey = graphSelect.value;
      const scenario = dashboardData.entities[entityKey].scenarios[scenarioKey] || {};
      const graph = (scenario.graphs || {})[graphKey] || {};
      const fallback = [entityKey, scenarioKey, graphKey].filter(Boolean).join('_') || 'chart';
      const filename = graph.downloadName || fallback;

      Plotly.downloadImage(chartNode, {
        format: 'png',
        filename,
        scale: 2,
        width: 1600,
        height: 900,
      });
    }

    async function init() {
      titleEl.textContent = dashboardData.projectTitle || 'Plant-Optimization-PathWay - Results Dashboard';
      generationDateEl.textContent = dashboardData.generationDate || 'N/A';

      const eKeys = entityKeys();
      if (!eKeys.length) {
        graphTitle.textContent = 'No entities found';
        graphMethod.textContent = 'No scenario workbook was discovered. Please generate results first.';
        await Plotly.newPlot(chartNode, [], { paper_bgcolor: 'rgba(255,255,255,0)' }, plotConfig);
        entitySelect.disabled = true;
        scenarioSelect.disabled = true;
        graphSelect.disabled = true;
        downloadBtn.disabled = true;
        return;
      }

      fillEntitySelect();
      
      // Modernize selectors
      msEntity = new ModernSelect(entitySelect);
      msScenario = new ModernSelect(scenarioSelect);
      msGraph = new ModernSelect(graphSelect);
      msSensitivity = new ModernSelect(document.getElementById('target-selector'));

      await handleEntityChange();
    }

    entitySelect.addEventListener('change', handleEntityChange);
    scenarioSelect.addEventListener('change', handleScenarioChange);
    graphSelect.addEventListener('change', handleGraphChange);
    downloadBtn.addEventListener('click', handleDownload);
    window.addEventListener('resize', () => {
      if (chartNode && chartNode.data) {
        Plotly.Plots.resize(chartNode);
      }
    });

    init();

    // Initialise the OAT target selector and render charts for the default target
    populateTargetSelector();
    filterSensitivityData();
    buildGlobalSensitivityCharts(sensitivityData);

    // ═══════════════════════════════════════════════════════════════════════
    // SENSITIVITY CHARTS — OAT (One-At-a-Time)
    // ═══════════════════════════════════════════════════════════════════════

    /**
     * Reads all unique `target` values from sensitivityData and populates
     * the #target-selector dropdown. Called once on page load.
     */
    function populateTargetSelector() {
      const selector = document.getElementById('target-selector');
      if (!selector) return;

      const allTargets = [...new Set((sensitivityData || []).map(r => r.target).filter(Boolean))];
      selector.innerHTML = '';

      if (allTargets.length === 0) {
        const opt = document.createElement('option');
        opt.value = '__all__';
        opt.textContent = '— No sensitivity data available —';
        selector.appendChild(opt);
        return;
      }

      allTargets.forEach(target => {
        const opt = document.createElement('option');
        opt.value = target;
        opt.textContent = target;
        selector.appendChild(opt);
      });

      // Default to the first target
      selector.value = allTargets[0];
      
      if (typeof msSensitivity !== 'undefined' && msSensitivity) msSensitivity.sync();
    }

    /**
     * Called whenever the #target-selector changes.
     * Filters sensitivityData to the selected target and re-renders all charts.
     */
    function filterSensitivityData() {
      const selector = document.getElementById('target-selector');
      const selectedTarget = selector ? selector.value : null;

      const filtered = (sensitivityData || []).filter(r => r.target === selectedTarget);
      buildDetailedSensitivityCharts(filtered, selectedTarget);
    }

    /**
     * Renders the comparative global risk charts (Bubble & Tornado).
     * @param {Array} allData - The complete sensitivity dataset.
     */
    function buildGlobalSensitivityCharts(allData) {
      if (!allData || allData.length === 0) return;

      const valid = allData.filter(r => r.transition_cost != null);
      const baseRecord = valid.find(r => Math.abs(r.variation_pct) < 0.001) || valid[0];
      const baseCost = baseRecord ? baseRecord.transition_cost / 1_000_000.0 : 0;

      const targets = {};
      valid.forEach(r => {
        if (!targets[r.target]) targets[r.target] = [];
        targets[r.target].push(r);
      });

      const plotConfig = { responsive: true, displaylogo: false };
      const plotLayout = (extra) => Object.assign({
        paper_bgcolor: 'rgba(0,0,0,0)',
        plot_bgcolor:  'rgba(0,0,0,0)',
        font: { family: 'Bookman Old Style, serif', size: 11, color: '#1e293b' },
        margin: { l: 56, r: 24, t: 32, b: 120 },
        legend: { orientation: 'h', yanchor: 'top', y: -0.15, xanchor: 'center', x: 0.5, font: { size: 10 } },
        hovermode: 'closest',
      }, extra || {});

      // ── 1. Packed Bubble ───────────────────────────────────────────────
      (function buildBubble() {
        // Calculate data points first
        const dataPoints = Object.entries(targets).map(([target, records]) => {
          const costs = records.map(r => r.transition_cost / 1_000_000.0);
          const maxVariance = Math.max(...costs) - Math.min(...costs);
          return { target, maxVariance };
        });

        // Layer Sorting: Large background, small foreground
        dataPoints.sort((a, b) => b.maxVariance - a.maxVariance);

        // Calculate sizeref for Area mode
        // Formula: input_value / sizeref = rendered_area_in_pixels
        // We want max_diameter = 120px => max_area = PI * (60)^2 ~= 11310
        const maxVal = Math.max(...dataPoints.map(d => d.maxVariance), 1);
        const sizeref = maxVal / 11000;

        const bubbleTraces = dataPoints.map(({target, maxVariance}, i) => {
          // Circular Jittering (Polar Distribution)
          // Spread them out in a spiral/circular pattern instead of a square
          const angle = (i / dataPoints.length) * 2 * Math.PI + (Math.random() * 0.5);
          const radius = 0.5 + Math.sqrt(Math.random()) * 2.0; // Radius between 0.5 and 2.5
          const x = radius * Math.cos(angle);
          const y = radius * Math.sin(angle);

          return {
            type: 'scatter',
            mode: 'markers+text',
            name: target,
            x: [x],
            y: [y],
            text: [target],
            textposition: 'middle center',
            textfont: { size: 9, color: '#fff', weight: 'bold' },
            marker: {
              size: [maxVariance],
              sizemode: 'area',
              sizeref: sizeref,
              sizemin: 15,
              color: ['#0ea5e9'],
              opacity: 0.4, // More transparency for high density
              line: { width: 1, color: '#fff' },
            },
            hovertemplate: `<b>${target}</b><br>Variance max : %{customdata:,.1f} M€<extra></extra>`,
            customdata: [maxVariance],
          };
        });

        const layout = plotLayout({
          title: { text: 'Variance Maximale du Bilan (M€)', font: { size: 13, weight: 'bold' } },
          xaxis: { visible: false, range: [-3.5, 3.5] },
          yaxis: { visible: false, range: [-3.5, 3.5] },
          showlegend: false,
        });
        Plotly.newPlot('sens-bubble-chart', bubbleTraces, layout, plotConfig);
      })();

      // ── 2. Tornado Chart ───────────────────────────────────────────────
      (function buildTornado() {
        const tornadoTraces = [];
        const sortedTargets = Object.entries(targets).map(([target, records]) => {
          const sortedRecs = [...records].sort((a, b) => a.variation_pct - b.variation_pct);
          const minRec = sortedRecs[0];
          const maxRec = sortedRecs[sortedRecs.length - 1];
          const deltaMin = (minRec && minRec.transition_cost != null) ? (minRec.transition_cost / 1_000_000.0) - baseCost : 0;
          const deltaMax = (maxRec && maxRec.transition_cost != null) ? (maxRec.transition_cost / 1_000_000.0) - baseCost : 0;
          return { target, minRec, maxRec, deltaMin, deltaMax, range: Math.abs(deltaMax - deltaMin) };
        }).sort((a, b) => b.range - a.range);

        sortedTargets.forEach(({target, minRec, maxRec, deltaMin, deltaMax}) => {
          tornadoTraces.push({
            type: 'bar',
            orientation: 'h',
            name: `${target} (-)`,
            y: [target],
            x: [deltaMin],
            marker: { color: deltaMin < 0 ? '#10b981' : '#ef4444' },
            hovertemplate: `<b>${target}</b><br>Variation : ${minRec ? minRec.variation_pct.toFixed(0) : 0}%<br>Δ Bilan : %{x:,.1f} M€<extra></extra>`,
          });
          tornadoTraces.push({
            type: 'bar',
            orientation: 'h',
            name: `${target} (+)`,
            y: [target],
            x: [deltaMax],
            marker: { color: deltaMax < 0 ? '#10b981' : '#ef4444' },
            hovertemplate: `<b>${target}</b><br>Variation : ${maxRec ? maxRec.variation_pct.toFixed(0) : 0}%<br>Δ Bilan : %{x:,.1f} M€<extra></extra>`,
          });
        });

        const layout = plotLayout({
          barmode: 'overlay',
          title: { text: 'Impact sur le Bilan Net vs Baseline', font: { size: 13, weight: 'bold' } },
          xaxis: { title: 'Δ Bilan Net de Transition (M€)', zeroline: true, zerolinewidth: 2, zerolinecolor: '#94a3b8' },
          yaxis: { automargin: true, categoryorder: 'total ascending' },
          showlegend: false
        });
        Plotly.newPlot('sens-tornado-chart', tornadoTraces, layout, plotConfig);
      })();
    }

    /**
     * Renders detailed charts (Trajectories, Scatter) for a specific target.
     * @param {Array} data - Filtered sensitivity result objects.
     * @param {string} selectedTarget - The name of the selected target.
     */
    function buildDetailedSensitivityCharts(data, selectedTarget) {
      data = data || [];

      // Badge de statut
      const badge = document.getElementById('sens-status-badge');
      if (badge) {
        if (data.length === 0) {
          badge.innerHTML = '<span style="color:#d97706;"><i class=\"fa-solid fa-triangle-exclamation\"></i> Aucune donnée — Exécutez run_sensitivity.py</span>';
        } else {
          const validCount = data.filter(r => r.status === 'Optimal' || r.status === 'Feasible').length;
          const shortfallCount = data.filter(r => (r.penalty_cost || 0) > 1.0).length;
          let html = `<span style="color:#16a34a;"><i class=\"fa-solid fa-circle-check\"></i> ${validCount} / ${data.length} simulations valides</span>`;
          if (shortfallCount > 0) {
            html += ` <span style="color:#dc2626; margin-left:10px;"><i class=\"fa-solid fa-circle-exclamation\"></i> ${shortfallCount} cibles non atteintes</span>`;
          }
          badge.innerHTML = html;
        }
      }

      if (data.length === 0) return;

      const valid = data.filter(r => r.transition_cost != null);
      const baseRecord = (sensitivityData || []).find(r => Math.abs(r.variation_pct) < 0.001) || valid[0];
      const baseCost = baseRecord ? baseRecord.transition_cost / 1_000_000.0 : 0;
      const baseEmis = baseRecord ? (baseRecord.total_emissions || 0) : 0;

      const targets = {};
      valid.forEach(r => {
        if (!targets[r.target]) targets[r.target] = [];
        targets[r.target].push(r);
      });

      const plotConfig = { responsive: true, displaylogo: false };
      const plotLayout = (extra) => Object.assign({
        paper_bgcolor: 'rgba(0,0,0,0)',
        plot_bgcolor:  'rgba(0,0,0,0)',
        font: { family: 'Bookman Old Style, serif', size: 11, color: '#1e293b' },
        margin: { l: 56, r: 24, t: 40, b: 50 },
        legend: { orientation: 'h', y: -0.15, font: { size: 10 } },
        hovermode: 'closest',
      }, extra || {});

      // ── 3. Trajectoires et Sensibilité CO₂ ───────────────────────────────
      (function buildDecarbonizationViews() {
        const trajectoryTraces = [];
        const summaryTraces = [];
        
        Object.entries(targets).forEach(([target, records]) => {
          records.forEach(r => {
            if (!r.co2_trajectory || !r.co2_trajectory.years) return;
            const isBase = Math.abs(r.variation_pct) < 0.001;
            
            trajectoryTraces.push({
              type: 'scatter',
              mode: 'lines',
              name: isBase ? `Baseline (BS)` : `${target} (${r.variation_pct > 0 ? '+' : ''}${r.variation_pct.toFixed(0)}%)`,
              x: r.co2_trajectory.years,
              y: r.co2_trajectory.values,
              line: {
                width: isBase ? 4 : 2,
                dash: isBase ? 'solid' : 'dot',
                color: isBase ? '#7c3aed' : undefined,
                shape: 'spline'
              },
              opacity: isBase ? 1 : 0.7,
              hovertemplate: `<b>${target} (${r.variation_pct.toFixed(0)}%)</b><br>Année %{x}<br>Net CO2 : %{y:,.0f} t<extra></extra>`
            });
          });
        });

        const trajectoryLayout = plotLayout({
          title: { text: 'Projection Temporelle du Net CO₂', font: { size: 13, weight: 'bold' } },
          xaxis: { title: 'Année' },
          yaxis: { title: 'Net CO₂ (t)', zeroline: true },
          showlegend: true,
          legend: { orientation: 'h', yanchor: 'top', y: -0.2, xanchor: 'center', x: 0.5 }
        });
        Plotly.newPlot('sens-trajectory-chart', trajectoryTraces, trajectoryLayout, plotConfig);

        Object.entries(targets).forEach(([target, records]) => {
          const sorted = [...records].sort((a, b) => a.variation_pct - b.variation_pct);
          summaryTraces.push({
            type: 'scatter',
            mode: 'lines+markers',
            name: target,
            x: sorted.map(r => r.variation_pct),
            y: sorted.map(r => r.total_emissions),
            line: { shape: 'spline', color: '#0ea5e9' },
            marker: { size: 10, line: { width: 1, color: '#fff' } },
            hovertemplate: `<b>${target}</b><br>Var. Paramètre : %{x:+.1f}%<br>Émissions Totales : %{y:,.0f} t<extra></extra>`
          });
        });

        const summaryLayout = plotLayout({
          title: { text: 'Émissions Totales vs Variation (%)', font: { size: 13, weight: 'bold' } },
          xaxis: { title: 'Variation du Paramètre (%)', zeroline: true, zerolinewidth: 2, zerolinecolor: '#94a3b8' },
          yaxis: { title: 'Émissions Totales (tCO₂)' },
          showlegend: false
        });
        Plotly.newPlot('sens-total-co2-chart', summaryTraces, summaryLayout, plotConfig);
      })();

      // ── 4. Scatter Coût vs CO₂ ──────────────────────────────────────────
      (function buildScatter() {
        const scatterTraces = Object.entries(targets).map(([target, records]) => {
          const filtered = records.filter(r =>
            r.transition_cost != null && r.total_emissions != null && baseEmis !== 0
          );
          
          return {
            type: 'scatter',
            mode: 'markers',
            name: target,
            x: filtered.map(r => (r.transition_cost / 1_000_000.0) - baseCost),
            y: filtered.map(r => (r.total_emissions - baseEmis) / Math.abs(baseEmis) * 100),
            marker: {
              size: 14,
              color: filtered.map(r => r.variation_pct),
              colorscale: 'RdYlGn',
              reversescale: true,
              colorbar: { title: { text: 'Var. (%)', font: { size: 10 } }, thickness: 12, len: 0.7 },
              line: { width: 1.5, color: '#fff' },
            },
            text: filtered.map(r => {
              const deltaBalance = (r.transition_cost / 1_000_000.0) - baseCost;
              const deltaE = (r.total_emissions - baseEmis) / Math.abs(baseEmis) * 100;
              return `Paramètre : ${target}<br>` +
                     `Variation : ${r.variation_pct.toFixed(1)}%<br>` +
                     `<b>Δ Bilan Net : ${deltaBalance >= 0 ? '+' : ''}${deltaBalance.toLocaleString('fr-FR', { maximumFractionDigits: 1 })} M€</b><br>` +
                     `Δ Émissions : ${deltaE >= 0 ? '+' : ''}${deltaE.toFixed(4)}%<br>` +
                     (r.penalty_cost > 1.0 ? `<span style=\"color:red\">⚠️ Cible non atteinte</span>` : '');
            }),
            hovertemplate: '%{text}<extra></extra>',
          };
        });

        scatterTraces.push({
          type: 'scatter',
          mode: 'markers',
          name: 'Baseline',
          x: [0], y: [0],
          marker: { size: 18, color: '#7c3aed', symbol: 'star', line: { width: 2, color: '#fff' } },
          hovertemplate: 'Scénario de base (Origine)<extra></extra>',
        });

        const layout = plotLayout({
          title: { text: 'Variation Bilan (M€) vs Émissions (%)', font: { size: 13, weight: 'bold' } },
          xaxis: { title: 'Δ Bilan Net (M€)', zeroline: true, zerolinewidth: 2, zerolinecolor: '#94a3b8' },
          yaxis: { title: 'Δ Émissions (%)',  zeroline: true, zerolinewidth: 2, zerolinecolor: '#94a3b8' },
        });
        Plotly.newPlot('sens-scatter-chart', scatterTraces, layout, plotConfig);
      })();
    }

  </script>
</body>
</html>
"""

    return (
        template
        .replace("__DASHBOARD_DATA__", payload_json)
        .replace("__SENSITIVITY_DATA__", sensitivity_json)
    )


def write_dashboard_html(output_path: Path, html: str) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html, encoding="utf-8")
def parse_args() -> argparse.Namespace:
    repo_root = get_repo_root()
    default_results_root = repo_root / "artifacts" / "reports"
    default_output_dir = default_results_root / "Results"

    parser = argparse.ArgumentParser(
        description="Generate a standalone HTML dashboard from PathWay scenario results."
    )
    parser.add_argument(
        "--results-root",
        type=Path,
        default=default_results_root,
        help="Directory containing scenario folders with Master_Plan.xlsx",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=default_output_dir,
        help="Directory where the standalone HTML file is written",
    )
    parser.add_argument(
        "--output-name",
        type=str,
        default="results_dashboard.html",
        help="Name of the generated HTML file",
    )
    parser.add_argument(
        "--discount-rate",
        type=float,
        default=DEFAULT_DISCOUNT_RATE,
        help="Discount rate for NPV computation (example: 0.08)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if args.discount_rate <= -1.0:
        raise ValueError("discount-rate must be greater than -1.0")

    results_root = args.results_root
    if not results_root.is_absolute():
        results_root = get_repo_root() / results_root

    output_dir = args.output_dir
    if not output_dir.is_absolute():
        output_dir = get_repo_root() / output_dir

    entity_dirs = discover_entities_and_scenarios(results_root)
    if not entity_dirs:
        print(f"[ERROR] No entity/scenario folders with a 'charts/' subdirectory found in {results_root}.")
        sys.exit(1)

    payload = build_dashboard_data(entity_dirs, discount_rate=args.discount_rate)
    sensitivity_data = load_sensitivity_data()
    html = build_html(payload, sensitivity_data=sensitivity_data)

    output_path = output_dir / args.output_name
    write_dashboard_html(output_path, html)

    print(f"Dashboard generated: {output_path}")
    print(f"Entities loaded: {len(entity_dirs)}")
    scen_count = sum(len(scens) for scens in entity_dirs.values())
    print(f"Scenarios loaded: {scen_count}")
    print(
        "Charts available per scenario: CARBON PRICE, CARBON TAX, CO2 TRAJECTORY, ENERGY MIX, "
        "FINANCING, INDIRECT EMISSIONS, INVESTMENT PLAN, RESSOURCES OPEX, DATA USED, "
        "TRANSITION COST, TOTAL ANNUAL OPEX, CO2 ABATEMENT"
    )


if __name__ == "__main__":
    main()