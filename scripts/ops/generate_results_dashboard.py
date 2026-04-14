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


def discover_scenarios(results_root: Path) -> Dict[str, Path]:
    """Scans the results directory for scenario folders containing a charts/ subdirectory."""
    roots = [results_root, results_root / "Results"]
    scenarios: Dict[str, Path] = {}

    for root in roots:
        if not root.exists() or not root.is_dir():
            continue
        for child in sorted(root.iterdir(), key=lambda p: p.name.lower()):
            if not child.is_dir():
                continue
            if child.name.lower() == "results":
                continue
            # We look for the charts directory which contains our Plotly JSON artifacts
            charts_dir = child / "charts"
            if charts_dir.exists() and charts_dir.is_dir():
                if child.name not in scenarios:
                    scenarios[child.name] = child

    return scenarios


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


def build_dashboard_data(scenario_dirs: Dict[str, Path], discount_rate: float) -> Dict[str, Any]:
    """Assembles the final dashboard payload by consuming pre-generated JSON artifacts."""
    if not scenario_dirs:
        print("[ERROR] No scenario directories with 'charts/' subfolder were found.")
        return {}

    # Determine generation date based on latest chart update
    latest_ts = 0.0
    for s_path in scenario_dirs.values():
        charts_dir = s_path / "charts"
        for f in charts_dir.glob("*.json"):
            latest_ts = max(latest_ts, f.stat().st_mtime)
    
    generation_date = datetime.fromtimestamp(latest_ts).strftime("%Y-%m-%d %H:%M:%S") if latest_ts > 0 else "Unknown"

    scenarios_payload: Dict[str, Any] = {}
    
    # Define mapping of dashboard keys to chart filenames
    # This must match self._save_plotly_figure calls in ReportingEngine
    chart_mapping = [
        ("co2_trajectory",    "CO2 TRAJECTORY",      "CO2_Trajectory"),
        ("indirect_emissions", "INDIRECT EMISSIONS", "Indirect_Emissions"),
        ("energy_mix",        "ENERGY MIX",          "Energy_Mix"),
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

    for scenario_name, scenario_path in scenario_dirs.items():
        print(f"  > Processing scenario: {scenario_name}")
        graphs = {}
        for key, label, filename in chart_mapping:
            graphs[key] = get_graph_payload(scenario_path, scenario_name, key, label, filename)

        scenarios_payload[scenario_name] = {
            "displayName": scenario_name.replace("_", " "),
            "sourcePath": str(scenario_path),
            "graphs": graphs,
        }

    payload = {
        "projectTitle": "Plant-Optimization-PathWay - Results Dashboard",
        "generationDate": generation_date,
        "discountRate": discount_rate,
        "scenarios": scenarios_payload,
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
  <link rel=\"stylesheet\" href=\"https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css\" integrity=\"sha512-SnH5WK+bZxgPHs44uWix+LLJAJ9/2PkPKZ5QiAj6Ta86w+fsb2TkR4j8f5Z5gDQL4x0XSLwWf2fQJKfG8d8gQw==\" crossorigin=\"anonymous\" referrerpolicy=\"no-referrer\" />
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

    .selector {
      width: 100%;
      border: 1px solid rgba(15, 23, 42, 0.12);
      border-radius: 0.9rem;
      padding: 0.75rem 0.9rem;
      background: #ffffff;
      color: #1e293b;
      font-size: 0.95rem;
      outline: none;
      transition: all 200ms ease;
    }

    .selector:focus {
      border-color: #0ea5e9;
      box-shadow: 0 0 0 3px rgba(14, 165, 233, 0.2);
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

      <section class=\"glass-card rounded-3xl p-5 md:p-6 mb-6 fade-in\">
        <div class=\"grid grid-cols-1 md:grid-cols-2 gap-4\">
          <label class=\"block\">
            <span class=\"text-sm font-bold text-slate-500 uppercase tracking-tight\">Scenario Selector</span>
            <select id=\"scenarioSelect\" class=\"selector mt-2\"></select>
          </label>

          <label class=\"block\">
            <span class=\"text-sm font-bold text-slate-500 uppercase tracking-tight\">Graph Selector</span>
            <select id=\"graphSelect\" class=\"selector mt-2\"></select>
          </label>
        </div>
      </section>

      <section class=\"glass-card rounded-3xl p-4 md:p-6 mb-6 fade-in\">
        <div class=\"flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 mb-3\">
          <h2 id=\"graphTitle\" class=\"font-heading text-lg md:text-2xl font-bold text-slate-800\"></h2>
          <button id=\"downloadBtn\" class=\"primary-btn inline-flex items-center justify-center gap-2\">
            <i class=\"fa-solid fa-download\"></i>
            Download Chart as Image
          </button>
        </div>
        <div id=\"chart\"></div>
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
            Analyse de Risque — One-At-a-Time (OAT)
          </p>
          <h2 class=\"font-heading text-2xl md:text-3xl font-bold mt-3\">Analyse de Sensibilité au Prix EUA</h2>
          <p class=\"text-slate-500 mt-2 text-sm\">
            Variation symétrique du prix du carbone (EUA) appliquée au scénario de référence (BS).
            Chaque point représente une simulation MILP indépendante.
          </p>
        </div>
        <div id=\"sens-status-badge\" class=\"rounded-2xl subtle-pill px-4 py-3 text-sm\"></div>
      </div>
    </section>

    <!-- Grille des 4 graphiques -->
    <div class=\"grid grid-cols-1 gap-6\">

      <!-- 1. Vue Globale des Risques (Packed Bubble) -->
      <section class=\"glass-card rounded-3xl p-5 md:p-6 fade-in\">
        <h3 class=\"font-heading text-lg font-bold text-slate-800 mb-1\">Vue Globale des Risques</h3>
        <p class=\"text-xs text-slate-500 mb-3\">
          Chaque bulle représente un paramètre perturbé. Le rayon est proportionnel à la variance maximale du coût de transition.
        </p>
        <div id=\"sens-bubble-chart\" style=\"height:500px;\"></div>
      </section>

      <!-- 2. Tornado Chart (Impact Financier) -->
      <section class=\"glass-card rounded-3xl p-5 md:p-6 fade-in\">
        <h3 class=\"font-heading text-lg font-bold text-slate-800 mb-1\">Impact Financier (Tornado)</h3>
        <p class=\"text-xs text-slate-500 mb-3\">
          Barres horizontales signant l'écart du coût de transition par rapport au scénario de base pour les variations extrêmes testées.
        </p>
        <div id=\"sens-tornado-chart\" style=\"height:500px;\"></div>
      </section>

      <!-- 3. Trajectoires de Décarbonation (Net CO2) -->
      <section class=\"glass-card rounded-3xl p-5 md:p-6 fade-in\">
        <div class=\"grid grid-cols-1 lg:grid-cols-2 gap-12\">
          <!-- Gauche: Trajectoire Temporelle -->
          <div>
            <h3 class=\"font-heading text-lg font-bold text-slate-800 mb-1\">Trajectoires de Décarbonation</h3>
            <p class=\"text-xs text-slate-500 mb-3\">
              Projection annuelle des émissions nettes (Scope 1 + Scope 2 - Captage DAC - Crédits) pour chaque scénario de variation.
            </p>
            <div id=\"sens-trajectory-chart\" style=\"height:500px;\"></div>
          </div>
          <!-- Droite: Sensibilité Totale -->
          <div>
            <h3 class=\"font-heading text-lg font-bold text-slate-800 mb-1\">Émissions Totales vs Variation</h3>
            <p class=\"text-xs text-slate-500 mb-3\">
              Impact cumulé sur les émissions totales sur l'horizon en fonction du pourcentage de variation du paramètre ciblé.
            </p>
            <div id=\"sens-total-co2-chart\" style=\"height:500px;\"></div>
          </div>
        </div>
      </section>

      <!-- 4. Scatter Coût vs CO₂ -->
      <section class=\"glass-card rounded-3xl p-5 md:p-6 fade-in\">
        <h3 class=\"font-heading text-lg font-bold text-slate-800 mb-1\">Coût vs Émissions CO₂</h3>
        <p class=\"text-xs text-slate-500 mb-3\">
          Chaque point est une simulation. L'axe X représente la variation du coût de transition (%) et l'axe Y la variation des émissions totales (%).
        </p>
        <div id=\"sens-scatter-chart\" style=\"height:500px;\"></div>
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

    function scenarioKeys() {
      return Object.keys(dashboardData.scenarios || {});
    }

    function graphKeysForScenario(scenarioKey) {
      const scenario = dashboardData.scenarios[scenarioKey] || {};
      return Object.keys(scenario.graphs || {});
    }

    function fillScenarioSelect() {
      const keys = scenarioKeys();
      scenarioSelect.innerHTML = '';
      keys.forEach((key) => {
        const option = document.createElement('option');
        option.value = key;
        option.textContent = dashboardData.scenarios[key].displayName || key;
        scenarioSelect.appendChild(option);
      });
    }

    function fillGraphSelect(selectedScenario) {
      const currentSelection = graphSelect.value;
      const keys = graphKeysForScenario(selectedScenario);
      graphSelect.innerHTML = '';
      keys.forEach((key) => {
        const graph = dashboardData.scenarios[selectedScenario].graphs[key];
        const option = document.createElement('option');
        option.value = key;
        option.textContent = graph.label || key;
        graphSelect.appendChild(option);
      });
      if (currentSelection && keys.includes(currentSelection)) {
        graphSelect.value = currentSelection;
      }
    }

    async function renderGraph() {
      const scenarioKey = scenarioSelect.value;
      const graphKey = graphSelect.value;
      const scenario = dashboardData.scenarios[scenarioKey] || { graphs: {} };
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

    async function handleScenarioChange() {
      fillGraphSelect(scenarioSelect.value);
      await renderGraph();
    }

    async function handleGraphChange() {
      await renderGraph();
    }

    function handleDownload() {
      const scenarioKey = scenarioSelect.value;
      const graphKey = graphSelect.value;
      const scenario = dashboardData.scenarios[scenarioKey] || {};
      const graph = (scenario.graphs || {})[graphKey] || {};
      const fallback = [scenarioKey, graphKey].filter(Boolean).join('_') || 'chart';
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

      const keys = scenarioKeys();
      if (!keys.length) {
        graphTitle.textContent = 'No scenarios found';
        graphMethod.textContent = 'No scenario workbook was discovered. Please generate results first.';
        await Plotly.newPlot(chartNode, [], { paper_bgcolor: 'rgba(255,255,255,0)' }, plotConfig);
        scenarioSelect.disabled = true;
        graphSelect.disabled = true;
        downloadBtn.disabled = true;
        return;
      }

      fillScenarioSelect();
      await handleScenarioChange();
    }

    scenarioSelect.addEventListener('change', handleScenarioChange);
    graphSelect.addEventListener('change', handleGraphChange);
    downloadBtn.addEventListener('click', handleDownload);
    window.addEventListener('resize', () => {
      if (chartNode && chartNode.data) {
        Plotly.Plots.resize(chartNode);
      }
    });

    init();

    // ═══════════════════════════════════════════════════════════════════════
    // GRAPHIQUES DE SENSIBILITÉ
    // ═══════════════════════════════════════════════════════════════════════

    (function buildSensitivityCharts() {
      const data = sensitivityData || [];

      // Badge de statut
      const badge = document.getElementById('sens-status-badge');
      if (badge) {
        if (data.length === 0) {
          badge.innerHTML = '<span style="color:#d97706;"><i class="fa-solid fa-triangle-exclamation"></i> Aucune donnée — Exécutez run_sensitivity.py</span>';
        } else {
          const validCount = data.filter(r => r.status === 'Optimal' || r.status === 'Feasible').length;
          const shortfallCount = data.filter(r => (r.penalty_cost || 0) > 1.0).length;
          let html = `<span style="color:#16a34a;"><i class="fa-solid fa-circle-check"></i> ${validCount} / ${data.length} simulations valides</span>`;
          if (shortfallCount > 0) {
            html += ` <span style="color:#dc2626; margin-left:10px;"><i class="fa-solid fa-circle-exclamation"></i> ${shortfallCount} cibles non atteintes (Pénalités réduites dans les graphiques)</span>`;
          }
          badge.innerHTML = html;
        }
      }

      if (data.length === 0) return;

      // ── Données valides uniquement ──────────────────────────────────────
      const valid = data.filter(r => r.transition_cost != null);

      // Paramètre de base (variation = 0)
      const baseRecord = valid.find(r => Math.abs(r.variation_pct) < 0.001) || valid[0];
      
      // transition_cost est désormais déjà le NET BALANCE (M€ ou €)
      // On convertit en M€ pour l'affichage
      const baseCost = baseRecord ? baseRecord.transition_cost / 1_000_000.0 : 0;
      const baseEmis = baseRecord ? (baseRecord.total_emissions || 0) : 0;

      // Groupement par cible
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
        margin: { l: 56, r: 24, t: 32, b: 50 },
        legend: { orientation: 'h', y: -0.15, font: { size: 10 } },
        hovermode: 'closest',
      }, extra || {});

      // ── 1. Packed Bubble (Vue Globale) ───────────────────────────────────
      (function buildBubble() {
        const bubbleTraces = Object.entries(targets).map(([target, records]) => {
          const costs = records.map(r => r.transition_cost / 1_000_000.0);
          const maxVariance = Math.max(...costs) - Math.min(...costs);
          
          return {
            type: 'scatter',
            mode: 'markers+text',
            name: target,
            x: [0],
            y: [0],
            text: [target],
            textposition: 'middle center',
            textfont: { size: 13, color: '#fff' },
            marker: {
              size: [Math.max(60, Math.min(180, maxVariance * 5))], // Scale factors
              sizemode: 'diameter',
              color: ['#0ea5e9'],
              opacity: 0.85,
              line: { width: 2, color: '#fff' },
            },
            hovertemplate: `<b>${target}</b><br>Variance max : %{customdata:,.1f} M€<extra></extra>`,
            customdata: [maxVariance],
          };
        });
        const layout = plotLayout({
          title: { text: 'Variance maximale du Bilan de Transition (M€)', font: { size: 13 } },
          xaxis: { visible: false, zeroline: false },
          yaxis: { visible: false, zeroline: false },
          showlegend: false,
        });
        Plotly.newPlot('sens-bubble-chart', bubbleTraces, layout, plotConfig);
      })();

      // ── 2. Tornado Chart (Impact Financier) ─────────────────────────────
      (function buildTornado() {
        const tornadoTraces = [];

        Object.entries(targets).forEach(([target, records], idx) => {
          const sorted = [...records].sort((a, b) => a.variation_pct - b.variation_pct);
          const minRec = sorted[0];
          const maxRec = sorted[sorted.length - 1];

          const deltaMin = (minRec && minRec.transition_cost != null) ? (minRec.transition_cost / 1_000_000.0) - baseCost : 0;
          const deltaMax = (maxRec && maxRec.transition_cost != null) ? (maxRec.transition_cost / 1_000_000.0) - baseCost : 0;

          tornadoTraces.push({
            type: 'bar',
            orientation: 'h',
            name: `${target} (variation basse ${minRec ? minRec.variation_pct.toFixed(0) : ''}%)`,
            y: [target],
            x: [deltaMin],
            marker: { color: deltaMin < 0 ? '#16a34a' : '#dc2626' },
            hovertemplate: `<b>${target}</b><br>Variation : ${minRec ? minRec.variation_pct.toFixed(1) : 0}%<br>Δ Bilan : %{x:,.1f} M€<extra></extra>`,
          });
          tornadoTraces.push({
            type: 'bar',
            orientation: 'h',
            name: `${target} (variation haute ${maxRec ? maxRec.variation_pct.toFixed(0) : ''}%)`,
            y: [target],
            x: [deltaMax],
            marker: { color: deltaMax < 0 ? '#16a34a' : '#dc2626' },
            hovertemplate: `<b>${target}</b><br>Variation : ${maxRec ? maxRec.variation_pct.toFixed(1) : 0}%<br>Δ Bilan : %{x:,.1f} M€<extra></extra>`,
          });
        });

        const layout = plotLayout({
          barmode: 'overlay',
          title: { text: 'Impact sur le Bilan Net / Scénario de Base', font: { size: 13 } },
          xaxis: { title: 'Δ Bilan Net de Transition (M€)', zeroline: true, zerolinewidth: 2, zerolinecolor: '#94a3b8' },
          yaxis: { automargin: true },
        });
        Plotly.newPlot('sens-tornado-chart', tornadoTraces, layout, plotConfig);
      })();

      // ── 3. Trajectoires et Sensibilité CO₂ ───────────────────────────────
      (function buildDecarbonizationViews() {
        const trajectoryTraces = [];
        const summaryTraces = [];
        
        // ── 3a. Trajectoires Temporelles (Gauche) ──────────────────────────
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
          title: { text: 'Projection Temporelle du Net CO₂', font: { size: 13 } },
          xaxis: { title: 'Année' },
          yaxis: { title: 'Net CO₂ (t)', zeroline: true },
          showlegend: true,
          legend: { orientation: 'h', y: -0.25 }
        });
        
        Plotly.newPlot('sens-trajectory-chart', trajectoryTraces, trajectoryLayout, plotConfig);

        // ── 3b. Émissions Totales vs Variation (Droite) ────────────────────
        Object.entries(targets).forEach(([target, records]) => {
          const sorted = [...records].sort((a, b) => a.variation_pct - b.variation_pct);
          
          summaryTraces.push({
            type: 'scatter',
            mode: 'markers',
            name: target,
            x: sorted.map(r => r.variation_pct),
            y: sorted.map(r => r.total_emissions),
            marker: { size: 10, line: { width: 1, color: '#fff' } },
            hovertemplate: `<b>${target}</b><br>Var. Paramètre : %{x:+.1f}%<br>Émissions Totales : %{y:,.0f} t<extra></extra>`
          });
        });

        const summaryLayout = plotLayout({
          title: { text: 'Sensibilité : Émissions Totales vs Variation Paramètre', font: { size: 13 } },
          xaxis: { title: 'Variation du Paramètre (%)', zeroline: true, zerolinewidth: 2, zerolinecolor: '#94a3b8' },
          yaxis: { title: 'Émissions Totales Horizon (tCO₂)', gridcolor: '#f1f5f9' },
          showlegend: true,
          legend: { orientation: 'h', y: -0.25 }
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
              colorbar: { title: { text: 'Var. EUA (%)' }, thickness: 12, len: 0.7 },
              line: { width: 1.5, color: '#fff' },
            },
            text: filtered.map(r => {
              const deltaBalance = (r.transition_cost / 1_000_000.0) - baseCost;
              const deltaE = (r.total_emissions - baseEmis) / Math.abs(baseEmis) * 100;
              return `${r.timed_out ? '⏱ Temps limité<br>' : ''}Paramètre : ${target}<br>` +
                     `Variation : ${r.variation_pct.toFixed(1)}%<br>` +
                     `<b>Δ Bilan Net de Transition : ${deltaBalance >= 0 ? '+' : ''}${deltaBalance.toLocaleString('fr-FR', { maximumFractionDigits: 1 })} M€</b><br>` +
                     `Δ Émissions : ${deltaE >= 0 ? '+' : ''}${deltaE.toFixed(4)}%<br>` +
                     `Émissions : ${r.total_emissions.toLocaleString('fr-FR')} tCO₂` +
                     (r.penalty_cost > 1.0 ? `<br><span style="color:red">⚠️ Cible non atteinte (Gap: ${r.gap_from_final_target.toFixed(0)}t)</span>` : '');
            }),
            hovertemplate: '%{text}<extra></extra>',
          };
        });

        // Ligne de base (Origine)
        scatterTraces.push({
          type: 'scatter',
          mode: 'markers',
          name: 'Baseline',
          x: [0], y: [0],
          marker: { size: 18, color: '#7c3aed', symbol: 'star', line: { width: 2, color: '#fff' } },
          hovertemplate: 'Scénario de base (Origine)<br>Δ Bilan : 0 M€<br>Δ Émissions : 0%<extra></extra>',
        });

        const layout = plotLayout({
          title: { text: 'Variation du Bilan Net (M€) vs Variation des Émissions (%)', font: { size: 13 } },
          xaxis: { title: 'Δ Bilan Net de Transition (M€ / Base)', zeroline: true, zerolinewidth: 2, zerolinecolor: '#94a3b8', tickformat: ',.0f' },
          yaxis: { title: 'Δ Émissions Totales (%)',  zeroline: true, zerolinewidth: 2, zerolinecolor: '#94a3b8', tickformat: '.3f' },
          shapes: [
            { type: 'line', x0: 0, x1: 0, y0: -1, y1: 1, line: { color: '#cbd5e1', width: 1, dash: 'dot' } },
          ],
        });
        Plotly.newPlot('sens-scatter-chart', scatterTraces, layout, plotConfig);
      })();

    })();

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

    scenario_dirs = discover_scenarios(results_root)
    if not scenario_dirs:
        print(f"[ERROR] No scenario folders with a 'charts/' subdirectory found in {results_root}.")
        sys.exit(1)

    payload = build_dashboard_data(scenario_dirs, discount_rate=args.discount_rate)
    sensitivity_data = load_sensitivity_data()
    html = build_html(payload, sensitivity_data=sensitivity_data)

    output_path = output_dir / args.output_name
    write_dashboard_html(output_path, html)

    print(f"Dashboard generated: {output_path}")
    print(f"Scenarios loaded: {len(scenario_dirs)}")
    print(
        "Charts available per scenario: CARBON PRICE, CARBON TAX, CO2 TRAJECTORY, ENERGY MIX, "
        "FINANCING, INDIRECT EMISSIONS, INVESTMENT PLAN, RESSOURCES OPEX, DATA USED, "
        "TRANSITION COST, TOTAL ANNUAL OPEX, CO2 ABATEMENT"
    )


if __name__ == "__main__":
    main()