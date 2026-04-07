"""
run_sensitivity.py — Script d'analyse de sensibilité au prix du carbone (EUA)
===============================================================================
Ce script exécute une analyse de sensibilité One-At-a-Time (OAT) sur le scénario
de référence (BS) en faisant varier le prix du carbone (EUA) selon les paramètres
définis dans le bloc SENSITIVITY du fichier Excel d'entrée.

Usage :
    python scripts/run_sensitivity.py --excel-path "data/raw/excel/PathFinder input.xlsx"
    python scripts/run_sensitivity.py --excel-path "PathFinder input.xlsx" --output artifacts/sensitivity/sensitivity_results.json --verbose

Sortie :
    Un fichier JSON contenant les KPIs extraits pour chaque simulation.
"""

import argparse
import copy
import json
import os
import sys
import time
import threading
from pathlib import Path
from typing import Any, Dict, List, Optional

# ── Bootstrapping du chemin source (compatible exécution depuis la racine) ────
_REPO_ROOT = Path(__file__).resolve().parents[1]
_SRC_DIR   = _REPO_ROOT / "src"
if str(_SRC_DIR) not in sys.path:
    sys.path.insert(0, str(_SRC_DIR))

import pulp
from rich import print
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TimeElapsedColumn, MofNCompleteColumn, TaskProgressColumn

from pathway.core.ingestion import PathFinderParser
from pathway.core.optimizer import PathFinderOptimizer
from pathway.core.model import SensitivityParams


# ─────────────────────────────────────────────────────────────────────────────
# Génération des multiplicateurs symétriques
# ─────────────────────────────────────────────────────────────────────────────

def _build_multipliers(params: SensitivityParams) -> List[float]:
    """
    Génère la liste complète des multiplicateurs EUA à partir des amplitudes
    de variation et de la direction spécifiée (P, N ou ALL).

    Exemples :
        variations=[0.05, 0.10], direction='ALL'
        → [1.0, 0.90, 0.95, 1.05, 1.10]

        variations=[0.05, 0.10], direction='P'
        → [1.0, 1.05, 1.10]

        variations=[0.05, 0.10], direction='N'
        → [1.0, 0.90, 0.95]

    Le multiplicateur 1.0 (scénario de base, variation = 0%) est toujours inclus.
    """
    multipliers: List[float] = []

    for v in sorted(params.variations):  # par amplitude croissante
        if params.direction in ('N', 'ALL'):
            multipliers.append(round(1.0 - v, 6))   # variation négative ex: 0.95
        if params.direction in ('P', 'ALL'):
            multipliers.append(round(1.0 + v, 6))   # variation positive ex: 1.05

    # Le scénario de base (multiplicateur 1.0) est toujours ajouté en premier
    multipliers = [1.0] + sorted(set(multipliers))

    return multipliers


def _pct_from_multiplier(multiplier: float) -> float:
    """Convertit un multiplicateur en pourcentage de variation (ex: 0.95 → -5)."""
    return round((multiplier - 1.0) * 100.0, 4)


# ─────────────────────────────────────────────────────────────────────────────
# Extraction des KPIs depuis le solver
# ─────────────────────────────────────────────────────────────────────────────

def _extract_kpis(optimizer: PathFinderOptimizer) -> Dict[str, Any]:
    """
    Extrait les 4 KPIs demandés après une résolution réussie.

    KPIs retournés :
    - transition_cost      : Valeur de la fonction objectif PuLP (€)
    - average_co2_abatement: Réduction moyenne des émissions directes (%) vs base
    - gap_from_final_target: Pénalité résiduelle totale (0 si toutes les cibles atteintes)
    - total_emissions      : Somme des émissions totales sur tout l'horizon (tCO2)
    """
    opt = optimizer

    # ── 1. Coût de transition = valeur objectif du solver ─────────────────
    try:
        transition_cost = float(opt.model.objective.value() or 0.0)
    except Exception:
        transition_cost = 0.0

    # ── 2. Abattement moyen CO2 (%) ───────────────────────────────────────
    base_emissions = opt.entity.base_emissions if opt.entity else 0.0
    abatements = []
    if base_emissions > 0:
        for t in opt.years:
            e_var = opt.emis_vars.get(t)
            if e_var is not None:
                e_val = getattr(e_var, 'varValue', None)
                if e_val is not None:
                    abatement_pct = (base_emissions - float(e_val)) / base_emissions * 100.0
                    abatements.append(abatement_pct)
    average_co2_abatement = round(sum(abatements) / len(abatements), 4) if abatements else 0.0

    # ── 3. Gap par rapport à la cible finale ─────────────────────────────
    # Somme de toutes les pénalités résiduelles (0 si toutes les cibles atteintes).
    # Les clés peuvent être int ou (int, year) selon le mode objectif.
    gap_from_final_target = 0.0
    for p_var in opt.penalty_vars.values():
        p_val = getattr(p_var, 'varValue', None)
        if p_val is not None and float(p_val) > 1e-6:
            gap_from_final_target += float(p_val)
    gap_from_final_target = round(gap_from_final_target, 2)

    # ── 4. Émissions totales sur l'horizon (tCO2) ─────────────────────────
    total_emissions = 0.0
    for t in opt.years:
        te_var = opt.total_emis_vars.get(t)
        if te_var is not None:
            te_val = getattr(te_var, 'varValue', None)
            if te_val is not None:
                total_emissions += float(te_val)
    total_emissions = round(total_emissions, 2)

    return {
        "transition_cost": round(transition_cost, 2),
        "average_co2_abatement": average_co2_abatement,
        "gap_from_final_target": gap_from_final_target,
        "total_emissions": total_emissions,
    }


def _get_model_solution(optimizer: PathFinderOptimizer) -> Dict[str, float]:
    """Extrait tous les noms et valeurs des variables du modèle après résolution."""
    return {v.name: v.varValue for v in optimizer.model.variables() if v.varValue is not None}


# ─────────────────────────────────────────────────────────────────────────────
# Boucle principale de simulation OAT
# ─────────────────────────────────────────────────────────────────────────────

def run_sensitivity(
    excel_path: str,
    output_path: Optional[str] = None,
    verbose: bool = False,
) -> List[Dict[str, Any]]:
    """
    Exécute l'analyse de sensibilité OAT selon les paramètres du bloc SENSITIVITY
    du fichier Excel, et retourne la liste des résultats.

    Stratégie de résolution par simulation :
        · Si disponible : MIP Start basé sur la solution du scénario de base (multiplier 1.0)
        · Tentative 1 : time_limit configuré dans SENSITIVITY
        · Si aucun nœud réalisable trouvé → Tentative 2 : 2 × time_limit
        · Si toujours rien → simulation ignorée ('Skipped')

    Paramètres :
        excel_path  : Chemin vers le fichier Excel PathFinder input
        output_path : Chemin de sortie JSON
                      (défaut : artifacts/sensitivity/sensitivity_results.json)
        verbose     : Activer les logs détaillés

    Retour :
        Liste de dicts de résultats (un par simulation)
    """
    # ── Détermination du chemin de sortie ─────────────────────────────────
    if output_path is None:
        output_path = str(
            _REPO_ROOT / "artifacts" / "sensitivity" / "sensitivity_results.json"
        )

    excel_path_abs = str(Path(excel_path).resolve())
    if not Path(excel_path_abs).exists():
        raise FileNotFoundError(f"Fichier Excel introuvable : {excel_path_abs}")

    print(f"\n[bold cyan]═══ Analyse de Sensibilité PathFinder ═══[/bold cyan]")
    print(f"  Fichier Excel : [green]{excel_path_abs}[/green]")
    print(f"  Sortie JSON   : [green]{output_path}[/green]\n")

    # ── Lecture des paramètres de sensibilité ─────────────────────────────
    parser = PathFinderParser(excel_path_abs, verbose=verbose)
    sens_params = parser.parse_sensitivity()

    if not sens_params.variations:
        raise ValueError(
            "Aucune amplitude de variation trouvée dans le bloc SENSITIVITY."
        )
    if not sens_params.scenarios:
        raise ValueError(
            "Aucun scénario cible trouvé dans le bloc SENSITIVITY. "
            "Ajoutez BS dans la ligne SIM."
        )

    # ── Vérification des cibles actives ───────────────────────────────────
    active_targets = {k for k, v in sens_params.targets.items() if v}
    if not active_targets:
        raise ValueError(
            "Aucune cible de sensibilité active (YES) trouvée. "
            "Activez au moins 'EUA' dans les lignes DATA? du bloc SENSITIVITY."
        )

    print(f"  Cibles actives  : [yellow]{active_targets}[/yellow]")
    print(f"  Scénarios       : [yellow]{sens_params.scenarios}[/yellow]")
    print(f"  Variations      : [yellow]{sens_params.variations}[/yellow]")
    print(f"  Direction       : [yellow]{sens_params.direction}[/yellow]")
    print(f"  Temps/simulation: [yellow]{sens_params.time_limit}s (max {sens_params.time_limit*2}s)[/yellow]\n")

    # ── Génération des multiplicateurs ────────────────────────────────────
    multipliers = _build_multipliers(sens_params)
    print(f"  [bold]Multiplicateurs générés ({len(multipliers)}) :[/bold] {multipliers}\n")

    # ── Boucle de simulation ──────────────────────────────────────────────
    results: List[Dict[str, Any]] = []
    total_sims = len(sens_params.scenarios) * len(multipliers)

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
        MofNCompleteColumn(),
        TimeElapsedColumn(),
    ) as progress:
        task = progress.add_task("Simulations en cours...", total=total_sims)

        for scenario_id in sens_params.scenarios:
            print(
                f"  [cyan]Chargement du scénario "
                f"[bold]{scenario_id}[/bold]...[/cyan]"
            )
            base_data = parser.parse(scenario_id=scenario_id)
            base_prices = dict(base_data.time_series.carbon_prices)
            base_kpis: Optional[Dict[str, Any]] = None
            base_solution_data: Optional[Dict[str, float]] = None

            for multiplier in multipliers:
                variation_pct = _pct_from_multiplier(multiplier)
                label = (
                    f"Scénario {scenario_id} | "
                    f"EUA × {multiplier:.3f} ({variation_pct:+.1f}%)"
                )
                progress.update(task, description=label)

                # ── Perturbation des données ──────────────────────────────
                data_copy = copy.deepcopy(base_data)
                if 'EUA' in active_targets:
                    new_prices = {
                        yr: round(base_prices.get(yr, 0.0) * multiplier, 4)
                        for yr in base_prices
                    }
                    data_copy.time_series.carbon_prices = new_prices

                # ── Résolution MILP avec stratégie de robustesse (3 étapes) ──
                # 1. Tentative : Temps configuré, Gap standard (config)
                # 2. Tentative : 2 × Temps, Gap standard
                # 3. Tentative : 2 × Temps, Gap relaxé (0.10) pour forcer une solution
                base_time_limit = float(sens_params.time_limit)
                base_gap = data_copy.parameters.mip_gap or 0.005 # fallback safe
                
                opt = None
                status = 'Infeasible'
                
                attempts = [
                    {"time": base_time_limit,   "gap": base_gap, "label": "standard"},
                    {"time": base_time_limit*2, "gap": base_gap, "label": "double temps"},
                    {"time": base_time_limit*2, "gap": 0.10,     "label": "gap relaxé (10%)"}
                ]

                for i, retry in enumerate(attempts, start=1):
                    data_copy.parameters.time_limit = retry["time"]
                    data_copy.parameters.mip_gap = retry["gap"]

                    if i > 1:
                        print(
                            f"  [yellow][⚠] Tentative {i} ({retry['label']}) "
                            f"mult={multiplier:.3f}[/yellow]"
                        )

                    try:
                        # 1. Étape de construction (hors chronomètre solver)
                        opt = PathFinderOptimizer(data_copy, verbose=False)
                        opt.build_model()
                        if base_solution_data:
                            opt.apply_warm_start(base_solution_data)

                        # 2. Étape de résolution avec chronomètre "live"
                        # On ne crée la tâche qu'ici pour que le temps affiché corresponde au solveur
                        solver_task = progress.add_task(
                            f"    [dim]└─ Solver ({retry['label']}) [cyan]{retry['time']}s max[/cyan][/dim]", 
                            total=retry["time"]
                        )
                        
                        # Utilisation d'un thread pour animer la barre pendant que le solveur bloque le thread principal
                        stop_event = threading.Event()
                        start_time = time.time()

                        def update_ticker():
                            while not stop_event.is_set():
                                elapsed = time.time() - start_time
                                progress.update(solver_task, completed=min(elapsed, retry["time"]))
                                time.sleep(0.5)

                        ticker = threading.Thread(target=update_ticker, daemon=True)
                        ticker.start()

                        try:
                            status = opt.solve(warm_start=(base_solution_data is not None))
                        finally:
                            stop_event.set()
                            ticker.join(timeout=1.0)
                            # Finalisation propre de la barre
                            final_elapsed = time.time() - start_time
                            progress.update(solver_task, completed=final_elapsed)
                            progress.remove_task(solver_task)

                    except Exception as e:
                        print(f"  [red][!] Erreur tentative {i} (mult={multiplier:.3f}): {e}[/red]")
                        status = 'Error'
                        break

                    # Nœud réalisable trouvé → inutile de réessayer
                    if status in ('Optimal', 'Feasible'):
                        break

                # ── Aucune solution après les 2 tentatives ────────────────
                if status not in ('Optimal', 'Feasible'):
                    if status == 'Error':
                        msg = f"Erreur solveur après {i} tentative(s)"
                    else:
                        msg = f"Aucune solution après {len(attempts)} tentatives (abandon)"
                    print(
                        f"  [red][✗] {msg} — "
                        f"ignoré (mult={multiplier:.3f})[/red]"
                    )
                    results.append({
                        "target": "EUA",
                        "scenario": scenario_id,
                        "multiplier": multiplier,
                        "variation_pct": variation_pct,
                        "status": "Skipped",
                        "timed_out": False,
                        "transition_cost": None,
                        "average_co2_abatement": None,
                        "gap_from_final_target": None,
                        "total_emissions": None,
                    })
                    progress.advance(task)
                    continue

                # ── Extraction des KPIs ───────────────────────────────────
                # 'Feasible' = CBC a atteint la limite de temps mais a trouvé
                # un nœud réalisable → on utilise la meilleure solution connue,
                # exactement comme le fait la simulation normale.
                timed_out = (status == 'Feasible')
                kpis = _extract_kpis(opt)

                if multiplier == 1.0 and base_kpis is None:
                    base_kpis = kpis.copy()
                    base_solution_data = _get_model_solution(opt)

                if verbose:
                    icon = "⏱" if timed_out else "✓"
                    suffix = " [meilleur nœud]" if timed_out else ""
                    print(
                        f"  [green][{icon}][/green] mult={multiplier:.3f} | "
                        f"Coût={kpis['transition_cost']:,.0f}€{suffix} | "
                        f"CO2 abattu={kpis['average_co2_abatement']:.1f}% | "
                        f"Émissions={kpis['total_emissions']:,.0f}tCO2"
                    )

                results.append({
                    "target": "EUA",
                    "scenario": scenario_id,
                    "multiplier": multiplier,
                    "variation_pct": variation_pct,
                    "status": status,
                    "timed_out": timed_out,
                    **kpis,
                })
                progress.advance(task)

    # ── Export JSON ───────────────────────────────────────────────────────
    output_dir = Path(output_path).parent
    output_dir.mkdir(parents=True, exist_ok=True)

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    valid_results = [r for r in results if r.get("status") in ('Optimal', 'Feasible')]
    skipped = [r for r in results if r.get("status") == 'Skipped']
    print(
        f"\n[bold green]✓ Analyse terminée — "
        f"{len(valid_results)}/{len(results)} simulations valides[/bold green]"
    )
    if skipped:
        print(
            f"  [yellow]⚠ {len(skipped)} simulation(s) ignorée(s) "
            f"(aucune solution après 2 tentatives)[/yellow]"
        )
    print(f"  Résultats exportés vers : [cyan]{output_path}[/cyan]\n")

    return results


# ─────────────────────────────────────────────────────────────────────────────
# Point d'entrée CLI
# ─────────────────────────────────────────────────────────────────────────────

def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Analyse de sensibilité OAT au prix du carbone (EUA) — PathFinder"
    )
    parser.add_argument(
        "--excel-path",
        type=str,
        default=str(_REPO_ROOT / "data" / "raw" / "excel" / "PathFinder input.xlsx"),
        help=(
            "Chemin vers le fichier Excel d'entrée PathFinder "
            "(défaut: data/raw/excel/PathFinder input.xlsx)"
        ),
    )
    parser.add_argument(
        "--output",
        type=str,
        default=None,
        help=(
            "Chemin de sortie du fichier JSON "
            "(défaut: artifacts/sensitivity/sensitivity_results.json)"
        ),
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Activer les logs détaillés",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = _parse_args()
    run_sensitivity(
        excel_path=args.excel_path,
        output_path=args.output,
        verbose=args.verbose,
    )
