"""
Core Sensitivity Analysis Engine
================================
Provides OAT (One-At-a-Time) sensitivity analysis for PathFinder simulations.
"""

import copy
import json
import os
import sys
import time
import threading
import concurrent.futures
from pathlib import Path
from typing import Any, Dict, List, Optional

import pulp
import pandas as pd
from rich import print
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TimeElapsedColumn, MofNCompleteColumn, TaskProgressColumn

from pathway.core.ingestion import PathFinderParser
from pathway.core.optimizer import PathFinderOptimizer
from pathway.core.model import SensitivityParams

def _build_multipliers(params: SensitivityParams) -> List[float]:
    """Generates the list of multipliers based on variations and direction (P, N, ALL)."""
    multipliers: List[float] = []
    for v in sorted(params.variations):
        if params.direction in ('N', 'ALL'):
            multipliers.append(round(1.0 - v, 6))
        if params.direction in ('P', 'ALL'):
            multipliers.append(round(1.0 + v, 6))
    return [1.0] + sorted(set(multipliers))

def _pct_from_multiplier(multiplier: float) -> float:
    return round((multiplier - 1.0) * 100.0, 4)

def _safe_val(value: Any, multiplier: float) -> Any:
    """Multiplies value if numeric, else returns original (for 'LINEAR INTER' keywords)."""
    try:
        return round(float(value) * multiplier, 4)
    except (ValueError, TypeError):
        return value

def _calculate_baseline_reference(optimizer: PathFinderOptimizer) -> float:
    """Calculates the BAU (Business-As-Usual) total cost for baseline reference."""
    opt = optimizer
    data = opt.data
    years = list(opt.years)
    total_baseline_cost = 0.0

    for idx, t in enumerate(years):
        b_emissions = 0.0
        b_consumptions = {res_id: 0.0 for res_id in data.resources}

        for p_id, process in opt.entity.processes.items():
            up_rate = 0.0
            if 'UP' in process.valid_technologies:
                up_tech = data.technologies['UP']
                up_imp = up_tech.impacts.get('ALL') or up_tech.impacts.get('CO2_EM')
                if up_imp and up_imp['type'] in ['variation', 'up']:
                    up_rate = abs(up_imp['value'])
            
            up_factor = (1.0 - up_rate) ** idx
            p_base_emis = (opt.entity.base_emissions * process.emission_shares.get('CO2_EM', 0.0)) * up_factor
            b_emissions += p_base_emis
            for res_id in data.resources:
                if res_id != 'CO2_EM':
                    p_base_cons = (opt.entity.base_consumptions.get(res_id, 0.0) * process.consumption_shares.get(res_id, 0.0)) * up_factor
                    b_consumptions[res_id] += p_base_cons

        allocated_emis_share = sum(p.emission_shares.get('CO2_EM', 0.0) for p in opt.entity.processes.values())
        b_emissions += opt.entity.base_emissions * (1.0 - allocated_emis_share)
        for res_id in data.resources:
            if res_id != 'CO2_EM':
                allocated_share = sum(p.consumption_shares.get(res_id, 0.0) for p in opt.entity.processes.values())
                b_consumptions[res_id] += opt.entity.base_consumptions.get(res_id, 0.0) * (1.0 - allocated_share)

        tax_price = data.time_series.carbon_prices.get(t, 0.0)
        mode = getattr(opt.entity, 'sv_act_mode', 'NORM')
        fq_pct = data.time_series.carbon_quotas_pi.get(t, 0.0) if mode == 'PI' else data.time_series.carbon_quotas_norm.get(t, 0.0)
        
        taxed_co2_b = b_emissions * (1.0 - fq_pct) if fq_pct <= 1.0 else max(0.0, b_emissions - fq_pct)
        total_baseline_cost += taxed_co2_b * tax_price
        for res_id, cons_val in b_consumptions.items():
            price = data.time_series.resource_prices.get(res_id, {}).get(t, 0.0)
            if price != 0: total_baseline_cost += cons_val * price

    return total_baseline_cost

def _extract_abatements(opt: PathFinderOptimizer) -> float:
    base_emissions = opt.entity.base_emissions if opt.entity else 0.0
    abatements = []
    if base_emissions > 0:
        for t in opt.years:
            val = getattr(opt.emis_vars.get(t), 'varValue', None)
            if val is not None:
                abatements.append((base_emissions - float(val)) / base_emissions * 100.0)
    return round(sum(abatements) / len(abatements), 4) if abatements else 0.0

def _extract_kpis(optimizer: PathFinderOptimizer) -> Dict[str, Any]:
    """Extracts summary KPIs from a solved optimizer instance."""
    opt = optimizer
    try: objective_val = float(opt.model.objective.value() or 0.0)
    except Exception: objective_val = 0.0

    penalty_cost = 0.0
    massive_penalty = getattr(opt, 'massive_penalty_cost', 1e9)
    for idx, obj in enumerate(opt.data.objectives):
        if obj.penalty_type == "NONE": continue
        
        is_realistic = (obj.penalty_type == "PENALTIES")
        
        p_var = opt.penalty_vars.get(idx)
        if p_var is not None and float(getattr(p_var, 'varValue', 0.0) or 0.0) > 1e-6:
            p_cost = massive_penalty
            if is_realistic:
                t_target = min(opt.years[-1], max(opt.years[0], obj.target_year))
                c_price = opt.data.time_series.carbon_prices.get(t_target, 0.0)
                p_fact = opt.data.time_series.carbon_penalties.get(t_target, 0.0)
                p_cost = c_price * (1.0 + p_fact)
            penalty_cost += float(p_var.varValue) * p_cost
            
        for t in opt.years:
            pt_var = opt.penalty_vars.get((idx, t))
            if pt_var is not None and float(getattr(pt_var, 'varValue', 0.0) or 0.0) > 1e-6:
                p_cost = massive_penalty
                if is_realistic:
                    c_price = opt.data.time_series.carbon_prices.get(t, 0.0)
                    p_fact = opt.data.time_series.carbon_penalties.get(t, 0.0)
                    p_cost = c_price * (1.0 + p_fact)
                penalty_cost += float(pt_var.varValue) * p_cost

    baseline_cost = _calculate_baseline_reference(opt)
    real_cost = objective_val - penalty_cost
    transition_balance = real_cost - baseline_cost

    carbon_tax = 0.0
    for t in opt.years:
        c_price = opt.data.time_series.carbon_prices.get(t, 0.0)
        p_fact = opt.data.time_series.carbon_penalties.get(t, 0.0)
        pq_v = getattr(opt.paid_quota_vars.get(t), 'varValue', 0.0) or 0.0
        pnq_v = getattr(opt.penalty_quota_vars.get(t), 'varValue', 0.0) or 0.0
        carbon_tax += (pq_v * c_price) + (pnq_v * c_price * (1.0 + p_fact))

    gap = sum(float(getattr(v, 'varValue', 0.0) or 0.0) for v in opt.penalty_vars.values() if float(getattr(v, 'varValue', 0.0) or 0.0) > 1e-6)
    total_emis = sum(float(getattr(v, 'varValue', 0.0) or 0.0) for v in opt.total_emis_vars.values())

    co2_traj = []
    for t in opt.years:
        net_t = (float(getattr(opt.total_emis_vars.get(t), 'varValue', 0.0) or 0.0) - 
                 float(getattr(opt.dac_captured_vars.get(t), 'varValue', 0.0) or 0.0) - 
                 float(getattr(opt.credit_purchased_vars.get(t), 'varValue', 0.0) or 0.0))
        co2_traj.append(round(net_t, 2))

    return {
        "transition_cost": round(transition_balance, 2),
        "real_cost": round(real_cost, 2),
        "baseline_cost": round(baseline_cost, 2),
        "total_objective": round(objective_val, 2),
        "penalty_cost": round(penalty_cost, 2),
        "carbon_tax_cost": round(carbon_tax, 2),
        "average_co2_abatement": _extract_abatements(opt),
        "gap_from_final_target": round(gap, 2),
        "total_emissions": round(total_emis, 2),
        "co2_trajectory": {"years": list(opt.years), "values": co2_traj}
    }

def _get_model_solution(optimizer: PathFinderOptimizer) -> Dict[str, float]:
    return {v.name: v.varValue for v in optimizer.model.variables() if v.varValue is not None}

def run_sensitivity(excel_path: str, output_path: Optional[str] = None, verbose: bool = False) -> List[Dict[str, Any]]:
    """Runs OAT sensitivity analysis and exports results to JSON."""
    parser = PathFinderParser(excel_path, verbose=verbose)
    sens_params = parser.parse_sensitivity()

    if not sens_params.variations or not sens_params.scenarios:
        if verbose: print("[yellow]No sensitivity parameters found. Skipping.[/yellow]")
        return []

    active_targets = {k for k, v in sens_params.targets.items() if v}
    if not active_targets:
        if verbose: print("[yellow]No active sensitivity targets. Skipping.[/yellow]")
        return []

    multipliers = _build_multipliers(sens_params)
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
        task = progress.add_task("Sensitivities...", total=total_sims)
        for scenario_id in sens_params.scenarios:
            base_data = parser.parse(scenario_id=scenario_id)
            base_sol = None
            for mult in multipliers:
                var_pct = _pct_from_multiplier(mult)
                label = f"S:{scenario_id} x{mult:.3f} ({var_pct:+.1f}%)"
                progress.update(task, description=label)

                data_copy = copy.deepcopy(base_data)
                start_yr = data_copy.parameters.start_year
                years_list = list(data_copy.time_series.carbon_prices.keys())

                if 'EUA' in active_targets:
                    new_anchors = {yr: (_safe_val(val, mult) if yr > start_yr else val) 
                                   for yr, val in data_copy.time_series.carbon_prices_anchors.items()}
                    data_copy.time_series.carbon_prices = parser._interpolate_dict(new_anchors, years_list)
                
                if 'RESOURCE PRICE' in active_targets:
                    for r_id, a in data_copy.time_series.resource_prices_anchors.items():
                        new_a = {yr: (_safe_val(v, mult) if yr > start_yr else v) for yr, v in a.items()}
                        data_copy.time_series.resource_prices[r_id] = parser._interpolate_dict(new_a, years_list)

                # CAPEX/OPEX simplified temporal sensitivity
                for tech in data_copy.technologies.values():
                    if 'CAPEX' in active_targets and tech.capex_anchors:
                        new_a = {yr: (_safe_val(v, mult) if yr > start_yr else v) for yr, v in tech.capex_anchors.items()}
                        tech.capex_by_year = parser._interpolate_dict(new_a, years_list)
                    if 'OPEX' in active_targets and tech.opex_anchors:
                        new_a = {yr: (_safe_val(v, mult) if yr > start_yr else v) for yr, v in tech.opex_anchors.items()}
                        tech.opex_by_year = parser._interpolate_dict(new_a, years_list)

                limit = float(sens_params.time_limit) * 2
                data_copy.parameters.time_limit = limit
                opt = PathFinderOptimizer(data_copy, verbose=False)
                opt.build_model()
                if base_sol: opt.apply_warm_start(base_sol)

                # Solver execution with concurrent timeout
                status = 'Error'
                is_hard_timeout = False
                try:
                    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
                        future = executor.submit(opt.solve, warm_start=(base_sol is not None))
                        try: status = future.result(timeout=limit + 20)
                        except concurrent.futures.TimeoutError:
                            if sys.platform == "win32": os.system("taskkill /F /IM cbc.exe /T >nul 2>&1")
                            else: os.system("pkill -9 cbc > /dev/null 2>&1")
                            is_hard_timeout, status = True, 'TimedOut'
                except Exception as e: print(f"[red]Solver Error: {e}[/red]")

                if status in ('Optimal', 'Feasible'):
                    kpis = _extract_kpis(opt)
                    if mult == 1.0: base_sol = _get_model_solution(opt)
                    results.append({"target": "EUA", "scenario": scenario_id, "multiplier": mult, 
                                    "variation_pct": var_pct, "status": status, "timed_out": (status=='Feasible'), **kpis})
                else:
                    results.append({"target": "EUA", "scenario": scenario_id, "multiplier": mult, 
                                    "variation_pct": var_pct, "status": "Skipped", "timed_out": is_hard_timeout})
                progress.advance(task)

    if output_path:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f: json.dump(results, f, ensure_ascii=False, indent=2)
    return results
