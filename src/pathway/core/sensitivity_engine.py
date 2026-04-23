"""
Core Sensitivity Analysis Engine
================================
Provides OAT (One-At-a-Time) sensitivity analysis for PathFinder simulations.

Each sensitivity target is varied independently while all other parameters
remain at their baseline values — this is the true OAT (One-At-a-Time) design.
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

def extract_kpis(optimizer: PathFinderOptimizer) -> Dict[str, Any]:
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

def get_model_solution(optimizer: PathFinderOptimizer) -> Dict[str, float]:
    return {v.name: v.varValue for v in optimizer.model.variables() if v.varValue is not None}


def _apply_target_mutation(data_copy, target: str, mult: float, start_yr: int, years_list: list, parser: PathFinderParser) -> None:
    """
    Applies the OAT mutation for a single target to the data_copy object.
    Only the specified target's data is perturbed; everything else stays at baseline.
    """
    if target == 'EUA':
        # Vary carbon prices via anchor re-interpolation
        new_anchors = {
            yr: (_safe_val(val, mult) if yr > start_yr else val)
            for yr, val in data_copy.time_series.carbon_prices_anchors.items()
        }
        data_copy.time_series.carbon_prices = parser._interpolate_dict(new_anchors, years_list)

    elif target.startswith('PRICE:'):
        # Specific resource price variation (e.g. "PRICE: EN_ELEC | Electricity")
        try:
            r_id = target.split('|')[0].replace('PRICE:', '').strip()
            if r_id in data_copy.time_series.resource_prices_anchors:
                anchors = data_copy.time_series.resource_prices_anchors[r_id]
                new_anchors = {
                    yr: (_safe_val(v, mult) if yr > start_yr else v)
                    for yr, v in anchors.items()
                }
                data_copy.time_series.resource_prices[r_id] = parser._interpolate_dict(new_anchors, years_list)
        except Exception as e:
            pass

    elif target == 'RESSOURCES PRICE':
        # Default fallback (legacy) - vary all
        for r_id, anchors in data_copy.time_series.resource_prices_anchors.items():
            new_anchors = {
                yr: (_safe_val(v, mult) if yr > start_yr else v)
                for yr, v in anchors.items()
            }
            data_copy.time_series.resource_prices[r_id] = parser._interpolate_dict(new_anchors, years_list)

    elif target == 'RESSOURCES EMISSIONS':
        # Vary indirect emission factors directly (already fully interpolated dicts)
        for r_id, factor_dict in data_copy.time_series.other_emissions_factors.items():
            data_copy.time_series.other_emissions_factors[r_id] = {
                yr: _safe_val(factor, mult)
                for yr, factor in factor_dict.items()
            }

    elif target == 'CAPEX/OPEX':
        # Vary CAPEX and OPEX for every technology that has anchors
        for tech in data_copy.technologies.values():
            if tech.capex_anchors:
                new_capex = {
                    yr: (_safe_val(v, mult) if yr > start_yr else v)
                    for yr, v in tech.capex_anchors.items()
                }
                tech.capex_by_year = parser._interpolate_dict(new_capex, years_list)
            if tech.opex_anchors:
                new_opex = {
                    yr: (_safe_val(v, mult) if yr > start_yr else v)
                    for yr, v in tech.opex_anchors.items()
                }
                tech.opex_by_year = parser._interpolate_dict(new_opex, years_list)

def _apply_structural_mutation(data_copy, target: str, state: str, parser: PathFinderParser) -> None:
    """
    Applies a categorical/structural state change to the data object.
    """
    target = target.upper().strip()
    state = state.upper().strip()
    
    if target == 'TAX_INDIRECT_EMISSIONS':
        is_taxed = (state == 'YES')
        for res in data_copy.resources.values():
            res.tax_indirect_emissions = is_taxed
            
    elif target == 'INTERPOLATION_CONSTRAINT':
        # Affects how TimeSeriesData handles gaps
        parser.interpolation_mode = state
        years = list(data_copy.time_series.carbon_prices.keys())
        # Force re-interpolation of carbon prices
        data_copy.time_series.carbon_prices = parser._interpolate_dict(data_copy.time_series.carbon_prices_anchors, years)
        # Force re-interpolation of all resource prices
        for r_id in data_copy.time_series.resource_prices:
            anchors = data_copy.time_series.resource_prices_anchors.get(r_id, {})
            data_copy.time_series.resource_prices[r_id] = parser._interpolate_dict(anchors, years)
        # Restore default for next OAT case
        parser.interpolation_mode = "LINEAR"
        
    elif target == 'OBJECTIVE_PENALTY':
        # Set penalty_type for all objectives
        for obj in data_copy.objectives:
            obj.penalty_type = state # NONE, PENALTIES, AT ALL COST

    elif target in ('GRANTS', 'GRANTS_ACTIVE'):
        data_copy.grant_params.active = (state == 'YES')

    elif target in ('CCFD', 'CCFD_ACTIVE'):
        data_copy.ccfd_params.active = (state == 'YES')
        
    elif target == 'SV_ACT_MODE':
        # Affects free quota calculation logic (PI vs NORM)
        for ent in data_copy.entities.values():
            ent.sv_act_mode = state

    elif target in ('CARBON CREDITS', 'CREDITS_ACTIVE'):
        data_copy.credit_params.active = (state == 'YES')

    elif target in ('DAC', 'DAC_ACTIVE'):
        data_copy.dac_params.active = (state == 'YES')

    elif target == 'RELAX_INTEGRALITY':
        data_copy.parameters.relax_integrality = (state == 'YES')


def run_sensitivity(excel_path: str, output_path: Optional[str] = None, verbose: bool = False, precomputed_base_sols: Optional[Dict[str, Dict[str, Any]]] = None) -> List[Dict[str, Any]]:
    """
    Runs a true OAT sensitivity analysis and exports results to JSON.
    """
    parser = PathFinderParser(excel_path, verbose=verbose)
    sens_params = parser.parse_sensitivity()

    if not sens_params.scenarios:
        if verbose: print("[yellow]No sensitivity scenarios found. Skipping.[/yellow]")
        return []

    active_targets = [k for k, v in sens_params.targets.items() if v]
    
    # --- Robust Target Expansion ---
    if 'RESSOURCES PRICE' in active_targets:
        try:
            sample_sc = sens_params.scenarios[0]
            sample_data = parser.parse(scenario_id=sample_sc)
            expanded = []
            for r_id, r_obj in sample_data.resources.items():
                if r_id in sample_data.time_series.resource_prices_anchors:
                    r_name = r_obj.name or r_id
                    expanded.append(f"PRICE: {r_id} | {r_name}")
            idx = active_targets.index('RESSOURCES PRICE')
            active_targets[idx:idx+1] = expanded
        except Exception: pass

    # Total simulations calculation
    multipliers = _build_multipliers(sens_params)
    structural_count = sum(len(states) for states in sens_params.structural_targets.values())
    total_sims = len(sens_params.scenarios) * (len(active_targets) * len(multipliers) + structural_count)
    if total_sims == 0:
        if verbose: print("[yellow]No sensitivity targets or multipliers selected. Skipping.[/yellow]")
        return []

    results: List[Dict[str, Any]] = []

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
        MofNCompleteColumn(),
        TimeElapsedColumn(),
    ) as progress:
        task = progress.add_task("OAT Sensitivity...", total=total_sims)
        sub_task = progress.add_task("Current Solve", total=100, visible=False)

        for scenario_id in sens_params.scenarios:
            base_data = parser.parse(scenario_id=scenario_id)

            # --- PART 1: Numeric Targets ---
            for target in active_targets:
                display_target = target.split('|')[-1].strip() if '|' in target else target

                for mult in multipliers:
                    var_pct = _pct_from_multiplier(mult)
                    label = f"S:{scenario_id} | {display_target} | x{mult:.3f} ({var_pct:+.1f}%)"
                    progress.update(task, description=label)

                    if mult == 1.0 and precomputed_base_sols and scenario_id in precomputed_base_sols:
                        kpis = precomputed_base_sols[scenario_id]["kpis"]
                        results.append({
                            "target": display_target, "scenario": scenario_id,
                            "multiplier": 1.0, "variation_pct": 0.0,
                            "status": "Optimal", "is_structural": False, **kpis
                        })
                        progress.advance(task)
                        continue

                    data_copy = copy.deepcopy(base_data)
                    _apply_target_mutation(data_copy, target, mult, data_copy.parameters.start_year, list(data_copy.time_series.carbon_prices.keys()), parser)
                    
                    status, kpis = _solve_sensitivity_case(data_copy, scenario_id, precomputed_base_sols, progress, sub_task, display_target)
                    
                    results.append({
                        "target": display_target, "scenario": scenario_id,
                        "multiplier": mult, "variation_pct": var_pct,
                        "status": status, "is_structural": False, **(kpis or {})
                    })
                    progress.advance(task)

            # --- PART 2: Structural Targets ---
            for target, states in sens_params.structural_targets.items():
                for state in states:
                    label = f"S:{scenario_id} | {target} | {state}"
                    progress.update(task, description=label)

                    data_copy = copy.deepcopy(base_data)
                    _apply_structural_mutation(data_copy, target, state, parser)
                    
                    status, kpis = _solve_sensitivity_case(data_copy, scenario_id, precomputed_base_sols, progress, sub_task, target)
                    
                    results.append({
                        "target": target, "scenario": scenario_id,
                        "state": state, "variation_pct": 0.0, # Categorical has no pct
                        "status": status, "is_structural": True, **(kpis or {})
                    })
                    progress.advance(task)

    if output_path:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)

    return results

def _solve_sensitivity_case(data, scenario_id, precomputed_base_sols, progress, sub_task, display_target):
    """Helper to handle the solver loop for a single sensitivity case."""
    limit = data.parameters.time_limit * 2
    opt = PathFinderOptimizer(data, verbose=False)
    opt.build_model()

    base_sol = precomputed_base_sols.get(scenario_id, {}).get("solution") if precomputed_base_sols else None
    if base_sol: opt.apply_warm_start(base_sol)

    status = 'Error'
    solve_done = threading.Event()
    solve_start = time.time()
    progress.update(sub_task, description=f"  [cyan]Solving {display_target}...[/cyan]", visible=True, completed=20)

    def _spinner():
        while not solve_done.is_set():
            elapsed = time.time() - solve_start
            pct = min(85.0, 20.0 + (elapsed / limit) * 65.0)
            progress.update(sub_task, completed=pct)
            time.sleep(0.5)

    spinner_thread = threading.Thread(target=_spinner, daemon=True)
    spinner_thread.start()

    try:
        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
            future = executor.submit(opt.solve, warm_start=(base_sol is not None))
            try:
                status = future.result(timeout=limit + 20)
            except concurrent.futures.TimeoutError:
                if sys.platform == "win32": os.system("taskkill /F /IM cbc.exe /T >nul 2>&1")
                else: os.system("pkill -9 cbc > /dev/null 2>&1")
                status = 'TimedOut'
    except Exception: status = 'Error'

    # Recovery pass
    if status not in ('Optimal', 'Feasible'):
        data.parameters.relax_integrality = True
        opt = PathFinderOptimizer(data, verbose=False)
        opt.build_model()
        if base_sol: opt.apply_warm_start(base_sol)
        try:
            status = opt.solve(warm_start=(base_sol is not None))
            if status in ('Optimal', 'Feasible'): status = 'Feasible (Relaxed)'
        except Exception: pass

    solve_done.set()
    spinner_thread.join(timeout=1.0)
    progress.update(sub_task, visible=False)
    
    kpis = extract_kpis(opt) if status in ('Optimal', 'Feasible', 'Feasible (Relaxed)') else None
    return status, kpis
