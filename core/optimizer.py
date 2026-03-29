import pulp
import pandas as pd
import time
from typing import Dict, Any, List
from rich import print
from tqdm import tqdm
from core.model import PathFinderData

MASSIVE_PENALTY_COST = 1e7

# Conversion factors between resource units
# Key: (source_unit, target_unit) -> multiplication factor
_UNIT_CONVERSIONS = {
    ('MWH', 'GJ'): 3.6,
    ('GJ', 'MWH'): 1.0 / 3.6,
    ('KWH', 'GJ'): 0.0036,
    ('GJ', 'KWH'): 1000.0 / 3.6,
    ('KWH', 'MWH'): 0.001,
    ('MWH', 'KWH'): 1000.0,
}

def _get_unit_conversion(ref_resource, target_resource) -> float:
    """Return the conversion factor to apply when a 'new' impact references a resource
    with a different unit than the produced/consumed resource.
    E.g. EN_FUEL (MWH) -> EN_H2_P (GJ): factor = 3.6"""
    if ref_resource is None or target_resource is None:
        return 1.0
    ref_unit = ref_resource.unit.strip().upper()
    tgt_unit = target_resource.unit.strip().upper()
    if ref_unit == tgt_unit:
        return 1.0
    return _UNIT_CONVERSIONS.get((ref_unit, tgt_unit), 1.0)

class PathFinderOptimizer:
    def __init__(self, data: PathFinderData, verbose: bool = False):
        self.data = data
        self.verbose = verbose
        self.model = pulp.LpProblem("PathFinder_Decarbonization", pulp.LpMinimize)
        
        # We will assume a single entity for now or aggregate
        self.entity = list(self.data.entities.values())[0] if self.data.entities else None

        self.years = list(range(self.data.parameters.start_year, self.data.parameters.start_year + self.data.parameters.duration + 1))
        
        # Variables
        self.invest_vars = {}
        self.active_vars = {}
        self.cons_vars = {}
        self.emis_vars = {}
        self.taxed_emis_vars = {}
        self.paid_quota_vars = {}
        self.penalty_quota_vars = {}
        self.indirect_emis_vars = {}
        self.total_emis_vars = {}
        self.project_vars = {}
        
        self.grant_used_vars = {}
        self.ccfd_used_vars = {}
        
        self.loan_vars = {} # (t, p_id, t_id, loan_id) -> amount borrowed locally for project
        
        self.dac_added_capacity_vars = {}
        self.dac_total_capacity_vars = {}
        self.dac_captured_vars = {}
        self.credit_purchased_vars = {}
    
    def build_model(self):
        if not self.entity:
            raise ValueError("No entity to optimize")
            
        total_cost = []
        if self.verbose:
            print("  [magenta][Optimizer][/magenta] [BUILD] Building MILP variables...")
        # 1. Decision Variables
        for t in tqdm(self.years, desc="Building Variables", disable=not self.verbose):
            self.emis_vars[t] = pulp.LpVariable(f"Emis_{t}", lowBound=0)
            self.indirect_emis_vars[t] = pulp.LpVariable(f"IndirectEmis_{t}", lowBound=None)
            self.total_emis_vars[t] = pulp.LpVariable(f"TotalEmis_{t}", lowBound=None)
            self.taxed_emis_vars[t] = pulp.LpVariable(f"TaxedEmis_{t}", lowBound=0)
            self.paid_quota_vars[t] = pulp.LpVariable(f"PaidQuota_{t}", lowBound=0)
            self.penalty_quota_vars[t] = pulp.LpVariable(f"PenaltyQuota_{t}", lowBound=0)
            
        self.h2_buy_resources = [r for r in self.data.resources if 'H2' in r.upper() and r in self.data.time_series.resource_prices and r not in ['EN_H2_P', 'EN_H2_C', 'EN_ELEC_FOR_H2']]
        self.h2_supply_vars = {t: {} for t in self.years}
        for t in self.years:
            self.h2_supply_vars[t]['PRODUCED_ON_SITE'] = pulp.LpVariable(f"H2_Produce_{t}", lowBound=0)
            for res in self.h2_buy_resources:
                self.h2_supply_vars[t][res] = pulp.LpVariable(f"H2_Buy_{res}_{t}", lowBound=0)
            for p_id, process in self.entity.processes.items():
                for t_id in process.valid_technologies:
                    if t_id == 'UP':  # UP is free continuous improvement — no decision variable needed
                        continue
                    # PERFORMANCE: Relax technologies if nb_units > 1 to speed up solver
                    # Integer decision to invest in a number of units of technology t_id at year t for process p_id
                    _H2_TECHS = {'FUEL_TO_H2', 'PEM_H2'}
                    # Relax if it's an H2 tech OR if nb_units > 1 (MIP complexity reduction)
                    should_relax = (t_id in _H2_TECHS) or (process.nb_units > 1)
                    
                    # Also check for global override from parameters if added
                    if hasattr(self.data.parameters, 'relax_integrality') and self.data.parameters.relax_integrality:
                        should_relax = True
                        
                    var_cat = pulp.LpContinuous if should_relax else pulp.LpInteger
                    self.invest_vars[(t, p_id, t_id)] = pulp.LpVariable(f"Invest_{t}_{p_id}_{t_id}", lowBound=0, upBound=process.nb_units, cat=var_cat)
                    # Same category for active units tracking
                    self.active_vars[(t, p_id, t_id)] = pulp.LpVariable(f"Active_{t}_{p_id}_{t_id}", lowBound=0, upBound=process.nb_units, cat=var_cat)
                    # Binary flag if an investment project (at least 1 unit) occurs this year
                    self.project_vars[(t, p_id, t_id)] = pulp.LpVariable(f"Project_{t}_{p_id}_{t_id}", cat=pulp.LpBinary)
                    
                    if self.data.grant_params.active:
                        self.grant_used_vars[(t, p_id, t_id)] = pulp.LpVariable(f"Grant_{t}_{p_id}_{t_id}", cat=pulp.LpBinary)
                    if self.data.ccfd_params.active:
                        self.ccfd_used_vars[(t, p_id, t_id)] = pulp.LpVariable(f"CCfD_{t}_{p_id}_{t_id}", cat=pulp.LpBinary)
                        
                    for l_id, loan in enumerate(self.data.bank_loans):
                        self.loan_vars[(t, p_id, t_id, l_id)] = pulp.LpVariable(f"Loan_{t}_{p_id}_{t_id}_{l_id}", lowBound=0)
                
            for res_id in self.data.resources:
                if res_id not in ['CO2_EM']:
                    res_obj = self.data.resources.get(res_id)
                    lb = 0.0 if res_obj and res_obj.type.strip().upper() != 'PRODUCTION' else None
                    self.cons_vars[(t, res_id)] = pulp.LpVariable(f"Cons_{t}_{res_id}", lowBound=lb)
                
            if self.data.dac_params.active:
                self.dac_added_capacity_vars[t] = pulp.LpVariable(f"DAC_Added_Cap_{t}", lowBound=0.0)
                self.dac_total_capacity_vars[t] = pulp.LpVariable(f"DAC_Total_Cap_{t}", lowBound=0.0)
                self.dac_captured_vars[t] = pulp.LpVariable(f"DAC_Captured_{t}", lowBound=0.0)
            if self.data.credit_params.active:
                self.credit_purchased_vars[t] = pulp.LpVariable(f"Credit_Purchased_{t}", lowBound=0.0)
                    
        if self.verbose:
            print("  [magenta][Optimizer][/magenta] [CONSTR] Adding Constraints...")
        # 2. Constraints
        
        # DAC & Credit constraints
        if self.data.dac_params.active:
            for t in self.years:
                start = self.data.dac_params.start_year or self.years[0]
                end = self.data.dac_params.end_year or self.years[-1]
                
                # Cumulative capacity = sum of past added capacity
                valid_install_years = [tau for tau in self.years if tau <= t]
                if valid_install_years:
                    self.model += self.dac_total_capacity_vars[t] == pulp.lpSum(self.dac_added_capacity_vars[tau] for tau in valid_install_years), f"DAC_Cumul_Cap_{t}"
                
                if not (start <= t <= end):
                    self.model += self.dac_added_capacity_vars[t] == 0, f"DAC_Added_Inactive_{t}"
                    self.model += self.dac_captured_vars[t] == 0, f"DAC_Cap_Inactive_{t}"
                    
                # We can't capture more than installed capacity
                self.model += self.dac_captured_vars[t] <= self.dac_total_capacity_vars[t], f"DAC_Capture_Limit_{t}"
                
                # Capture volume limit (percentage of reference year emissions)
                if self.data.dac_params.max_volume_pct < 1.0:
                    ref_yr = self.data.dac_params.ref_year
                    if ref_yr in self.years:
                        ref_emis_expr = self.emis_vars[ref_yr]
                    else:
                        # Historical or external reference from Entities sheet
                        yr_map = self.entity.ref_baselines.get('CO2_EM', {})
                        if ref_yr in yr_map:
                            ref_emis_expr = yr_map[ref_yr]
                        elif yr_map:
                            # Fallback to any available reference value
                            ref_emis_expr = next(iter(yr_map.values()))
                        else:
                            ref_emis_expr = self.entity.base_emissions
                    self.model += self.dac_captured_vars[t] <= ref_emis_expr * self.data.dac_params.max_volume_pct, f"DAC_Volume_Limit_{t}"
                    
        if self.data.credit_params.active:
            ref_year = self.data.credit_params.ref_year
            if ref_year in self.years:
                ref_emis_expr = self.emis_vars[ref_year]
            else:
                # Use historical/external reference from Entities sheet
                yr_map = self.entity.ref_baselines.get('CO2_EM', {})
                if ref_year in yr_map:
                    ref_emis_expr = yr_map[ref_year]
                elif yr_map:
                    # Fallback to any available reference value
                    ref_emis_expr = next(iter(yr_map.values()))
                else:
                    ref_emis_expr = self.entity.base_emissions
            for t in self.years:
                start = self.data.credit_params.start_year or self.years[0]
                end = self.data.credit_params.end_year or self.years[-1]
                if not (start <= t <= end):
                    self.model += self.credit_purchased_vars[t] == 0, f"Credit_Inactive_{t}"
                    
                self.model += self.credit_purchased_vars[t] <= ref_emis_expr * self.data.credit_params.max_volume_pct, f"Credit_Limit_{t}"
        
        
        # Active logic: maximum capacity and cumulative investments
        for p_id, process in self.entity.processes.items():
            major_techs = [t_id for t_id in process.valid_technologies if t_id != 'UP']
            
            # --- Directed Technology Precedence Logic ---
            for t1 in major_techs:
                # Max units for a single technology over the horizon cannot exceed process units
                self.model += pulp.lpSum(self.invest_vars[(t, p_id, t1)] for t in self.years) <= process.nb_units, f"Max_{process.nb_units}_Invest_{p_id}_{t1}"
                
                comp_list = self.data.tech_compatibilities.get(t1, [])
                for t2 in major_techs:
                    if t1 == t2: continue
                    
                    if t2 not in comp_list:
                        # T2 cannot follow T1: The sum of active units of T1 and T2 cannot exceed total process units
                        for t in self.years:
                            self.model += self.active_vars[(t, p_id, t1)] + self.active_vars[(t, p_id, t2)] <= process.nb_units, f"Mutual_Capacity_Exclusion_{p_id}_{t1}_{t2}_at_{t}"
                    else:
                        # T2 follows T1: We can only have T2 units if T1 was already deployed there (T2 upgrades T1).
                        # Meaning, active T2 units at year t cannot exceed the number of previously invested T1 units.
                        for t in self.years:
                            prev_years_t2 = [tau for tau in self.years if tau <= t - self.data.technologies[t2].implementation_time]
                            
                            if prev_years_t2:
                                # For T2, it can only reach the level of T1 installed previously. 
                                # Assuming T1 also has lead time, active T1 at t must be >= active T2 at t.
                                self.model += self.active_vars[(t, p_id, t2)] <= self.active_vars[(t, p_id, t1)], f"Precedence_Capacity_{p_id}_{t1}_{t2}_at_{t}"
                            else:
                                self.model += self.active_vars[(t, p_id, t2)] == 0, f"No_T2_Precedence_{p_id}_{t1}_{t2}_at_{t}"

            for t_id in process.valid_technologies:
                if t_id == 'UP':  # UP is handled as automatic continuous improvement, not a decision
                    continue
                
                tech = self.data.technologies[t_id]
                delay = tech.implementation_time
                for t in self.years:
                    valid_invest_years = [tau for tau in self.years if tau <= t - delay]
                    if valid_invest_years:
                        self.model += self.active_vars[(t, p_id, t_id)] == pulp.lpSum(self.invest_vars[(tau, p_id, t_id)] for tau in valid_invest_years), f"Active_Logic_{t}_{p_id}_{t_id}"
                    else:
                        self.model += self.active_vars[(t, p_id, t_id)] == 0, f"Active_Logic_{t}_{p_id}_{t_id}"
                        
                    # Project variable link (1 if Invest >= 1, else 0)
                    self.model += self.invest_vars[(t, p_id, t_id)] <= process.nb_units * self.project_vars[(t, p_id, t_id)], f"Project_Link_UB_{t}_{p_id}_{t_id}"
                    self.model += self.project_vars[(t, p_id, t_id)] <= self.invest_vars[(t, p_id, t_id)], f"Project_Link_LB_{t}_{p_id}_{t_id}"

        # Public Aids Exclusivity and Limits
        for t in self.years:
            for p_id, process in self.entity.processes.items():
                for t_id in process.valid_technologies:
                    if t_id == 'UP': continue
                    expr = []
                    if self.data.grant_params.active:
                        expr.append(self.grant_used_vars[(t, p_id, t_id)])
                    if self.data.ccfd_params.active:
                        expr.append(self.ccfd_used_vars[(t, p_id, t_id)])
                    if expr:
                        self.model += pulp.lpSum(expr) <= self.project_vars[(t, p_id, t_id)], f"Max_1_Aid_{t}_{p_id}_{t_id}"

        # Grant Spacing Constraint
        if self.data.grant_params.active and self.data.grant_params.renew_time > 0:
            renew_int = int(self.data.grant_params.renew_time)
            for t in self.years:
                window_years = [tau for tau in self.years if t <= tau < t + renew_int]
                if window_years:
                    self.model += pulp.lpSum(self.grant_used_vars[(tau, p_id, t_id)] 
                                             for tau in window_years 
                                             for p_id, proc in self.entity.processes.items() 
                                             for t_id in proc.valid_technologies if t_id != 'UP') <= 1, f"Grant_Spacing_{t}"

        # CCfD Max Contracts
        if self.data.ccfd_params.active and self.data.ccfd_params.nb_contracts > 0:
            self.model += pulp.lpSum(self.ccfd_used_vars[(t, p_id, t_id)] 
                                     for t in self.years
                                     for p_id, proc in self.entity.processes.items()
                                     for t_id in proc.valid_technologies if t_id != 'UP') <= self.data.ccfd_params.nb_contracts, "CCfD_Max_Contracts"

        # --- Subsidies Values (Linearization of Cap) ---
        self.grant_amt_vars = {}
        for t in self.years:
            for p_id, process in self.entity.processes.items():
                for t_id in process.valid_technologies:
                    if t_id == 'UP': continue
                    if self.data.grant_params.active and self.data.grant_params.rate > 0:
                        self.grant_amt_vars[(t, p_id, t_id)] = pulp.LpVariable(f"GrantAmt_{t}_{p_id}_{t_id}", lowBound=0.0)
                        
                        tech = self.data.technologies[t_id]
                        cap_capex = 1.0
                        if tech.capex_per_unit:
                            if tech.capex_unit == 'tCO2': cap_capex = self.entity.base_emissions * process.emission_shares.get('CO2_EM', 0.0)
                            elif 'MW' in tech.capex_unit.upper(): 
                                base_mwh = self.entity.base_consumptions.get('EN_FUEL', 0.0) * process.consumption_shares.get('EN_FUEL', 0.0)
                                cap_capex = base_mwh / self.entity.annual_operating_hours
                        true_capex_per_unit = (tech.capex * cap_capex) / process.nb_units
                        
                        max_possible_grant = min(true_capex_per_unit * process.nb_units * self.data.grant_params.rate, self.data.grant_params.cap)
                        
                        # Grant amount is bonded by:
                        # 1. Total proportional grant (rate * capex * N_invest)
                        self.model += self.grant_amt_vars[(t, p_id, t_id)] <= (true_capex_per_unit * self.invest_vars[(t, p_id, t_id)]) * self.data.grant_params.rate, f"GrantAmt_Proportional_{t}_{p_id}_{t_id}"
                        # 2. Hard Cap if grant is active
                        self.model += self.grant_amt_vars[(t, p_id, t_id)] <= self.data.grant_params.cap * self.grant_used_vars[(t, p_id, t_id)], f"GrantAmt_Cap_{t}_{p_id}_{t_id}"
                        # 3. Only positive when grant used
                        M_grant = true_capex_per_unit * process.nb_units * self.data.grant_params.rate
                        self.model += self.grant_amt_vars[(t, p_id, t_id)] >= (true_capex_per_unit * self.invest_vars[(t, p_id, t_id)]) * self.data.grant_params.rate - M_grant * (1 - self.grant_used_vars[(t, p_id, t_id)]), f"GrantAmt_LB_{t}_{p_id}_{t_id}"
                        self.model += self.grant_amt_vars[(t, p_id, t_id)] >= 0, f"GrantAmt_Positive_{t}_{p_id}_{t_id}"

        # --- NEW: Process State Tracking Variables ---
        self.process_state_vars = pulp.LpVariable.dicts("State", 
            ((t, p_id, res_id) for t in self.years for p_id in self.entity.processes for res_id in self.data.resources),
            lowBound=None, cat='Continuous')
            
        self.tech_lock_vars = pulp.LpVariable.dicts("TechLock",
            ((t, p_id, t_id, res_id) for t in self.years for p_id, p in self.entity.processes.items() for t_id in p.valid_technologies for res_id in self.data.technologies[t_id].impacts),
            lowBound=None, cat='Continuous')
            
        # Penalty variables for soft constraints on objectives
        self.penalty_vars = {}
        for i, obj in enumerate(self.data.objectives):
            if obj.mode == 'LINEAR':
                t_target = min(obj.target_year, self.years[-1])
                # Find previous objective in the same group
                prev_year = self.years[0]
                if obj.group:
                    same_group_objs = [
                        o for o in self.data.objectives 
                        if o.group == obj.group and o.resource == obj.resource and o.entity == obj.entity
                        and o.target_year < obj.target_year
                    ]
                    if same_group_objs:
                        prev_year = min(max(same_group_objs, key=lambda o: o.target_year).target_year, self.years[-1])
                
                # We need vars for years > prev_year up to t_target
                for t in range(prev_year + 1, t_target + 1):
                    # Make sure the year is in the actual simulation duration
                    if t in self.years:
                        self.penalty_vars[(i, t)] = pulp.LpVariable(f"Penalty_Obj_{i}_{t}", lowBound=0.0)
            else:
                self.penalty_vars[i] = pulp.LpVariable(f"Penalty_Obj_{i}", lowBound=0.0)

        # --- Base State Initialization & Evolution ---
        for p_id, process in self.entity.processes.items():
            for res_id in self.data.resources:
                if res_id == 'CO2_EM':
                    initial_val = self.entity.base_emissions * process.emission_shares.get('CO2_EM', 0.0)
                else:
                    initial_val = self.entity.base_consumptions.get(res_id, 0.0) * process.consumption_shares.get(res_id, 0.0)
                
                # Big-M value for linearizations: must be large enough to encapsulate shifting cross-dimensional references
                # Now calculated dynamically per process/resource to improve numerical stability.
                # We use a safety factor of 1.5x the initial value or a sensible minimum.
                M = max(1000.0, abs(initial_val) * 2.0)
                
                # Evolution of the state year over year
                # Determine UP continuous improvement rate for this process
                up_rate = 0.0
                if 'UP' in process.valid_technologies:
                    up_tech = self.data.technologies['UP']
                    # UP impacts ALL resources with a variation type
                    up_imp = up_tech.impacts.get('ALL') or up_tech.impacts.get(res_id)
                    if up_imp and (up_imp['type'] == 'variation' or up_imp['type'] == 'up'):
                        up_rate = abs(up_imp['value'])  # e.g. 0.01 for 1% per year
                
                for i, t in enumerate(self.years):
                    # Accumulate all technologies active at year t applied from the UP-adjusted baseline
                    # UP is excluded from this loop (handled above via up_adjusted_val)
                    impacts_t = []
                    
                    for t_id in process.valid_technologies:
                        if t_id == 'UP':  # UP already applied via up_adjusted_val — skip to avoid double-counting
                            continue
                        tech = self.data.technologies[t_id]
                        
                        # Find exactly matching impacts for res_id
                        imp = tech.impacts.get(res_id)
                        # Special handling for indirect 'CO2' keys mapping to CO2_EM
                        if not imp and res_id == 'CO2_EM':
                            co2_keys = [k for k in tech.impacts.keys() if 'CO2' in k]
                            if co2_keys: imp = tech.impacts[co2_keys[0]]
                        
                        if imp:
                            act_var_current = self.active_vars[(t, p_id, t_id)]
                            
                            reference = imp.get('reference', 'INITIAL')
                            ref_res = imp.get('ref_resource')
                            if ref_res == 'CO2_EM':
                                base_ref_amount = self.entity.base_emissions * process.emission_shares.get('CO2_EM', 0.0)
                            elif ref_res and ref_res in self.entity.base_consumptions:
                                base_ref_amount = self.entity.base_consumptions.get(ref_res, 0.0) * process.consumption_shares.get(ref_res, 0.0)
                            else:
                                base_ref_amount = 1.0
                                
                            if imp.get('reference', '') == 'AVOIDED' and ref_res:
                                ref_imp = tech.impacts.get(ref_res)
                                if not ref_imp and ref_res == 'CO2_EM':
                                    co2_keys = [k for k in tech.impacts.keys() if 'CO2' in k]
                                    if co2_keys: ref_imp = tech.impacts[co2_keys[0]]
                                if ref_imp and ref_imp['type'] in ['variation', 'up'] and ref_imp['value'] < 0:
                                    base_ref = abs(ref_imp['value']) * base_ref_amount
                                else:
                                    base_ref = base_ref_amount
                            else:
                                base_ref = base_ref_amount

                            reference = imp.get('reference', 'INITIAL')
                            if reference == 'ACTUAL':
                                # Without recurvise tracking, 'ACTUAL' must approximate an order of operations
                                # Let's assume targeted reductions refer to the dynamic sequence. 
                                # Since exact continuous state tracking collapses linearly, 
                                # we map ACTUAL to its dynamic sequence via sequenced continuous bounds.
                                
                                # To avoid circular paradox, we calculate the continuous locked impact
                                impact_var = pulp.LpVariable(f"ActImpact_{t}_{p_id}_{t_id}_{res_id}", lowBound=None, cat='Continuous')
                                
                                # We evaluate the 'previous state' abstractly as what the state would be without this technology
                                # Mathematically, this is roughly just bounded by M
                                
                                if imp['type'] == 'variation' or imp['type'] == 'up':
                                    # Target reduction scaled proportionally to the number of active units out of nb_units
                                    scaling_factor = 1.0 / process.nb_units
                                    target_reduction = imp['value'] * initial_val * scaling_factor
                                elif imp['type'] == 'new':
                                    scaling_factor = 1.0 / process.nb_units
                                    # Apply unit conversion if ref_resource and produced resource have different units
                                    # e.g. EN_FUEL (MWH) -> EN_H2_P (GJ): multiply by 3.6
                                    conv = _get_unit_conversion(
                                        self.data.resources.get(ref_res),
                                        self.data.resources.get(res_id)
                                    )
                                    target_reduction = imp['value'] * base_ref * scaling_factor * conv
                                else:
                                    target_reduction = 0
                                    
                                # Lock bounds directly to the active modifier
                                # If act_var_current is 0 -> impact_var = 0 (bounded tightly between -0 and 0)
                                # If act_var_current is > 0 -> target_reduction is enforced (scaled by N units)
                                self.model += impact_var <= M * act_var_current, f"ImpUB1_{t}_{p_id}_{t_id}_{res_id}"
                                self.model += impact_var >= -M * act_var_current, f"ImpLB1_{t}_{p_id}_{t_id}_{res_id}"
                                
                                # We scale the target explicitly per unit active
                                exact_reduction_expr = target_reduction * act_var_current
                                
                                # Strict assignment since it's an additive continuous effect, bounded correctly.
                                self.model += impact_var == exact_reduction_expr, f"Impact_Exact_{t}_{p_id}_{t_id}_{res_id}"
                                
                                impacts_t.append(impact_var)

                            else:
                                # INITIAL formulation binds to constant
                                if imp['type'] == 'variation' or imp['type'] == 'up':
                                    scaling_factor = 1.0 / process.nb_units
                                    target_reduction = imp['value'] * initial_val * scaling_factor
                                elif imp['type'] == 'new':
                                    scaling_factor = 1.0 / process.nb_units
                                    # Apply unit conversion if ref_resource and produced resource have different units
                                    # e.g. EN_FUEL (MWH) -> EN_H2_P (GJ): multiply by 3.6
                                    conv = _get_unit_conversion(
                                        self.data.resources.get(ref_res),
                                        self.data.resources.get(res_id)
                                    )
                                    target_reduction = imp['value'] * base_ref * scaling_factor * conv
                                else:
                                    target_reduction = 0
                                    
                                impacts_t.append(target_reduction * act_var_current)
                                    
                    # Apply compounding continuous improvement (UP) to the entire state
                    # Formula: State(t) = (Initial + Sum of Tech Impacts) * (1 - up_rate)^t
                    up_factor = (1.0 - up_rate) ** i
                    
                    net_base_val = initial_val + pulp.lpSum(impacts_t)
                    res_obj = self.data.resources.get(res_id)
                    is_prod = res_obj.type.strip().upper() == 'PRODUCTION' if res_obj else False
                    
                    # We NO LONGER need the State_NonNeg_Slack here as we added lowBound=0 to cons_vars
                    # which propagates back through the mass balance.
                    # However, for CO2_EM (which is not in cons_vars), we use the emis_vars bounds.
                    self.model += self.process_state_vars[(t, p_id, res_id)] == net_base_val * up_factor, f"State_Evol_{t}_{p_id}_{res_id}"

        # Resource Mass Balance via mapped dynamic states
        for t in self.years:
            for res_id in self.data.resources:
                if res_id == 'CO2_EM':
                    continue
                    
                # The total consumption is exactly the sum of the dynamic process states
                total_process_cons = pulp.lpSum([self.process_state_vars[(t, p_id, res_id)] for p_id in self.entity.processes])

                allocated_share = sum(p.consumption_shares.get(res_id, 0.0) for p in self.entity.processes.values())
                frac_unallocated = 1.0 - allocated_share
                if abs(frac_unallocated) < 1e-4: frac_unallocated = 0.0
                unallocated = self.entity.base_consumptions.get(res_id, 0.0) * frac_unallocated
                
                dac_cons = 0.0
                if self.data.dac_params.active and res_id == 'EN_ELEC':
                    dac_cons = self.dac_captured_vars[t] * self.data.dac_params.elec_by_year.get(t, 0.0)
                
                h2_additional = 0.0
                # NOTE: We NO LONGER add h2_additional to self.cons_vars for H2 resources
                # because it causes an identity cancellation in the Balancer equation.
                # Cons_vars should represent the PHYSICAL consumption of the refinery processes.
                if res_id == 'EN_ELEC_FOR_H2':
                    h2_additional = self.h2_supply_vars[t]['PRODUCED_ON_SITE'] * 2.0
                    
                self.model += self.cons_vars[(t, res_id)] == total_process_cons + unallocated + dac_cons + h2_additional, f"Cons_Balance_{t}_{res_id}"
                
                # Granular Reporting: Link specific H2 buy variables to their cons_vars
                if res_id in self.h2_buy_resources:
                    self.model += self.cons_vars[(t, res_id)] == self.h2_supply_vars[t][res_id], f"H2_Market_Link_{t}_{res_id}"
            # H2 Balancer equation
            # Physical demand (H2_C) vs physical supply from techs (H2_P)
            h2_cons = self.cons_vars.get((t, 'EN_H2_C'), 0.0)
            h2_prod_tech = self.cons_vars.get((t, 'EN_H2_P'), 0.0)
            
            # Additional supply from market or proxy electrolyzer
            h2_market_and_proxy = pulp.lpSum([v for v in self.h2_supply_vars[t].values()])
            
            # Eq: ProcessDemand + Unallocated == MarketPurchase + ProxyElectrolysis + TechProduction
            # Note: total_process_cons for H2_C is what we want here.
            # However, cons_vars[EN_H2_C] already includes unallocated.
            self.model += h2_cons == h2_market_and_proxy + h2_prod_tech, f"H2_Demand_Fulfillment_{t}"
            # Emissions computation
            direct_process_emis = pulp.lpSum([self.process_state_vars[(t, p_id, 'CO2_EM')] for p_id in self.entity.processes])
            
            # Account for unallocated emissions
            allocated_emis_share = sum(p.emission_shares.get('CO2_EM', 0.0) for p in self.entity.processes.values())
            frac_unallocated_emis = 1.0 - allocated_emis_share
            if abs(frac_unallocated_emis) < 1e-4: frac_unallocated_emis = 0.0
            unallocated_co2 = self.entity.base_emissions * frac_unallocated_emis
            
            self.model += self.emis_vars[t] == direct_process_emis + unallocated_co2, f"Emis_Balance_{t}"
            
            # Indirect and Total emissions
            indirect_expr = []
            for r_id in self.data.resources:
                if r_id in self.data.time_series.other_emissions_factors:
                    factor = self.data.time_series.other_emissions_factors[r_id].get(t, 0.0)
                    if factor > 0:
                        indirect_expr.append(self.cons_vars[(t, r_id)] * factor)
                        
            if indirect_expr:
                self.model += self.indirect_emis_vars[t] == pulp.lpSum(indirect_expr), f"Indirect_Emis_Balance_{t}"
            else:
                self.model += self.indirect_emis_vars[t] == 0, f"Indirect_Emis_Dummy_{t}"
                
            self.model += self.total_emis_vars[t] == self.emis_vars[t] + self.indirect_emis_vars[t], f"Total_Emis_Balance_{t}"
            
            # Net emissions after DAC physically capturing CO2 (Voluntary Credits do not lower taxable emissions)
            net_emis_for_tax = self.emis_vars[t]
            if self.data.dac_params.active:
                net_emis_for_tax -= self.dac_captured_vars[t]
            
            # Quotas (Applies strictly to Direct Emissions)
            if self.entity.sv_act_mode == "PI":
                fq_pct = self.data.time_series.carbon_quotas_pi.get(t, 0.0)
            else:
                fq_pct = self.data.time_series.carbon_quotas_norm.get(t, 0.0)

            if fq_pct <= 1.0:
                self.model += self.taxed_emis_vars[t] >= net_emis_for_tax * (1.0 - fq_pct), f"Taxed_Emis_Limit_{t}"
            else:
                self.model += self.taxed_emis_vars[t] >= net_emis_for_tax - fq_pct, f"Taxed_Emis_Limit_{t}"

            # Split taxed emissions into paid vs penalty portions
            self.model += self.taxed_emis_vars[t] == self.paid_quota_vars[t] + self.penalty_quota_vars[t], f"Split_Taxed_Emis_{t}"

        # Company Objectives Constraints
        for i, obj in enumerate(self.data.objectives):
            t_target = obj.target_year
            
            # Use specific target year or the end year
            if t_target not in self.years:
                if t_target < self.years[0]: continue
                t_target = min(t_target, self.years[-1])

            # compute base comparison value
            if obj.comparison_year and obj.comparison_year < self.years[0]:
                # assume baseline from the first year
                pass
                
            base_val = 0.0
            if obj.resource in self.entity.ref_baselines:
                yr_map = self.entity.ref_baselines[obj.resource]
                ref_yr = obj.comparison_year or 2025
                if ref_yr in yr_map:
                    base_val = yr_map[ref_yr]
                elif yr_map:
                    base_val = next(iter(yr_map.values()))
            elif obj.resource == 'CO2_EM':
                base_val = self.entity.base_emissions
            elif obj.resource == 'INDIRECT_CO2_EM':
                base_indir = 0.0
                for r_id in self.data.resources:
                    if r_id in self.data.time_series.other_emissions_factors:
                        factor = self.data.time_series.other_emissions_factors[r_id].get(self.years[0], 0.0)
                        if factor > 0:
                            base_indir += self.entity.base_consumptions.get(r_id, 0.0) * factor
                base_val = base_indir
            elif obj.resource == 'TOTAL_CO2_EM':
                base_indir = 0.0
                for r_id in self.data.resources:
                    if r_id in self.data.time_series.other_emissions_factors:
                        factor = self.data.time_series.other_emissions_factors[r_id].get(self.years[0], 0.0)
                        if factor > 0:
                            base_indir += self.entity.base_consumptions.get(r_id, 0.0) * factor
                base_val = self.entity.base_emissions + base_indir
            else:
                base_val = self.entity.base_consumptions.get(obj.resource, 0.0)
                
            # If comparison year exists, logic might be a relative reduction (e.g. -0.5 for 50% reduction)
            # If cap_value is very small (between -1 and 1), it's probably percentage. Use base_val * (1 + cap_value)
            # Else it's absolute limit
            if obj.comparison_year and -1.0 <= obj.cap_value <= 1.0:
                limit_val = base_val * (1 + obj.cap_value)
            else:
                limit_val = obj.cap_value

            if obj.mode == 'LINEAR':
                t_target = min(obj.target_year, self.years[-1])
                start_year = self.years[0]
                prev_year = start_year
                prev_val = base_val
                
                if obj.group:
                    same_group_objs = [
                        o for o in self.data.objectives 
                        if o.group == obj.group and o.resource == obj.resource and o.entity == obj.entity
                        and o.target_year < obj.target_year
                    ]
                    if same_group_objs:
                        prev_obj = max(same_group_objs, key=lambda o: o.target_year)
                        prev_year = min(prev_obj.target_year, self.years[-1])
                        # compute prev_val
                        if prev_obj.comparison_year and -1.0 <= prev_obj.cap_value <= 1.0:
                            prev_val = base_val * (1 + prev_obj.cap_value)
                        else:
                            prev_val = prev_obj.cap_value
                
                # Enforce limit for all valid years from prev_year + 1 up to t_target
                for t in range(prev_year + 1, t_target + 1):
                    if t not in self.years:
                        continue
                    
                    if t_target > prev_year:
                        limit_t = prev_val + (limit_val - prev_val) * (t - prev_year) / (t_target - prev_year)
                    else:
                        limit_t = limit_val
                    
                    if obj.resource == 'CO2_EM':
                        net_obj_emis = self.emis_vars[t]
                        if self.data.dac_params.active:
                            net_obj_emis -= self.dac_captured_vars[t]
                        if self.data.credit_params.active:
                            net_obj_emis -= self.credit_purchased_vars[t]
                        var_expr = net_obj_emis
                    elif obj.resource == 'INDIRECT_CO2_EM':
                        var_expr = self.indirect_emis_vars[t]
                    elif obj.resource == 'TOTAL_CO2_EM':
                        net_obj_emis = self.total_emis_vars[t]
                        if self.data.dac_params.active:
                            net_obj_emis -= self.dac_captured_vars[t]
                        if self.data.credit_params.active:
                            net_obj_emis -= self.credit_purchased_vars[t]
                        var_expr = net_obj_emis
                    else:
                        var_expr = self.cons_vars[(t, obj.resource)]
                        
                    penalty_t = self.penalty_vars[(i, t)]
                    
                    if obj.limit_type == 'CAP' or obj.limit_type == 'MAX':
                        self.model += var_expr - penalty_t <= limit_t, f"Objective_{i}_CAP_{obj.resource}_{t}"
                    elif obj.limit_type == 'MIN':
                        self.model += var_expr + penalty_t >= limit_t, f"Objective_{i}_MIN_{obj.resource}_{t}"
            else:
                if obj.resource == 'CO2_EM':
                    net_obj_emis = self.emis_vars[t_target]
                    if self.data.dac_params.active:
                        net_obj_emis -= self.dac_captured_vars[t_target]
                    if self.data.credit_params.active:
                        net_obj_emis -= self.credit_purchased_vars[t_target]
                    var_expr = net_obj_emis
                elif obj.resource == 'INDIRECT_CO2_EM':
                    var_expr = self.indirect_emis_vars[t_target]
                elif obj.resource == 'TOTAL_CO2_EM':
                    net_obj_emis = self.total_emis_vars[t_target]
                    if self.data.dac_params.active:
                        net_obj_emis -= self.dac_captured_vars[t_target]
                    if self.data.credit_params.active:
                        net_obj_emis -= self.credit_purchased_vars[t_target]
                    var_expr = net_obj_emis
                else:
                    var_expr = self.cons_vars[(t_target, obj.resource)]
                    
                penalty = self.penalty_vars[i]
                
                if obj.limit_type == 'CAP' or obj.limit_type == 'MAX':
                    self.model += var_expr - penalty <= limit_val, f"Objective_{i}_CAP_{obj.resource}_{t_target}"
                elif obj.limit_type == 'MIN':
                    self.model += var_expr + penalty >= limit_val, f"Objective_{i}_MIN_{obj.resource}_{t_target}"

        # CA-based Rolling Budget Constraints
        ca_vars = {}
        yearly_budget_vars = {}
        capex_spent_vars = {}
        out_of_pocket_capex_vars = {}
        
        # Setup variables for CA and budget
        for t in self.years:
            ca_vars[t] = pulp.LpVariable(f"CA_{t}")
            yearly_budget_vars[t] = pulp.LpVariable(f"YearlyBudget_{t}")
            capex_spent_vars[t] = pulp.LpVariable(f"CapexSpent_{t}", lowBound=0.0)
            out_of_pocket_capex_vars[t] = pulp.LpVariable(f"OutOfPocketCapex_{t}", lowBound=0.0)
            
            # Calculate total Capex spent this year
            capex_expr = []
            for p_id, process in self.entity.processes.items():
                for t_id in process.valid_technologies:
                    if t_id == 'UP': continue
                    tech = self.data.technologies[t_id]
                    if tech.capex > 0:
                        cap = 1.0
                        if tech.capex_per_unit:
                            if tech.capex_unit == 'tCO2': cap = self.entity.base_emissions * process.emission_shares.get('CO2_EM', 0.0)
                            elif 'MW' in tech.capex_unit.upper(): 
                                base_mwh = self.entity.base_consumptions.get('EN_FUEL', 0.0) * process.consumption_shares.get('EN_FUEL', 0.0)
                                cap = base_mwh / self.entity.annual_operating_hours
                        true_capex = tech.capex * cap
                        
                        if self.data.grant_params.active and self.data.grant_params.rate > 0:
                            project_net_capex = (true_capex / process.nb_units) * self.invest_vars[(t, p_id, t_id)] - self.grant_amt_vars[(t, p_id, t_id)]
                        else:
                            project_net_capex = (true_capex / process.nb_units) * self.invest_vars[(t, p_id, t_id)]
                            
                        capex_expr.append(project_net_capex)
                        
                        self.model += pulp.lpSum(self.loan_vars[(t, p_id, t_id, l_id)] for l_id in range(len(self.data.bank_loans))) <= project_net_capex, f"Loan_Limit_CAPEX_{t}_{p_id}_{t_id}"
                            
            if self.data.dac_params.active:
                capex_expr.append(self.dac_added_capacity_vars[t] * self.data.dac_params.capex_by_year.get(t, 0.0))
                            
            self.model += capex_spent_vars[t] == pulp.lpSum(capex_expr), f"Capex_Spent_Calc_{t}"
            
            # Out of pocket = Capex - Loans
            total_loans_t = pulp.lpSum(self.loan_vars[(t, p_id, t_id, l_id)] 
                                       for p_id, p in self.entity.processes.items() 
                                       for t_id in p.valid_technologies if t_id != 'UP'
                                       for l_id in range(len(self.data.bank_loans)))
            self.model += out_of_pocket_capex_vars[t] == capex_spent_vars[t] - total_loans_t, f"OutOfPocket_Calc_{t}"

            # Calculate CA for year t only if limit is active
            if self.entity.ca_percentage_limit > 0 and self.entity.sold_resources:
                ca_expr = []
                for res_id in self.entity.sold_resources:
                    price = self.data.time_series.resource_prices.get(res_id, {}).get(t, 0.0)
                    if price > 0:
                        res_obj = self.data.resources.get(res_id)
                        # We want REVENUE (Positive). 
                        # If cons_vars is negative (Refinery production), we multiply by -1.
                        # If cons_vars is positive (New tech production), we multiply by +1.
                        # The heuristic: if base_consumption is 0, it's a new tech (positive impact).
                        # If not, we follow the resource type.
                        r_base = self.entity.base_consumptions.get(res_id, 0.0)
                        if res_obj and res_obj.type.strip().upper() == 'PRODUCTION':
                            if r_base < 0: # Already negative baseline
                                ca_expr.append(-price * self.cons_vars[(t, res_id)])
                            else: # Likely positive impact production
                                ca_expr.append(price * self.cons_vars[(t, res_id)])
                        else:
                            # Consumption resources sold (should be negative in net, but rare)
                            ca_expr.append(-price * self.cons_vars[(t, res_id)])
                
                if ca_expr:
                    # Turnover (CA) is the sum of revenues.
                    self.model += ca_vars[t] == pulp.lpSum(ca_expr), f"CA_Calc_{t}"
                else:
                    self.model += ca_vars[t] == 0, f"CA_Calc_{t}"
                    
                self.model += yearly_budget_vars[t] == ca_vars[t] * self.entity.ca_percentage_limit, f"Budget_Calc_{t}"
            else:
                self.model += ca_vars[t] == 0, f"CA_Dummy_{t}"
                self.model += yearly_budget_vars[t] == 1e12, f"Budget_Dummy_{t}" # Infinite budget if not defined

        if self.entity.ca_percentage_limit > 0 and self.entity.sold_resources:
            for t in self.years:
                # STRICT ANNUAL BUDGET: Out-of-pocket CAPEX cannot exceed annual turnover-based limit
                # If total CAPEX > yearly_budget, the difference MUST be covered by a bank loan.
                self.model += out_of_pocket_capex_vars[t] <= yearly_budget_vars[t], f"Strict_Annual_Budget_{t}"
                
                # STRICT CAPEX LIMIT WITH LOANS: Total CAPEX including loans cannot exceed 1.5x the budget limit
                self.model += capex_spent_vars[t] <= 1.5 * yearly_budget_vars[t], f"Strict_Max_Capex_With_Loans_{t}"

        # 3. Objective Function (Economic TCO)
        if self.verbose:
            print("  [magenta][Optimizer][/magenta] [OBJ] Building Objective...")
        for t in self.years:
            # 1. Financial Flows (Capex Out-of-pocket + Loan Annuities)
            if self.entity.ca_percentage_limit > 0 and self.entity.sold_resources:
                total_cost.append(out_of_pocket_capex_vars[t])
            else:
                # If no budget constraint, we still need to account for CAPEX in objective
                # But wait, without budget limit, the user might still want to take loans?
                # Usually loans are taken BECAUSE of the budget limit.
                # If no budget limit, we use raw capex.
                # Actually, let's always use the financing logic if loans are available.
                # Let's define capex_spent_vars even if no budget limit.
                pass
            
            # Calculate Repayments (Annuities) for all loans taken up to year t
            for tau in self.years:
                if tau > t: continue
                for l_id, loan in enumerate(self.data.bank_loans):
                    # If loan taken at tau, it impacts years [tau, tau + duration - 1]
                    if tau <= t < tau + loan.duration:
                        # Annuity = P * (r / (1 - (1+r)^-d))
                        # If r=0, Annuity = P / d
                        r = loan.rate
                        d = loan.duration
                        if r > 0:
                            annuity_factor = r / (1 - (1 + r)**(-d))
                        else:
                            annuity_factor = 1.0 / d
                        
                        total_loan_amount = pulp.lpSum(self.loan_vars[(tau, p_id, t_id, l_id)]
                                                       for p_id, p in self.entity.processes.items()
                                                       for t_id in p.valid_technologies if t_id != 'UP')
                        total_cost.append(total_loan_amount * annuity_factor)
                        
                        # Add a tiny penalty to loan selection to strictly prioritize out-of-pocket usage
                        # Even if rate is 0%, loans shouldn't be taken if budget is available.
                        total_cost.append(total_loan_amount * 1e-4)

            # CAPEX and OPEX
            for p_id, process in self.entity.processes.items():
                for t_id in process.valid_technologies:
                    if t_id == 'UP':  # UP is free — no CAPEX/OPEX cost
                        continue
                    tech = self.data.technologies[t_id]
                    
                    # OPEX and CCfD (CAPEX handled separately if budget active)
                    cap_calc = 1.0
                    if tech.capex_per_unit or tech.opex_per_unit:
                        if tech.capex_unit == 'tCO2' or tech.opex_unit == 'tCO2': 
                            cap_calc = self.entity.base_emissions * process.emission_shares.get('CO2_EM', 0.0)
                        elif 'MW' in str(tech.capex_unit).upper() or 'MW' in str(tech.opex_unit).upper(): 
                            base_mwh = self.entity.base_consumptions.get('EN_FUEL', 0.0) * process.consumption_shares.get('EN_FUEL', 0.0)
                            cap_calc = base_mwh / self.entity.annual_operating_hours
                    
                    true_capex = tech.capex * cap_calc
                    true_opex = tech.opex * cap_calc
                    
                    # If NO budget constraint, we add CAPEX here directly.
                    # If budget constraint ACTIVE, CAPEX is already in out_of_pocket_capex_vars[t].
                    if not (self.entity.ca_percentage_limit > 0 and self.entity.sold_resources):
                        if self.data.grant_params.active and self.data.grant_params.rate > 0:
                            total_cost.append((true_capex / process.nb_units) * self.invest_vars[(t, p_id, t_id)] - self.grant_amt_vars[(t, p_id, t_id)])
                        else:
                            total_cost.append((true_capex / process.nb_units) * self.invest_vars[(t, p_id, t_id)])

                    total_cost.append((true_opex / process.nb_units) * self.active_vars[(t, p_id, t_id)])
                    
                    # CCfD Calculation
                    if self.data.ccfd_params.active and self.data.ccfd_params.duration > 0:
                        ccfd_p = self.data.ccfd_params
                        imp = tech.impacts.get('CO2_EM')
                        if not imp:
                            co2_keys = [k for k in tech.impacts.keys() if 'CO2' in k]
                            if co2_keys: imp = tech.impacts[co2_keys[0]]
                            
                        if imp and (imp['type'] == 'variation' or imp['type'] == 'up'):
                            # Avoided emissions: The absolute reduction per year for ONE unit out of nb_units
                            reduction_frac_per_unit = (-imp['value'] if imp['value'] < 0 else 0) / process.nb_units
                            if reduction_frac_per_unit > 0:
                                max_emis_for_proc = self.entity.base_emissions * process.emission_shares.get('CO2_EM', 0.0)
                                avoided_tons_per_year_per_unit = reduction_frac_per_unit * max_emis_for_proc
                                
                                start_yr = t + tech.implementation_time
                                end_yr = start_yr + ccfd_p.duration
                                
                                ccfd_revenue_expr = []
                                for tau in range(start_yr, end_yr):
                                    if tau in self.years:
                                        c_price = self.data.time_series.carbon_prices.get(tau, 0.0)
                                        strike_price = (1.0 + ccfd_p.eua_price_pct) * self.data.time_series.carbon_prices.get(t, 0.0)
                                        
                                        if ccfd_p.contract_type == 2:
                                            subsidy_per_ton = strike_price - c_price
                                        else: # type == 1
                                            subsidy_per_ton = max(0, strike_price - c_price)
                                            
                                        # Subsidy is applied to the number of units invested in this project
                                        ccfd_revenue_expr.append(subsidy_per_ton * avoided_tons_per_year_per_unit * self.invest_vars[(t, p_id, t_id)])
                                        
                                if ccfd_revenue_expr:
                                    total_subsidy = pulp.lpSum(ccfd_revenue_expr)
                                    # Dynamic Big-M for CCfD subsidy: 1.5x the max possible subsidy
                                    M_subsidy = max(1000.0, 1.5 * sum(
                                        abs( ( (1.0 + ccfd_p.eua_price_pct) * self.data.time_series.carbon_prices.get(t, 0.0) ) - self.data.time_series.carbon_prices.get(tau, 0.0) ) * avoided_tons_per_year_per_unit * process.nb_units
                                        for tau in range(start_yr, end_yr) if tau in self.years
                                    ))
                                    ccfd_amt_var = pulp.LpVariable(f"CCfDAmt_{t}_{p_id}_{t_id}", lowBound=0.0)
                                    self.model += ccfd_amt_var <= total_subsidy, f"CCfDAmt_Prop_{t}_{p_id}_{t_id}"
                                    self.model += ccfd_amt_var <= M_subsidy * self.ccfd_used_vars[(t, p_id, t_id)], f"CCfDAmt_Cap_{t}_{p_id}_{t_id}"
                                    self.model += ccfd_amt_var >= total_subsidy - M_subsidy * (1 - self.ccfd_used_vars[(t, p_id, t_id)]), f"CCfDAmt_LB_{t}_{p_id}_{t_id}"
                                    
                                    total_cost.append(-ccfd_amt_var)
                
            if self.data.dac_params.active:
                total_cost.append(self.dac_total_capacity_vars[t] * self.data.dac_params.opex_by_year.get(t, 0.0))
            if self.data.credit_params.active:
                total_cost.append(self.credit_purchased_vars[t] * self.data.credit_params.cost_by_year.get(t, 0.0))
                
            # Resource Costs / Revenues
            for r_id in self.data.resources:
                if r_id == 'CO2_EM': continue
                price = 0.0
                if r_id in self.data.time_series.resource_prices:
                    price = self.data.time_series.resource_prices[r_id].get(t, 0.0)
                
                if price > 0:
                    # EXCEPTION: We do not add direct revenue for EN_H2_P in the objective
                    # because its primary benefit is reducing market H2 purchases in the balancer.
                    if r_id == 'EN_H2_P':
                        continue
                        
                    res_obj = self.data.resources.get(r_id)
                    if res_obj and res_obj.type.strip().upper() == 'PRODUCTION':
                        # PRODUCTION resources: revenue depends on the sign of cons_var.
                        # We use simple price * Quantity. 
                        # Quantity is +cons_var if impacts are positive, -cons_var if negative.
                        r_base = self.entity.base_consumptions.get(r_id, 0.0)
                        if r_base < 0: # Refinery style: negative is production
                            total_cost.append(price * self.cons_vars[(t, r_id)])
                        else: # Tech style: positive is production
                            total_cost.append(-price * self.cons_vars[(t, r_id)])
                    else:
                        total_cost.append(price * self.cons_vars[(t, r_id)])
                    
            # Carbon Taxes
            c_price = self.data.time_series.carbon_prices.get(t, 0.0)
            penalty_factor = self.data.time_series.carbon_penalties.get(t, 0.0)
            if c_price > 0:
                # Paid quotas are at market price
                total_cost.append(c_price * self.paid_quota_vars[t])
                # Unpaid/Penalty emissions are at penalized price (1 + x)
                total_cost.append(c_price * (1.0 + penalty_factor) * self.penalty_quota_vars[t])
                
        # Calculate dynamic massive penalty cost: 
        # Needs to be significantly higher than any real cost (Capex, Opex, Carbon Tax, Subsidies)
        # to ensure objectives are prioritized.
        # Find max tech opex/capex or carbon price.
        max_possible_cost = 0.0
        for t_id, tech in self.data.technologies.items():
            # Check tech costs (scaled by process units)
            max_possible_cost = max(max_possible_cost, tech.capex, tech.opex)
        
        c_prices = self.data.time_series.carbon_prices.values()
        if c_prices:
            max_possible_cost = max(max_possible_cost, max(c_prices))
            
        # We also need to account for multi-ton impacts (tons * price)
        # A conservative estimate is 100x the max price or max capex.
        # Given we saw coefficients of 30M in objective, let's target 10^9 if needed.
        # But to avoid numerical blowup, we'll use a 1000x factor over the max price/unit cost.
        MASSIVE_PENALTY_COST = max(1e6, max_possible_cost * 1000) 
        
        for i in self.penalty_vars:
            total_cost.append(self.penalty_vars[i] * MASSIVE_PENALTY_COST)
            
        self.model += pulp.lpSum(total_cost), "Total_Cost_Objective"

    def solve(self):
        if self.verbose:
            print("  [magenta][Optimizer][/magenta] [SOLVE] Solving model (this may take a moment)...")
        solver = pulp.PULP_CBC_CMD(
            msg=False,  # Always silence raw CBC output to prevent console deadlocks
            timeLimit=self.data.parameters.time_limit,
            gapRel=self.data.parameters.mip_gap,
            threads=4,
            presolve=True
        )
        self.model.solve(solver)
        status = pulp.LpStatus[self.model.status]

        # Normalize status: CBC reports 'Not Solved' / 'Undefined' when the time limit
        # is hit, but sol_status==1 means a feasible incumbent was found — treat as Feasible.
        if status in ['Not Solved', 'Undefined']:
            if self.model.sol_status == 1:
                status = 'Feasible'
                if self.verbose:
                    print("  [magenta][Optimizer][/magenta] [!] Time limit reached with a FEASIBLE incumbent — reporting best solution.")
            else:
                status = 'Infeasible'
        elif status == 'Optimal':
            pass  # already correct
        elif status == 'Infeasible':
            if self.model.sol_status == 1:
                # Rare numerical edge-case: solver claims infeasible but has a solution
                status = 'Feasible'

        if self.verbose:
            print(f"  [magenta][Optimizer][/magenta] [OK] Solver Status: [bold cyan]{status}[/bold cyan]")
        return status
