import pulp
import pandas as pd
import time
from typing import Dict, Any, List, Optional, Set
from rich import print
from tqdm import tqdm
from .model import PathFinderData
from .solver_factory import get_solver

MASSIVE_PENALTY_COST = 1e7

class PathFinderOptimizer:
    def __init__(self, data: PathFinderData, verbose: bool = False, solver_name: str = 'HIGHS'):
        self.data = data
        self.verbose = verbose
        self.solver_name = solver_name
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

        # Semantic caches resolved from resource/technology names and metadata.
        self.primary_emission_id: Optional[str] = None
        self.continuous_improvement_tech_ids: Set[str] = set()

        self.fuel_resource_id: Optional[str] = None
        self.electricity_resource_id: Optional[str] = None
        self.ccs_tech_ids: Set[str] = set()
        self.co2_storage_resource_id: Optional[str] = None
        self.co2_transport_resource_id: Optional[str] = None

    @staticmethod
    def _norm(raw_val: Any) -> str:
        return str(raw_val).strip().upper() if raw_val is not None else ""

    def _resource_name_upper(self, res_id: str) -> str:
        res = self.data.resources.get(res_id)
        return self._norm(res.name if res is not None else res_id)

    def _resource_type_upper(self, res_id: str) -> str:
        res = self.data.resources.get(res_id)
        return self._norm(res.type if res is not None else "")

    def _tech_name_upper(self, t_id: str) -> str:
        tech = self.data.technologies.get(t_id)
        return self._norm(tech.name if tech is not None else t_id)

    def _is_primary_emission_candidate(self, res_id: str) -> bool:
        res = self.data.resources.get(res_id)
        return res is not None and res.resource_type == 'CO2' and 'EMISS' in self._norm(res.type)

    def _is_continuous_improvement_tech(self, t_id: str) -> bool:
        tech_name = self._tech_name_upper(t_id)
        return tech_name in {'PROGRESS', 'CONTINUOUS IMPROVEMENT', 'UPGRADE'}

    def _is_continuous_improvement_tech_id(self, t_id: str) -> bool:
        return t_id in self.continuous_improvement_tech_ids

    def _is_ccs_tech(self, t_id: str) -> bool:
        tech_name = self._tech_name_upper(t_id)
        return 'CCS' in tech_name or 'CCU' in tech_name or 'CAPTURE' in tech_name

    def _is_electric_resource(self, res_id: str) -> bool:
        res = self.data.resources.get(res_id)
        return res is not None and res.resource_type == 'ELECTRICITY'

    def _is_fuel_resource(self, res_id: str) -> bool:
        res = self.data.resources.get(res_id)
        return res is not None and res.resource_type == 'FOSSIL FUEL'

    def _pick_largest_base_resource(self, resource_ids: List[str]) -> Optional[str]:
        if not resource_ids:
            return None
        return max(resource_ids, key=lambda r_id: abs(self.entity.base_consumptions.get(r_id, 0.0)))

    def _find_named_resource(self, required_any: List[str], required_all: List[str]) -> Optional[str]:
        candidates = []
        for r_id in self.data.resources:
            r_name = self._resource_name_upper(r_id)
            if not any(k in r_name for k in required_any):
                continue
            if not all(k in r_name for k in required_all):
                continue
            candidates.append(r_id)
        return self._pick_largest_base_resource(candidates)

    def _resolve_semantic_mappings(self) -> None:
        emission_candidates = [r_id for r_id in self.data.resources if self._is_primary_emission_candidate(r_id)]
        if len(emission_candidates) != 1:
            raise ValueError(
                "Primary emission resource resolution failed. "
                f"Expected exactly one CO2/CARBON emission resource, found {len(emission_candidates)}: {emission_candidates}"
            )
        self.primary_emission_id = emission_candidates[0]

        self.continuous_improvement_tech_ids = {
            t_id for t_id in self.data.technologies if self._is_continuous_improvement_tech(t_id)
        }

        self.ccs_tech_ids = {t_id for t_id in self.data.technologies if self._is_ccs_tech(t_id)}
        self.fuel_resource_id = self._pick_largest_base_resource([
            r_id for r_id in self.data.resources
            if self._is_fuel_resource(r_id) and self._resource_type_upper(r_id) != 'PRODUCTION'
        ])
        self.electricity_resource_id = self._pick_largest_base_resource([
            r_id for r_id in self.data.resources
            if self._is_electric_resource(r_id)
        ])
        self.co2_storage_resource_id = self._find_named_resource(['CO2', 'CARBON'], ['STORAGE'])
        self.co2_transport_resource_id = self._find_named_resource(['CO2', 'CARBON'], ['TRANSPORT'])

    def _get_unit_conversion(self, ref_resource_id: Optional[str], target_resource_id: Optional[str]) -> float:
        if ref_resource_id is None or target_resource_id is None:
            return 1.0
        ref_resource = self.data.resources.get(ref_resource_id)
        target_resource = self.data.resources.get(target_resource_id)
        if ref_resource is None or target_resource is None:
            return 1.0

        ref_unit = self._norm(ref_resource.unit)
        tgt_unit = self._norm(target_resource.unit)
        if not ref_unit or not tgt_unit or ref_unit == tgt_unit:
            return 1.0
        return self.data.unit_conversions.get((ref_unit, tgt_unit), 1.0)

    def _find_impact_for_resource(self, tech: Any, resource_id: str) -> Optional[Dict[str, Any]]:
        imp = tech.impacts.get(resource_id)
        if imp:
            return imp

        if resource_id == self.primary_emission_id:
            for impact_resource_id, impact_data in tech.impacts.items():
                if impact_resource_id in self.data.resources and self._is_primary_emission_candidate(impact_resource_id):
                    return impact_data
                impact_token = self._norm(impact_resource_id)
                if 'CO2' in impact_token or 'CARBON' in impact_token:
                    return impact_data
        return None

    def _is_primary_emission_objective(self, resource_token: str) -> bool:
        token_upper = self._norm(resource_token)
        if resource_token == self.primary_emission_id:
            return True
        if resource_token in self.data.resources:
            return self._is_primary_emission_candidate(resource_token)
        return ('CO2' in token_upper or 'CARBON' in token_upper) and 'INDIRECT' not in token_upper and 'TOTAL' not in token_upper

    def _is_indirect_emission_objective(self, resource_token: str) -> bool:
        token_upper = self._norm(resource_token)
        return 'INDIRECT' in token_upper and ('CO2' in token_upper or 'CARBON' in token_upper)

    def _is_total_emission_objective(self, resource_token: str) -> bool:
        token_upper = self._norm(resource_token)
        return 'TOTAL' in token_upper and ('CO2' in token_upper or 'CARBON' in token_upper)
    
    def build_model(self):
        if not self.entity:
            raise ValueError("No entity to optimize")

        self._resolve_semantic_mappings()
            
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
            
        self.market_trade_vars = {t: {} for t in self.years}
        for t in self.years:
            for res_id in self.data.resources:
                if res_id in self.data.time_series.resource_prices:
                    res_obj = self.data.resources.get(res_id)
                    can_buy = res_obj.can_buy if res_obj else False
                    can_sell = res_obj.can_sell if res_obj else False
                    
                    if can_buy and can_sell:
                        self.market_trade_vars[t][res_id] = pulp.LpVariable(f"Market_Trade_{res_id}_{t}", lowBound=None, upBound=None)
                    elif can_buy and not can_sell:
                        self.market_trade_vars[t][res_id] = pulp.LpVariable(f"Market_Trade_{res_id}_{t}", lowBound=0.0, upBound=None)
                    elif not can_buy and can_sell:
                        self.market_trade_vars[t][res_id] = pulp.LpVariable(f"Market_Trade_{res_id}_{t}", lowBound=None, upBound=0.0)
                    
            for p_id, process in self.entity.processes.items():
                for t_id in process.valid_technologies:
                    if self._is_continuous_improvement_tech_id(t_id):
                        continue
                    # PERFORMANCE: Relax technologies if nb_units > 1 to speed up solver
                    # Integer decision to invest in a number of units of technology t_id at year t for process p_id
                    # Relax if it's an H2 tech OR if nb_units > 1 (MIP complexity reduction)
                    should_relax = (process.nb_units > 1)
                    
                    # Also check for global override from parameters if added
                    if hasattr(self.data.parameters, 'relax_integrality') and self.data.parameters.relax_integrality:
                        should_relax = True
                        
                    var_cat = pulp.LpContinuous if should_relax else pulp.LpInteger
                    self.invest_vars[(t, p_id, t_id)] = pulp.LpVariable(f"Invest_{t}_{p_id}_{t_id}", lowBound=0, upBound=process.nb_units, cat=var_cat)
                    # Same category for active units tracking
                    self.active_vars[(t, p_id, t_id)] = pulp.LpVariable(f"Active_{t}_{p_id}_{t_id}", lowBound=0, upBound=process.nb_units, cat=var_cat)
                    # Binary flag if an investment project (at least 1 unit) occurs this year
                    self.project_vars[(t, p_id, t_id)] = pulp.LpVariable(f"Project_{t}_{p_id}_{t_id}", cat=pulp.LpBinary)
                    
                    if self.data.grant_params.active and t_id not in self.data.grant_params.excluded_technologies:
                        self.grant_used_vars[(t, p_id, t_id)] = pulp.LpVariable(f"Grant_{t}_{p_id}_{t_id}", cat=pulp.LpBinary)
                    if self.data.ccfd_params.active and t_id not in self.data.grant_params.excluded_technologies:
                        self.ccfd_used_vars[(t, p_id, t_id)] = pulp.LpVariable(f"CCfD_{t}_{p_id}_{t_id}", cat=pulp.LpBinary)
                        
                    for l_id, loan in enumerate(self.data.bank_loans):
                        self.loan_vars[(t, p_id, t_id, l_id)] = pulp.LpVariable(f"Loan_{t}_{p_id}_{t_id}_{l_id}", lowBound=0)
                
            for res_id in self.data.resources:
                if res_id != self.primary_emission_id:
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
                        yr_map = self.entity.ref_baselines.get(self.primary_emission_id, {})
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
                yr_map = self.entity.ref_baselines.get(self.primary_emission_id, {})
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
            major_techs = [
                t_id for t_id in process.valid_technologies
                if not self._is_continuous_improvement_tech_id(t_id)
            ]
            
            free_groups = []
            tech_to_free_group = {}
            for t1 in major_techs:
                comp_dict = self.data.tech_compatibilities.get(t1, {})
                for t2 in major_techs:
                    if t1 == t2: continue
                    if comp_dict.get(t2) == 'FREE':
                        if t1 not in tech_to_free_group and t2 not in tech_to_free_group:
                            new_group_id = len(free_groups)
                            free_groups.append([t1, t2])
                            tech_to_free_group[t1] = new_group_id
                            tech_to_free_group[t2] = new_group_id
                        elif t1 in tech_to_free_group and t2 not in tech_to_free_group:
                            free_groups[tech_to_free_group[t1]].append(t2)
                            tech_to_free_group[t2] = tech_to_free_group[t1]
                        elif t2 in tech_to_free_group and t1 not in tech_to_free_group:
                            free_groups[tech_to_free_group[t2]].append(t1)
                            tech_to_free_group[t1] = tech_to_free_group[t2]
                        elif tech_to_free_group[t1] != tech_to_free_group[t2]:
                            g1 = tech_to_free_group[t1]
                            g2 = tech_to_free_group[t2]
                            free_groups[g1].extend(free_groups[g2])
                            for t3 in free_groups[g2]:
                                tech_to_free_group[t3] = g1
                            free_groups[g2] = []
                            
            free_groups = [set(g) for g in free_groups if g]
            
            # --- Directed Technology Precedence Logic ---
            for t1 in major_techs:
                # Max units for a single technology over the horizon cannot exceed process units
                self.model += pulp.lpSum(self.invest_vars[(t, p_id, t1)] for t in self.years) <= process.nb_units, f"Max_{process.nb_units}_Invest_{p_id}_{t1}"
                
                comp_dict = self.data.tech_compatibilities.get(t1, {})
                for t2 in major_techs:
                    if t1 == t2: continue
                    val = comp_dict.get(t2, '')
                    
                    if val == 'X':
                        # Mutual Exclusion: The sum of active units of T1 and T2 cannot exceed total process units
                        for t in self.years:
                            self.model += self.active_vars[(t, p_id, t1)] + self.active_vars[(t, p_id, t2)] <= process.nb_units, f"Mutual_Capacity_Exclusion_{p_id}_{t1}_{t2}_at_{t}"

            for t_id in process.valid_technologies:
                if self._is_continuous_improvement_tech_id(t_id):
                    continue
                
                tech = self.data.technologies[t_id]
                delay = tech.implementation_time
                
                is_free = False
                for g in free_groups:
                    if t_id in g:
                        is_free = True
                        break
                        
                for t in self.years:
                    valid_invest_years = [tau for tau in self.years if tau <= t - delay]
                    if not is_free:
                        if valid_invest_years:
                            self.model += self.active_vars[(t, p_id, t_id)] == pulp.lpSum(self.invest_vars[(tau, p_id, t_id)] for tau in valid_invest_years), f"Active_Logic_{t}_{p_id}_{t_id}"
                        else:
                            self.model += self.active_vars[(t, p_id, t_id)] == 0, f"Active_Logic_{t}_{p_id}_{t_id}"
                        
                    # Project variable link (1 if Invest >= 1, else 0)
                    self.model += self.invest_vars[(t, p_id, t_id)] <= process.nb_units * self.project_vars[(t, p_id, t_id)], f"Project_Link_UB_{t}_{p_id}_{t_id}"
                    self.model += self.project_vars[(t, p_id, t_id)] <= self.invest_vars[(t, p_id, t_id)], f"Project_Link_LB_{t}_{p_id}_{t_id}"

            for g_idx, group in enumerate(free_groups):
                for t in self.years:
                    sum_usage = pulp.lpSum(self.active_vars[(t, p_id, t_id)] for t_id in group)
                    sum_invest = []
                    for t_id in group:
                        delay = self.data.technologies[t_id].implementation_time
                        sum_invest.extend(self.invest_vars[(tau, p_id, t_id)] for tau in self.years if tau <= t - delay)
                    
                    self.model += sum_usage <= pulp.lpSum(sum_invest), f"FREE_Capacity_Group_{p_id}_{g_idx}_{t}"

        # Public Aids Exclusivity and Limits
        for t in self.years:
            for p_id, process in self.entity.processes.items():
                for t_id in process.valid_technologies:
                    if self._is_continuous_improvement_tech_id(t_id):
                        continue
                    expr = []
                    if self.data.grant_params.active and t_id not in self.data.grant_params.excluded_technologies:
                        expr.append(self.grant_used_vars[(t, p_id, t_id)])
                    if self.data.ccfd_params.active and t_id not in self.data.grant_params.excluded_technologies:
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
                                             for t_id in proc.valid_technologies
                                             if not self._is_continuous_improvement_tech_id(t_id)
                                             and t_id not in self.data.grant_params.excluded_technologies) <= 1, f"Grant_Spacing_{t}"

        # CCfD Max Contracts
        if self.data.ccfd_params.active and self.data.ccfd_params.nb_contracts > 0:
            self.model += pulp.lpSum(self.ccfd_used_vars[(t, p_id, t_id)] 
                                     for t in self.years
                                     for p_id, proc in self.entity.processes.items()
                                     for t_id in proc.valid_technologies
                                     if not self._is_continuous_improvement_tech_id(t_id)
                                     and t_id not in self.data.grant_params.excluded_technologies) <= self.data.ccfd_params.nb_contracts, "CCfD_Max_Contracts"

        # --- Subsidies Values (Linearization of Cap) ---
        self.grant_amt_vars = {}
        for t in self.years:
            for p_id, process in self.entity.processes.items():
                for t_id in process.valid_technologies:
                    if self._is_continuous_improvement_tech_id(t_id):
                        continue
                    if self.data.grant_params.active and self.data.grant_params.rate > 0 and t_id not in self.data.grant_params.excluded_technologies:
                        self.grant_amt_vars[(t, p_id, t_id)] = pulp.LpVariable(f"GrantAmt_{t}_{p_id}_{t_id}", lowBound=0.0)
                        
                        tech = self.data.technologies[t_id]
                        cap_capex = 1.0
                        if tech.capex_per_unit:
                            if tech.capex_unit == 'tCO2':
                                cap_capex = self.entity.base_emissions * process.emission_shares.get(self.primary_emission_id, 0.0)
                            elif 'MW' in tech.capex_unit.upper(): 
                                base_mwh = 0.0
                                if self.fuel_resource_id is not None:
                                    base_mwh = (
                                        self.entity.base_consumptions.get(self.fuel_resource_id, 0.0)
                                        * process.consumption_shares.get(self.fuel_resource_id, 0.0)
                                    )
                                cap_capex = base_mwh / self.entity.annual_operating_hours
                        current_capex = tech.capex_by_year.get(t, tech.capex)
                        true_capex_per_unit = (current_capex * cap_capex) / process.nb_units
                        
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
                if res_id == self.primary_emission_id:
                    initial_val = self.entity.base_emissions * process.emission_shares.get(self.primary_emission_id, 0.0)
                else:
                    initial_val = self.entity.base_consumptions.get(res_id, 0.0) * process.consumption_shares.get(res_id, 0.0)
                
                # Big-M value for linearizations: must be large enough to encapsulate shifting cross-dimensional references
                # Now calculated dynamically per process/resource to improve numerical stability.
                # We use a safety factor of 1.5x the initial value or a sensible minimum.
                M = max(1000.0, abs(initial_val) * 2.0)
                
                # Evolution of the state year over year
                # Determine continuous improvement rate for this process
                up_rate = 0.0
                for ci_t_id in process.valid_technologies:
                    if not self._is_continuous_improvement_tech_id(ci_t_id):
                        continue
                    ci_tech = self.data.technologies[ci_t_id]
                    ci_imp = ci_tech.impacts.get('ALL') or ci_tech.impacts.get(res_id)
                    if ci_imp and ci_imp['type'] in ['variation', 'up']:
                        up_rate += abs(ci_imp['value'])
                up_rate = min(up_rate, 0.99)
                
                for i, t in enumerate(self.years):
                    # Accumulate all technologies active at year t from the adjusted baseline
                    impacts_t = []
                    
                    for t_id in process.valid_technologies:
                        if self._is_continuous_improvement_tech_id(t_id):
                            continue
                        tech = self.data.technologies[t_id]
                        
                        imp = self._find_impact_for_resource(tech, res_id)
                        
                        if imp:
                            act_var_current = self.active_vars[(t, p_id, t_id)]
                            
                            reference = imp.get('reference', 'INITIAL')
                            ref_res = imp.get('ref_resource')
                            if ref_res == self.primary_emission_id:
                                base_ref_amount = self.entity.base_emissions * process.emission_shares.get(self.primary_emission_id, 0.0)
                            elif ref_res and ref_res in self.entity.base_consumptions:
                                base_ref_amount = self.entity.base_consumptions.get(ref_res, 0.0) * process.consumption_shares.get(ref_res, 0.0)
                            else:
                                base_ref_amount = 1.0
                                
                            if imp.get('reference', '') == 'AVOIDED' and ref_res:
                                ref_imp = self._find_impact_for_resource(tech, ref_res)
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
                                    conv = self._get_unit_conversion(ref_res, res_id)
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
                                    conv = self._get_unit_conversion(ref_res, res_id)
                                    target_reduction = imp['value'] * base_ref * scaling_factor * conv
                                else:
                                    target_reduction = 0
                                    
                                impacts_t.append(target_reduction * act_var_current)
                                    
                    # Apply compounding continuous improvement to the entire state
                    # Formula: State(t) = (Initial + Sum of Tech Impacts) * (1 - up_rate)^t
                    up_factor = (1.0 - up_rate) ** i
                    
                    net_base_val = initial_val + pulp.lpSum(impacts_t)
                    res_obj = self.data.resources.get(res_id)
                    is_prod = res_obj.type.strip().upper() == 'PRODUCTION' if res_obj else False
                    
                    # We NO LONGER need the State_NonNeg_Slack here as we added lowBound=0 to cons_vars
                    # which propagates back through the mass balance.
                    self.model += self.process_state_vars[(t, p_id, res_id)] == net_base_val * up_factor, f"State_Evol_{t}_{p_id}_{res_id}"

        # Resource Mass Balance via mapped dynamic states
        for t in self.years:
            for res_id in self.data.resources:
                if res_id == self.primary_emission_id:
                    continue
                    
                # The total consumption is exactly the sum of the dynamic process states
                total_process_cons = pulp.lpSum([self.process_state_vars[(t, p_id, res_id)] for p_id in self.entity.processes])

                allocated_share = sum(p.consumption_shares.get(res_id, 0.0) for p in self.entity.processes.values())
                frac_unallocated = 1.0 - allocated_share
                if abs(frac_unallocated) < 1e-4: frac_unallocated = 0.0
                unallocated = self.entity.base_consumptions.get(res_id, 0.0) * frac_unallocated
                
                dac_cons = 0.0
                if self.data.dac_params.active and self.electricity_resource_id is not None and res_id == self.electricity_resource_id:
                    dac_cons = self.dac_captured_vars[t] * self.data.dac_params.elec_by_year.get(t, 0.0)
                
                self.model += self.cons_vars[(t, res_id)] == total_process_cons + unallocated + dac_cons + self.market_trade_vars[t].get(res_id, 0.0), f"Cons_Balance_{t}_{res_id}"
                
            # Emissions computation
            direct_process_emis = pulp.lpSum([
                self.process_state_vars[(t, p_id, self.primary_emission_id)]
                for p_id in self.entity.processes
            ])
            
            # Account for unallocated emissions
            allocated_emis_share = sum(
                p.emission_shares.get(self.primary_emission_id, 0.0)
                for p in self.entity.processes.values()
            )
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
                target_limit = net_emis_for_tax * fq_pct
                self.model += self.taxed_emis_vars[t] >= net_emis_for_tax * (1.0 - fq_pct), f"Taxed_Emis_Limit_{t}"
            else:
                target_limit = fq_pct
                self.model += self.taxed_emis_vars[t] >= net_emis_for_tax - fq_pct, f"Taxed_Emis_Limit_{t}"

            # Fix: Enforce an upper bound on paid quotas equal to the defined emission limit (target).
            # This ensures any excess emission is funneled into the penalty tracking variable.
            self.model += self.paid_quota_vars[t] <= target_limit, f"Paid_Quota_Upper_Bound_{t}"

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
            elif self._is_primary_emission_objective(obj.resource):
                base_val = self.entity.base_emissions
            elif self._is_indirect_emission_objective(obj.resource):
                base_indir = 0.0
                for r_id in self.data.resources:
                    if r_id in self.data.time_series.other_emissions_factors:
                        factor = self.data.time_series.other_emissions_factors[r_id].get(self.years[0], 0.0)
                        if factor > 0:
                            base_indir += self.entity.base_consumptions.get(r_id, 0.0) * factor
                base_val = base_indir
            elif self._is_total_emission_objective(obj.resource):
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
                    
                    if self._is_primary_emission_objective(obj.resource):
                        net_obj_emis = self.emis_vars[t]
                        if self.data.dac_params.active:
                            net_obj_emis -= self.dac_captured_vars[t]
                        if self.data.credit_params.active:
                            net_obj_emis -= self.credit_purchased_vars[t]
                        var_expr = net_obj_emis
                    elif self._is_indirect_emission_objective(obj.resource):
                        var_expr = self.indirect_emis_vars[t]
                    elif self._is_total_emission_objective(obj.resource):
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
                if self._is_primary_emission_objective(obj.resource):
                    net_obj_emis = self.emis_vars[t_target]
                    if self.data.dac_params.active:
                        net_obj_emis -= self.dac_captured_vars[t_target]
                    if self.data.credit_params.active:
                        net_obj_emis -= self.credit_purchased_vars[t_target]
                    var_expr = net_obj_emis
                elif self._is_indirect_emission_objective(obj.resource):
                    var_expr = self.indirect_emis_vars[t_target]
                elif self._is_total_emission_objective(obj.resource):
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
                    if self._is_continuous_improvement_tech_id(t_id):
                        continue
                    tech = self.data.technologies[t_id]
                    if tech.capex > 0:
                        cap = 1.0
                        if tech.capex_per_unit:
                            if tech.capex_unit == 'tCO2':
                                cap = self.entity.base_emissions * process.emission_shares.get(self.primary_emission_id, 0.0)
                            elif 'MW' in tech.capex_unit.upper(): 
                                base_mwh = 0.0
                                if self.fuel_resource_id is not None:
                                    base_mwh = (
                                        self.entity.base_consumptions.get(self.fuel_resource_id, 0.0)
                                        * process.consumption_shares.get(self.fuel_resource_id, 0.0)
                                    )
                                cap = base_mwh / self.entity.annual_operating_hours
                        current_capex = tech.capex_by_year.get(t, tech.capex)
                        true_capex = current_capex * cap
                        
                        if self.data.grant_params.active and self.data.grant_params.rate > 0 and t_id.upper() not in [x.upper() for x in self.data.grant_params.excluded_technologies]:
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
                                       for t_id in p.valid_technologies
                                       if not self._is_continuous_improvement_tech_id(t_id)
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
            df = (1.0 + self.data.parameters.discount_rate) ** (t - self.years[0])
            
            # 1. Financial Flows (Capex Out-of-pocket + Loan Annuities)
            if self.entity.ca_percentage_limit > 0 and self.entity.sold_resources:
                total_cost.append(out_of_pocket_capex_vars[t] / df)
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
                                                       for t_id in p.valid_technologies
                                                       if not self._is_continuous_improvement_tech_id(t_id))
                        total_cost.append((total_loan_amount * annuity_factor) / df)
                        
                        # Add a tiny penalty to loan selection to strictly prioritize out-of-pocket usage
                        # Even if rate is 0%, loans shouldn't be taken if budget is available.
                        total_cost.append((total_loan_amount * 1e-4) / df)

            # CAPEX and OPEX
            for p_id, process in self.entity.processes.items():
                for t_id in process.valid_technologies:
                    if self._is_continuous_improvement_tech_id(t_id):
                        continue
                    tech = self.data.technologies[t_id]
                    
                    # OPEX and CCfD (CAPEX handled separately if budget active)
                    cap_calc = 1.0
                    if tech.capex_per_unit or tech.opex_per_unit:
                        if tech.capex_unit == 'tCO2' or tech.opex_unit == 'tCO2': 
                            cap_calc = self.entity.base_emissions * process.emission_shares.get(self.primary_emission_id, 0.0)
                        elif 'MW' in str(tech.capex_unit).upper() or 'MW' in str(tech.opex_unit).upper(): 
                            base_mwh = 0.0
                            if self.fuel_resource_id is not None:
                                base_mwh = (
                                    self.entity.base_consumptions.get(self.fuel_resource_id, 0.0)
                                    * process.consumption_shares.get(self.fuel_resource_id, 0.0)
                                )
                            cap_calc = base_mwh / self.entity.annual_operating_hours
                    
                    current_capex = tech.capex_by_year.get(t, tech.capex)
                    current_opex = tech.opex_by_year.get(t, tech.opex)
                    true_capex = current_capex * cap_calc
                    true_opex = current_opex * cap_calc
                    
                    # If NO budget constraint, we add CAPEX here directly.
                    # If budget constraint ACTIVE, CAPEX is already in out_of_pocket_capex_vars[t].
                    if not (self.entity.ca_percentage_limit > 0 and self.entity.sold_resources):
                        if self.data.grant_params.active and self.data.grant_params.rate > 0 and t_id not in self.data.grant_params.excluded_technologies:
                            total_cost.append(((true_capex / process.nb_units) * self.invest_vars[(t, p_id, t_id)] - self.grant_amt_vars[(t, p_id, t_id)]) / df)
                        else:
                            total_cost.append(((true_capex / process.nb_units) * self.invest_vars[(t, p_id, t_id)]) / df)

                    total_cost.append(((true_opex / process.nb_units) * self.active_vars[(t, p_id, t_id)]) / df)
                    
                    # CCfD Calculation
                    if self.data.ccfd_params.active and self.data.ccfd_params.duration > 0 and t_id not in self.data.grant_params.excluded_technologies:
                        ccfd_p = self.data.ccfd_params
                        imp = self._find_impact_for_resource(tech, self.primary_emission_id)
                            
                        if imp and (imp['type'] == 'variation' or imp['type'] == 'up'):
                            # Avoided emissions: The absolute reduction per year for ONE unit out of nb_units
                            reduction_frac_per_unit = (-imp['value'] if imp['value'] < 0 else 0) / process.nb_units
                            if reduction_frac_per_unit > 0:
                                max_emis_for_proc = self.entity.base_emissions * process.emission_shares.get(self.primary_emission_id, 0.0)
                                avoided_scope1_per_unit = reduction_frac_per_unit * max_emis_for_proc
                                
                                # NET DECARBONIZATION: Subtract added Scope 3 from H2 or other fuels
                                added_scope3_per_unit = 0.0
                                for res_id, r_imp in tech.impacts.items():
                                    if r_imp['type'] == 'new' and r_imp['value'] > 0:
                                        # If this tech consumes a resource with an emission factor (like H2), 
                                        # its added footprint reduces the CCfD eligibility.
                                        # We use a representative factor or link to market variables.
                                        # For simplicity and linearity, we'll use the weighted emission factor of the H2 pool?
                                        # No, let's just subtract the max possible footprint of Grey H2 to be conservative?
                                        # Actually, let's just subtract the specific impact * current pool factor.
                                        pass # Handled below in the year loop
                                
                                avoided_tons_per_year_per_unit = avoided_scope1_per_unit # Base
                                
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
                                            
                                        # Adjusted Avoided Tons: Avoided Scope 1 - Added Scope 3 (from H2 choice)
                                        # Since we want to reward clean H2, the CCfD is proportional to:
                                        # (Avoided S1) - (H2_Cons * Factor_Chosen_H2)
                                        # In the MILP, we can just subtract the total Scope 3 H2 emissions from the total CCfD pot if this tech is active.
                                        
                                        # Simplified: Avoided tons = Avoided S1. 
                                        # The 'Choice' is already favored by 'total_cost' including H2 prices.
                                        # BUT to follow user request of 'Choice in CCfD', we subtract the H2 footprint.
                                        
                                        h2_footprint_penalty = 0.0
                                        # We calculate the average emission factor of the H2 supply at year tau
                                        # This is non-linear. Let's use a linear approximation: 
                                        # If H2_Buy_Grey is used, it adds its own footprint to the costs anyway.
                                        
                                        # FIXED: Use the actual indirect emissions from the H2 balancer in the year tau
                                        # We'll subtract the total indirect H2 emissions from the CCfD revenue line.
                                        # This effectively means "You get CCfD on what you avoid, minus what you emit indirectly to do so".
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
                                    
                        total_cost.append(-ccfd_amt_var / df)
                        
                    # CCS-related transport and storage variable costs
                    if t_id in self.ccs_tech_ids:
                        imp = self._find_impact_for_resource(tech, self.primary_emission_id)
                        if imp and (imp['type'] == 'variation' or imp['type'] == 'up'):
                            # Captured tonnage per unit: original emissions that are NO LONGER emitted
                            reduction_frac_per_unit = (-imp['value'] if imp['value'] < 0 else 0) / process.nb_units
                            if reduction_frac_per_unit > 0:
                                max_emis_for_proc = self.entity.base_emissions * process.emission_shares.get(self.primary_emission_id, 0.0)
                                captured_tons_per_unit = reduction_frac_per_unit * max_emis_for_proc
                                
                                s_price = 0.0
                                if self.co2_storage_resource_id is not None:
                                    s_price = self.data.time_series.resource_prices.get(self.co2_storage_resource_id, {}).get(t, 0.0)

                                tr_price = 0.0
                                if self.co2_transport_resource_id is not None:
                                    tr_price = self.data.time_series.resource_prices.get(self.co2_transport_resource_id, {}).get(t, 0.0)

                                if s_price > 0 or tr_price > 0:
                                    total_cost.append(((s_price + tr_price) * captured_tons_per_unit * self.active_vars[(t, p_id, t_id)]) / df)

                total_cost.append((self.dac_total_capacity_vars[t] * self.data.dac_params.opex_by_year.get(t, 0.0)) / df)
            if self.data.credit_params.active:
                total_cost.append((self.credit_purchased_vars[t] * self.data.credit_params.cost_by_year.get(t, 0.0)) / df)
                
            # Resource Costs / Revenues
            for r_id in self.data.resources:
                if r_id == self.primary_emission_id:
                    continue
                price = 0.0
                if r_id in self.data.time_series.resource_prices:
                    price = self.data.time_series.resource_prices[r_id].get(t, 0.0)
                
                if price > 0:
                    res_obj = self.data.resources.get(r_id)
                    if res_obj and res_obj.type.strip().upper() == 'PRODUCTION':
                        # PRODUCTION resources: revenue depends on the sign of cons_var.
                        # We use simple price * Quantity. 
                        # Quantity is +cons_var if impacts are positive, -cons_var if negative.
                        r_base = self.entity.base_consumptions.get(r_id, 0.0)
                        if r_base < 0: # Refinery style: negative is production
                            total_cost.append((price * self.cons_vars[(t, r_id)]) / df)
                        else: # Tech style: positive is production
                            total_cost.append((-price * self.cons_vars[(t, r_id)]) / df)
                    elif r_id in self.market_trade_vars[t]:
                        # Purchase/Sell from the market
                        total_cost.append((price * self.market_trade_vars[t][r_id]) / df)
                        
                        # NET CCfD PENALTY: Scope 3 emissions reduce the 'Avoided Tons' benefit
                        if r_id in self.data.time_series.other_emissions_factors:
                            factor = self.data.time_series.other_emissions_factors[r_id].get(t, 0.0)
                            if factor > 0 and self.data.ccfd_params.active:
                                ccfd_p = self.data.ccfd_params
                                c_price = self.data.time_series.carbon_prices.get(t, 0.0)
                                strike_price = (1.0 + ccfd_p.eua_price_pct) * self.data.time_series.carbon_prices.get(t, 0.0)
                                subsidy_per_ton = strike_price - c_price
                                if ccfd_p.contract_type == 1: subsidy_per_ton = max(0, subsidy_per_ton)
                                
                                # Penalty reduces the negative cost (subsidy)
                                total_cost.append((self.market_trade_vars[t][r_id] * factor * subsidy_per_ton) / df)
                    
            # Carbon Taxes
            c_price = self.data.time_series.carbon_prices.get(t, 0.0)
            penalty_factor = self.data.time_series.carbon_penalties.get(t, 0.0)
            if c_price > 0:
                # Paid quotas are at market price
                total_cost.append((c_price * self.paid_quota_vars[t]) / df)
                # Unpaid/Penalty emissions are at penalized price (1 + x)
                total_cost.append((c_price * (1.0 + penalty_factor) * self.penalty_quota_vars[t]) / df)
                
        # Calculate dynamic massive penalty cost: 
        # Needs to be significantly higher than any real cost (Capex, Opex, Carbon Tax, Subsidies)
        # to ensure objectives are prioritized.
        # Find max tech opex/capex or carbon price.
        max_possible_cost = 0.0
        for t_id, tech in self.data.technologies.items():
            # Check tech costs (scaled by process units)
            max_tech_capex = max(tech.capex_by_year.values()) if tech.capex_by_year else tech.capex
            max_tech_opex = max(tech.opex_by_year.values()) if tech.opex_by_year else tech.opex
            max_possible_cost = max(max_possible_cost, max_tech_capex, max_tech_opex)
        
        c_prices = self.data.time_series.carbon_prices.values()
        if c_prices:
            max_possible_cost = max(max_possible_cost, max(c_prices))
            
        # We also need to account for multi-ton impacts (tons * price)
        # A conservative estimate is 100x the max price or max capex.
        # Given we saw coefficients of 30M in objective, let's target 10^9 if needed.
        # But to avoid numerical blowup, we'll use a 1000x factor over the max price/unit cost.
        self.massive_penalty_cost = max(1e6, max_possible_cost * 1000) 
        
        for idx, obj in enumerate(self.data.objectives):
            if obj.penalty_type == "NONE":
                continue
            
            is_realistic = (obj.penalty_type == "PENALTIES")
                
            # Handle both scalar and time-indexed penalty variables
            if idx in self.penalty_vars:
                # Use target year price if realistic, else massive penalty
                p_cost = self.massive_penalty_cost
                t_target = obj.target_year
                # Clamp target year to simulation range
                if t_target not in self.years:
                    t_target = min(self.years[-1], max(self.years[0], t_target))
                
                df_target = (1.0 + self.data.parameters.discount_rate) ** (t_target - self.years[0])
                
                if is_realistic:
                    c_price = self.data.time_series.carbon_prices.get(t_target, 0.0)
                    p_fact = self.data.time_series.carbon_penalties.get(t_target, 0.0)
                    p_cost = c_price * (1.0 + p_fact)
                
                total_cost.append((self.penalty_vars[idx] * p_cost) / df_target)
            
            # Check for (idx, t) keys (usually for LINEAR objectives)
            for t in self.years:
                if (idx, t) in self.penalty_vars:
                    p_cost = self.massive_penalty_cost
                    df = (1.0 + self.data.parameters.discount_rate) ** (t - self.years[0])
                    if is_realistic:
                        c_price = self.data.time_series.carbon_prices.get(t, 0.0)
                        p_fact = self.data.time_series.carbon_penalties.get(t, 0.0)
                        p_cost = c_price * (1.0 + p_fact)
                    
                    total_cost.append((self.penalty_vars[(idx, t)] * p_cost) / df)
            
        self.model += pulp.lpSum(total_cost), "Total_Cost_Objective"

    def apply_warm_start(self, solution_data: Dict[str, float]):
        """Sets the initial values of variables for a MIP start (warm start)."""
        if not solution_data:
            return
        
        var_map = {v.name: v for v in self.model.variables()}
        count = 0
        for name, value in solution_data.items():
            if name in var_map and value is not None:
                var_map[name].varValue = value
                count += 1
        
        if self.verbose and count > 0:
            print(f"  [magenta][Optimizer][/magenta] [WARM START] Applied {count} initial values.")

    def solve(self, warm_start: bool = False):
        if self.verbose:
            print(f"  [magenta][Optimizer][/magenta] [SOLVE] Solving model with [bold cyan]{self.solver_name}[/bold cyan]...")
        
        solver = get_solver(
            solver_name=self.solver_name,
            time_limit=self.data.parameters.time_limit,
            gap_rel=self.data.parameters.mip_gap,
            msg=False,  # Silence raw solver output to prevent console deadlocks
            threads=4,
            warm_start=warm_start,
        )
        self.model.solve(solver)
        status = pulp.LpStatus[self.model.status]

        objective_value = None
        try:
            if self.model.objective is not None:
                objective_value = self.model.objective.value()
        except Exception:
            objective_value = None

        vars_with_values = sum(1 for v in self.model.variables() if v.varValue is not None)
        has_incumbent = (vars_with_values > 0) and (objective_value is not None)

        # Normalize status: CBC reports 'Not Solved' / 'Undefined' when the time limit
        # is hit, but sol_status==1 means a feasible incumbent was found — treat as Feasible.
        if status in ['Not Solved', 'Undefined']:
            if self.model.sol_status == 1 or has_incumbent:
                status = 'Feasible'
                if self.verbose:
                    print("  [magenta][Optimizer][/magenta] [!] Time limit reached with a FEASIBLE incumbent — reporting best solution.")
            else:
                status = 'Infeasible'
        elif status == 'Optimal':
            pass  # already correct
        elif status == 'Infeasible':
            if self.model.sol_status == 1 or has_incumbent:
                # Rare numerical edge-case: solver claims infeasible but has a solution
                status = 'Feasible'

        if self.verbose:
            print(f"  [magenta][Optimizer][/magenta] [OK] Solver Status: [bold cyan]{status}[/bold cyan]")
        return status

